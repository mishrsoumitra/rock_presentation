// The render tree a slide builds into, and what drawing it does: fills, strokes, text as glyph outlines,
// animation overrides and the base transform the editor draws under.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { FontRegistry } from '@pptx/text';
import { deckBytes } from '../../resolve/test/deck.ts';
import { buildSlide, drawShape, drawSlide, shapeTextLayout, type RenderShape, type RenderSlide } from '../src/index.ts';
import { installStubs, recorder, type Call } from './canvas-stubs.ts';

installStubs();

const FONTS = join(import.meta.dirname, '../../../assets/fonts');
const fonts = new FontRegistry();
for (const dir of ['carlito', 'caladea', 'liberation']) {
  for (const f of readdirSync(join(FONTS, dir))) if (f.endsWith('.ttf')) fonts.addBytes(new Uint8Array(readFileSync(join(FONTS, dir, f))));
}

const SLIDE = '/ppt/slides/slide1.xml';

function build(): RenderSlide {
  const doc = new PresentationDocument(OpcPackage.open(deckBytes()));
  return buildSlide(resolveSlide(doc, SLIDE), { fonts });
}

const find = (shapes: RenderShape[], name: string): RenderShape | undefined => {
  for (const s of shapes) {
    if (s.name === name) return s;
    const inner = find(s.children, name);
    if (inner) return inner;
  }
  return undefined;
};

describe('buildSlide', () => {
  test('shapes come in drawing order: master, layout, then the slide', () => {
    const slide = build();
    expect(slide.shapes.map(s => `${s.layer}:${s.name}`)).toEqual([
      'master:Logo', 'slide:Title 1', 'slide:Content Placeholder 2', 'slide:Slide Number Placeholder 3', 'slide:Callout', 'slide:Group',
    ]);
    expect(slide.size).toEqual({ cx: 12192000, cy: 6858000 });
  });

  test('a styled shape gets its fill and line from the theme; its text is laid out with the theme font', () => {
    const callout = find(build().shapes, 'Callout')!;
    const path = callout.prims.find(p => p.kind === 'path');
    expect(path?.kind === 'path' && path.fill).toMatchObject({ kind: 'solid', color: { r: 0x44 / 255, g: 0x72 / 255, b: 0xc4 / 255, a: 1 } });
    expect(path?.kind === 'path' && path.stroke?.width).toBe(12700);
    const text = callout.prims.find(p => p.kind === 'text');
    expect(text?.kind === 'text' && text.layout.lines.map(l => l.runs.map(r => r.match.face.family).join())).toEqual(['Carlito']);
  });

  test('a group carries its members, transformed through its child space', () => {
    const group = find(build().shapes, 'Group')!;
    expect(group.children.map(c => c.name)).toEqual(['Member']);
    // the member is drawn in the group's rotated box: 4000000 x 2000000 child units into 2000000 x 1000000 EMU, turned 90°
    const m = group.children[0]!.transform;
    expect(m.map(v => Math.round(v * 1000) / 1000).slice(0, 4)).toEqual([0, 0.5, -0.5, 0]);
  });

  test('empty text still lays out (for the caret), but draws nothing', () => {
    const title = find(build().shapes, 'Title 1')!;
    expect(title.prims.some(p => p.kind === 'text')).toBe(true);
    const doc = new PresentationDocument(OpcPackage.open(deckBytes()));
    const resolved = resolveSlide(doc, SLIDE).shapes.find(s => s.name === 'Title 1')!;
    const empty = { ...resolved.text!, paragraphs: [{ ...resolved.text!.paragraphs[0]!, runs: [] }] };
    const t = shapeTextLayout(resolved, fonts, empty)!;
    expect(t.layout.lines).toHaveLength(1);
    expect(t.layout.lines[0]!.height).toBeGreaterThan(0);
  });
});

describe('drawing', () => {
  const ops = (calls: Call[]) => calls.map(c => c.op);

  test('the background is filled first, then every shape, text as glyph outlines', () => {
    const { ctx, calls } = recorder();
    drawSlide(ctx as never, build(), { scale: 0.0001, image: () => undefined });
    expect(calls[0]!.op).toBe('fill');
    expect(ops(calls).filter(o => o === 'fill').length).toBeGreaterThan(5);
    expect(ops(calls)).toContain('stroke');
    // nothing is drawn with a non-finite transform
    for (const c of calls) for (const v of c.transform) expect(Number.isFinite(v)).toBe(true);
  });

  test('a base transform comes before every transform the renderer sets', () => {
    const slide = build();
    const callout = find(slide.shapes, 'Callout')!;
    const plain = recorder(), based = recorder();
    drawShape(plain.ctx as never, callout, { scale: 1, image: () => undefined });
    drawShape(based.ctx as never, callout, { scale: 1, base: [2, 0, 0, 2, 10, 20], image: () => undefined });
    expect(based.calls.length).toBe(plain.calls.length);
    const a = plain.calls[0]!.transform, b = based.calls[0]!.transform;
    expect(b).toEqual([a[0] * 2, a[1] * 2, a[2] * 2, a[3] * 2, a[4] * 2 + 10, a[5] * 2 + 20]);
  });

  test('animation overrides hide, fade and move shapes without touching them', () => {
    const slide = build();
    const callout = find(slide.shapes, 'Callout')!;
    const hidden = recorder();
    drawShape(hidden.ctx as never, callout, { scale: 1, image: () => undefined, override: () => ({ hidden: true }) });
    expect(hidden.calls).toEqual([]);
    const faded = recorder();
    drawShape(faded.ctx as never, callout, { scale: 1, image: () => undefined, override: () => ({ opacity: 0.25 }) });
    expect(faded.calls.every(c => c.alpha === 0.25)).toBe(true);
    const moved = recorder(), still = recorder();
    drawShape(still.ctx as never, callout, { scale: 1, image: () => undefined });
    drawShape(moved.ctx as never, callout, { scale: 1, image: () => undefined, override: () => ({ matrix: [1, 0, 0, 1, 1000, 0] }) });
    expect(moved.calls[0]!.transform[4]).toBeCloseTo(still.calls[0]!.transform[4] + 1000, 6);
  });
});
