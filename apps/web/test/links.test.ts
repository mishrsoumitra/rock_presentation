// What a click in the slide show hits: a link on a run of text, then a link on a shape.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { buildSlide } from '@pptx/render';
import { resolveSlide } from '@pptx/resolve';
import { FontRegistry } from '@pptx/text';
import { deckBytes } from '../../../packages/resolve/test/deck.ts';
import { isSafeUrl, linkAt } from '../src/ui/links.ts';

const fonts = new FontRegistry();
const FONTS = join(import.meta.dirname, '../../../assets/fonts');
for (const dir of ['carlito', 'caladea', 'liberation']) {
  for (const f of readdirSync(join(FONTS, dir))) if (f.endsWith('.ttf')) fonts.addBytes(new Uint8Array(readFileSync(join(FONTS, dir, f))));
}

test('a click on a linked run of text finds its link; a click elsewhere finds none', () => {
  const doc = new PresentationDocument(OpcPackage.open(deckBytes()));
  const slide = buildSlide(resolveSlide(doc, '/ppt/slides/slide1.xml'), { fonts });
  const body = slide.shapes.find(s => s.name === 'Content Placeholder 2')!;
  const text = body.prims.find(p => p.kind === 'text')!;
  if (text.kind !== 'text') throw new Error('no text');
  const line = text.layout.lines.find(l => l.runs.some(r => r.props.hyperlink))!;
  const run = line.runs.find(r => r.props.hyperlink)!;
  const m = text.transform;
  const u = run.x + run.advance / 2, v = line.top + line.height / 2;
  const at = (x: number, y: number) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]] as const;
  expect(linkAt(slide, ...at(u, v))).toEqual({ url: 'https://example.com/' });
  expect(linkAt(slide, ...at(u, text.layout.lines[0]!.top + 1))).toBeUndefined();
});

test('only web and mail links are followed', () => {
  expect(isSafeUrl('https://example.com/')).toBe(true);
  expect(isSafeUrl('mailto:someone@example.com')).toBe(true);
  expect(isSafeUrl('javascript:alert(1)')).toBe(false);
  expect(isSafeUrl('file:///etc/passwd')).toBe(false);
  expect(isSafeUrl('C:\\\\Windows\\\\calc.exe')).toBe(false);
});
