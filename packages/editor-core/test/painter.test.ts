// The format painter: what it picks up from a shape, and what putting that on another shape changes and leaves.
import { checkOrder, child, children, runsOf } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  addSlide, copyFormat, copyTextFormat, formatBody, formatParagraphs, formatRuns, insertPicture, insertShape, insertTable, insertTextBox, pasteFormat,
  prepareImage, setFill, setLine, setPictureFill, setShadow, setTextLink, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0, 123, 64, 232, 221, 0, 0, 0, 15, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 86, 199, 45, 179, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);
const ALL = { start: { p: 0, offset: 0 }, end: { p: 0, offset: 99 } };

const shape = (ed: Editor, id: number, part = SLIDE) => ed.shape(part, id)!;
const drawn = (ed: Editor, id: number, part = SLIDE) => resolveSlide(ed.doc, part).shapes.find(s => s.id === id && s.layer === 'slide')!;

/**
 * A styled source: a red rectangle with a dashed 3 pt blue line and a shadow, whose text "Source" is bold italic
 * 28 pt green Georgia, right-aligned and anchored at the bottom; and a plain ellipse saying "Target".
 */
function styled(): { ed: Editor; src: number; dst: number } {
  const ed = openEditor(deckBytes());
  const a = insertShape(ed, SLIDE, 'rect', { x: 1000000, y: 1000000, w: 2000000, h: 1000000 });
  ed.run(a.command);
  const b = insertShape(ed, SLIDE, 'ellipse', { x: 4000000, y: 1000000, w: 2000000, h: 1000000 });
  ed.run(b.command);
  const s = () => shape(ed, a.id);
  ed.run(setFill(ed, SLIDE, [s()], { kind: 'srgb', hex: 'FF0000' }));
  ed.run(setLine(ed, SLIDE, [s()], { color: { kind: 'srgb', hex: '0000FF' }, width: 38100, dash: 'dash' }));
  ed.run(setShadow(ed, SLIDE, [s()], 'bottom-right'));
  for (const [id, text] of [[a.id, 'Source'], [b.id, 'Target']] as const) {
    const body = shape(ed, id).txBody!;
    const p = children(body, 'a:p')[0]!;
    p.children.splice(p.children.length - 1, 0, { kind: 'element', prefix: 'a', local: 'r', ns: p.ns, attrs: [], parent: p, selfClosing: false, children: [] } as never);
    const r = children(p, 'a:r')[0]!;
    r.children.push({ kind: 'element', prefix: 'a', local: 'rPr', ns: p.ns, attrs: [{ prefix: null, local: 'lang', value: id === a.id ? 'en-US' : 'fr-FR', ns: '' }], parent: r, selfClosing: true, children: [] } as never,
      { kind: 'element', prefix: 'a', local: 't', ns: p.ns, attrs: [], parent: r, selfClosing: false, children: [{ kind: 'text', value: text }] } as never);
    ed.changed([SLIDE]);
  }
  ed.run(formatRuns(ed, SLIDE, s().txBody!, ALL, { b: true, i: true, sz: 2800, latin: 'Georgia', fill: { kind: 'srgb', hex: '00AA00' } }));
  ed.run(formatParagraphs(ed, SLIDE, s().txBody!, ALL, { algn: 'r' }));
  ed.run(formatBody(ed, SLIDE, s().txBody!, { anchor: 'b' }));
  ed.log.clear();
  return { ed, src: a.id, dst: b.id };
}

describe('picking up and putting on a shape', () => {
  test('the target takes the fill, line, shadow and text look of the source, and keeps its own shape, place, text and language', () => {
    const { ed, src, dst } = styled();
    const clip = copyFormat(shape(ed, src));
    const before = shape(ed, dst).xfrm;
    expect(cycle(ed, () => pasteFormat(ed, SLIDE, [shape(ed, dst)], clip))).toEqual([]);
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, dst)], clip));
    const s = drawn(ed, src), d = drawn(ed, dst);
    expect(d.fill).toEqual(s.fill);
    expect(d.line).toEqual(s.line);
    expect(d.effects.outerShadow).toEqual(s.effects.outerShadow);
    const run = d.text!.paragraphs[0]!.runs[0]!;
    expect(run.text).toBe('Target');
    expect(run.props).toMatchObject({ bold: true, italic: true, size: 28, latin: 'Georgia' });
    expect(run.props.fill).toEqual(s.text!.paragraphs[0]!.runs[0]!.props.fill);
    expect(d.text!.paragraphs[0]!.props.align).toBe('r');
    expect(d.text!.body.anchor).toBe('b');
    // what is not a look stays the target's own
    expect(shape(ed, dst).xfrm).toEqual(before);
    expect(serializeElement(shape(ed, dst).spPr!)).toContain('prst="ellipse"');
    expect(serializeElement(child(children(children(shape(ed, dst).txBody!, 'a:p')[0]!, 'a:r')[0]!, 'a:rPr')!)).toContain('lang="fr-FR"');
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Format Painter');
  });

  test('several targets at once are one undo step; a look also goes onto a shape of another slide', () => {
    const { ed, src, dst } = styled();
    const second = addSlide(ed, LAYOUT);
    ed.run(second.command);
    const other = insertShape(ed, second.part, 'rect', { x: 1000000, y: 1000000, w: 1000000, h: 1000000 });
    ed.run(other.command);
    const third = insertShape(ed, SLIDE, 'triangle', { x: 7000000, y: 1000000, w: 1000000, h: 1000000 });
    ed.run(third.command);
    ed.log.clear();
    const clip = copyFormat(shape(ed, src));
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, dst), shape(ed, third.id)], clip));
    expect(drawn(ed, third.id).fill).toEqual(drawn(ed, src).fill);
    ed.undo();
    expect(drawn(ed, dst).fill).not.toEqual(drawn(ed, src).fill);
    expect(ed.log.canUndo).toBe(false);
    ed.run(pasteFormat(ed, second.part, [shape(ed, other.id, second.part)], clip));
    expect(drawn(ed, other.id, second.part).fill).toEqual(drawn(ed, src).fill);
    expect(checkOrder(ed.doc.root(second.part))).toEqual([]);
  });

  test('a source with no fill of its own gives the target its theme style back; no fill is carried as no fill', () => {
    const { ed, src, dst } = styled();
    const plain = insertShape(ed, SLIDE, 'rect', { x: 1000000, y: 3000000, w: 1000000, h: 500000 });
    ed.run(plain.command);
    // onto the styled shape: the red fill, the line and the shadow go, the theme's look comes back
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, src)], copyFormat(shape(ed, plain.id))));
    expect(serializeElement(shape(ed, src).spPr!)).not.toMatch(/solidFill|a:ln|effectLst/);
    expect(drawn(ed, src).fill).toEqual(drawn(ed, plain.id).fill);
    ed.run(setFill(ed, SLIDE, [shape(ed, plain.id)], 'none'));
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, dst)], copyFormat(shape(ed, plain.id))));
    expect(drawn(ed, dst).fill.kind).toBe('none');
  });
});

describe('what the painter leaves alone', () => {
  test('a picture fill is not carried, and one the target has stays when the clip brings no fill', async () => {
    const { ed, src, dst } = styled();
    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100 });
    ed.run(setPictureFill(ed, SLIDE, [shape(ed, src)], img));
    const clip = copyFormat(shape(ed, src));
    expect(clip.shape!.props.join('')).not.toContain('blip');
    expect(clip.shape!.props.join('')).toContain('<a:ln');
    ed.run(setPictureFill(ed, SLIDE, [shape(ed, dst)], img));
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, dst)], clip));
    expect(drawn(ed, dst).fill.kind).toBe('image');                      // kept: the clip brought a line, no fill
    expect(drawn(ed, dst).line).toEqual(drawn(ed, src).line);
  });

  test('a picture and a table keep their look; their text, where they have any, still takes the text look', async () => {
    const { ed, src } = styled();
    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100 });
    const pic = insertPicture(ed, SLIDE, img);
    ed.run(pic.command);
    const table = insertTable(ed, SLIDE, 2, 2, { x: 1000000, y: 4000000, w: 4000000, h: 800000 });
    ed.run(table.command);
    const before = [serializeElement(shape(ed, pic.id).el), serializeElement(shape(ed, table.id).el)];
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, pic.id), shape(ed, table.id)], copyFormat(shape(ed, src))));
    expect([serializeElement(shape(ed, pic.id).el), serializeElement(shape(ed, table.id).el)]).toEqual(before);
  });

  test('links on the target text stay links; the source link is not copied', () => {
    const { ed, src, dst } = styled();
    ed.run(setTextLink(ed, SLIDE, shape(ed, src).txBody!, ALL, { url: 'https://source.example' }));
    ed.run(setTextLink(ed, SLIDE, shape(ed, dst).txBody!, ALL, { url: 'https://target.example' }));
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, dst)], copyFormat(shape(ed, src))));
    const run = drawn(ed, dst).text!.paragraphs[0]!.runs[0]!;
    expect(run.props.hyperlink?.url).toBe('https://target.example');
    expect(run.props.bold).toBe(true);
  });
});

describe('text only', () => {
  test('a look picked up from text carries no shape formatting: only the text of the target changes', () => {
    const { ed, src, dst } = styled();
    const rPr = runsOf(children(shape(ed, src).txBody!, 'a:p')[0]!)[0]!.rPr;
    const clip = copyTextFormat(rPr);
    expect(clip.shape).toBeUndefined();
    const fill = drawn(ed, dst).fill;
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, dst)], clip));
    expect(drawn(ed, dst).fill).toEqual(fill);
    expect(drawn(ed, dst).text!.paragraphs[0]!.runs[0]!.props).toMatchObject({ bold: true, italic: true, size: 28 });
    expect(drawn(ed, dst).text!.paragraphs[0]!.props.align).not.toBe('r');
    expect(copyTextFormat(undefined)).toEqual({});
  });

  test('a text box takes the text look and stays unfilled when the source is a text box', () => {
    const { ed, src } = styled();
    const box = insertTextBox(ed, SLIDE, { x: 1000000, y: 5000000, w: 3000000, h: 400000 }, 'In a box');
    ed.run(box.command);
    ed.run(pasteFormat(ed, SLIDE, [shape(ed, box.id)], copyTextFormat(runsOf(children(shape(ed, src).txBody!, 'a:p')[0]!)[0]!.rPr)));
    expect(drawn(ed, box.id).fill.kind).toBe('none');
    expect(drawn(ed, box.id).text!.paragraphs[0]!.runs[0]!.props.size).toBe(28);
  });
});
