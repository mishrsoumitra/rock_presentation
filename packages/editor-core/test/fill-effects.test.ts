// Fills and effects beyond a plain colour: gradients, picture fills, transparency, arrowheads and shadows. Each is
// checked in the XML it writes, in what the resolver then paints, and by exact undo and redo.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, child, PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  fillColorOf, insertPicture, insertShape, prepareImage, setFill, setFillTransparency, setGradientFill, setLine, setPictureFill, setShadow, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0, 123, 64, 232, 221, 0, 0, 0, 15, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 86, 199, 45, 179, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);

/** A deck with a rectangle (filled by its theme style) and a straight line on its slide. */
function withShapes(): { ed: Editor; rect: number; line: number } {
  const ed = openEditor(deckBytes());
  const rect = insertShape(ed, SLIDE, 'rect', { x: 1000000, y: 1000000, w: 2000000, h: 1000000 });
  ed.run(rect.command);
  const line = insertShape(ed, SLIDE, 'line', { x: 1000000, y: 3000000, w: 2000000, h: 0 });
  ed.run(line.command);
  ed.log.clear();
  return { ed, rect: rect.id, line: line.id };
}
const shape = (ed: Editor, id: number) => ed.shape(SLIDE, id)!;
const spPr = (ed: Editor, id: number) => serializeElement(shape(ed, id).spPr!);
const drawn = (ed: Editor, id: number) => resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!;
const media = (ed: Editor) => ed.pkg.partNames().filter(n => n.startsWith('/ppt/media/')).length;

describe('gradient fill', () => {
  test('two colours along an angle, turning with the shape; the resolver paints a linear gradient of those stops', () => {
    const { ed, rect } = withShapes();
    const g = { from: { kind: 'srgb', hex: 'FFFFFF' } as const, to: { kind: 'scheme', val: 'accent1' } as const, angle: 90 };
    expect(cycle(ed, () => setGradientFill(ed, SLIDE, [shape(ed, rect)], g))).toEqual([]);
    ed.run(setGradientFill(ed, SLIDE, [shape(ed, rect)], g));
    expect(spPr(ed, rect)).toContain('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent1"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>');
    const fill = drawn(ed, rect).fill;
    expect(fill.kind).toBe('linear');
    if (fill.kind !== 'linear') return;
    expect(fill.angle).toBe(90);
    expect(fill.stops.map(s => s.pos)).toEqual([0, 1]);
    expect(fill.stops[0]!.color).toMatchObject({ r: 1, g: 1, b: 1 });
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Shape Fill');
  });

  test('angles wrap into a full turn; a gradient replaces a solid fill and a solid fill replaces a gradient', () => {
    const { ed, rect } = withShapes();
    const red = { kind: 'srgb', hex: 'FF0000' } as const, blue = { kind: 'srgb', hex: '0000FF' } as const;
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], red));
    ed.run(setGradientFill(ed, SLIDE, [shape(ed, rect)], { from: red, to: blue, angle: -45 }));
    expect(spPr(ed, rect)).toContain('<a:lin ang="18900000"');
    expect(spPr(ed, rect)).not.toContain('solidFill');
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], blue));
    expect(spPr(ed, rect)).not.toContain('gradFill');
    expect(drawn(ed, rect).fill.kind).toBe('solid');
  });
});

describe('picture fill', () => {
  test('the shape is filled with the picture, stored once however many shapes use it; replacing it drops the old picture', async () => {
    const { ed, rect, line } = withShapes();
    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100 });
    expect(cycle(ed, () => setPictureFill(ed, SLIDE, [shape(ed, rect)], img))).toEqual([]);
    expect(media(ed)).toBe(0);
    ed.run(setPictureFill(ed, SLIDE, [shape(ed, rect), shape(ed, line)], img));
    expect(media(ed)).toBe(1);
    expect(spPr(ed, rect)).toMatch(/<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId\d+"\/><a:srcRect\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>/);
    expect(drawn(ed, rect).fill.kind).toBe('image');
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    // a second picture replaces the first, which nothing uses any more
    const other = PNG.slice();
    other[other.length - 13] = other[other.length - 13]! ^ 1;
    const img2 = await prepareImage(ed, { bytes: other, contentType: 'image/png', ext: 'png', width: 200, height: 100 });
    ed.run(setPictureFill(ed, SLIDE, [shape(ed, rect), shape(ed, line)], img2));
    expect(media(ed)).toBe(1);
    // back to a colour: the picture is still in the deck until save prunes nothing it should not
    const back = new PresentationDocument(OpcPackage.open(ed.save()));
    expect(back.slide(SLIDE).shapes.find(s => s.id === rect)).toBeDefined();
  });

  test('a picture itself has no fill to set, and its own image is untouched', async () => {
    const { ed } = withShapes();
    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100 });
    const pic = insertPicture(ed, SLIDE, img);
    ed.run(pic.command);
    const before = serializeElement(shape(ed, pic.id).el);
    ed.run(setPictureFill(ed, SLIDE, [shape(ed, pic.id)], img));
    expect(serializeElement(shape(ed, pic.id).el)).toBe(before);
  });
});

describe('transparency', () => {
  test('a solid fill gets an alpha; 0 takes it off again', () => {
    const { ed, rect } = withShapes();
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], { kind: 'srgb', hex: 'FF0000' }));
    expect(cycle(ed, () => setFillTransparency(ed, SLIDE, [shape(ed, rect)], 25))).toEqual([]);
    ed.run(setFillTransparency(ed, SLIDE, [shape(ed, rect)], 25));
    expect(spPr(ed, rect)).toContain('<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="75000"/></a:srgbClr></a:solidFill>');
    const fill = drawn(ed, rect).fill;
    expect(fill.kind === 'solid' && fill.color.a).toBeCloseTo(0.75);
    ed.run(setFillTransparency(ed, SLIDE, [shape(ed, rect)], 60));
    expect(spPr(ed, rect).match(/a:alpha/g)?.length).toBe(1);             // replaced, not added twice
    ed.run(setFillTransparency(ed, SLIDE, [shape(ed, rect)], 0));
    expect(spPr(ed, rect)).toContain('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>');
    expect(ed.log.undoLabel).toBe('Transparency');
  });

  test("a shape filled by its theme style gets that colour as its own fill, to carry the alpha", () => {
    const { ed, rect } = withShapes();
    expect(spPr(ed, rect)).not.toContain('Fill');
    const before = drawn(ed, rect).fill;
    ed.run(setFillTransparency(ed, SLIDE, [shape(ed, rect)], 50));
    expect(spPr(ed, rect)).toContain('<a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill>');
    const after = drawn(ed, rect).fill;
    expect(before.kind === 'solid' && after.kind === 'solid' && [after.color.r, after.color.g, after.color.b]).toEqual(before.kind === 'solid' && [before.color.r, before.color.g, before.color.b]);
    expect(after.kind === 'solid' && after.color.a).toBeCloseTo(0.5);
  });

  test('every stop of a gradient gets the alpha; no fill stays no fill; out-of-range values are clamped', () => {
    const { ed, rect } = withShapes();
    ed.run(setGradientFill(ed, SLIDE, [shape(ed, rect)], { from: { kind: 'srgb', hex: 'FFFFFF' }, to: { kind: 'srgb', hex: '000000' }, angle: 0 }));
    ed.run(setFillTransparency(ed, SLIDE, [shape(ed, rect)], 40));
    expect(spPr(ed, rect).match(/<a:alpha val="60000"\/>/g)?.length).toBe(2);
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], 'none'));
    ed.run(setFillTransparency(ed, SLIDE, [shape(ed, rect)], 40));
    expect(spPr(ed, rect)).toContain('<a:noFill/>');
    expect(spPr(ed, rect)).not.toContain('alpha');
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], { kind: 'srgb', hex: '00FF00' }));
    ed.run(setFillTransparency(ed, SLIDE, [shape(ed, rect)], 250));
    expect(spPr(ed, rect)).toContain('<a:alpha val="0"/>');
  });
});

describe('the colour a shape is filled with', () => {
  test("its own colour, or its theme style's; none for no fill or a gradient", () => {
    const { ed, rect } = withShapes();
    expect(fillColorOf(shape(ed, rect))).toEqual({ kind: 'scheme', val: 'accent1' });
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], { kind: 'scheme', val: 'accent2', lumMod: 40000, lumOff: 60000 }));
    expect(fillColorOf(shape(ed, rect))).toEqual({ kind: 'scheme', val: 'accent2', lumMod: 40000, lumOff: 60000 });
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], { kind: 'srgb', hex: '#12ab34' }));
    expect(fillColorOf(shape(ed, rect))).toEqual({ kind: 'srgb', hex: '12AB34' });
    ed.run(setFill(ed, SLIDE, [shape(ed, rect)], 'none'));
    expect(fillColorOf(shape(ed, rect))).toBeUndefined();
    ed.run(setGradientFill(ed, SLIDE, [shape(ed, rect)], { from: { kind: 'srgb', hex: 'FFFFFF' }, to: { kind: 'srgb', hex: '000000' }, angle: 0 }));
    expect(fillColorOf(shape(ed, rect))).toBeUndefined();
  });
});

describe('arrowheads', () => {
  test('a line gets an arrow at its end, its start, or both, in schema order with its other line properties', () => {
    const { ed, line } = withShapes();
    expect(cycle(ed, () => setLine(ed, SLIDE, [shape(ed, line)], { tailEnd: 'triangle' }))).toEqual([]);
    ed.run(setLine(ed, SLIDE, [shape(ed, line)], { tailEnd: 'triangle', width: 25400, dash: 'dash', color: { kind: 'srgb', hex: '000000' } }));
    ed.run(setLine(ed, SLIDE, [shape(ed, line)], { headEnd: 'oval' }));
    expect(serializeElement(child(shape(ed, line).spPr!, 'a:ln')!)).toBe('<a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/><a:headEnd type="oval"/><a:tailEnd type="triangle"/></a:ln>');
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    const ends = drawn(ed, line).line!;
    expect(ends.head?.type).toBe('oval');
    expect(ends.tail?.type).toBe('triangle');
    ed.run(setLine(ed, SLIDE, [shape(ed, line)], { headEnd: null, tailEnd: 'none' }));
    const ln = serializeElement(child(shape(ed, line).spPr!, 'a:ln')!);
    expect(ln).not.toContain('headEnd');
    expect(ln).toContain('<a:tailEnd type="none"/>');
  });
});

describe('shadow', () => {
  test('a preset shadow is painted by the resolver; no shadow is an empty effect list; other effects are kept', () => {
    const { ed, rect } = withShapes();
    expect(cycle(ed, () => setShadow(ed, SLIDE, [shape(ed, rect)], 'bottom-right'))).toEqual([]);
    ed.run(setShadow(ed, SLIDE, [shape(ed, rect)], 'bottom-right'));
    expect(spPr(ed, rect)).toContain('<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>');
    expect(drawn(ed, rect).effects.outerShadow).toMatchObject({ blurRad: 50800, dist: 38100 });
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    ed.run(setShadow(ed, SLIDE, [shape(ed, rect)], 'center'));
    expect(spPr(ed, rect).match(/outerShdw /g)?.length).toBe(1);
    expect(drawn(ed, rect).effects.outerShadow?.dist ?? 0).toBe(0);
    ed.run(setShadow(ed, SLIDE, [shape(ed, rect)], null));
    expect(spPr(ed, rect)).toContain('<a:effectLst/>');
    expect(drawn(ed, rect).effects.outerShadow).toBeUndefined();
    expect(ed.log.undoLabel).toBe('Shadow');
    expect(() => setShadow(ed, SLIDE, [shape(ed, rect)], 'sideways' as never)).toThrow(/no such shadow/);
  });
});
