// Slide backgrounds: a theme colour, a gradient, a picture; on one slide and on all of them.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, child, PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { addSlide, applyBackgroundToAll, deleteSlide, prepareImage, setSlideBackground, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0, 123, 64, 232, 221, 0, 0, 0, 15, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 86, 199, 45, 179, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);

/** A deck of three slides. */
function three(): { ed: Editor; slides: string[] } {
  const ed = openEditor(deckBytes());
  for (let i = 0; i < 2; i++) ed.run(addSlide(ed, LAYOUT).command);
  ed.log.clear();
  return { ed, slides: ed.doc.slideParts() };
}
const bg = (ed: Editor, part: string) => { const e = child(child(ed.doc.root(part), 'p:cSld'), 'p:bg'); return e ? serializeElement(e) : undefined; };
const drawn = (ed: Editor, part: string) => resolveSlide(ed.doc, part).background;
const media = (ed: Editor) => ed.pkg.partNames().filter(n => n.startsWith('/ppt/media/')).length;

describe('one slide', () => {
  test('a theme colour: written as the theme colour (so it follows the theme), drawn as that colour, undone exactly', () => {
    const { ed } = three();
    const spec = { color: { kind: 'scheme', val: 'accent2', lumMod: 75000 } as const };
    expect(cycle(ed, () => setSlideBackground(ed, [SLIDE], spec))).toEqual([]);
    ed.run(setSlideBackground(ed, [SLIDE], spec));
    expect(bg(ed, SLIDE)).toBe('<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>');
    expect(drawn(ed, SLIDE).kind).toBe('solid');
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Format Background');
  });

  test('a gradient, then no background of its own again', () => {
    const { ed } = three();
    const before = drawn(ed, SLIDE);
    ed.run(setSlideBackground(ed, [SLIDE], { gradient: { from: { kind: 'srgb', hex: 'FFFFFF' }, to: { kind: 'srgb', hex: '003366' }, angle: 90 } }));
    const fill = drawn(ed, SLIDE);
    expect(fill.kind).toBe('linear');
    expect(fill.kind === 'linear' && fill.angle).toBe(90);
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    ed.run(setSlideBackground(ed, [SLIDE], null));
    expect(bg(ed, SLIDE)).toBeUndefined();
    expect(drawn(ed, SLIDE)).toEqual(before);
  });

  test('a picture: stored once, related from the slide; replaced by a colour, the picture leaves the deck', async () => {
    const { ed } = three();
    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100 });
    expect(cycle(ed, () => setSlideBackground(ed, [SLIDE], { picture: img }))).toEqual([]);
    ed.run(setSlideBackground(ed, [SLIDE], { picture: img }));
    expect(bg(ed, SLIDE)).toMatch(/<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId\d+"\/><a:srcRect\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>/);
    expect(drawn(ed, SLIDE).kind).toBe('image');
    expect(media(ed)).toBe(1);
    expect(new PresentationDocument(OpcPackage.open(ed.save())).targets(SLIDE, 'image').length).toBe(1);
    ed.run(setSlideBackground(ed, [SLIDE], { color: { kind: 'srgb', hex: '112233' } }));
    expect(media(ed)).toBe(0);
    expect(ed.pkg.relationships(SLIDE).some(r => r.type.endsWith('/image'))).toBe(false);
  });
});

describe('all slides', () => {
  test('several slides at once are one undo step', () => {
    const { ed, slides } = three();
    ed.run(setSlideBackground(ed, slides, { color: { kind: 'srgb', hex: 'FFEEDD' } }));
    expect(slides.map(p => drawn(ed, p).kind)).toEqual(['solid', 'solid', 'solid']);
    ed.undo();
    expect(slides.map(p => bg(ed, p))).toEqual([undefined, undefined, undefined]);
    expect(ed.log.canUndo).toBe(false);
  });

  test('Apply to All gives every slide the background one has; a picture is shared, each slide with its own relationship to it', async () => {
    const { ed, slides } = three();
    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100 });
    ed.run(setSlideBackground(ed, [slides[1]!], { picture: img }));
    ed.log.seal();
    expect(cycle(ed, () => applyBackgroundToAll(ed, slides[1]!))).toEqual([]);
    ed.run(applyBackgroundToAll(ed, slides[1]!));
    expect(slides.map(p => drawn(ed, p).kind)).toEqual(['image', 'image', 'image']);
    expect(media(ed)).toBe(1);
    for (const p of slides) {
      const ids = new Set(ed.pkg.relationships(p).map(r => r.id));
      const id = /r:embed="([^"]+)"/.exec(bg(ed, p)!)![1]!;
      expect(ids.has(id), p).toBe(true);
      expect(checkOrder(ed.doc.root(p))).toEqual([]);
    }
    // a slide with the picture deleted: the others still show it
    ed.run(deleteSlide(ed, slides[1]!));
    expect(media(ed)).toBe(1);
    expect(ed.log.undoLabel).toBe('Delete Slide');
  });

  test('from a slide with no background of its own, Apply to All takes the others\' own backgrounds away', () => {
    const { ed, slides } = three();
    ed.run(setSlideBackground(ed, [slides[0]!, slides[2]!], { color: { kind: 'srgb', hex: 'FF0000' } }));
    ed.run(applyBackgroundToAll(ed, slides[1]!));
    expect(slides.map(p => bg(ed, p))).toEqual([undefined, undefined, undefined]);
  });
});
