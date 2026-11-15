// Video and audio: the media part, the poster picture that stands for it, and the two relationships it is found by.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { deleteShapes, duplicateShapes, duplicateSlide, insertMedia, prepareImage, prepareMedia, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0, 123, 64, 232, 221, 0, 0, 0, 15, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 86, 199, 45, 179, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);
/** Bytes that stand for a media file: the editor stores them, it never reads them. */
const CLIP = new Uint8Array(2048).map((_, i) => (i * 7) & 255);

const mediaParts = (ed: Editor) => ed.pkg.partNames().filter(n => n.startsWith('/ppt/media/'));
const drawn = (ed: Editor, id: number) => resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!;
async function prepared(ed: Editor) {
  return {
    video: await prepareMedia(ed, { bytes: CLIP, contentType: 'video/mp4', ext: 'mp4' }),
    audio: await prepareMedia(ed, { bytes: CLIP.slice(0, 512), contentType: 'audio/mpeg', ext: 'mp3' }),
    poster: await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 640, height: 360 }),
  };
}

describe('inserting media', () => {
  test('a video: its file as a part, a poster picture in its place, and both relationships to the file; undone exactly', async () => {
    const ed = openEditor(deckBytes());
    const { video, poster } = await prepared(ed);
    expect(cycle(ed, () => insertMedia(ed, SLIDE, 'video', video, poster).command)).toEqual([]);
    expect(mediaParts(ed)).toEqual([]);
    const m = insertMedia(ed, SLIDE, 'video', video, poster, undefined, 'launch.mp4');
    ed.run(m.command);
    expect(mediaParts(ed).sort()).toEqual(['/ppt/media/image1.png', '/ppt/media/media1.mp4']);
    expect(ed.pkg.getPart('/ppt/media/media1.mp4')!.contentType).toBe('video/mp4');
    const rels = ed.pkg.relationships(SLIDE).filter(r => r.target.endsWith('media1.mp4')).map(r => r.type);
    expect(rels.sort()).toEqual(['http://schemas.microsoft.com/office/2007/relationships/media', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video']);
    const shape = ed.shape(SLIDE, m.id)!;
    expect(shape.kind).toBe('pic');
    expect(shape.name).toBe('launch.mp4');
    expect(serializeElement(shape.el)).toMatch(/<p:nvPr><a:videoFile r:link="rId\d+"\/><p:extLst><p:ext uri="\{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230\}"><p14:media[^>]* r:embed="rId\d+"\/><\/p:ext><\/p:extLst><\/p:nvPr>/);
    // 640 by 360 pixels at 96 dpi, centred
    expect(shape.xfrm).toMatchObject({ cx: 6096000, cy: 3429000, x: (12192000 - 6096000) / 2 });
    expect(drawn(ed, m.id).media).toEqual({ kind: 'video', part: '/ppt/media/media1.mp4' });
    expect(drawn(ed, m.id).fill.kind).toBe('image');
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Insert Video');
  });

  test('a sound: an audio relationship, and the box it is given', async () => {
    const ed = openEditor(deckBytes());
    const { audio, poster } = await prepared(ed);
    const m = insertMedia(ed, SLIDE, 'audio', audio, poster, { x: 100, y: 200, w: 609600, h: 609600 });
    ed.run(m.command);
    expect(drawn(ed, m.id).media).toEqual({ kind: 'audio', part: '/ppt/media/media1.mp3' });
    expect(ed.pkg.relationships(SLIDE).some(r => r.type.endsWith('/audio'))).toBe(true);
    expect(ed.shape(SLIDE, m.id)!.xfrm).toMatchObject({ x: 100, y: 200, cx: 609600, cy: 609600 });
    expect(ed.shape(SLIDE, m.id)!.name).toMatch(/^Audio \d+$/);
    expect(ed.log.undoLabel).toBe('Insert Audio');
  });

  test('the same file inserted twice is stored once; a saved deck opens with the media and its bytes intact', async () => {
    const ed = openEditor(deckBytes());
    const { video, poster } = await prepared(ed);
    ed.run(insertMedia(ed, SLIDE, 'video', video, poster).command);
    ed.run(insertMedia(ed, SLIDE, 'video', await prepareMedia(ed, { bytes: CLIP.slice(), contentType: 'video/mp4', ext: 'mp4' }), poster).command);
    expect(mediaParts(ed).filter(n => n.endsWith('.mp4')).length).toBe(1);
    const back = OpcPackage.open(ed.save());
    expect(Array.from(back.getPart('/ppt/media/media1.mp4')!.bytes)).toEqual(Array.from(CLIP));
    expect(new PresentationDocument(back).slide(SLIDE).shapes.filter(s => s.media).length).toBe(2);
  });
});

describe('media and what holds it', () => {
  test('deleting the last shape that plays a file takes the file with it; a copy on the slide or of the slide shares it', async () => {
    const ed = openEditor(deckBytes());
    const { video, poster } = await prepared(ed);
    const m = insertMedia(ed, SLIDE, 'video', video, poster);
    ed.run(m.command);
    const dup = duplicateShapes(ed, SLIDE, [ed.shape(SLIDE, m.id)!]);
    ed.run(dup.command);
    expect(mediaParts(ed).filter(n => n.endsWith('.mp4')).length).toBe(1);
    expect(drawn(ed, dup.ids[0]!).media?.part).toBe('/ppt/media/media1.mp4');
    const copy = duplicateSlide(ed, SLIDE);
    ed.run(copy.command);
    expect(mediaParts(ed).filter(n => n.endsWith('.mp4')).length).toBe(1);
    ed.run(deleteShapes(ed, SLIDE, [ed.shape(SLIDE, m.id)!, ed.shape(SLIDE, dup.ids[0]!)!]));
    expect(mediaParts(ed).filter(n => n.endsWith('.mp4')).length).toBe(1);            // the copied slide still plays it
    ed.run(deleteShapes(ed, copy.part!, ed.doc.slide(copy.part!).shapes.filter(s => s.media)));
    expect(mediaParts(ed).filter(n => n.endsWith('.mp4'))).toEqual([]);
  });
});
