// Copying and pasting whole slides: in the deck they came from, after the slide was cut, and into another deck.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, PresentationDocument } from '@pptx/model';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  addSlide, bodyPlainText, copySlide, deleteSlide, insertPicture, insertTextBox, notesText, pasteSlide, prepareImage, setNotes, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const BOX = { x: 1000000, y: 1000000, w: 4000000, h: 800000 };
/** A tiny PNG; pictures are compared by their bytes, never decoded here. */
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0, 123, 64, 232, 221, 0, 0, 0, 15, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 86, 199, 45, 179, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);

/** The texts of the text boxes of each slide, in order: what tells one slide from another. */
const boxes = (ed: Editor) => ed.doc.slideParts().map(p => ed.doc.slide(p).shapes.filter(s => s.name.startsWith('TextBox')).map(s => bodyPlainText(s.txBody!)).join('+'));
const media = (ed: Editor) => ed.pkg.partNames().filter(n => n.startsWith('/ppt/media/'));
const pictures = (ed: Editor, part: string) => ed.doc.slide(part).shapes.filter(s => s.kind === 'pic').length;

/** Three slides marked "one", "two", "three"; the first has a picture and speaker notes. */
async function deck(): Promise<Editor> {
  const ed = openEditor(deckBytes());
  ed.run(insertTextBox(ed, SLIDE, BOX, 'one').command);
  ed.run(insertPicture(ed, SLIDE, await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100 })).command);
  ed.run(setNotes(ed, SLIDE, 'notes of one'));
  for (const mark of ['two', 'three']) {
    const s = addSlide(ed, LAYOUT);
    ed.run(s.command);
    ed.run(insertTextBox(ed, s.part, BOX, mark).command);
  }
  ed.log.clear();
  return ed;
}

describe('copy and paste a slide in its own deck', () => {
  test('the copy goes where asked, with the same layout, its own notes, and the picture shared, not stored twice', async () => {
    const ed = await deck();
    const clip = copySlide(ed, SLIDE);
    expect(boxes(ed)).toEqual(['one', 'two', 'three']);                  // copying changes nothing
    const pics = media(ed).length;
    expect(cycle(ed, () => pasteSlide(ed, clip, 2).command)).toEqual([]);
    const pasted = pasteSlide(ed, clip, 2);
    ed.run(pasted.command);
    const part = pasted.part!;
    expect(boxes(ed)).toEqual(['one', 'two', 'one', 'three']);
    expect(ed.doc.slideParts()[2]).toBe(part);
    expect(ed.doc.layoutOf(part)).toBe(LAYOUT);
    expect(pictures(ed, part)).toBe(1);
    expect(media(ed).length).toBe(pics);
    expect(notesText(ed, part)).toBe('notes of one');
    expect(ed.doc.notesOf(part)).not.toBe(ed.doc.notesOf(SLIDE));
    ed.run(setNotes(ed, part, 'changed on the copy'));
    expect(notesText(ed, SLIDE)).toBe('notes of one');
    expect(checkOrder(ed.presentationRoot)).toEqual([]);
    expect(ed.log.undoLabel).toBe('Notes');
  });

  test('without a position it goes last; the same clip pastes any number of times', async () => {
    const ed = await deck();
    const clip = copySlide(ed, ed.doc.slideParts()[1]!);
    ed.run(pasteSlide(ed, clip).command);
    ed.run(pasteSlide(ed, clip, 0).command);
    expect(boxes(ed)).toEqual(['two', 'one', 'two', 'three', 'two']);
  });

  test('a saved deck opens with the pasted slide, its notes and its picture', async () => {
    const ed = await deck();
    ed.run(pasteSlide(ed, copySlide(ed, SLIDE)).command);
    const back = new PresentationDocument(OpcPackage.open(ed.save()));
    const last = back.slideParts().at(-1)!;
    expect(back.slide(last).shapes.filter(s => s.kind === 'pic').length).toBe(1);
    expect(back.notesOf(last)).toBeDefined();
    expect(back.notesMaster()).toBeDefined();
    expect(back.slideParts().length).toBe(4);
  });
});

describe('cut and paste a slide', () => {
  test('the slide comes back whole after it was deleted, with the picture only it used', async () => {
    const ed = await deck();
    const clip = copySlide(ed, SLIDE);
    ed.run(deleteSlide(ed, SLIDE));
    expect(boxes(ed)).toEqual(['two', 'three']);
    expect(media(ed).length).toBe(0);                                    // the picture went with its only slide
    const pasted = pasteSlide(ed, clip);
    ed.run(pasted.command);
    expect(boxes(ed)).toEqual(['two', 'three', 'one']);
    expect(pictures(ed, pasted.part!)).toBe(1);
    expect(media(ed).length).toBe(1);
    expect(notesText(ed, pasted.part!)).toBe('notes of one');
    // undo the paste, then the cut: the deck is as it was at the start
    ed.undo();
    ed.undo();
    expect(boxes(ed)).toEqual(['one', 'two', 'three']);
    expect(notesText(ed, SLIDE)).toBe('notes of one');
  });

  test('the clip is a copy from when it was taken: later changes to the slide are not in it', async () => {
    const ed = await deck();
    const clip = copySlide(ed, SLIDE);
    ed.run(insertTextBox(ed, SLIDE, BOX, 'added later').command);
    ed.run(pasteSlide(ed, clip).command);
    expect(boxes(ed)).toEqual(['one+added later', 'two', 'three', 'one']);
  });
});

describe('links between slides', () => {
  /** Give the first slide's text box a link to another slide. */
  function link(ed: Editor, to: string): void {
    const id = ed.pkg.addRelationship(SLIDE, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', to);
    const box = ed.doc.slide(SLIDE).shapes.find(s => s.name.startsWith('TextBox'))!;
    const rPr = box.txBody!.children.flatMap(c => (c.kind === 'element' && c.local === 'p' ? c.children : [])).flatMap(c => (c.kind === 'element' && c.local === 'r' ? c.children : []))
      .find(c => c.kind === 'element' && c.local === 'rPr');
    if (!rPr || rPr.kind !== 'element') throw new Error('no run');
    rPr.selfClosing = false;
    rPr.children.push({ kind: 'element', prefix: 'a', local: 'hlinkClick', ns: rPr.ns, attrs: [{ prefix: 'r', local: 'id', value: id, ns: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }, { prefix: null, local: 'action', value: 'ppaction://hlinksldjump', ns: '' }], children: [], parent: rPr, selfClosing: true } as never);
    ed.changed([SLIDE]);
  }
  const linksOf = (ed: Editor, part: string) => ed.doc.targets(part, 'slide');
  const hlinks = (ed: Editor, part: string) => serializeElement(ed.doc.root(part)).match(/<a:hlinkClick/g)?.length ?? 0;

  test('a link to a slide that is still there is kept; to one that is gone, removed, the text staying', async () => {
    const ed = await deck();
    const two = ed.doc.slideParts()[1]!;
    link(ed, two);
    const clip = copySlide(ed, SLIDE);
    const kept = pasteSlide(ed, clip);
    ed.run(kept.command);
    expect(linksOf(ed, kept.part!)).toEqual([two]);
    expect(hlinks(ed, kept.part!)).toBe(hlinks(ed, SLIDE));                 // every link of the original, the new one among them

    ed.run(deleteSlide(ed, two));
    const dropped = pasteSlide(ed, clip);
    ed.run(dropped.command);
    expect(linksOf(ed, dropped.part!)).toEqual([]);
    expect(hlinks(ed, dropped.part!)).toBe(hlinks(ed, SLIDE));              // like the original, which lost that link when the slide was deleted
    expect(boxes(ed).at(-1)).toBe('one');
    // every relationship the pasted slide's XML names exists
    const back = OpcPackage.open(ed.save());
    for (const p of new PresentationDocument(back).slideParts()) {
      const ids = new Set(back.relationships(p).map(r => r.id));
      const xml = new TextDecoder().decode(back.getPart(p)!.bytes);
      for (const m of xml.matchAll(/r:(?:id|embed|link)="([^"]+)"/g)) expect(ids.has(m[1]!), `${p} ${m[1]}`).toBe(true);
    }
  });
});

describe('paste into another deck', () => {
  test('the slide takes the first layout there and brings its picture and notes with it', async () => {
    const from = await deck();
    const clip = copySlide(from, SLIDE);
    const to = openEditor(deckBytes());
    expect(cycle(to, () => pasteSlide(to, clip).command)).toEqual([]);
    const pasted = pasteSlide(to, clip);
    to.run(pasted.command);
    const part = pasted.part!;
    expect(to.doc.slideParts().length).toBe(2);
    expect(to.doc.layoutOf(part)).toBe(to.doc.layoutsOf(to.doc.masterParts()[0]!)[0]);
    expect(pictures(to, part)).toBe(1);
    expect(media(to).length).toBe(1);
    expect(notesText(to, part)).toBe('notes of one');
    expect(to.doc.notesMaster()).toBeDefined();
    const back = new PresentationDocument(OpcPackage.open(to.save()));
    expect(back.slideParts().length).toBe(2);
    // the deck it came from is untouched
    expect(boxes(from)).toEqual(['one', 'two', 'three']);
  });
});
