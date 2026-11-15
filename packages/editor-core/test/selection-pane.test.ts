// What the selection pane does to a shape: rename it, give it alt text, hide it and show it again.
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { group, insertShape, rename, setAltText, setShapesHidden, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';

/** A deck with two rectangles drawn on its slide; returns their ids. */
function withShapes(): { ed: Editor; a: number; b: number } {
  const ed = openEditor(deckBytes());
  const a = insertShape(ed, SLIDE, 'rect', { x: 1000000, y: 1000000, w: 1000000, h: 500000 });
  ed.run(a.command);
  const b = insertShape(ed, SLIDE, 'ellipse', { x: 3000000, y: 1000000, w: 1000000, h: 500000 });
  ed.run(b.command);
  ed.log.clear();
  return { ed, a: a.id, b: b.id };
}
const drawnIds = (ed: Editor) => resolveSlide(ed.doc, SLIDE).shapes.filter(s => s.layer === 'slide').map(s => s.id);
const cNvPr = (ed: Editor, id: number) => serializeElement(ed.shape(SLIDE, id)!.cNvPr!);

describe('rename and alt text', () => {
  test('a new name replaces the old one, with characters XML must escape kept as typed', () => {
    const { ed, a } = withShapes();
    expect(cycle(ed, () => rename(ed, SLIDE, ed.shape(SLIDE, a)!, 'Logo <left> & "top"'))).toEqual([]);
    ed.run(rename(ed, SLIDE, ed.shape(SLIDE, a)!, 'Logo <left> & "top"'));
    expect(ed.shape(SLIDE, a)!.name).toBe('Logo <left> & "top"');
    const back = new PresentationDocument(OpcPackage.open(ed.save()));
    expect(back.slide(SLIDE).shapes.map(s => s.name)).toContain('Logo <left> & "top"');
    expect(ed.log.undoLabel).toBe('Rename');
  });

  test('alt text is set, changed, and removed when emptied', () => {
    const { ed, a } = withShapes();
    expect(ed.shape(SLIDE, a)!.descr).toBeUndefined();
    expect(cycle(ed, () => setAltText(ed, SLIDE, ed.shape(SLIDE, a)!, 'A blue rectangle'))).toEqual([]);
    ed.run(setAltText(ed, SLIDE, ed.shape(SLIDE, a)!, 'A blue rectangle'));
    expect(ed.shape(SLIDE, a)!.descr).toBe('A blue rectangle');
    ed.run(setAltText(ed, SLIDE, ed.shape(SLIDE, a)!, 'Company logo'));
    expect(cNvPr(ed, a).match(/descr=/g)?.length).toBe(1);
    expect(new PresentationDocument(OpcPackage.open(ed.save())).slide(SLIDE).shapes.find(s => s.id === a)!.descr).toBe('Company logo');
    ed.run(setAltText(ed, SLIDE, ed.shape(SLIDE, a)!, ''));
    expect(cNvPr(ed, a)).not.toContain('descr');
  });
});

describe('hiding shapes', () => {
  test('a hidden shape is not drawn, keeps its place in the order, and comes back when shown', () => {
    const { ed, a, b } = withShapes();
    const order = ed.slide(SLIDE).shapes.map(s => s.id);
    expect(drawnIds(ed)).toContain(a);
    expect(cycle(ed, () => setShapesHidden(ed, SLIDE, [ed.shape(SLIDE, a)!], true))).toEqual([]);
    ed.run(setShapesHidden(ed, SLIDE, [ed.shape(SLIDE, a)!], true));
    expect(ed.shape(SLIDE, a)!.hidden).toBe(true);
    expect(drawnIds(ed)).not.toContain(a);
    expect(drawnIds(ed)).toContain(b);
    expect(ed.slide(SLIDE).shapes.map(s => s.id)).toEqual(order);
    expect(ed.log.undoLabel).toBe('Hide');
    // hidden survives a save
    expect(new PresentationDocument(OpcPackage.open(ed.save())).slide(SLIDE).shapes.find(s => s.id === a)!.hidden).toBe(true);
    ed.run(setShapesHidden(ed, SLIDE, [ed.shape(SLIDE, a)!], false));
    expect(cNvPr(ed, a)).not.toContain('hidden');
    expect(drawnIds(ed)).toContain(a);
    expect(ed.log.undoLabel).toBe('Show');
  });

  test('several at once is one undo step; hiding twice is the same as once', () => {
    const { ed, a, b } = withShapes();
    const both = () => [ed.shape(SLIDE, a)!, ed.shape(SLIDE, b)!];
    ed.run(setShapesHidden(ed, SLIDE, both(), true));
    ed.run(setShapesHidden(ed, SLIDE, both(), true));
    expect(cNvPr(ed, a).match(/hidden=/g)?.length).toBe(1);
    expect(drawnIds(ed)).not.toContain(a);
    expect(drawnIds(ed)).not.toContain(b);
    ed.undo();
    ed.undo();
    expect(drawnIds(ed)).toEqual(expect.arrayContaining([a, b]));
    expect(ed.log.canUndo).toBe(false);
  });

  test('a shape inside a group can be hidden on its own; hiding the group hides them all', () => {
    const { ed, a, b } = withShapes();
    ed.run(group(ed, SLIDE, [ed.shape(SLIDE, a)!, ed.shape(SLIDE, b)!]));
    const g = ed.slide(SLIDE).shapes.find(s => s.kind === 'grpSp' && s.children.some(c => c.id === a))!;
    const members = () => resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === g.id && s.layer === 'slide')?.children.map(c => c.id);
    expect(members()).toEqual([a, b]);
    ed.run(setShapesHidden(ed, SLIDE, [ed.shape(SLIDE, a)!], true));
    expect(members()).toEqual([b]);
    ed.run(setShapesHidden(ed, SLIDE, [ed.shape(SLIDE, g.id)!], true));
    expect(drawnIds(ed)).not.toContain(g.id);
  });
});
