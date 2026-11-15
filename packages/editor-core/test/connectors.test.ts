// Connectors glued to shapes: where a shape's connection points are, attaching an end, and connectors following
// the shapes they are attached to when those move, resize, turn, align or flip.
import { checkOrder, child } from '@pptx/model';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  align, attachments, connectionSites, connectorEnds, flip, frameBetween, insertShape, insertTextBox, isConnector, moveBy, nearestSite, rotateQuarter,
  setConnectorEnds, setTransforms, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';

/** Two rectangles side by side (A at x 1M, B at x 5M, both 2M by 1M at y 1M) and a straight connector off on its own. */
function scene(): { ed: Editor; a: number; b: number; c: number } {
  const ed = openEditor(deckBytes());
  const a = insertShape(ed, SLIDE, 'rect', { x: 1000000, y: 1000000, w: 2000000, h: 1000000 });
  ed.run(a.command);
  const b = insertShape(ed, SLIDE, 'rect', { x: 5000000, y: 1000000, w: 2000000, h: 1000000 });
  ed.run(b.command);
  const c = insertShape(ed, SLIDE, 'straightConnector1', { x: 500000, y: 4000000, w: 1000000, h: 500000 });
  ed.run(c.command);
  ed.log.clear();
  return { ed, a: a.id, b: b.id, c: c.id };
}
const shape = (ed: Editor, id: number) => ed.shape(SLIDE, id)!;
const sites = (ed: Editor, id: number) => connectionSites(shape(ed, id)).map(s => [Math.round(s.x), Math.round(s.y)]);
const ends = (ed: Editor, id: number) => { const e = connectorEnds(shape(ed, id)); return [[Math.round(e.start.x), Math.round(e.start.y)], [Math.round(e.end.x), Math.round(e.end.y)]]; };
/** Attach the connector from A's right side (point 3) to B's left side (point 1). */
function connect(ed: Editor, a: number, b: number, c: number): void {
  const from = connectionSites(shape(ed, a))[3]!, to = connectionSites(shape(ed, b))[1]!;
  ed.run(setConnectorEnds(ed, SLIDE, shape(ed, c), from, to, { start: { id: a, index: 3 }, end: { id: b, index: 1 } }));
  ed.log.seal();
}

describe('connection points', () => {
  test('a rectangle has four, at the middles of its sides: top, left, bottom, right', () => {
    const { ed, a } = scene();
    expect(sites(ed, a)).toEqual([[2000000, 1000000], [1000000, 1500000], [2000000, 2000000], [3000000, 1500000]]);
  });

  test('they turn and flip with the shape', () => {
    const { ed, a } = scene();
    ed.run(rotateQuarter(ed, SLIDE, [shape(ed, a)], 1));                  // a quarter turn clockwise about (2M, 1.5M)
    expect(sites(ed, a)).toEqual([[2500000, 1500000], [2000000, 500000], [1500000, 1500000], [2000000, 2500000]]);
    ed.undo();
    ed.run(flip(ed, SLIDE, [shape(ed, a)], 'h'));
    expect(sites(ed, a)).toEqual([[2000000, 1000000], [3000000, 1500000], [2000000, 2000000], [1000000, 1500000]]);
  });

  test('other shapes have the points their geometry gives; a connector has none; a text box has its four sides', () => {
    const ed = openEditor(deckBytes());
    const o = insertShape(ed, SLIDE, 'ellipse', { x: 0, y: 0, w: 2000000, h: 1000000 });
    ed.run(o.command);
    expect(connectionSites(shape(ed, o.id)).length).toBe(8);
    const line = insertShape(ed, SLIDE, 'line', { x: 0, y: 0, w: 1000000, h: 0 });
    ed.run(line.command);
    expect(connectionSites(shape(ed, line.id))).toEqual([]);
    const box = insertTextBox(ed, SLIDE, { x: 1000000, y: 3000000, w: 2000000, h: 400000 }, 'x');
    ed.run(box.command);
    expect(connectionSites(shape(ed, box.id)).length).toBe(4);
  });

  test('the nearest point within reach, never one of the shape being dragged', () => {
    const { ed, a, b } = scene();
    const all = [shape(ed, a), shape(ed, b)];
    expect(nearestSite(all, { x: 3050000, y: 1480000 }, 100000)).toMatchObject({ site: { index: 3, x: 3000000, y: 1500000 } });
    expect(nearestSite(all, { x: 3050000, y: 1480000 }, 100000)!.shape.id).toBe(a);
    expect(nearestSite(all, { x: 4000000, y: 1500000 }, 100000)).toBeUndefined();
    expect(nearestSite(all, { x: 4950000, y: 1500000 }, 100000, shape(ed, b).el)).toBeUndefined();
    expect(nearestSite(all, { x: 4000000, y: 1500000 }, 1100000)!.shape.id).toBe(a);        // both within reach: the nearer... A's is 1M away, B's 1M: the first found
  });
});

describe('a connector and its ends', () => {
  test('ends are the corners of its frame, which corner being said by its flips; a frame is rebuilt from two points', () => {
    const { ed, c } = scene();
    expect(isConnector(shape(ed, c))).toBe(true);
    expect(ends(ed, c)).toEqual([[500000, 4000000], [1500000, 4500000]]);
    expect(frameBetween({ x: 300, y: 100 }, { x: 100, y: 400 })).toEqual({ x: 100, y: 100, cx: 200, cy: 300, rot: 0, flipH: true, flipV: false });
    expect(frameBetween({ x: 100, y: 400 }, { x: 300, y: 100 })).toEqual({ x: 100, y: 100, cx: 200, cy: 300, rot: 0, flipH: false, flipV: true });
  });

  test('attaching: the ends go to the points, and the file records which shape and which point', () => {
    const { ed, a, b, c } = scene();
    const from = connectionSites(shape(ed, a))[3]!, to = connectionSites(shape(ed, b))[1]!;
    expect(cycle(ed, () => setConnectorEnds(ed, SLIDE, shape(ed, c), from, to, { start: { id: a, index: 3 }, end: { id: b, index: 1 } }))).toEqual([]);
    connect(ed, a, b, c);
    expect(ends(ed, c)).toEqual([[3000000, 1500000], [5000000, 1500000]]);
    expect(attachments(shape(ed, c))).toEqual({ start: { id: a, index: 3 }, end: { id: b, index: 1 } });
    expect(serializeElement(child(child(shape(ed, c).el, 'p:nvCxnSpPr')!, 'p:cNvCxnSpPr')!)).toBe(`<p:cNvCxnSpPr><a:stCxn id="${a}" idx="3"/><a:endCxn id="${b}" idx="1"/></p:cNvCxnSpPr>`);
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Move End Point');
  });

  test('an end dragged away is detached; the other end stays attached', () => {
    const { ed, a, b, c } = scene();
    connect(ed, a, b, c);
    ed.run(setConnectorEnds(ed, SLIDE, shape(ed, c), connectorEnds(shape(ed, c)).start, { x: 6000000, y: 4000000 }, { end: null }));
    expect(attachments(shape(ed, c))).toEqual({ start: { id: a, index: 3 } });
    expect(ends(ed, c)).toEqual([[3000000, 1500000], [6000000, 4000000]]);
  });

  test('a drag of many moves with one key is one undo step', () => {
    const { ed, c } = scene();
    const before = ends(ed, c);
    for (const x of [2000000, 2500000, 3000000]) ed.run(setConnectorEnds(ed, SLIDE, shape(ed, c), { x: 500000, y: 4000000 }, { x, y: 4500000 }, {}, 'end:1'));
    ed.undo();
    expect(ends(ed, c)).toEqual(before);
    expect(ed.log.canUndo).toBe(false);
  });
});

describe('connectors follow their shapes', () => {
  test('moving a shape moves the attached end with it, in the same undo step; the other end stays on its shape', () => {
    const { ed, a, b, c } = scene();
    connect(ed, a, b, c);
    expect(cycle(ed, () => moveBy(ed, SLIDE, [shape(ed, b)], 500000, 1000000))).toEqual([]);
    ed.run(moveBy(ed, SLIDE, [shape(ed, b)], 500000, 1000000));
    expect(ends(ed, c)).toEqual([[3000000, 1500000], [5500000, 2500000]]);
    ed.undo();
    expect(ends(ed, c)).toEqual([[3000000, 1500000], [5000000, 1500000]]);
    expect(ed.log.canUndo).toBe(true);                                   // only the move was undone, not the attaching
  });

  test('a shape moved past the other: the connector flips to run the other way and stays attached', () => {
    const { ed, a, b, c } = scene();
    connect(ed, a, b, c);
    ed.run(moveBy(ed, SLIDE, [shape(ed, b)], -6000000, 2000000));        // B now left of and below A
    expect(ends(ed, c)).toEqual([[3000000, 1500000], [-1000000, 3500000]]);
    expect(shape(ed, c).xfrm).toMatchObject({ flipH: true, flipV: false, x: -1000000, cx: 4000000, cy: 2000000 });
  });

  test('resizing, turning, flipping and aligning a shape all carry the end to where the point now is', () => {
    const { ed, a, b, c } = scene();
    connect(ed, a, b, c);
    const x = shape(ed, a).xfrm!;
    ed.run(setTransforms(ed, SLIDE, [{ shape: shape(ed, a), xfrm: { ...x, cx: 3000000, cy: 2000000 } }], 'Resize'));
    expect(ends(ed, c)[0]).toEqual([4000000, 2000000]);                  // the right side's middle of the bigger rectangle
    ed.run(rotateQuarter(ed, SLIDE, [shape(ed, a)], 1));
    expect(ends(ed, c)[0]).toEqual(sites(ed, a)[3]);
    ed.run(flip(ed, SLIDE, [shape(ed, a)], 'v'));
    expect(ends(ed, c)[0]).toEqual(sites(ed, a)[3]);
    ed.run(align(ed, SLIDE, [shape(ed, a)], 'left'));
    expect(ends(ed, c)[0]).toEqual(sites(ed, a)[3]);
    expect(ends(ed, c)[1]).toEqual([5000000, 1500000]);
  });

  test('both shapes moved together take the whole connector along; a connector that is itself moved is left as moved', () => {
    const { ed, a, b, c } = scene();
    connect(ed, a, b, c);
    ed.run(moveBy(ed, SLIDE, [shape(ed, a), shape(ed, b)], 100000, 200000));
    expect(ends(ed, c)).toEqual([[3100000, 1700000], [5100000, 1700000]]);
    ed.run(moveBy(ed, SLIDE, [shape(ed, a), shape(ed, b), shape(ed, c)], 100000, 0));
    expect(ends(ed, c)).toEqual([[3200000, 1700000], [5200000, 1700000]]);
  });

  test('a shape with no connector attached moves as before, and unattached connectors stay where they are', () => {
    const { ed, a, c } = scene();
    const before = ends(ed, c);
    ed.run(moveBy(ed, SLIDE, [shape(ed, a)], 123456, 0));
    expect(ends(ed, c)).toEqual(before);
    expect(shape(ed, a).xfrm!.x).toBe(1123456);
  });
});
