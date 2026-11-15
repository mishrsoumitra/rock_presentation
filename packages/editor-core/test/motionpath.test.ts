// Motion paths drawn by hand: the points of a drag become a Custom Path animation of the shape.
import { listEffects, parseMotionPath, parseTiming, pointAt } from '@pptx/anim';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { addMotionPath, insertShape, motionPathsOf, pathFromPoints, simplifyPath, type Editor, type PathPoint } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const SIZE = { cx: 12192000, cy: 6858000 };

/** A drag as a mouse gives it: many points along the corners given, a little off the line. */
function drag(corners: PathPoint[], steps = 40, wobble = 2000): PathPoint[] {
  const out: PathPoint[] = [];
  for (let i = 0; i + 1 < corners.length; i++) {
    const a = corners[i]!, b = corners[i + 1]!;
    for (let k = 0; k < steps; k++) {
      const t = k / steps;
      out.push({ x: a.x + (b.x - a.x) * t + Math.sin(k * 1.7) * wobble, y: a.y + (b.y - a.y) * t + Math.cos(k * 2.3) * wobble });
    }
  }
  out.push(corners[corners.length - 1]!);
  return out;
}

describe('simplifying a drawn line', () => {
  test('a wobbly line through three corners keeps the corners and little else', () => {
    const corners = [{ x: 1e6, y: 1e6 }, { x: 5e6, y: 1e6 }, { x: 5e6, y: 4e6 }];
    const kept = simplifyPath(drag(corners), 20000);
    expect(kept.length).toBeLessThanOrEqual(5);
    expect(kept[0]).toEqual({ x: 1e6 + 0, y: 1e6 + 2000 });               // the first point as drawn
    expect(kept.at(-1)).toEqual(corners[2]);
    expect(kept.some(p => Math.hypot(p.x - 5e6, p.y - 1e6) < 150000)).toBe(true);   // the corner survives
  });

  test('a curve keeps enough points to stay within the tolerance of what was drawn', () => {
    const arc = Array.from({ length: 200 }, (_, i) => ({ x: 3e6 + Math.cos((i / 199) * Math.PI) * 2e6, y: 3e6 - Math.sin((i / 199) * Math.PI) * 2e6 }));
    const kept = simplifyPath(arc, 30000);
    expect(kept.length).toBeGreaterThan(6);
    expect(kept.length).toBeLessThan(40);
    for (const p of kept) expect(Math.abs(Math.hypot(p.x - 3e6, p.y - 3e6) - 2e6)).toBeLessThan(1);
  });

  test('one or two points are left as they are', () => {
    expect(simplifyPath([], 10)).toEqual([]);
    expect(simplifyPath([{ x: 1, y: 2 }], 10)).toEqual([{ x: 1, y: 2 }]);
    expect(simplifyPath([{ x: 1, y: 2 }, { x: 3, y: 4 }], 10)).toEqual([{ x: 1, y: 2 }, { x: 3, y: 4 }]);
  });
});

describe('a path from the points drawn', () => {
  test('offsets from the first point, as fractions of the slide, wherever it was drawn', () => {
    const here = pathFromPoints([{ x: 1000000, y: 1000000 }, { x: 1000000 + SIZE.cx / 4, y: 1000000 }, { x: 1000000 + SIZE.cx / 4, y: 1000000 + SIZE.cy / 2 }], SIZE);
    expect(here).toBe('M 0 0 L 0.25 0 L 0.25 0.5 E');
    const there = pathFromPoints([{ x: 7e6, y: 3e6 }, { x: 7e6 + SIZE.cx / 4, y: 3e6 }, { x: 7e6 + SIZE.cx / 4, y: 3e6 + SIZE.cy / 2 }], SIZE);
    expect(there).toBe(here);
    expect(pathFromPoints([{ x: 5e6, y: 5e6 }, { x: 5e6 - SIZE.cx / 8, y: 5e6 - SIZE.cy / 8 }], SIZE)).toBe('M 0 0 L -0.125 -0.125 E');
  });

  test('a click or a jitter goes nowhere: no path', () => {
    expect(pathFromPoints([], SIZE)).toBeUndefined();
    expect(pathFromPoints([{ x: 1e6, y: 1e6 }], SIZE)).toBeUndefined();
    expect(pathFromPoints([{ x: 1e6, y: 1e6 }, { x: 1e6 + 500, y: 1e6 - 300 }, { x: 1e6 - 200, y: 1e6 + 100 }], SIZE)).toBeUndefined();
  });

  test('what is written reads back as a path that ends where the drawing ended', () => {
    const corners = [{ x: 2e6, y: 2e6 }, { x: 6e6, y: 2e6 }, { x: 6e6, y: 5e6 }, { x: 3e6, y: 5e6 }];
    const path = pathFromPoints(drag(corners), SIZE)!;
    expect(path).toMatch(/^M 0 0 (L -?[\d.]+ -?[\d.]+ )+E$/);
    expect(path.split('L').length - 1).toBeLessThan(8);
    const parsed = parseMotionPath(path)!;
    const end = pointAt(parsed, 1);
    expect(end.x * SIZE.cx).toBeCloseTo(1e6, -5);                         // 3e6 - 2e6, to the wobble
    expect(end.y * SIZE.cy).toBeCloseTo(3e6, -5);
  });
});

describe('animating a shape along a drawn path', () => {
  function withShape(): { ed: Editor; id: number } {
    const ed = openEditor(deckBytes());
    const s = insertShape(ed, SLIDE, 'ellipse', { x: 1e6, y: 3e6, w: 1e6, h: 1e6 });
    ed.run(s.command);
    ed.log.clear();
    return { ed, id: s.id };
  }
  const L = [{ x: 4e6, y: 1e6 }, { x: 4e6 + SIZE.cx / 4, y: 1e6 }, { x: 4e6 + SIZE.cx / 4, y: 1e6 + SIZE.cy / 4 }];

  test('a Custom Path effect on a click, two seconds long, with the path drawn; undone exactly', () => {
    const { ed, id } = withShape();
    expect(cycle(ed, () => addMotionPath(ed, SLIDE, ed.shape(SLIDE, id)!, L))).toEqual([]);
    ed.run(addMotionPath(ed, SLIDE, ed.shape(SLIDE, id)!, L));
    const effects = listEffects(ed.slide(SLIDE));
    expect(effects.map(e => [e.spid, e.name, e.presetClass, e.presetID, e.trigger, e.duration])).toEqual([[id, 'Custom Path', 'path', 0, 'click', 2000]]);
    expect(serializeElement(ed.slide(SLIDE).timing!)).toContain('path="M 0 0 L 0.25 0 L 0.25 0.25 E"');
    expect(ed.log.undoLabel).toBe('Draw Motion Path');
    expect(() => parseTiming(ed.slide(SLIDE).timing!)).not.toThrow();
  });

  test('the shape\'s paths read back as offsets on the slide; a second path is a second effect; other shapes have none', () => {
    const { ed, id } = withShape();
    ed.run(addMotionPath(ed, SLIDE, ed.shape(SLIDE, id)!, L));
    ed.run(addMotionPath(ed, SLIDE, ed.shape(SLIDE, id)!, [{ x: 0, y: 0 }, { x: 0, y: SIZE.cy / 2 }]));
    const paths = motionPathsOf(ed, SLIDE, id);
    expect(paths.map(p => [p.index, p.path])).toEqual([[0, 'M 0 0 L 0.25 0 L 0.25 0.25 E'], [1, 'M 0 0 L 0 0.5 E']]);
    expect(paths[0]!.points).toEqual([{ x: 0, y: 0 }, { x: SIZE.cx / 4, y: 0 }, { x: SIZE.cx / 4, y: SIZE.cy / 4 }]);
    expect(motionPathsOf(ed, SLIDE, 2)).toEqual([]);
    const back = openEditor(ed.save());
    expect(motionPathsOf(back, SLIDE, id).map(p => p.path)).toEqual(['M 0 0 L 0.25 0 L 0.25 0.25 E', 'M 0 0 L 0 0.5 E']);
  });

  test('a drawing that goes nowhere adds nothing', () => {
    const { ed, id } = withShape();
    expect(() => addMotionPath(ed, SLIDE, ed.shape(SLIDE, id)!, [{ x: 5, y: 5 }, { x: 6, y: 6 }])).toThrow(/goes nowhere/);
    expect(listEffects(ed.slide(SLIDE))).toEqual([]);
  });
});
