// Dragging adjust handles: adjustFromHandle finds the adjust values that put a handle where it is dragged.
import { describe, expect, test } from 'vitest';
import {
  adjustFromHandle, evaluateGeometry, GeometryError, PRESET_NAMES, presetGeometry, TEXT_WARP_NAMES, textWarpGeometry,
  type AdjustHandle, type Guide, type ShapeGeometry,
} from '../src/index.ts';
import { adj, custGeom, valueOf } from './helpers.ts';

const W = 1000, H = 500;
const value = (adjust: Guide[], name: string) => {
  const g = adjust.find(a => a.name === name);
  return g ? Number(/^val (-?[0-9.]+)$/.exec(g.fmla)![1]) : undefined;
};

describe('adjustFromHandle', () => {
  const roundRect = presetGeometry('roundRect')!;

  test('roundRect: the handle on the top edge sets the corner radius, x / ss', () => {
    expect(adjustFromHandle(roundRect, W, H, [], 0, { x: 100, y: 0 })).toEqual([{ name: 'adj', fmla: 'val 20000' }]);
    // it only moves along x, so y does not matter
    expect(adjustFromHandle(roundRect, W, H, [], 0, { x: 100, y: 400 })).toEqual([{ name: 'adj', fmla: 'val 20000' }]);
  });

  test('values stay within the handle limits and are whole numbers', () => {
    expect(adjustFromHandle(roundRect, W, H, [], 0, { x: 900, y: 0 })).toEqual([{ name: 'adj', fmla: 'val 50000' }]);
    expect(adjustFromHandle(roundRect, W, H, [], 0, { x: -50, y: 0 })).toEqual([{ name: 'adj', fmla: 'val 0' }]);
    expect(adjustFromHandle(roundRect, W, H, [], 0, { x: 100.3, y: 0 })).toEqual([{ name: 'adj', fmla: 'val 20060' }]);
  });

  test('a polar handle follows the direction of the point (pie on a circle)', () => {
    const pie = presetGeometry('pie')!;
    for (const deg of [30, 135, 200, 350]) {
      const a = (deg * Math.PI) / 180;
      const far = { x: 500 + 2000 * Math.cos(a), y: 500 + 2000 * Math.sin(a) };   // well outside the circle
      expect(value(adjustFromHandle(pie, 1000, 1000, [], 0, far), 'adj1')).toBeCloseTo(deg * 60000, -1);
    }
  });

  test("the shape's other adjust values are kept, in avLst order, and unknown ones stay at the end", () => {
    const rightArrow = presetGeometry('rightArrow')!;
    const given = [{ name: 'extra', fmla: 'val 1' }, { name: 'adj2', fmla: 'val 70000' }];
    // handle 0 moves adj1 vertically: the shaft's top edge at y = vc - h * adj1 / 200000
    expect(adjustFromHandle(rightArrow, W, H, given, 0, { x: 0, y: 200 })).toEqual([
      { name: 'adj1', fmla: 'val 20000' }, { name: 'adj2', fmla: 'val 70000' }, { name: 'extra', fmla: 'val 1' },
    ]);
  });

  test('limits that are guides are evaluated for the size (rightArrow: maxAdj2 = 100000 w / ss)', () => {
    const rightArrow = presetGeometry('rightArrow')!;
    expect(value(adjustFromHandle(rightArrow, W, H, [], 1, { x: -5000, y: 0 }), 'adj2')).toBe(200000);
    expect(value(adjustFromHandle(rightArrow, 500, 500, [], 1, { x: -5000, y: 0 }), 'adj2')).toBe(100000);
  });

  test('a handle with no adjust value to move, or no limits, changes nothing', () => {
    const g = custGeom(`<a:avLst><a:gd name="adj" fmla="val 5"/></a:avLst><a:gdLst><a:gd name="g" fmla="val 1"/></a:gdLst>
      <a:ahLst><a:ahXY gdRefX="g" minX="0" maxX="10"><a:pos x="g" y="0"/></a:ahXY><a:ahXY gdRefX="adj"><a:pos x="adj" y="0"/></a:ahXY></a:ahLst>
      <a:pathLst/>`);
    expect(adjustFromHandle(g, W, H, [], 0, { x: 7, y: 0 })).toEqual([]);
    expect(adjustFromHandle(g, W, H, adj({ adj: 3 }), 1, { x: 7, y: 0 })).toEqual(adj({ adj: 3 }));
  });

  test('an index with no handle is an error', () => {
    expect(() => adjustFromHandle(roundRect, W, H, [], 1, { x: 0, y: 0 })).toThrow(GeometryError);
    expect(() => adjustFromHandle(roundRect, W, H, [], -1, { x: 0, y: 0 })).toThrow(GeometryError);
  });
});

/** The adjust values a handle moves and their limits: [gdRef, min, max, how it moves]. */
function axesOf(h: AdjustHandle): [string, string, string, 'x' | 'y' | 'polar'][] {
  const all: [string | undefined, string | undefined, string | undefined, 'x' | 'y' | 'polar'][] = h.kind === 'xy'
    ? [[h.gdRefX, h.minX, h.maxX, 'x'], [h.gdRefY, h.minY, h.maxY, 'y']]
    : [[h.gdRefAng, h.minAng, h.maxAng, 'polar'], [h.gdRefR, h.minR, h.maxR, 'polar']];
  return all.filter((a): a is [string, string, string, 'x' | 'y' | 'polar'] => a[0] !== undefined && a[1] !== undefined && a[2] !== undefined);
}

describe.each([
  ...PRESET_NAMES.map((n): [string, ShapeGeometry] => [n, presetGeometry(n)!]),
  ...TEXT_WARP_NAMES.map((n): [string, ShapeGeometry] => [n, textWarpGeometry(n)!]),
].filter(([, g]) => g.ahLst.length > 0))('%s', (_, geom) => {
  const w = 1600000, h = 1100000;
  const avNames = new Set(geom.avLst.map(g => g.name));

  test('moving each handle to where adjust value A puts it gives back A', () => {
    geom.ahLst.forEach((handle, i) => {
      const axes = axesOf(handle).filter(a => avNames.has(a[0]));
      for (const fraction of [0.3, 0.7]) {
        // choose A axis by axis, each within the limits the values before it give
        let target: Guide[] = [];
        const chosen: { ref: string; a: number; span: number; along: 'x' | 'y' | 'polar' }[] = [];
        for (const [ref, min, max, along] of axes) {
          const e = evaluateGeometry(geom, w, h, target);
          const lo = Math.min(valueOf(e, min), valueOf(e, max)), hi = Math.max(valueOf(e, min), valueOf(e, max));
          const unbounded = hi - lo > 21600000;   // callout tails: +-2147483647
          const a = Math.round(unbounded ? e.guides.get(ref)! + 25000 : lo + fraction * (hi - lo));
          target = [...target, { name: ref, fmla: `val ${a}` }];
          chosen.push({ ref, a, span: unbounded ? 100000 : hi - lo, along });
        }
        const at = (adjust: Guide[]) => evaluateGeometry(geom, w, h, adjust).handles[i]!;
        const goal = at(target);
        const result = adjustFromHandle(geom, w, h, [], i, goal);
        const reached = at(result);
        expect(Math.hypot(reached.x - goal.x, reached.y - goal.y), `handle ${i} at ${fraction}: ${JSON.stringify(result)}`)
          .toBeLessThan(1e-4 * Math.max(w, h));
        // where the handle moves when its value does, the value itself must come back too
        for (const { ref, a, span, along } of chosen) {
          const nudged = (d: number) => at(target.map(t => (t.name === ref ? { name: ref, fmla: `val ${a + d}` } : t)));
          const p = nudged(-span / 200), q = nudged(span / 200);
          const moves = along === 'x' ? Math.abs(p.x - q.x) : along === 'y' ? Math.abs(p.y - q.y) : Math.hypot(p.x - q.x, p.y - q.y);
          if (moves < 1e-3 * Math.max(w, h)) continue;
          expect(Math.abs(value(result, ref)! - a), `handle ${i} ${ref}: ${JSON.stringify(result)} for ${a}`).toBeLessThanOrEqual(Math.max(2, 1e-4 * span));
        }
      }
    });
  });
});
