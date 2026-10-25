// Exact geometry for shapes whose outline is known, and arcTo (ECMA-376 Part 1, 20.1.9.3) in every direction.
import { describe, expect, test } from 'vitest';
import { evaluateGeometry, pathBounds, presetGeometry, type GeometryPath, type PathCommand } from '../src/index.ts';
import { adj, cubicAt, custGeom, ends, fixed, fixedRect, near } from './helpers.ts';

const W = 1000, H = 500;
const preset = (name: string, w = W, h = H, adjust = {}) => evaluateGeometry(presetGeometry(name)!, w, h, adj(adjust));
const K = (4 / 3) * Math.tan(Math.PI / 8);   // control distance of a quarter-circle Bezier, per unit radius

/** Every curve of a path, with the point it starts from. */
function curves(path: GeometryPath): [[number, number], Extract<PathCommand, { op: 'C' }>][] {
  const out: [[number, number], Extract<PathCommand, { op: 'C' }>][] = [];
  let at: [number, number] = [0, 0];
  for (const c of path.commands) {
    if (c.op === 'C') out.push([at, c]);
    if (c.op !== 'Z') at = [c.x, c.y];
  }
  return out;
}

/** How far curves stray from the ellipse (cx, cy, rx, ry), as a fraction of its radius, sampled along each. */
function offEllipse(path: GeometryPath | ReturnType<typeof curves>, cx: number, cy: number, rx: number, ry: number): number {
  let worst = 0;
  for (const [p0, c] of Array.isArray(path) ? path : curves(path)) {
    for (let t = 0; t <= 1; t += 0.125) {
      const [x, y] = cubicAt(p0, c, t);
      worst = Math.max(worst, Math.abs(Math.hypot((x - cx) / rx, (y - cy) / ry) - 1));
    }
  }
  return worst;
}

describe('presets with a known outline', () => {
  test('rect', () => {
    const e = preset('rect');
    expect(e.paths).toEqual([{
      commands: [{ op: 'M', x: 0, y: 0 }, { op: 'L', x: W, y: 0 }, { op: 'L', x: W, y: H }, { op: 'L', x: 0, y: H }, { op: 'Z' }],
      fill: 'norm', stroke: true, extrusionOk: true,
    }]);
    expect(e.textRect).toEqual({ x: 0, y: 0, w: W, h: H });
    expect(e.connectionSites).toEqual([
      { x: 500, y: 0, angle: 16200000 }, { x: 0, y: 250, angle: 10800000 }, { x: 500, y: 500, angle: 5400000 }, { x: 1000, y: 250, angle: 0 },
    ]);
    expect(e.handles).toEqual([]);
  });

  test('roundRect at its default adjust value: corner radius ss * 16667 / 100000', () => {
    const e = preset('roundRect');
    const r = (Math.min(W, H) * 16667) / 100000;
    expect(e.guides.get('x1')).toBeCloseTo(r, 9);
    const path = e.paths[0]!;
    const expected: [number, number][] = [[0, r], [r, 0], [W - r, 0], [W, r], [W, H - r], [W - r, H], [r, H], [0, H - r]];
    ends(path).forEach((p, i) => near(p, expected[i]!, 1e-9));
    expect(path.commands.map(c => c.op).join('')).toBe('MCLCLCLCZ');
    // each corner is a quarter circle of radius r: one Bezier with its controls K r along the tangents
    const first = curves(path)[0]![1];
    near([first.x1, first.y1], [0, r - K * r], 1e-9);
    near([first.x2, first.y2], [r - K * r, 0], 1e-9);
    const corners: [number, number][] = [[r, r], [W - r, r], [W - r, H - r], [r, H - r]];
    curves(path).forEach((curve, i) => expect(offEllipse([curve], corners[i]![0], corners[i]![1], r, r)).toBeLessThan(3e-4));
    expect(e.handles).toEqual([{ x: r, y: 0 }]);
    const il = (r * 29289) / 100000;
    expect(fixedRect(e.textRect)).toEqual([il, il, W - 2 * il, H - 2 * il].map(fixed));
  });

  test('roundRect with the largest corner is a stadium: two half circles joined by straight lines', () => {
    const path = preset('roundRect', W, H, { adj: 50000 }).paths[0]!;
    const expected: [number, number][] = [[0, 250], [250, 0], [750, 0], [1000, 250], [1000, 250], [750, 500], [250, 500], [0, 250]];
    ends(path).forEach((p, i) => near(p, expected[i]!, 1e-9));
    const [a, b, c, d] = curves(path);
    expect(offEllipse([a!, d!], 250, 250, 250, 250)).toBeLessThan(3e-4);
    expect(offEllipse([b!, c!], 750, 250, 250, 250)).toBeLessThan(3e-4);
  });

  test('ellipse passes through the four edge midpoints and stays on the ellipse', () => {
    const e = preset('ellipse');
    const path = e.paths[0]!;
    expect(path.commands.map(c => c.op).join('')).toBe('MCCCCZ');
    const expected: [number, number][] = [[0, 250], [500, 0], [1000, 250], [500, 500], [0, 250]];
    ends(path).forEach((p, i) => near(p, expected[i]!, 1e-9));
    expect(offEllipse(path, 500, 250, 500, 250)).toBeLessThan(3e-4);   // the usual quarter-turn Bezier error
    const idx = 500 * Math.cos(Math.PI / 4), idy = 250 * Math.sin(Math.PI / 4);
    expect(fixedRect(e.textRect)).toEqual([500 - idx, 250 - idy, 2 * idx, 2 * idy].map(fixed));
  });

  test.each([[50000, 500], [25000, 250], [0, 0], [100000, 1000], [150000, 1000]])('triangle with adj %d has its apex at x = %d', (a, apex) => {
    const e = preset('triangle', W, H, { adj: a });
    expect(ends(e.paths[0]!)).toEqual([[0, H], [apex, 0], [W, H]]);
    expect(e.handles).toEqual([{ x: apex, y: 0 }]);
  });

  test('rightArrow', () => {
    const e = preset('rightArrow');
    // shaft half the height, head ss * 50% long: maxAdj2 = 200000, dx1 = 250, dy1 = 125
    expect(ends(e.paths[0]!)).toEqual([[0, 125], [750, 125], [750, 0], [1000, 250], [750, 500], [750, 375], [0, 375]]);
    expect(e.handles).toEqual([{ x: 0, y: 125 }, { x: 750, y: 0 }]);
    // the text rectangle reaches into the head, where the head is still as tall as the shaft
    expect(e.textRect).toEqual({ x: 0, y: 125, w: 750 + (125 * 250) / 250, h: 250 });
  });

  test('star5: points alternate between two ellipses, 36 degrees apart, the first point straight up', () => {
    for (const [w, h] of [[1000, 1000], [1600, 900]] as const) {
      const e = preset('star5', w, h);
      const g = (n: string) => e.guides.get(n)!;
      const [hc, svc, swd2, shd2] = [g('hc'), g('svc'), g('swd2'), g('shd2')];
      expect(swd2).toBeCloseTo((w / 2) * 1.05146, 9);   // hf and vf stretch the star to fill its box
      expect(shd2).toBeCloseTo((h / 2) * 1.10557, 9);
      const inner = 19098 / 50000;
      const pts = ends(e.paths[0]!);
      expect(pts).toHaveLength(10);
      pts.forEach((p, k) => {
        const a = ((198 + 36 * k) * Math.PI) / 180;
        const s = k % 2 === 0 ? 1 : inner;
        near(p, [hc + s * swd2 * Math.cos(a), svc + s * shd2 * Math.sin(a)], 1e-9);
      });
      near(pts[2]!, [w / 2, 0], 1e-9);
    }
  });

  test('pie with given angles', () => {
    // 0 to 90 degrees: a quarter, from the right edge to the bottom
    let e = preset('pie', W, H, { adj1: 0, adj2: 5400000 });
    expect(e.paths[0]!.commands.map(c => c.op).join('')).toBe('MCLZ');
    ([[1000, 250], [500, 500], [500, 250]] as const).forEach((p, i) => near(ends(e.paths[0]!)[i]!, p, 1e-9));
    // 90 back round to 0: three quarters, clockwise
    e = preset('pie', W, H, { adj1: 5400000, adj2: 0 });
    expect(e.paths[0]!.commands.map(c => c.op).join('')).toBe('MCCCLZ');
    near(ends(e.paths[0]!)[0]!, [500, 500], 1e-9);
    near(ends(e.paths[0]!)[3]!, [1000, 250], 1e-9);
    // angles off the axes: the arc must end where the preset's own guides put the end point
    e = preset('pie', W, H, { adj1: 2700000, adj2: 8100000 });
    const pts = ends(e.paths[0]!);
    near(pts[0]!, [e.guides.get('x1')!, e.guides.get('y1')!], 1e-9);
    near(pts[pts.length - 2]!, [e.guides.get('x2')!, e.guides.get('y2')!], 1e-6);
    near(pts[pts.length - 1]!, [500, 250], 1e-9);
    expect(offEllipse(e.paths[0]!, 500, 250, 500, 250)).toBeLessThan(3e-4);
    // a 45-degree ray from the centre meets the ellipse where the arc starts
    expect(Math.atan2(pts[0]![1] - 250, pts[0]![0] - 500)).toBeCloseTo(Math.PI / 4, 9);
  });

  test('upArrow is downArrow upside down', () => {
    const flip = ([x, y]: [number, number], h: number): [number, number] => [x, h - y];
    const sorted = (pts: [number, number][]) => pts.map(p => p.map(fixed)).sort((a, b) => a[0]! - b[0]! || a[1]! - b[1]!);
    for (const [w, h] of [[1000, 500], [500, 1000], [300, 300]] as const) {
      for (const values of [{}, { adj1: 20000, adj2: 80000 }, { adj1: 100000, adj2: 400000 }, { adj1: 0, adj2: 0 }]) {
        const up = preset('upArrow', w, h, values);
        const down = preset('downArrow', w, h, values);
        expect(sorted(ends(up.paths[0]!))).toEqual(sorted(ends(down.paths[0]!).map(p => flip(p, h))));
        up.handles.forEach((p, i) => near(p, flip([down.handles[i]!.x, down.handles[i]!.y], h), 1e-9));
        const d = down.textRect;
        expect(fixedRect(up.textRect)).toEqual([d.x, h - d.y - d.h, d.w, d.h].map(fixed));
        // connection sites: positions flipped, and directions too (down becomes up)
        const sites = (e: typeof up, mirror: boolean) => e.connectionSites
          .map(s => [s.x, mirror ? h - s.y : s.y, mirror ? (21600000 - s.angle) % 21600000 : s.angle].map(fixed).join(' '))
          .sort();
        expect(sites(up, false)).toEqual(sites(down, true));
      }
    }
  });
});

describe('arcTo', () => {
  // A circle of radius 100 x 50 around (100, 50), drawn in a 200 x 100 shape.
  const arc = (stAng: number, swAng: number, start: [number, number], path = '') => evaluateGeometry(custGeom(`<a:pathLst><a:path ${path}>
      <a:moveTo><a:pt x="${start[0]}" y="${start[1]}"/></a:moveTo>
      <a:arcTo wR="100" hR="50" stAng="${stAng}" swAng="${swAng}"/>
    </a:path></a:pathLst>`), 200, 100).paths[0]!;
  const pointAt = (deg: number): [number, number] => {
    const a = (deg * Math.PI) / 180;
    const t = Math.atan2(100 * Math.sin(a), 50 * Math.cos(a));
    return [100 + 100 * Math.cos(t), 50 + 50 * Math.sin(t)];
  };

  test('a 360-degree sweep is a full ellipse, not nothing', () => {
    const path = arc(0, 21600000, [200, 50]);
    expect(path.commands.map(c => c.op).join('')).toBe('MCCCC');
    const pts = ends(path);
    [[200, 50], [100, 100], [0, 50], [100, 0], [200, 50]].forEach((p, i) => near(pts[i]!, p as [number, number], 1e-9));
    expect(offEllipse(path, 100, 50, 100, 50)).toBeLessThan(3e-4);
    expect(pathBounds([path])).toEqual({ x: expect.closeTo(0, 9), y: expect.closeTo(0, 9), w: expect.closeTo(200, 9), h: expect.closeTo(100, 9) });
  });

  test('a full turn from an angle off the axes ends where it started', () => {
    const start = pointAt(37);
    const path = arc(37 * 60000, 21600000, start);
    expect(curves(path)).toHaveLength(4);
    near(ends(path).at(-1)!, start, 1e-9);
    expect(offEllipse(path, 100, 50, 100, 50)).toBeLessThan(3e-4);
  });

  test('two turns go round twice; more is bounded but ends in the same place', () => {
    expect(curves(arc(0, 43200000, [200, 50]))).toHaveLength(8);
    const many = arc(0, 21600000 * 1000 + 5400000, [200, 50]);
    expect(curves(many).length).toBeLessThanOrEqual(9);
    near(ends(many).at(-1)!, [100, 100], 1e-6);
  });

  test('positive sweeps go clockwise (y down), negative ones counter-clockwise', () => {
    near(ends(arc(0, 5400000, [200, 50])).at(-1)!, [100, 100], 1e-9);
    near(ends(arc(0, -5400000, [200, 50])).at(-1)!, [100, 0], 1e-9);
    const back = arc(0, -21600000, [200, 50]);
    expect(curves(back)).toHaveLength(4);
    near(ends(back)[1]!, [100, 0], 1e-9);     // the first quarter goes up
    near(ends(back).at(-1)!, [200, 50], 1e-9);
    expect(curves(back)[0]![1].y1).toBeLessThan(50);
  });

  test('negative sweeps off the axes', () => {
    const start = pointAt(45);
    const path = arc(2700000, -8100000, start);   // 45 back to -90 degrees: the top
    near(ends(path).at(-1)!, [100, 0], 1e-9);
    expect(curves(path)).toHaveLength(2);
    expect(offEllipse(path, 100, 50, 100, 50)).toBeLessThan(3e-4);
    const other = arc(2700000, -5400000, start);  // to -45 degrees
    near(ends(other).at(-1)!, pointAt(-45), 1e-9);
  });

  test('a sweep of 0 draws nothing and leaves the pen where it was', () => {
    const g = custGeom(`<a:pathLst><a:path><a:moveTo><a:pt x="10" y="20"/></a:moveTo>
      <a:arcTo wR="5" hR="5" stAng="0" swAng="0"/><a:lnTo><a:pt x="30" y="20"/></a:lnTo></a:path></a:pathLst>`);
    expect(evaluateGeometry(g, 100, 100).paths[0]!.commands).toEqual([{ op: 'M', x: 10, y: 20 }, { op: 'L', x: 30, y: 20 }]);
  });

  test('the arc is drawn from the pen and continues from its end', () => {
    const g = custGeom(`<a:pathLst><a:path><a:moveTo><a:pt x="0" y="50"/></a:moveTo>
      <a:arcTo wR="50" hR="50" stAng="cd2" swAng="cd2"/><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst>`);
    const path = evaluateGeometry(g, 100, 100).paths[0]!;
    expect(path.commands.map(c => c.op).join('')).toBe('MCCLZ');
    near(ends(path)[2]!, [100, 50], 1e-9);
  });

  test('a flat ellipse (one radius 0) still lands on the right points', () => {
    const g = custGeom(`<a:pathLst><a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo>
      <a:arcTo wR="0" hR="50" stAng="3cd4" swAng="cd2"/></a:path></a:pathLst>`);
    near(ends(evaluateGeometry(g, 100, 100).paths[0]!).at(-1)!, [0, 100], 1e-9);
  });

  test('in a path with its own w and h the arc is computed in that space, then scaled', () => {
    // a unit circle in a 2000 x 2000 path space, scaled into a 1000 x 500 shape: 45 degrees there is not 45 degrees here
    const g = custGeom(`<a:pathLst><a:path w="2000" h="2000"><a:moveTo><a:pt x="2000" y="1000"/></a:moveTo>
      <a:arcTo wR="1000" hR="1000" stAng="0" swAng="2700000"/></a:path></a:pathLst>`);
    const path = evaluateGeometry(g, 1000, 500).paths[0]!;
    const c = Math.SQRT1_2;
    near(ends(path).at(-1)!, [500 * (1 + c), 250 * (1 + c)], 1e-9);
    expect(offEllipse(path, 500, 250, 500, 250)).toBeLessThan(3e-4);
  });
});

describe('paths', () => {
  test('every subpath starts with M, also after Z; the pen starts at the origin', () => {
    const g = custGeom(`<a:pathLst><a:path>
      <a:lnTo><a:pt x="10" y="0"/></a:lnTo><a:close/><a:close/>
      <a:lnTo><a:pt x="20" y="5"/></a:lnTo>
      <a:quadBezTo><a:pt x="30" y="0"/><a:pt x="40" y="5"/></a:quadBezTo>
      <a:cubicBezTo><a:pt x="50" y="0"/><a:pt x="60" y="10"/><a:pt x="70" y="5"/></a:cubicBezTo>
    </a:path></a:pathLst>`);
    expect(evaluateGeometry(g, 100, 100).paths[0]!.commands).toEqual([
      { op: 'M', x: 0, y: 0 }, { op: 'L', x: 10, y: 0 }, { op: 'Z' },
      { op: 'M', x: 0, y: 0 }, { op: 'L', x: 20, y: 5 },
      { op: 'Q', x1: 30, y1: 0, x: 40, y: 5 },
      { op: 'C', x1: 50, y1: 0, x2: 60, y2: 10, x: 70, y: 5 },
    ]);
  });

  test('path attributes come through, and its own space scales to the shape', () => {
    const g = custGeom(`<a:pathLst><a:path w="10" h="20" fill="darkenLess" stroke="0" extrusionOk="false">
      <a:moveTo><a:pt x="5" y="10"/></a:moveTo><a:lnTo><a:pt x="10" y="20"/></a:lnTo></a:path></a:pathLst>`);
    expect(evaluateGeometry(g, 1000, 500).paths).toEqual([{
      commands: [{ op: 'M', x: 500, y: 250 }, { op: 'L', x: 1000, y: 500 }], fill: 'darkenLess', stroke: false, extrusionOk: false,
    }]);
  });
});

describe('pathBounds', () => {
  const box = (commands: PathCommand[]) => pathBounds([{ commands, fill: 'norm', stroke: true, extrusionOk: true }]);

  test('curves count by their extremes, not their control points', () => {
    expect(box([{ op: 'M', x: 0, y: 0 }, { op: 'C', x1: 0, y1: 100, x2: 100, y2: 100, x: 100, y: 0 }])).toEqual({ x: 0, y: 0, w: 100, h: 75 });
    expect(box([{ op: 'M', x: 0, y: 0 }, { op: 'Q', x1: 50, y1: 100, x: 100, y: 0 }])).toEqual({ x: 0, y: 0, w: 100, h: 50 });
  });

  test('a Z returns to the start of the subpath', () => {
    expect(box([{ op: 'M', x: 10, y: 10 }, { op: 'L', x: 20, y: 10 }, { op: 'Z' }, { op: 'Q', x1: 10, y1: 30, x: 10, y: 10 }]))
      .toEqual({ x: 10, y: 10, w: 10, h: 10 });
  });

  test('presets: the ellipse fills its box; a callout tail reaches out of it', () => {
    const b = pathBounds(preset('ellipse').paths);
    expect(fixedRect(b)).toEqual([0, 0, W, H]);
    const tail = pathBounds(preset('wedgeRectCallout').paths);   // its tail points below the box
    expect(tail.y + tail.h).toBeCloseTo(250 + (62500 / 100000) * H, 6);
  });

  test('nothing drawn: an empty box', () => {
    expect(pathBounds([])).toEqual({ x: 0, y: 0, w: 0, h: 0 });
  });
});
