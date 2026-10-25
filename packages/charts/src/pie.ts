// packages/charts/src/pie.ts  -- pie, doughnut and pie-of-pie charts; 3-D pies tilted with a side wall
//
// Slices run clockwise from c:firstSliceAng (degrees from 12 o'clock) in point order, sized by the absolute values
// of the first series (a doughnut draws every series as a ring, the first innermost). c:explosion moves a slice out
// along its middle by that percentage of the radius. A pie of pie (ofPieChart) is drawn as one pie of all points.
import type { Paint } from '@pptx/resolve';
import type { Primitive } from '@pptx/render';
import type { Chart, Group } from './model.ts';
import { pathPrim, polyline, wedge, type Cmd, type Rect } from './draw.ts';
import type { LabelSite } from './series.ts';
import { filledStyle, type Env } from './style.ts';
import { flag, valOf } from './xml.ts';

const TAU = Math.PI * 2;
type Curve = Extract<Cmd, { op: 'C' }>;

/**
 * The side of a 3-D slice: its colour darkened, more so to the right, as if lit from the upper left. `from` and `to`
 * are where the piece starts and ends across the pie (0 left edge, 1 right edge), so the shading runs on unbroken.
 */
function side(p: Paint, from: number, to: number): Paint {
  // a gradient's side takes its last stop, the colour nearest the rim
  const gradient = p.kind === 'linear' || p.kind === 'path';
  const base = p.kind === 'solid' ? p.color : gradient ? p.stops.at(-1)?.color : undefined;
  if (!base) return p;
  const d = (f: number) => {
    const k = 0.82 - 0.32 * Math.max(0, Math.min(1, f));
    return { r: base.r * k, g: base.g * k, b: base.b * k, a: base.a };
  };
  return {
    kind: 'linear', angle: 0, scaled: false, rotWithShape: true, tileRect: { l: 0, t: 0, r: 0, b: 0 }, flip: 'none',
    stops: [{ pos: 0, color: d(from) }, { pos: 1, color: d(to) }],
  };
}

/** The part of [a0, a1] (radians, clockwise from 3 o'clock) on the front half of a tilted pie, where its side shows. */
function frontParts(a0: number, a1: number): [number, number][] {
  const out: [number, number][] = [];
  for (let k = -1; k <= 1; k++) {
    const lo = Math.max(a0, k * TAU);
    const hi = Math.min(a1, k * TAU + Math.PI);
    if (hi > lo + 1e-6) out.push([lo, hi]);
  }
  return out;
}

/** The side of a tilted pie between angles b0 and b1: the rim's arc, down by the depth, and back along the bottom. */
function sideWall(cx: number, cy: number, r: number, ry: number, depth: number, b0: number, b1: number): Cmd[] {
  const arc = (y: number) => wedge(cx, y, r, b0, b1, 0, ry).filter((c): c is Curve => c.op === 'C');
  const bottom = arc(cy + depth);
  const out: Cmd[] = [{ op: 'M', x: cx + r * Math.cos(b0), y: cy + ry * Math.sin(b0) }, ...arc(cy)];
  out.push({ op: 'L', x: cx + r * Math.cos(b1), y: cy + depth + ry * Math.sin(b1) });
  // back along the bottom edge: the same arc reversed, each curve ending where the one before it began
  for (let i = bottom.length - 1; i >= 0; i--) {
    const c = bottom[i]!;
    const end = bottom[i - 1] ?? { x: cx + r * Math.cos(b0), y: cy + depth + ry * Math.sin(b0) };
    out.push({ op: 'C', x1: c.x2, y1: c.y2, x2: c.x1, y2: c.y1, x: end.x, y: end.y });
  }
  out.push({ op: 'Z' });
  return out;
}

const SHOWN = ['c:showVal', 'c:showPercent', 'c:showCatName', 'c:showSerName'];

/** Whether a group's labels sit outside the slices, which makes Office draw the pie smaller. */
export function labelsOutside(g: Group): boolean {
  return [g.dLbls, ...g.series.map(s => s.dLbls)].some(el => {
    if (!el || !SHOWN.some(name => flag(el, name, false))) return false;
    const p = valOf(el, 'c:dLblPos');
    return p === undefined || p === 'outEnd' || p === 'bestFit';
  });
}

/** Draw a pie or doughnut group in the plot area, collecting a label site per slice. */
export function drawPie(env: Env, chart: Chart, g: Group, plot: Rect, out: Primitive[], sites: LabelSite[]): void {
  const doughnut = g.kind === 'doughnut';
  const series = doughnut ? g.series : g.series.slice(0, 1);
  if (!series.length) return;
  const tilt = g.threeD ? Math.max(10, Math.min(90, chart.rotX || 30)) : 90;
  // room for the farthest exploded slice
  const explosions = series.flatMap(s => [s.explosion, ...[...s.points.values()].map(p => p.explosion ?? 0)]);
  const maxExplosion = Math.max(0, ...explosions);
  const shrink = (labelsOutside(g) ? 0.82 : 1) / (1 + Math.min(400, maxExplosion) / 100);
  const cx = plot.x + plot.w / 2;
  const squash = Math.sin((tilt * Math.PI) / 180);
  // a tilted pie is 2r wide and 2r*squash tall, plus its side, which is about a third of the radius at 30 degrees
  const thick = g.threeD ? 0.33 * Math.cos((tilt * Math.PI) / 180) : 0;
  const r = Math.max(0, Math.min(plot.w / 2, plot.h / (2 * squash + thick)) * shrink);
  const depth = r * thick;
  const cy = plot.y + plot.h / 2 - depth / 2;
  const hole = doughnut ? (r * g.holeSize) / 100 : 0;
  const ring = (r - hole) / series.length;
  const start = ((g.firstSliceAng - 90) * Math.PI) / 180;

  const tops: Primitive[] = [];
  const walls: { prim: Primitive; dist: number }[] = [];
  const faces: { prim: Primitive; dist: number }[] = [];
  series.forEach((s, j) => {
    const vals = (s.val?.pts ?? []).map(v => Math.abs(v ?? 0));
    const sum = vals.reduce((a, b) => a + b, 0);
    if (!(sum > 0)) return;
    const n = vals.length;
    const rOut = doughnut ? hole + ring * (j + 1) : r;
    const rIn = doughnut ? hole + ring * j : 0;
    let a = start;
    for (let i = 0; i < n; i++) {
      const v = vals[i]!;
      if (!v) continue;
      const sweep = (TAU * v) / sum;
      const a0 = a;
      const a1 = a + sweep;
      a = a1;
      const mid = (a0 + a1) / 2;
      // only the outer ring of a doughnut explodes
      const ex = doughnut && j < series.length - 1 ? 0 : ((s.points.get(i)?.explosion ?? s.explosion) / 100) * r;
      const sx = cx + ex * Math.cos(mid);
      const sy = cy + ex * Math.sin(mid) * squash;
      // varied colours run by point, the same in every ring of a doughnut
      const [k, max] = g.varyColors ? [i, n - 1] : [s.index, env.maxIndex];
      const style = filledStyle(env, k, max, g.threeD, [s.spPr, s.points.get(i)?.spPr]);
      const slice = wedge(sx, sy, rOut, a0, a1, rIn, rOut * squash, rIn * squash);
      const top = pathPrim(slice, style.fill, style.line, style.effects);
      if (top) tops.push(top);
      if (depth > 0 && !doughnut && sweep < TAU - 1e-6) {
        // cut faces show where they turn towards the viewer: a slice's first edge on the left, its last on the right
        for (const [edge, visible] of [[a0, Math.cos(a0) < -1e-6], [a1, Math.cos(a1) > 1e-6]] as const) {
          if (!visible) continue;
          const rim = { x: sx + rOut * Math.cos(edge), y: sy + rOut * squash * Math.sin(edge) };
          const cut = polyline([{ x: sx, y: sy }, rim, { x: rim.x, y: rim.y + depth }, { x: sx, y: sy + depth }], true);
          const face = pathPrim(cut, side(style.fill, 0.5, 0.5), style.line);
          if (face) faces.push({ prim: face, dist: -Math.sin(edge) });
        }
      }
      if (depth > 0) {
        // normalise the slice into [0, 2pi) before finding its front part
        const shift = Math.floor(a0 / TAU) * TAU;
        const across = (b: number) => (1 + Math.cos(b)) / 2;
        for (const [b0, b1] of frontParts(a0 - shift, a1 - shift)) {
          const sideCmds = sideWall(sx, sy, rOut, rOut * squash, depth, b0, b1);
          const wall = pathPrim(sideCmds, side(style.fill, across(b1), across(b0)), style.line);
          if (wall) walls.push({ prim: wall, dist: Math.abs((b0 + b1) / 2 - Math.PI / 2) });
        }
      }
      const rMid = (rIn + rOut) / 2;
      sites.push({
        g, s, i, value: s.val?.pts[i] ?? 0, percent: v / sum,
        x: sx + rMid * Math.cos(mid), y: sy + rMid * squash * Math.sin(mid),
        slice: { cx: sx, cy: sy, r: rOut, ry: rOut * squash, inner: rIn, mid }, style,
      });
    }
  });
  // the cut faces and the sides farthest back first, then every top
  faces.sort((p, q) => q.dist - p.dist);
  walls.sort((p, q) => q.dist - p.dist);
  out.push(...faces.map(f => f.prim), ...walls.map(w => w.prim), ...tops);
}
