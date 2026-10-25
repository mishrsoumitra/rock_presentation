// packages/geometry/src/handles.ts  -- dragging an adjust handle: from where the handle should be to adjust values
import { finite } from './arc.ts';
import { GeometryError } from './errors.ts';
import { overridesOf, program, run, type Axis, type Program } from './evaluate.ts';
import type { Guide, ShapeGeometry } from './types.ts';

const SAMPLES = 32;
const GOLDEN = (Math.sqrt(5) - 1) / 2;

const samples = (lo: number, hi: number) => Array.from({ length: SAMPLES + 1 }, (_, i) => (i === SAMPLES ? hi : lo + ((hi - lo) * i) / SAMPLES));

/** Where in [lo, hi] g is least: a golden-section search, never returning anything worse than `start`. */
function minimize(g: (v: number) => number, lo: number, hi: number, start: number): number {
  let a = lo, b = hi;
  let c = b - GOLDEN * (b - a), d = a + GOLDEN * (b - a);
  let gc = g(c), gd = g(d);
  for (let i = 0; i < 100 && b - a > 1e-6 * (Math.abs(a) + 1); i++) {
    if (gc < gd) { b = d; d = c; gd = gc; c = b - GOLDEN * (b - a); gc = g(c); }
    else { a = c; c = d; gc = gd; d = a + GOLDEN * (b - a); gd = g(d); }
  }
  const m = (a + b) / 2;
  return g(m) <= g(start) ? m : start;
}

/** The least g over [lo, hi]: the best of the samples, refined between its neighbours. */
function nearest(g: (v: number) => number, lo: number, hi: number): number {
  const vs = samples(lo, hi);
  const gs = vs.map(g);
  let best = 0;
  for (let i = 1; i <= SAMPLES; i++) if (gs[i]! < gs[best]!) best = i;
  return minimize(g, vs[Math.max(0, best - 1)]!, vs[Math.min(SAMPLES, best + 1)]!, vs[best]!);
}

/**
 * Where in [lo, hi] f is zero: the sample interval where f changes sign (the one nearest `current` if there are
 * several), narrowed by bisection. When f has no zero in range, where |f| is least: the handle as near as it gets.
 */
function solveZero(f: (v: number) => number, lo: number, hi: number, current: number): number {
  const vs = samples(lo, hi);
  const fs = vs.map(f);
  const gap = (i: number) => Math.max(0, vs[i]! - current, current - vs[i + 1]!);
  let pick = -1;
  for (let i = 0; i < SAMPLES; i++) {
    const fa = fs[i]!, fb = fs[i + 1]!;
    if (((fa <= 0 && fb >= 0) || (fa >= 0 && fb <= 0)) && (pick < 0 || gap(i) < gap(pick))) pick = i;
  }
  if (pick < 0) return nearest(v => Math.abs(f(v)), lo, hi);
  let a = vs[pick]!, b = vs[pick + 1]!, fa = fs[pick]!;
  if (fa === 0) return a;
  if (fs[pick + 1] === 0) return b;
  for (let i = 0; i < 100 && b - a > 1e-6 * (Math.abs(a) + 1); i++) {
    const m = (a + b) / 2;
    const fm = f(m);
    if (fm === 0) return m;
    if ((fm < 0) === (fa < 0)) { a = m; fa = fm; } else b = m;
  }
  return (a + b) / 2;
}

type Handle = Program['handles'][number];

/** Solve one axis of a handle with the other adjust values held, and store the result, rounded, in `values`. */
function solveAxis(p: Program, handle: Handle, axis: Axis, w: number, h: number, values: Map<string, string | number>, point: { x: number; y: number }): void {
  const v = run(p, w, h, values);
  let lo = v[axis.min]!, hi = v[axis.max]!;
  if (lo > hi) [lo, hi] = [hi, lo];
  const current = v[p.adjustSlots.get(axis.ref)!]!;
  const evaluate = (value: number) => {
    values.set(axis.ref, value);
    return run(p, w, h, values);
  };
  let value = lo;
  if (lo < hi && axis.along === 'polar') {
    value = nearest(x => {
      const r = evaluate(x);
      return Math.hypot(r[handle.x]! - point.x, r[handle.y]! - point.y);
    }, lo, hi);
  } else if (lo < hi) {
    const slot = axis.along === 'x' ? handle.x : handle.y;
    const goal = axis.along === 'x' ? point.x : point.y;
    value = solveZero(x => evaluate(x)[slot]! - goal, lo, hi, current);
  }
  // PowerPoint stores adjust values as integers. Round, and stay within the limits.
  let rounded = Math.round(value);
  if (Math.ceil(lo) <= Math.floor(hi)) rounded = Math.min(Math.floor(hi), Math.max(Math.ceil(lo), rounded));
  values.set(axis.ref, finite(rounded));
}

/**
 * The adjust values that bring handle `handleIndex` (an index into geom.ahLst) to `point`, in the shape's own space
 * in EMU, or as near as the handle's limits allow. Returns `adjust` with the values the handle moves set to 'val N',
 * in avLst order; the shape's other adjust values are kept as they were.
 *
 * A handle's position comes out of the guide formulas, so the adjust value is solved for numerically rather than by
 * inverting formulas, one axis at a time with the others held: an ahXY axis by bisection over [min, max] until the
 * handle's x (or y) is the point's; an ahPolar axis, angle or radius, by bringing the handle as close to the point
 * as it goes, since a polar handle's path has no centre the geometry states.
 */
export function adjustFromHandle(
  geom: ShapeGeometry, w: number, h: number, adjust: readonly Guide[], handleIndex: number, point: { x: number; y: number },
): Guide[] {
  const p = program(geom);
  const handle = p.handles[handleIndex];
  if (!handle) throw new GeometryError(`no adjust handle ${handleIndex}`);
  // Only adjust values can move, and an axis without both limits cannot (sections 20.1.9.2 and 20.1.9.3).
  const axes = handle.axes.filter(a => p.adjustSlots.has(a.ref) && a.min >= 0 && a.max >= 0);
  if (axes.length === 0) return [...adjust];
  w = finite(w);
  h = finite(h);
  const target = { x: finite(point.x), y: finite(point.y) };
  const values = overridesOf(adjust);
  // Axes can depend on each other (a polar handle's angle and radius both move it), so go round until a round
  // changes nothing. The circular arrows couple their two values strongly and need more than a few rounds.
  for (let round = axes.length > 1 ? 16 : 1; round > 0; round--) {
    const before = axes.map(a => values.get(a.ref));
    for (const a of axes) solveAxis(p, handle, a, w, h, values, target);
    if (axes.every((a, i) => values.get(a.ref) === before[i])) break;
  }

  const given = overridesOf(adjust);
  const moved = new Set(axes.map(a => a.ref));
  const out: Guide[] = [];
  const seen = new Set<string>();
  for (const { name } of geom.avLst) {
    if (seen.has(name)) continue;
    seen.add(name);
    const fmla = moved.has(name) ? `val ${values.get(name)}` : given.get(name);
    if (fmla !== undefined) out.push({ name, fmla: String(fmla) });
  }
  for (const g of adjust) if (!seen.has(g.name)) out.push(g);
  return out;
}
