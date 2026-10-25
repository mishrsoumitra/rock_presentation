// packages/geometry/src/arc.ts  -- a:arcTo (ECMA-376 Part 1, 20.1.9.3) as cubic Bezier curves
import type { PathCommand } from './types.ts';

const RAD = Math.PI / 10800000;     // radians per 60000th of a degree
const TURN = 2 * Math.PI;

const wrapToPi = (a: number) => a - TURN * Math.round(a / TURN);
/** NaN and the infinities as 0, so that no non-finite number leaves this package. */
export const finite = (v: number) => (Number.isFinite(v) ? v : 0);

/**
 * The ellipse's parametric angle t, where the point is (wR cos t, hR sin t) from the centre, for the geometric angle
 * `ang` (60000ths of a degree, the direction of the ray from the centre, clockwise since y points down).
 * atan2 alone answers in (-pi, pi], so a sweep of 360 degrees would end where it started and draw nothing. The angle
 * is unwrapped against `ang` instead: t and the geometric angle never differ by more than a quarter turn, so
 * ang + 360 degrees gives t + 2 pi, and a negative or multi-turn sweep keeps its direction and length.
 */
export function parametricAngle(ang: number, wR: number, hR: number): number {
  const a = ang * RAD;
  // On an axis the two angles agree exactly. That also keeps a flat ellipse (wR or hR 0) right, where atan2 cannot.
  if (ang % 5400000 === 0) return a;
  return a + wrapToPi(Math.atan2(wR * Math.sin(a), hR * Math.cos(a)) - a);
}

/**
 * Append the arc that starts at the pen (x0, y0) as cubic Beziers, at most a quarter turn each. The arc is
 * computed in the path's own coordinate space, where its angles are measured; only the resulting points are scaled
 * by (sx, sy) into the shape. Returns the end of the arc, which becomes the pen, in path units.
 */
export function appendArc(
  out: PathCommand[], x0: number, y0: number, wR: number, hR: number, stAng: number, swAng: number, sx: number, sy: number,
): [number, number] {
  const t0 = parametricAngle(stAng, wR, hR);
  let sweep = parametricAngle(stAng + swAng, wR, hR) - t0;
  // Past two turns an arc only retraces itself. Bound the work an untrusted file can ask for, keeping the end point.
  if (Math.abs(sweep) > 2 * TURN) sweep = Math.sign(sweep) * (TURN + (Math.abs(sweep) % TURN));
  const n = Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9);
  if (!(n > 0)) return [x0, y0];
  const cx = x0 - wR * Math.cos(t0);
  const cy = y0 - hR * Math.sin(t0);
  const delta = sweep / n;
  const k = (4 / 3) * Math.tan(delta / 4);
  let x = x0, y = y0, cos1 = Math.cos(t0), sin1 = Math.sin(t0);
  for (let i = 1; i <= n; i++) {
    const t = t0 + delta * i;
    const cos2 = Math.cos(t), sin2 = Math.sin(t);
    const x3 = cx + wR * cos2, y3 = cy + hR * sin2;
    out.push({
      op: 'C',
      x1: finite((x - k * wR * sin1) * sx), y1: finite((y + k * hR * cos1) * sy),
      x2: finite((x3 + k * wR * sin2) * sx), y2: finite((y3 - k * hR * cos2) * sy),
      x: finite(x3 * sx), y: finite(y3 * sy),
    });
    x = x3; y = y3; cos1 = cos2; sin1 = sin2;
  }
  return [x, y];
}
