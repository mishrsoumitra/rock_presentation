// packages/geometry/src/bounds.ts
import type { GeometryPath, Rect } from './types.ts';

/** Parameters in (0, 1) where a quadratic a t^2 + b t + c is zero: where a Bezier coordinate turns around. */
function roots(a: number, b: number, c: number): number[] {
  const found: number[] = [];
  if (Math.abs(a) < 1e-12 * (Math.abs(b) + Math.abs(c) + 1)) {
    if (b !== 0) found.push(-c / b);
  } else {
    const disc = b * b - 4 * a * c;
    if (disc >= 0) {
      const s = Math.sqrt(disc);
      found.push((-b + s) / (2 * a), (-b - s) / (2 * a));
    }
  }
  return found.filter(t => t > 0 && t < 1);
}

/** Where a cubic coordinate p0 p1 p2 p3 has its extremes inside the segment. */
function cubicExtremes(p0: number, p1: number, p2: number, p3: number): number[] {
  // B'(t) / 3 = (1-t)^2 d0 + 2 (1-t) t d1 + t^2 d2
  const d0 = p1 - p0, d1 = p2 - p1, d2 = p3 - p2;
  return roots(d0 - 2 * d1 + d2, 2 * (d1 - d0), d0).map(t => {
    const u = 1 - t;
    return u * u * u * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t * p3;
  });
}

function quadExtremes(p0: number, p1: number, p2: number): number[] {
  const den = p0 - 2 * p1 + p2;
  if (den === 0) return [];
  const t = (p0 - p1) / den;
  return t > 0 && t < 1 ? [(1 - t) * (1 - t) * p0 + 2 * (1 - t) * t * p1 + t * t * p2] : [];
}

/**
 * The box around evaluated paths: the end points and the extremes of every curve, so it is tight, not just the
 * hull of the control points. Paths that are neither filled nor stroked count too; leave them out to skip them.
 */
export function pathBounds(paths: readonly GeometryPath[]): Rect {
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  const xs = (x: number) => { if (x < minX) minX = x; if (x > maxX) maxX = x; };
  const ys = (y: number) => { if (y < minY) minY = y; if (y > maxY) maxY = y; };
  for (const path of paths) {
    let x = 0, y = 0, x0 = 0, y0 = 0;
    for (const c of path.commands) {
      if (c.op === 'Z') {
        x = x0; y = y0;
        continue;
      }
      if (c.op === 'M') {
        x0 = c.x; y0 = c.y;
      } else if (c.op === 'Q') {
        quadExtremes(x, c.x1, c.x).forEach(xs);
        quadExtremes(y, c.y1, c.y).forEach(ys);
      } else if (c.op === 'C') {
        cubicExtremes(x, c.x1, c.x2, c.x).forEach(xs);
        cubicExtremes(y, c.y1, c.y2, c.y).forEach(ys);
      }
      x = c.x; y = c.y;
      xs(x); ys(y);
    }
  }
  if (minX > maxX) return { x: 0, y: 0, w: 0, h: 0 };
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}
