// packages/metafile/src/geom.ts  -- transforms, paths in output space, and elliptical arcs as Bezier curves
import type { Matrix, Rect } from './types.ts';

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** a after b: apply b first, then a. */
export function compose(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
  ];
}

export const det = (m: Matrix) => m[0] * m[3] - m[1] * m[2];

export function invert(m: Matrix): Matrix | undefined {
  const d = det(m);
  if (!d || !Number.isFinite(d)) return undefined;
  return [m[3] / d, -m[1] / d, -m[2] / d, m[0] / d, (m[2] * m[5] - m[3] * m[4]) / d, (m[1] * m[4] - m[0] * m[5]) / d];
}

export const apply = (m: Matrix, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

/** True when the matrix keeps axis-aligned rectangles axis-aligned. */
export const axisAligned = (m: Matrix) => (m[1] === 0 && m[2] === 0) || (m[0] === 0 && m[3] === 0);

/** A usable transform: finite and invertible. */
export const usable = (m: Matrix) => m.every(Number.isFinite) && det(m) !== 0 && Number.isFinite(1 / det(m));

export const normalize = (r: Rect): Rect => ({ l: Math.min(r.l, r.r), t: Math.min(r.t, r.b), r: Math.max(r.l, r.r), b: Math.max(r.t, r.b) });

const M = 0, L = 1, C = 2, Z = 3;

/**
 * A path in output coordinates: flat commands (M x y, L x y, C x1 y1 x2 y2 x y, Z), so that it can be kept for
 * clipping and replayed onto any context.
 */
export class Path {
  ops: number[] = [];
  /** Where the open subpath ends (and whether there is one): figures continue from here. */
  open = false;
  x = 0;
  y = 0;

  moveTo(x: number, y: number): void {
    this.ops.push(M, x, y);
    this.open = true;
    this.x = x; this.y = y;
  }

  lineTo(x: number, y: number): void {
    if (!this.open) this.moveTo(this.x, this.y);
    this.ops.push(L, x, y);
    this.x = x; this.y = y;
  }

  curveTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): void {
    if (!this.open) this.moveTo(this.x, this.y);
    this.ops.push(C, x1, y1, x2, y2, x, y);
    this.x = x; this.y = y;
  }

  close(): void {
    if (this.open) this.ops.push(Z);
    this.open = false;
  }

  /** A closed polygon from flat output coordinates. */
  polygon(xy: ArrayLike<number>, closed = true): void {
    if (xy.length < 2) return;
    this.moveTo(xy[0]!, xy[1]!);
    for (let i = 2; i + 1 < xy.length; i += 2) this.lineTo(xy[i]!, xy[i + 1]!);
    if (closed) this.close();
  }

  append(p: Path): void {
    for (const op of p.ops) this.ops.push(op);
    this.open = p.open;
    this.x = p.x; this.y = p.y;
  }

  get empty(): boolean { return this.ops.length === 0; }

  /** Put every vertex and curve end on whole units (device pixels), as GDI's integer device coordinates are. */
  snap(): void {
    const o = this.ops;
    for (let i = 0; i < o.length;) {
      if (o[i] === M || o[i] === L) { o[i + 1] = Math.round(o[i + 1]!); o[i + 2] = Math.round(o[i + 2]!); i += 3; }
      else if (o[i] === C) { o[i + 5] = Math.round(o[i + 5]!); o[i + 6] = Math.round(o[i + 6]!); i += 7; }
      else i += 1;
    }
    this.x = Math.round(this.x);
    this.y = Math.round(this.y);
  }

  /** Number of points, for the work budget. */
  get size(): number { return this.ops.length >> 1; }

  /** Replay onto a context as its current path. */
  trace(ctx: PathSink): void {
    const o = this.ops;
    for (let i = 0; i < o.length;) {
      switch (o[i]) {
        case M: ctx.moveTo(o[i + 1]!, o[i + 2]!); i += 3; break;
        case L: ctx.lineTo(o[i + 1]!, o[i + 2]!); i += 3; break;
        case C: ctx.bezierCurveTo(o[i + 1]!, o[i + 2]!, o[i + 3]!, o[i + 4]!, o[i + 5]!, o[i + 6]!); i += 7; break;
        default: ctx.closePath(); i += 1; break;
      }
    }
  }

  /** The same path through a transform. */
  map(m: Matrix): Path {
    const out = new Path();
    const o = this.ops;
    const xy = (i: number) => apply(m, o[i]!, o[i + 1]!);
    for (let i = 0; i < o.length;) {
      switch (o[i]) {
        case M: out.moveTo(...xy(i + 1)); i += 3; break;
        case L: out.lineTo(...xy(i + 1)); i += 3; break;
        case C: { const [a, b] = xy(i + 1); const [c, d] = xy(i + 3); const [e, f] = xy(i + 5); out.curveTo(a, b, c, d, e, f); i += 7; break; }
        default: out.close(); i += 1; break;
      }
    }
    out.open = this.open;
    return out;
  }

  /** Bounding box of the points (control points included). */
  bounds(): Rect | undefined {
    const o = this.ops;
    let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
    for (let i = 0; i < o.length;) {
      const n = o[i] === M || o[i] === L ? 1 : o[i] === C ? 3 : 0;
      for (let k = 0; k < n; k++) {
        const x = o[i + 1 + k * 2]!, y = o[i + 2 + k * 2]!;
        l = Math.min(l, x); t = Math.min(t, y); r = Math.max(r, x); b = Math.max(b, y);
      }
      i += 1 + n * 2;
    }
    return l <= r ? { l, t, r, b } : undefined;
  }
}

/** The path-building methods of a 2D context. */
export interface PathSink {
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  bezierCurveTo(cp1x: number, cp1y: number, cp2x: number, cp2y: number, x: number, y: number): void;
  closePath(): void;
}

/**
 * An elliptical arc as cubic Beziers, in the ellipse's own coordinates: centre (cx, cy), radii rx and ry, from
 * parameter angle t0 through `sweep` radians (positive: increasing angle, which is clockwise on a y-down screen).
 * Calls `to` with each curve's control points and end; the start point is (cx + rx cos t0, cy + ry sin t0).
 */
export function arcCurves(cx: number, cy: number, rx: number, ry: number, t0: number, sweep: number,
  to: (x1: number, y1: number, x2: number, y2: number, x: number, y: number) => void): void {
  const n = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  const d = sweep / n;
  const k = (4 / 3) * Math.tan(d / 4);
  let t = t0;
  for (let i = 0; i < n; i++) {
    const c0 = Math.cos(t), s0 = Math.sin(t), c1 = Math.cos(t + d), s1 = Math.sin(t + d);
    to(cx + rx * (c0 - k * s0), cy + ry * (s0 + k * c0), cx + rx * (c1 + k * s1), cy + ry * (s1 - k * c1), cx + rx * c1, cy + ry * s1);
    t += d;
  }
}
