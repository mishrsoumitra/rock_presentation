// packages/charts/src/draw.ts  -- path primitives for chart pieces: rectangles, polylines, wedges, markers
import type { Paint, ResolvedEffects, Stroke } from '@pptx/resolve';
import type { PathPrim, PathSpec } from '@pptx/render';

export type Cmd = PathSpec['commands'][number];
export interface Pt { x: number; y: number; }
export interface Rect { x: number; y: number; w: number; h: number; }

const fin = (v: number) => Number.isFinite(v);

/** Whether every coordinate of a command is a finite number. */
function finite(c: Cmd): boolean {
  if (c.op === 'Z') return true;
  if (!fin(c.x) || !fin(c.y)) return false;
  if (c.op === 'C') return fin(c.x1) && fin(c.y1) && fin(c.x2) && fin(c.y2);
  return c.op !== 'Q' || (fin(c.x1) && fin(c.y1));
}

/**
 * A path primitive in chart space. Coordinates are moved into the path's own bounding box, which is the box a
 * gradient or picture fill stretches over, as it does for a shape. Non-finite points are dropped; nothing left
 * gives undefined.
 */
export function pathPrim(
  cmds: Cmd[], fill: Paint, stroke: Stroke | null, effects: ResolvedEffects = {},
): PathPrim | undefined {
  const ok = cmds.filter(finite);
  if (!ok.some(c => c.op !== 'Z')) return undefined;
  if (ok[0]?.op !== 'M') {
    const first = ok.find(c => c.op !== 'Z') as Exclude<Cmd, { op: 'Z' }>;
    ok.unshift({ op: 'M', x: first.x, y: first.y });
  }
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  const see = (x: number, y: number) => {
    if (x < x0) x0 = x;
    if (x > x1) x1 = x;
    if (y < y0) y0 = y;
    if (y > y1) y1 = y;
  };
  for (const c of ok) {
    if (c.op === 'Z') continue;
    see(c.x, c.y);
    if (c.op === 'C') { see(c.x1, c.y1); see(c.x2, c.y2); }
    if (c.op === 'Q') see(c.x1, c.y1);
  }
  const shifted: Cmd[] = ok.map(c => {
    switch (c.op) {
      case 'Z': return c;
      case 'C': return { ...c, x1: c.x1 - x0, y1: c.y1 - y0, x2: c.x2 - x0, y2: c.y2 - y0, x: c.x - x0, y: c.y - y0 };
      case 'Q': return { ...c, x1: c.x1 - x0, y1: c.y1 - y0, x: c.x - x0, y: c.y - y0 };
      default: return { op: c.op, x: c.x - x0, y: c.y - y0 };
    }
  });
  const filled = fill.kind !== 'none';
  if (!filled && !stroke) return undefined;
  return {
    kind: 'path', transform: [1, 0, 0, 1, x0, y0], box: { w: Math.max(1, x1 - x0), h: Math.max(1, y1 - y0) },
    paths: [{ commands: shifted, fill: filled ? 'norm' : 'none', stroke: !!stroke }], fill, stroke, effects,
  };
}

/** A path in a fill, line and effects, such as a point's or a box's style. */
export const styledPath = (cmds: Cmd[], st: { fill: Paint; line: Stroke | null; effects?: ResolvedEffects }) =>
  pathPrim(cmds, st.fill, st.line, st.effects);

export function rect(r: Rect): Cmd[] {
  const x = Math.min(r.x, r.x + r.w);
  const y = Math.min(r.y, r.y + r.h);
  const w = Math.abs(r.w);
  const h = Math.abs(r.h);
  return polyline([{ x, y }, { x: x + w, y }, { x: x + w, y: y + h }, { x, y: y + h }], true);
}

/** A rectangle with rounded corners (c:roundedCorners). */
export function roundRect(r: Rect, radius: number): Cmd[] {
  const k = Math.min(radius, Math.abs(r.w) / 2, Math.abs(r.h) / 2);
  if (k <= 0) return rect(r);
  const { x, y, w, h } = r;
  const out: Cmd[] = [{ op: 'M', x: x + k, y }];
  const H = Math.PI / 2;
  // each side, then a quarter circle round the next corner
  out.push({ op: 'L', x: x + w - k, y });
  arcTo(out, x + w - k, y + k, k, k, -H, 0);
  out.push({ op: 'L', x: x + w, y: y + h - k });
  arcTo(out, x + w - k, y + h - k, k, k, 0, H);
  out.push({ op: 'L', x: x + k, y: y + h });
  arcTo(out, x + k, y + h - k, k, k, H, 2 * H);
  out.push({ op: 'L', x, y: y + k });
  arcTo(out, x + k, y + k, k, k, 2 * H, 3 * H);
  out.push({ op: 'Z' });
  return out;
}

export function polyline(pts: Pt[], close = false): Cmd[] {
  const out: Cmd[] = pts.map((p, i) => ({ op: i === 0 ? 'M' : 'L', x: p.x, y: p.y }));
  if (close && out.length) out.push({ op: 'Z' });
  return out;
}

export const line = (a: Pt, b: Pt): Cmd[] => [{ op: 'M', x: a.x, y: a.y }, { op: 'L', x: b.x, y: b.y }];

/**
 * A smooth curve through the points (c:smooth): Catmull-Rom segments as cubic Beziers. Office's smoothing is a
 * similar interpolating spline; the curve passes through every point.
 */
export function smooth(pts: Pt[]): Cmd[] {
  if (pts.length < 3) return polyline(pts);
  const out: Cmd[] = [{ op: 'M', x: pts[0]!.x, y: pts[0]!.y }];
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[Math.max(0, i - 1)]!;
    const p1 = pts[i]!;
    const p2 = pts[i + 1]!;
    const p3 = pts[Math.min(pts.length - 1, i + 2)]!;
    out.push({
      op: 'C', x1: p1.x + (p2.x - p0.x) / 6, y1: p1.y + (p2.y - p0.y) / 6,
      x2: p2.x - (p3.x - p1.x) / 6, y2: p2.y - (p3.y - p1.y) / 6, x: p2.x, y: p2.y,
    });
  }
  return out;
}

/** Arc points of a circle from angle a0 to a1 (radians, clockwise from 3 o'clock since y points down) as Beziers. */
function arcTo(out: Cmd[], cx: number, cy: number, rx: number, ry: number, a0: number, a1: number): void {
  const n = Math.max(1, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 2) - 1e-9));
  const d = (a1 - a0) / n;
  const k = (4 / 3) * Math.tan(d / 4);
  for (let i = 0; i < n; i++) {
    const s = a0 + d * i;
    const e = s + d;
    out.push({
      op: 'C',
      x1: cx + rx * (Math.cos(s) - k * Math.sin(s)), y1: cy + ry * (Math.sin(s) + k * Math.cos(s)),
      x2: cx + rx * (Math.cos(e) + k * Math.sin(e)), y2: cy + ry * (Math.sin(e) - k * Math.cos(e)),
      x: cx + rx * Math.cos(e), y: cy + ry * Math.sin(e),
    });
  }
}

export function ellipse(cx: number, cy: number, rx: number, ry: number): Cmd[] {
  const out: Cmd[] = [{ op: 'M', x: cx + rx, y: cy }];
  arcTo(out, cx, cy, rx, ry, 0, Math.PI * 2);
  out.push({ op: 'Z' });
  return out;
}

/**
 * A pie slice or doughnut segment between angles a0 < a1 (radians, clockwise from 3 o'clock). `ry` and `innerY`
 * flatten it into an ellipse for tilted (3-D) pies.
 */
export function wedge(
  cx: number, cy: number, r: number, a0: number, a1: number, inner = 0, ry = r, innerY = inner,
): Cmd[] {
  const full = a1 - a0 >= Math.PI * 2 - 1e-9;
  const out: Cmd[] = [];
  if (full) {
    out.push(...ellipse(cx, cy, r, ry));
    if (inner > 0) {
      // the hole, wound the other way so it stays empty under the non-zero rule
      out.push({ op: 'M', x: cx + inner, y: cy });
      arcTo(out, cx, cy, inner, innerY, Math.PI * 2, 0);
      out.push({ op: 'Z' });
    }
    return out;
  }
  if (inner > 0) {
    out.push({ op: 'M', x: cx + inner * Math.cos(a0), y: cy + innerY * Math.sin(a0) });
    out.push({ op: 'L', x: cx + r * Math.cos(a0), y: cy + ry * Math.sin(a0) });
    arcTo(out, cx, cy, r, ry, a0, a1);
    out.push({ op: 'L', x: cx + inner * Math.cos(a1), y: cy + innerY * Math.sin(a1) });
    arcTo(out, cx, cy, inner, innerY, a1, a0);
  } else {
    out.push({ op: 'M', x: cx, y: cy });
    out.push({ op: 'L', x: cx + r * Math.cos(a0), y: cy + ry * Math.sin(a0) });
    arcTo(out, cx, cy, r, ry, a0, a1);
  }
  out.push({ op: 'Z' });
  return out;
}

/** The symbols Office cycles through for series whose markers are automatic. */
export const AUTO_SYMBOLS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dot', 'dash'];

/** A marker symbol of size s (EMU) centred at (x, y); `open` symbols are lines only and take no fill. */
export function marker(symbol: string, x: number, y: number, s: number): { cmds: Cmd[]; open: boolean } {
  const h = s / 2;
  const shut = (cmds: Cmd[]) => ({ cmds, open: false });
  const seg = ([dx0, dy0, dx1, dy1]: number[]) => line({ x: x + dx0!, y: y + dy0! }, { x: x + dx1!, y: y + dy1! });
  const lines = (...segs: number[][]) => ({ cmds: segs.flatMap(seg), open: true });
  switch (symbol) {
    case 'circle': return shut(ellipse(x, y, h, h));
    case 'square': return shut(rect({ x: x - h, y: y - h, w: s, h: s }));
    case 'triangle': return shut(polyline([{ x, y: y - h }, { x: x + h, y: y + h }, { x: x - h, y: y + h }], true));
    case 'dash': return shut(rect({ x: x - h, y: y - s / 10, w: s, h: s / 5 }));
    case 'dot': return shut(ellipse(x, y, s / 5, s / 5));
    case 'x': return lines([-h, -h, h, h], [-h, h, h, -h]);
    case 'plus': return lines([-h, 0, h, 0], [0, -h, 0, h]);
    case 'star': return lines([-h, -h, h, h], [-h, h, h, -h], [0, -h, 0, h]);
    default: return shut(polyline([{ x, y: y - h }, { x: x + h, y }, { x, y: y + h }, { x: x - h, y }], true));
  }
}
