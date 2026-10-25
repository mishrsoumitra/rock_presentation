// packages/charts/src/series.ts  -- data series on cartesian axes: bars, lines, areas, scatter, bubbles and stock
import type { Paint, Stroke } from '@pptx/resolve';
import type { Primitive } from '@pptx/render';
import type { Chart, Group, Series } from './model.ts';
import {
  at, baseValue, categoryCount, pointAt, PT, slotOf, slotSize, stack, valueAt, xValues, yValues, type AxisRt,
} from './axes.ts';
import {
  AUTO_SYMBOLS, ellipse, line as lineCmds, marker, pathPrim, polyline, rect, smooth, type Cmd, type Pt, type Rect,
} from './draw.ts';
import {
  fillOf, filledStyle, markerStyle, NONE, otherLineColor, seriesLine, solid, strokeOf, themedLine, upDownFill,
  type Env, type PointStyle,
} from './style.ts';

/** Where a data label goes, with what it says: collected while the series are drawn, drawn on top afterwards. */
export interface LabelSite {
  g: Group;
  s: Series;
  i: number;
  value?: number;
  percent?: number;
  size?: number;
  x: number;
  y: number;
  /** Bars: the bar's rectangle and which way it grows. */
  bar?: { r: Rect; dir: 'up' | 'down' | 'left' | 'right' };
  /** Pie slices: centre, radii and the slice's middle angle (radians, clockwise from 3 o'clock). */
  slice?: { cx: number; cy: number; r: number; ry: number; inner: number; mid: number };
  /** Markers and bubbles: the radius around the point that labels keep clear of. */
  radius?: number;
  style: PointStyle | { color: PointStyle['color']; fill: Paint; line: Stroke | null };
}

/** The formatting index and its range for a point: per point when a single-series group varies colours. */
export function colorIndex(env: Env, g: Group, s: Series, i: number, n: number): { k: number; max: number } {
  return g.varyColors && g.series.length === 1 ? { k: i, max: Math.max(0, n - 1) } : { k: s.index, max: env.maxIndex };
}

/** A value held within an axis's scale, so bars and areas past its ends stop at the plot's edge. */
const clampT = (a: AxisRt, v: number) => {
  const { min, max } = a.scale;
  return Math.max(Math.min(min, max), Math.min(Math.max(min, max), v));
};
const pointSpPr = (s: Series, i: number) => s.points.get(i)?.spPr;
const isStacked = (g: Group) => g.grouping === 'stacked' || g.grouping === 'percentStacked';

/** A plain one-pixel-ish (0.75 pt) line in a paint, for outlines the formatting leaves out. */
const hairline = (paint: Paint, join: Stroke['join']): Stroke =>
  ({ width: 9525, paint, cap: 'flat', join, miterLimit: 8, dash: null, compound: 'sng', align: 'ctr' });
const BG1 = { base: { kind: 'scheme' as const, name: 'bg1' }, transforms: [] };

// --------------------------------------------------------------------------------------------------------- bars

function drawBars(env: Env, g: Group, cat: AxisRt, val: AxisRt, out: Primitive[], sites: LabelSite[]): void {
  const n = categoryCount(g);
  const stacked = isStacked(g);
  const clusters = stacked ? 1 : Math.max(1, g.series.length);
  const ov = g.overlap / 100;
  const gw = g.gapWidth / 100;
  const slotT = 1 / Math.max(1, cat.count);
  const barT = slotT / (clusters - (clusters - 1) * ov + gw);
  const st = stacked ? stack(g, n) : undefined;
  const base = baseValue(cat, val);
  g.series.forEach((s, j) => {
    const vals = s.val?.pts ?? [];
    for (let i = 0; i < n; i++) {
      const v = vals[i];
      if (v === undefined) continue;
      const v0 = st ? st.base[j]![i]! : base;
      const v1 = st ? st.top[j]![i]! : v;
      const t0 = slotOf(cat, i) * slotT + (gw * barT) / 2 + (stacked ? 0 : j * barT * (1 - ov));
      const p0 = at(cat, t0);
      const p1 = at(cat, t0 + barT);
      const q0 = valueAt(val, clampT(val, v0));
      const q1 = valueAt(val, clampT(val, v1));
      const r: Rect = cat.dir === 'h'
        ? { x: Math.min(p0, p1), y: Math.min(q0, q1), w: Math.abs(p1 - p0), h: Math.abs(q1 - q0) }
        : { x: Math.min(q0, q1), y: Math.min(p0, p1), w: Math.abs(q1 - q0), h: Math.abs(p1 - p0) };
      const { k, max } = colorIndex(env, g, s, i, n);
      let style = filledStyle(env, k, max, g.threeD, [s.spPr, pointSpPr(s, i)]);
      const invert = s.points.get(i)?.invertIfNegative ?? s.invertIfNegative;
      if (invert && v < 0 && style.fill.kind === 'solid') {
        // inverted: a white bar, outlined in the series colour when it has no line of its own
        style = { ...style, fill: solid(env, BG1), line: style.line ?? hairline(style.fill, 'miter') };
      }
      const prim = pathPrim(rect(r), style.fill, style.line, style.effects);
      if (prim) out.push(prim);
      const grows = q1 >= q0 ? 1 : -1;
      const dir = cat.dir === 'h' ? (grows < 0 ? 'up' : 'down') : grows > 0 ? 'right' : 'left';
      const percent = st && g.grouping === 'percentStacked' ? { percent: v1 - v0 } : {};
      sites.push({ g, s, i, value: v, ...percent, x: r.x + r.w / 2, y: r.y + r.h / 2, bar: { r, dir }, style });
    }
  });
}

// -------------------------------------------------------------------------------------------- lines and areas

/** Runs of consecutive points a line connects: blanks break the line ("gap"), or are skipped ("span", "zero"). */
function runs(points: (Pt | undefined)[], blanks: Chart['dispBlanksAs']): Pt[][] {
  const out: Pt[][] = [];
  let cur: Pt[] = [];
  for (const p of points) {
    if (p) cur.push(p);
    else if (blanks === 'gap' && cur.length) {
      out.push(cur);
      cur = [];
    }
  }
  if (cur.length) out.push(cur);
  return out;
}

/** The symbol of point i's marker: its own, the series', or `dflt`; automatic symbols cycle by series. */
function symbolOf(s: Series, i: number, dflt: 'auto' | 'none'): string {
  const sym = s.points.get(i)?.marker?.symbol ?? s.marker?.symbol ?? dflt;
  // picture markers are drawn as squares
  return sym === 'auto' ? AUTO_SYMBOLS[s.index % AUTO_SYMBOLS.length]! : sym === 'picture' ? 'square' : sym;
}

/** A point with its index in the series. */
export interface Mark { p: Pt; i: number; }

/** Draw a series' markers; returns the largest marker's radius, which its data labels keep clear of. */
export function drawMarkers(
  env: Env, g: Group, s: Series, marks: Mark[], dflt: 'auto' | 'none', n: number, out: Primitive[],
): number {
  let radius = 0;
  for (const { p, i } of marks) {
    const sym = symbolOf(s, i, dflt);
    if (sym === 'none') continue;
    const own = s.points.get(i)?.marker;
    const size = (own?.size ?? s.marker?.size ?? 5) * PT;
    const { k, max } = colorIndex(env, g, s, i, n);
    const st = markerStyle(env, seriesLine(env, k, max, []).color, [s.marker?.spPr, own?.spPr]);
    const m = marker(sym, p.x, p.y, size);
    // open symbols (x, plus, star) are lines only: drawn in the marker's line, or its fill when it has none
    const prim = m.open
      ? pathPrim(m.cmds, NONE, st.line ?? hairline(st.fill, 'round'))
      : pathPrim(m.cmds, st.fill, st.line);
    if (prim) out.push(prim);
    radius = Math.max(radius, size / 2);
  }
  return radius;
}

function drawLines(
  env: Env, chart: Chart, g: Group, cat: AxisRt, val: AxisRt, out: Primitive[], sites: LabelSite[],
): void {
  const n = categoryCount(g);
  const st = isStacked(g) ? stack(g, n) : undefined;
  const blanks = chart.dispBlanksAs;
  g.series.forEach((s, j) => {
    const vals = s.val?.pts ?? [];
    const pts: (Pt | undefined)[] = [];
    const marks: Mark[] = [];
    for (let i = 0; i < n; i++) {
      // a stacked blank still has a running total, which the line follows unless blanks make gaps
      let v = st ? (vals[i] === undefined && blanks === 'gap' ? undefined : st.top[j]![i]) : vals[i];
      if (v === undefined && blanks === 'zero' && i < vals.length) v = st ? st.top[j]![i] : 0;
      if (v === undefined) {
        pts.push(undefined);
        continue;
      }
      const p = { x: pointAt(cat, i), y: valueAt(val, v) };
      pts.push(cat.dir === 'h' ? p : { x: p.y, y: p.x });
      if (vals[i] !== undefined) marks.push({ p: pts.at(-1)!, i });
    }
    // stock series have no automatic line (ECMA-376 21.2.3.46, Table 1): only one their spPr draws
    const line = g.kind === 'stock'
      ? strokeOf(env, s.spPr, null)
      : seriesLine(env, s.index, env.maxIndex, [s.spPr]).line;
    if (line) {
      const cmds: Cmd[] = runs(pts, chart.dispBlanksAs).flatMap(r => (s.smooth ? smooth(r) : polyline(r)));
      const prim = pathPrim(cmds, NONE, line);
      if (prim) out.push(prim);
    }
    const dflt = g.kind === 'stock' ? 'none' : g.showMarker ? 'auto' : 'none';
    const radius = drawMarkers(env, g, s, marks, dflt, n, out);
    const color = seriesLine(env, s.index, env.maxIndex, []).color;
    const style = { color, fill: solid(env, color), line };
    for (const { p, i } of marks) sites.push({ g, s, i, value: vals[i]!, x: p.x, y: p.y, radius, style });
  });
}

function drawAreas(env: Env, g: Group, cat: AxisRt, val: AxisRt, out: Primitive[], sites: LabelSite[]): void {
  const n = categoryCount(g);
  if (!n) return;
  const stacked = isStacked(g);
  const st = stack(g, n);
  const base = baseValue(cat, val);
  const place = (i: number, v: number): Pt => {
    const along = pointAt(cat, i);
    const across = valueAt(val, clampT(val, v));
    return cat.dir === 'h' ? { x: along, y: across } : { x: across, y: along };
  };
  // a 3-D chart's first series stands in front: draw it last
  const order = g.series.map((s, j) => ({ s, j }));
  if (g.threeD && !stacked) order.reverse();
  for (const { s, j } of order) {
    const vals = s.val?.pts ?? [];
    const upper: Pt[] = [];
    const lower: Pt[] = [];
    for (let i = 0; i < n; i++) {
      upper.push(place(i, stacked ? st.top[j]![i]! : vals[i] ?? 0));
      lower.push(place(i, stacked ? st.base[j]![i]! : base));
    }
    const style = filledStyle(env, s.index, env.maxIndex, g.threeD, [s.spPr]);
    const prim = pathPrim(polyline([...upper, ...lower.reverse()], true), style.fill, style.line, style.effects);
    if (prim) out.push(prim);
    for (let i = 0; i < n; i++) {
      if (vals[i] === undefined) continue;
      const top = upper[i]!;
      const bottom = place(i, stacked ? st.base[j]![i]! : base);
      sites.push({ g, s, i, value: vals[i]!, x: (top.x + bottom.x) / 2, y: (top.y + bottom.y) / 2, style });
    }
  }
}

// ------------------------------------------------------------------------------------------- scatter, bubbles

function drawScatter(
  env: Env, chart: Chart, g: Group, xa: AxisRt, ya: AxisRt, out: Primitive[], sites: LabelSite[],
): void {
  const smoothAll = g.scatterStyle === 'smooth' || g.scatterStyle === 'smoothMarker';
  const lines = smoothAll || g.scatterStyle === 'line' || g.scatterStyle === 'lineMarker';
  for (const s of g.series) {
    const xs = xValues(s);
    const ys = yValues(g, s);
    const n = Math.max(xs.length, ys.length);
    const pts: (Pt | undefined)[] = [];
    const marks: Mark[] = [];
    for (let i = 0; i < n; i++) {
      const x = xs[i];
      let y = ys[i];
      if (y === undefined && chart.dispBlanksAs === 'zero' && x !== undefined) y = 0;
      if (x === undefined || y === undefined) {
        pts.push(undefined);
        continue;
      }
      const p = { x: valueAt(xa, x), y: valueAt(ya, y) };
      pts.push(p);
      marks.push({ p, i });
    }
    const { line, color } = seriesLine(env, s.index, env.maxIndex, [s.spPr]);
    if (lines && line) {
      const cmds = runs(pts, chart.dispBlanksAs).flatMap(r => (s.smooth || smoothAll ? smooth(r) : polyline(r)));
      const prim = pathPrim(cmds, NONE, line);
      if (prim) out.push(prim);
    }
    const dflt = g.scatterStyle === 'line' || g.scatterStyle === 'smooth' ? 'none' : 'auto';
    const radius = drawMarkers(env, g, s, marks, dflt, n, out);
    const style = { color, fill: solid(env, color), line };
    // a blank shown as zero gets a marker but no value to label
    for (const { p, i } of marks) {
      sites.push({ g, s, i, ...(ys[i] !== undefined ? { value: ys[i] } : {}), x: p.x, y: p.y, radius, style });
    }
  }
}

function drawBubbles(
  env: Env, g: Group, xa: AxisRt, ya: AxisRt, plot: Rect, out: Primitive[], sites: LabelSite[],
): void {
  // the largest bubble is a quarter of the plot area's smaller side across, scaled by c:bubbleScale
  let maxSize = 0;
  for (const s of g.series) {
    for (const v of s.size?.pts ?? []) if (v !== undefined) maxSize = Math.max(maxSize, Math.abs(v));
  }
  const rMax = (Math.min(plot.w, plot.h) * 0.25 * (g.bubbleScale / 100)) / 2;
  for (const s of g.series) {
    const xs = xValues(s);
    const ys = yValues(g, s);
    const sizes = s.size?.pts ?? [];
    const n = Math.max(xs.length, ys.length);
    for (let i = 0; i < n; i++) {
      const x = xs[i];
      const y = ys[i];
      const z = sizes[i];
      if (x === undefined || y === undefined || z === undefined || !maxSize) continue;
      if (z < 0 && !g.showNegBubbles) continue;
      const f = Math.abs(z) / maxSize;
      const r = rMax * (g.sizeRepresents === 'w' ? f : Math.sqrt(f));
      const cx = valueAt(xa, x);
      const cy = valueAt(ya, y);
      const { k, max } = colorIndex(env, g, s, i, n);
      const style = filledStyle(env, k, max, false, [s.spPr, pointSpPr(s, i)]);
      // negative bubbles are drawn hollow, outlined in the series colour
      const hollow = z < 0;
      const line = hollow ? style.line ?? strokeOf(env, undefined, themedLine(env, style.color)) : style.line;
      const prim = pathPrim(ellipse(cx, cy, r, r), hollow ? NONE : style.fill, line, style.effects);
      if (prim) out.push(prim);
      sites.push({ g, s, i, value: y, size: z, x: cx, y: cy, radius: r, style });
    }
  }
}

// ---------------------------------------------------------------------------------------------------- stock

/**
 * The extra lines of line and stock charts: high-low lines across each category's values, up and down bars from the
 * first series to the last (open to close), and drop lines from each point to the category axis.
 */
function drawExtras(env: Env, g: Group, cat: AxisRt, val: AxisRt, out: Primitive[]): void {
  const n = categoryCount(g);
  const series = g.series;
  const other = () => themedLine(env, otherLineColor(env.style));
  if (g.hiLowLines) {
    const stroke = strokeOf(env, g.hiLowLines.spPr, other());
    const cmds: Cmd[] = [];
    for (let i = 0; i < n; i++) {
      const vs = series.map(s => s.val?.pts[i]).filter((v): v is number => v !== undefined);
      if (vs.length < 2) continue;
      const x = pointAt(cat, i);
      cmds.push(...lineCmds({ x, y: valueAt(val, Math.min(...vs)) }, { x, y: valueAt(val, Math.max(...vs)) }));
    }
    const prim = stroke ? pathPrim(cmds, NONE, stroke) : undefined;
    if (prim) out.push(prim);
  }
  if (g.dropLines) {
    const stroke = strokeOf(env, g.dropLines.spPr, other());
    const base = valueAt(val, baseValue(cat, val));
    const cmds: Cmd[] = [];
    for (let i = 0; i < n; i++) {
      const vs = series.map(s => s.val?.pts[i]).filter((v): v is number => v !== undefined);
      if (!vs.length) continue;
      // from the axis to the highest point, or to the top of the stack
      const top = g.grouping === 'standard' ? Math.max(...vs) : vs.reduce((a, b) => a + b, 0);
      const x = pointAt(cat, i);
      cmds.push(...lineCmds({ x, y: base }, { x, y: valueAt(val, clampT(val, top)) }));
    }
    const prim = stroke ? pathPrim(cmds, NONE, stroke) : undefined;
    if (prim) out.push(prim);
  }
  if (g.upDownBars && series.length >= 2) {
    const open = series[0]!;
    const close = series.at(-1)!;
    const w = slotSize(cat) / (1 + g.upDownBars.gapWidth / 100);
    for (let i = 0; i < n; i++) {
      const o = open.val?.pts[i];
      const c = close.val?.pts[i];
      if (o === undefined || c === undefined) continue;
      const up = c >= o;
      const x = pointAt(cat, i);
      const y0 = valueAt(val, o);
      const y1 = valueAt(val, c);
      const spPr = up ? g.upDownBars.up : g.upDownBars.down;
      const fill = fillOf(env, spPr, () => upDownFill(env, up));
      // Table 4: the bars of the plainer styles are outlined
      const outlined = env.style <= 16 || (env.style >= 33 && env.style <= 40);
      const line = strokeOf(env, spPr, outlined ? other() : null);
      const prim = pathPrim(rect({ x: x - w / 2, y: Math.min(y0, y1), w, h: Math.abs(y1 - y0) }), fill, line);
      if (prim) out.push(prim);
    }
  }
}

/** Draw one cartesian chart group on its two axes. */
export function drawCartesian(
  env: Env, chart: Chart, g: Group, a: AxisRt, b: AxisRt, plot: Rect, out: Primitive[], sites: LabelSite[],
): void {
  switch (g.kind) {
    case 'bar': drawBars(env, g, a, b, out, sites); break;
    case 'line': case 'stock':
      drawExtras(env, g, a, b, out);
      drawLines(env, chart, g, a, b, out, sites);
      break;
    case 'area':
      drawAreas(env, g, a, b, out, sites);
      drawExtras(env, g, a, b, out);
      break;
    case 'scatter': drawScatter(env, chart, g, a, b, out, sites); break;
    case 'bubble': drawBubbles(env, g, a, b, plot, out, sites); break;
    default: break;
  }
}
