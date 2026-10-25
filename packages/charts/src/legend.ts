// packages/charts/src/legend.ts  -- the legend: a key and a name per series, or per category for varied colours
//
// Upright legends (right, left, top right) list horizontal bars and stacked columns or areas bottom series last, so
// the list reads in the order the series are stacked on the chart, as Office does.
import type { Paint, ResolvedEffects, Stroke } from '@pptx/resolve';
import type { Primitive } from '@pptx/render';
import type { Chart, Group, Series } from './model.ts';
import { PT } from './axes.ts';
import { AUTO_SYMBOLS, line as lineCmds, marker, pathPrim, rect, type Rect } from './draw.ts';
import { formatValue } from './numfmt.ts';
import { effectsOf, fillOf, filledStyle, markerStyle, NONE, seriesLine, strokeOf, type Env } from './style.ts';
import { block, placeAt, plainBody, textPrim, textStyle, type Block } from './text.ts';

/** One legend entry: its text and its key, a filled square or a line with an optional marker. */
export interface Entry {
  text: string;
  key: 'fill' | 'line';
  fill: Paint;
  line: Stroke | null;
  effects?: ResolvedEffects;
  marker?: { symbol: string; fill: Paint; line: Stroke | null; size: number };
}

/** The text of category i: a number formatted with the cache's format, or the label as cached. */
export function categoryText(env: Env, s: Series | undefined, i: number): string {
  const cat = s?.cat;
  const n = cat?.nums?.[i];
  if (n !== undefined) return formatValue(n, cat?.formatCode ?? 'General', env.date1904);
  return cat?.pts[i] ?? String(i + 1);
}

const PIE = new Set(['pie', 'doughnut', 'ofPie']);
const LINE_LIKE = new Set(['line', 'scatter', 'stock']);
const SCATTER_LINES = new Set(['line', 'lineMarker', 'smooth', 'smoothMarker']);

/** The marker a group's series show when their own c:marker says nothing. */
function defaultSymbol(g: Group): 'auto' | 'none' {
  if (g.kind === 'line') return g.showMarker ? 'auto' : 'none';
  if (g.kind === 'scatter') return g.scatterStyle === 'line' || g.scatterStyle === 'smooth' ? 'none' : 'auto';
  if (g.kind === 'radar') return g.radarStyle === 'marker' ? 'auto' : 'none';
  return 'none';
}

/** A series' entry: a filled key for bars, areas and slices, its line and marker for line-like series. */
function seriesEntry(env: Env, g: Group, s: Series): Entry {
  if (!LINE_LIKE.has(g.kind) && !(g.kind === 'radar' && g.radarStyle !== 'filled')) {
    const st = filledStyle(env, s.index, env.maxIndex, g.threeD, [s.spPr]);
    return { text: s.name, key: 'fill', fill: st.fill, line: st.line, effects: st.effects };
  }
  const { line, color } = seriesLine(env, s.index, env.maxIndex, [s.spPr]);
  let symbol = s.marker?.symbol ?? defaultSymbol(g);
  if (symbol === 'auto') symbol = AUTO_SYMBOLS[s.index % AUTO_SYMBOLS.length]!;
  // the key shows the series' line: none for unconnected scatter points, only an explicit one for stock
  let shown = line;
  if (g.kind === 'stock') shown = strokeOf(env, s.spPr, null);
  else if (g.kind === 'scatter' && !SCATTER_LINES.has(g.scatterStyle)) shown = null;
  const e: Entry = { text: s.name, key: 'line', fill: NONE, line: shown };
  if (symbol !== 'none') {
    const ms = markerStyle(env, color, [s.marker?.spPr]);
    const size = (s.marker?.size ?? 5) * PT;
    e.marker = { symbol: symbol === 'picture' ? 'square' : symbol, fill: ms.fill, line: ms.line, size };
  }
  return e;
}

/** Legend entries in display order (before deleted entries are dropped). */
export function legendEntries(env: Env, chart: Chart): Entry[] {
  const groups = chart.plotArea.groups;
  const total = groups.reduce((n, g) => n + g.series.length, 0);
  const out: Entry[] = [];
  for (const g of groups) {
    const s0 = g.series[0];
    // varied colours give an entry per point: always for pies, for bars and bubbles when the chart has one series
    const single = total === 1 && (g.kind === 'bar' || g.kind === 'bubble');
    const perPoint = s0 && g.varyColors && (PIE.has(g.kind) || single);
    if (perPoint) {
      const n = Math.max(s0.val?.pts.length ?? 0, s0.cat?.pts.length ?? 0, (s0.y ?? s0.val)?.pts.length ?? 0);
      for (let i = 0; i < n; i++) {
        const st = filledStyle(env, i, n - 1, g.threeD, [s0.spPr, s0.points.get(i)?.spPr]);
        out.push({ text: categoryText(env, s0, i), key: 'fill', fill: st.fill, line: st.line, effects: st.effects });
      }
      continue;
    }
    if (PIE.has(g.kind) && g.kind !== 'doughnut') {
      if (s0) out.push(seriesEntry(env, g, s0));
      continue;
    }
    for (const s of g.series) out.push(seriesEntry(env, g, s));
  }
  return out;
}

/** Whether an upright legend lists the series last to first (see the top of this file). */
function reversedFor(chart: Chart): boolean {
  const g = chart.plotArea.groups[0];
  if (!g) return false;
  const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
  return (g.kind === 'bar' && (g.barDir === 'bar' || stacked)) || (g.kind === 'area' && stacked);
}

/** A legend placed in the chart: its box, and how to draw it once the rest of the chart is drawn. */
export interface LegendLayout { rect: Rect; draw: (paths: Primitive[], texts: Primitive[]) => void; }

/** Lay the legend out inside `area` (the chart area within its margins); undefined when it has nothing to show. */
export function layoutLegend(
  env: Env, chart: Chart, area: Rect, chartBox: { w: number; h: number }, fontSize: number,
): LegendLayout | undefined {
  const lg = chart.legend;
  if (!lg) return undefined;
  const upright = lg.pos === 'r' || lg.pos === 'l' || lg.pos === 'tr';
  let entries = legendEntries(env, chart).map((e, i) => ({ e, i })).filter(({ i }) => !lg.deleted.has(i));
  if (upright && reversedFor(chart)) entries = entries.reverse();
  // a legend shows what fits: no need to lay out thousands of entries
  entries = entries.slice(0, 500);
  if (!entries.length) return undefined;
  const em = fontSize * PT;
  const styleFor = (i: number) =>
    textStyle(env, [lg.entryTxPr.get(i), lg.txPr, chart.txPr], { size: fontSize, bold: false });
  const base = styleFor(-1);
  // Office's legend keys: a small square, or a short line with the marker on it
  const keyH = base.size * PT * 0.4;
  const hasLine = entries.some(({ e }) => e.key === 'line');
  const keyW = hasLine ? base.size * PT * 1.25 : keyH;
  const gap = base.size * PT * (hasLine ? 0.1 : 0.25);
  const pad = em * 0.5;
  const maxText = upright ? chartBox.w * 0.3 : area.w - 2 * pad - keyW - gap;
  const items = entries.map(({ e, i }) => {
    const style = lg.entryTxPr.has(i) ? styleFor(i) : base;
    const b: Block = block(env, plainBody(style, e.text, 'l'), Math.max(em, maxText));
    return { e, b, w: keyW + gap + b.w, h: Math.max(b.h, keyH) };
  });
  // rows: one entry per row when upright, packed left to right otherwise
  const rows: (typeof items)[] = [];
  if (upright) for (const it of items) rows.push([it]);
  else {
    const spacing = em * 0.8;
    let cur: typeof items = [];
    let width = 0;
    for (const it of items) {
      if (cur.length && width + spacing + it.w > area.w - 2 * pad) {
        rows.push(cur);
        cur = [];
        width = 0;
      }
      width += (cur.length ? spacing : 0) + it.w;
      cur.push(it);
    }
    if (cur.length) rows.push(cur);
  }
  const spacing = em * 0.8;
  const rowW = (r: typeof items) => r.reduce((s, it) => s + it.w, 0) + spacing * (r.length - 1);
  const rowH = (r: typeof items) => Math.max(...r.map(it => it.h));
  const vgap = em * 0.3;
  let w = Math.max(...rows.map(rowW)) + 2 * pad;
  let h = rows.reduce((s, r) => s + rowH(r), 0) + vgap * (rows.length - 1) + 2 * pad;
  // a legend taller than the chart shows the entries that fit
  while (rows.length > 1 && h > area.h) {
    const r = rows.pop()!;
    h -= rowH(r) + vgap;
  }
  let x: number;
  let y: number;
  switch (lg.pos) {
    case 'l': x = area.x; y = area.y + (area.h - h) / 2; break;
    case 't': x = area.x + (area.w - w) / 2; y = area.y; break;
    case 'b': x = area.x + (area.w - w) / 2; y = area.y + area.h - h; break;
    case 'tr': x = area.x + area.w - w; y = area.y; break;
    default: x = area.x + area.w - w; y = area.y + (area.h - h) / 2;
  }
  const m = lg.layout;
  if (m) {
    // edge: a fraction of the chart from its top left; factor: a fraction of the chart moved from where it was
    const at = (v: number | undefined, mode: string, dflt: number, size: number) =>
      (v === undefined ? dflt : mode === 'edge' ? v * size : dflt + v * size);
    x = at(m.x, m.xMode, x, chartBox.w);
    y = at(m.y, m.yMode, y, chartBox.h);
    if (m.w !== undefined && m.w > 0) w = m.wMode === 'edge' ? m.w * chartBox.w - x : m.w * chartBox.w;
    if (m.h !== undefined && m.h > 0) h = m.hMode === 'edge' ? m.h * chartBox.h - y : m.h * chartBox.h;
    // a legend moved by hand stays inside the chart area
    x = Math.max(0, Math.min(chartBox.w - w, x));
    y = Math.max(0, Math.min(chartBox.h - h, y));
  }
  const box: Rect = { x, y, w, h };
  const draw = (paths: Primitive[], texts: Primitive[]) => {
    const spPr = lg.spPr;
    const bg = pathPrim(rect(box), fillOf(env, spPr, () => NONE), strokeOf(env, spPr, null), effectsOf(env, spPr));
    if (bg) paths.push(bg);
    const used = rows.reduce((s, r) => s + rowH(r), 0) + vgap * (rows.length - 1);
    let cy = y + (h - used) / 2;
    for (const r of rows) {
      const rh = rowH(r);
      let cx = upright ? x + pad : x + (w - rowW(r)) / 2;
      for (const it of r) {
        const mid = cy + rh / 2;
        drawKey(it.e, cx, mid, keyW, keyH, paths);
        texts.push(textPrim(it.b, placeAt(cx + keyW + gap, mid, 0, it.b.h / 2)));
        cx += it.w + spacing;
      }
      cy += rh + vgap;
    }
  };
  return { rect: box, draw };
}

/** An entry's key, keyW wide and centred on `mid`: a square keyH across, or a line with the marker on it. */
function drawKey(e: Entry, x: number, mid: number, keyW: number, keyH: number, out: Primitive[]): void {
  if (e.key === 'fill') {
    const square = rect({ x: x + (keyW - keyH) / 2, y: mid - keyH / 2, w: keyH, h: keyH });
    const k = pathPrim(square, e.fill, e.line);
    if (k) out.push(k);
    return;
  }
  if (e.line) {
    const l = pathPrim(lineCmds({ x, y: mid }, { x: x + keyW, y: mid }), NONE, e.line);
    if (l) out.push(l);
  }
  if (e.marker) {
    const m = marker(e.marker.symbol, x + keyW / 2, mid, Math.min(e.marker.size, keyH * 2));
    const p = pathPrim(m.cmds, m.open ? NONE : e.marker.fill, e.marker.line);
    if (p) out.push(p);
  }
}
