// packages/charts/src/chart.ts  -- a whole chart: automatic layout of title, legend and plot area, then drawing
//
// Without a manual layout Office puts the title at the top centre, the legend on its side (right by default), and
// gives the plot area what is left, minus the room its axes' labels and titles need. A c:manualLayout fixes the
// plot area's inner rectangle (layoutTarget inner) or the rectangle including its labels (outer).
import type { Primitive } from '@pptx/render';
import type { Stroke } from '@pptx/resolve';
import type { Chart, Layout } from './model.ts';
import {
  at, buildAxes, crossFraction, drawAxisLine, drawAxisText, drawGridlines, layoutLabels, layoutTitle, prepareAxes, PT,
  reach, scaleAxis, setEnds, titleExtent, type AxisRt,
} from './axes.ts';
import { pathPrim, rect, roundRect, styledPath, type Rect } from './draw.ts';
import { drawLabels } from './labels.ts';
import { layoutLegend } from './legend.ts';
import { drawPie } from './pie.ts';
import { drawRadar } from './radar.ts';
import { drawCartesian, type LabelSite } from './series.ts';
import { drawUserShapes } from './shapes.ts';
import { boxStyle, chartAreaFill, plotAreaFill, type Env } from './style.ts';
import {
  block, chartFontSize, isEmpty, placeAt, plainBody, richBody, textPrim, textStyle, type Block,
} from './text.ts';

type Size = { w: number; h: number };
const PIES = new Set(['pie', 'doughnut', 'ofPie']);

/** The chart title's text, or undefined when there is none to show (ECMA-376 21.2.2.7 autoTitleDeleted). */
function titleText(env: Env, chart: Chart, fontSize: number, maxW: number): Block | undefined {
  const series = chart.plotArea.groups.flatMap(g => g.series);
  const t = chart.title;
  if (!t && (chart.autoTitleDeleted || series.length !== 1 || !series[0]!.named)) return undefined;
  const d = { size: fontSize * 1.2, bold: true, scaled: true };
  const chain = [t?.txPr, chart.txPr];
  if (t?.rich) {
    const r = richBody(env, t.rich, chain, d);
    return isEmpty(r.body) ? undefined : block(env, r.body, maxW);
  }
  // a title without text shows the only series' name, else Office's placeholder text
  const only = series.length === 1 && series[0]!.named ? series[0]!.name : undefined;
  const text = t?.text ?? only ?? 'Chart Title';
  if (!text.trim()) return undefined;
  return block(env, plainBody(textStyle(env, chain, d), text), maxW);
}

/** A manual layout's rectangle in chart space, with `auto` the automatic one it offsets in factor mode. */
function manualRect(m: Layout, auto: Rect, box: Size): Rect {
  const x = m.x === undefined ? auto.x : m.xMode === 'edge' ? m.x * box.w : auto.x + m.x * box.w;
  const y = m.y === undefined ? auto.y : m.yMode === 'edge' ? m.y * box.h : auto.y + m.y * box.h;
  const w = m.w === undefined ? auto.w : m.wMode === 'edge' ? m.w * box.w - x : m.w * box.w;
  const h = m.h === undefined ? auto.h : m.hMode === 'edge' ? m.h * box.h - y : m.h * box.h;
  return { x, y, w: Math.max(1, w), h: Math.max(1, h) };
}

/** The side of the plot area an axis's labels are on. */
const sideOf = (a: AxisRt): 'l' | 'r' | 't' | 'b' =>
  (a.dir === 'h' ? (a.side > 0 ? 'b' : 't') : a.side > 0 ? 'r' : 'l');

/**
 * Lay out the cartesian axes in `outer` (or with their inner rectangle fixed) and return the inner plot rectangle.
 * Labels depend on the scale, which depends on the axis length, which depends on the labels: a few passes settle it.
 */
function layoutCartesian(env: Env, axes: AxisRt[], outer: Rect, fixedInner: Rect | undefined, box: Size): Rect {
  let plot = fixedInner ?? { ...outer };
  for (let pass = 0; pass < 4; pass++) {
    for (const a of axes) setEnds(a, plot);
    for (const a of axes) scaleAxis(env, a);
    // each axis lies where it crosses its partner; one without a partner lies along the plot's bottom or left
    for (const a of axes) a.line = a.other ? at(a.other, crossFraction(a)) : a.dir === 'h' ? plot.y + plot.h : plot.x;
    for (const a of axes) layoutLabels(env, a, plot, box.w, box.h);
    if (fixedInner) break;
    const need = { l: 0, r: 0, t: 0, b: 0 };
    for (const a of axes) {
      if (a.m.deleted) continue;
      const side = sideOf(a);
      const edge = side === 'b' ? plot.y + plot.h : side === 't' ? plot.y : side === 'r' ? plot.x + plot.w : plot.x;
      const atEdge = Math.abs(a.labelLine - edge) < Math.max(plot.w, plot.h) * 1e-3;
      need[side] = Math.max(need[side], (atEdge ? a.extent : 0) + titleExtent(a));
      // the first and last labels hang over the ends of the axis
      for (const l of a.labels) {
        if (a.dir === 'h') {
          need.l = Math.max(need.l, plot.x - (l.x - reach(l, 'v', -1)));
          need.r = Math.max(need.r, l.x + reach(l, 'v', 1) - (plot.x + plot.w));
        } else {
          need.t = Math.max(need.t, plot.y - (l.y - reach(l, 'h', -1)));
          need.b = Math.max(need.b, l.y + reach(l, 'h', 1) - (plot.y + plot.h));
        }
      }
    }
    const next = {
      x: outer.x + need.l, y: outer.y + need.t, w: outer.w - need.l - need.r, h: outer.h - need.t - need.b,
    };
    // never squeeze the plot area below a fifth of what there is
    next.w = Math.max(next.w, outer.w * 0.2);
    next.h = Math.max(next.h, outer.h * 0.2);
    const moved = (['x', 'y', 'w', 'h'] as const).reduce((sum, k) => sum + Math.abs(next[k] - plot[k]), 0);
    const same = moved < 1000;
    plot = next;
    if (same && pass > 0) break;
  }
  return plot;
}

const grey = (v: number) => ({ kind: 'solid' as const, color: { r: v, g: v, b: v, a: 1 } });
const DASHED: Stroke = {
  width: 9525, paint: grey(0.63), cap: 'flat', join: 'miter', miterLimit: 8, dash: [4, 3], compound: 'sng',
  align: 'ctr',
};

/** A labelled box where a chart type is not drawn (surface charts). */
function placeholder(
  env: Env, chart: Chart, r: Rect, fontSize: number, label: string, paths: Primitive[], texts: Primitive[],
): void {
  const box = pathPrim(rect(r), grey(0.92), DASHED);
  if (box) paths.push(box);
  const b = block(env, plainBody(textStyle(env, [chart.txPr], { size: fontSize, bold: false }), label));
  texts.push(textPrim(b, placeAt(r.x + r.w / 2, r.y + r.h / 2, b.w / 2, b.h / 2)));
}

/** Draw a parsed chart into a box (EMU); primitives in the box's own space. */
export function drawChart(chart: Chart, env: Env, box: Size): Primitive[] {
  const paths: Primitive[] = [];
  const texts: Primitive[] = [];
  const W = box.w;
  const H = box.h;
  const fs = chartFontSize(chart.txPr);
  const em = fs * PT;
  const whole: Rect = { x: 0, y: 0, w: W, h: H };
  const outline = chart.roundedCorners ? roundRect(whole, Math.min(W, H) * 0.06) : rect(whole);
  const bg = styledPath(outline, boxStyle(env, chart.spPr, () => chartAreaFill(env)));
  if (bg) paths.push(bg);

  // Office keeps a margin of 7 px (at 96 dpi) inside the chart area
  const pad = 7 * 9525;
  let free: Rect = { x: pad, y: pad, w: W - 2 * pad, h: H - 2 * pad };

  // title
  const title = titleText(env, chart, fs, W * 0.8);
  let titleBox: Rect | undefined;
  if (title) {
    const size = title.layout.lines[0]?.runs[0]?.size ?? em * 1.2;
    const auto = { x: (W - title.w) / 2, y: pad + size * 0.2, w: title.w, h: title.h };
    titleBox = chart.title?.layout ? { ...manualRect(chart.title.layout, auto, box), w: title.w, h: title.h } : auto;
    const below = titleBox.y + title.h + em * 0.6;
    if (!chart.title?.overlay && !chart.title?.layout) free = { ...free, y: below, h: free.h - (below - free.y) };
  }
  const gap = em;

  // legend
  const legend = layoutLegend(env, chart, free, box, fs);
  if (legend && chart.legend && !chart.legend.overlay && !chart.legend.layout) {
    const r = legend.rect;
    switch (chart.legend.pos) {
      case 'l': free = { ...free, x: r.x + r.w + gap, w: free.w - (r.x + r.w + gap - free.x) }; break;
      case 't': free = { ...free, y: r.y + r.h + gap, h: free.h - (r.y + r.h + gap - free.y) }; break;
      case 'b': free = { ...free, h: r.y - gap - free.y }; break;
      default: free = { ...free, w: r.x - gap - free.x };
    }
  }
  free = { ...free, w: Math.max(free.w, W * 0.1), h: Math.max(free.h, H * 0.1) };

  // plot area
  const groups = chart.plotArea.groups;
  const m = chart.plotArea.layout;
  const manual = m && (m.x !== undefined || m.w !== undefined) ? manualRect(m, free, box) : undefined;
  const outer = manual && m?.target === 'outer' ? manual : free;
  const fixedInner = manual && m?.target === 'inner' ? manual : undefined;
  const { axes, of } = buildAxes(chart);
  const sites: LabelSite[] = [];
  let plot = fixedInner ?? outer;
  if (axes.length) {
    prepareAxes(env, chart, axes, fs);
    for (const a of axes) layoutTitle(env, a, chart, fs);
    plot = layoutCartesian(env, axes, outer, fixedInner, box);
  }
  const threeD = groups.some(g => g.threeD);
  const plotBg = styledPath(rect(plot), boxStyle(env, chart.plotArea.spPr, () => plotAreaFill(env, threeD)));
  if (plotBg) paths.push(plotBg);
  for (const a of axes) drawGridlines(env, a, plot, paths);
  // area charts at the back, then bars, then lines, then points, as Office stacks chart types
  const rank = (k: string) => (k === 'area' ? 0 : k === 'bar' ? 1 : k === 'line' || k === 'stock' ? 2 : 3);
  const ordered = groups.map((g, i) => ({ g, i })).sort((p, q) => rank(p.g.kind) - rank(q.g.kind) || p.i - q.i);
  // charts without axes take the plot area as laid out, labels included
  const area = fixedInner ?? outer;
  for (const { g } of ordered) {
    const pair = of.get(g);
    if (pair) drawCartesian(env, chart, g, pair.a, pair.b, plot, paths, sites);
    else if (PIES.has(g.kind)) drawPie(env, chart, g, area, paths, sites);
    else if (g.kind === 'radar') drawRadar(env, chart, g, area, fs, { out: paths, texts, sites });
    else if (g.kind === 'surface') placeholder(env, chart, area, fs, 'Surface chart', paths, texts);
  }
  for (const a of axes) drawAxisLine(env, a, paths);
  for (const a of axes) drawAxisText(a, plot, texts);
  drawLabels(env, chart, sites, box, fs, paths, texts);
  legend?.draw(paths, texts);
  if (title && titleBox && chart.title?.spPr) {
    const { x, y, w, h } = titleBox;
    const frame = rect({ x: x - em * 0.2, y: y - em * 0.1, w: w + em * 0.4, h: h + em * 0.2 });
    const tb = styledPath(frame, boxStyle(env, chart.title.spPr));
    if (tb) paths.push(tb);
  }
  if (title && titleBox) texts.push(textPrim(title, placeAt(titleBox.x, titleBox.y, 0, 0)));
  // shapes the author drew over the chart; a broken drawing part leaves the chart as it is
  try {
    drawUserShapes(env, box, paths, texts);
  } catch {
    // nothing to add
  }
  return [...paths, ...texts];
}
