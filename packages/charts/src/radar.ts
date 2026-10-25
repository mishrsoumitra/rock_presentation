// packages/charts/src/radar.ts  -- radar charts: one spoke per category clockwise from 12 o'clock, values outward
//
// The value axis runs up the first spoke with its labels to its left; its major gridlines are polygons through the
// spokes, and the category axis's gridlines are the spokes themselves. c:radarStyle filled fills each series.
import type { Primitive } from '@pptx/render';
import type { Axis, Chart, Group } from './model.ts';
import { categoryCount, labelBlock, PT } from './axes.ts';
import { line as lineCmds, pathPrim, polyline, type Cmd, type Pt, type Rect } from './draw.ts';
import { formatValue } from './numfmt.ts';
import { autoScale, fraction, ticks, type ScaleInput } from './scale.ts';
import { drawMarkers, type LabelSite, type Mark } from './series.ts';
import { axisColor, filledStyle, NONE, seriesLine, solid, strokeOf, themedLine, type Env } from './style.ts';
import { placeAt, textPrim, textStyle } from './text.ts';

/** The group's axis with an id: of the kind asked for if two share it, else whichever has it. */
function axisOf(chart: Chart, id: number | undefined, kind: 'cat' | 'val'): Axis | undefined {
  const withId = chart.plotArea.axes.filter(a => a.id === id);
  return withId.find(a => (kind === 'val') === (a.kind === 'val')) ?? withId[0];
}

/** Where the drawing goes: shapes, text above them, and label sites. */
export interface RadarOut { out: Primitive[]; texts: Primitive[]; sites: LabelSite[]; }

/** Draw a radar group in the plot area, with its own axes. */
export function drawRadar(env: Env, chart: Chart, g: Group, plot: Rect, fontSize: number, to: RadarOut): void {
  const { out, texts, sites } = to;
  const n = categoryCount(g);
  if (n < 1) return;
  const catM = axisOf(chart, g.axIds[0], 'cat');
  const valM = axisOf(chart, g.axIds[1], 'val');
  const catStyle = textStyle(env, [catM?.txPr, chart.txPr], { size: fontSize, bold: false });
  const valStyle = textStyle(env, [valM?.txPr, chart.txPr], { size: fontSize, bold: false });
  const source = g.series.find(s => s.cat)?.cat;
  const catFormat = catM?.numFmt && !catM.numFmt.linked ? catM.numFmt.code : source?.formatCode ?? 'General';
  const catTexts = Array.from({ length: n }, (_, i) => {
    const num = source?.nums?.[i];
    return num !== undefined ? formatValue(num, catFormat, env.date1904) : source?.pts[i] ?? String(i + 1);
  });
  const showCats = catM ? !catM.deleted && catM.tickLblPos !== 'none' : true;
  const catBlocks = showCats ? catTexts.map(t => labelBlock(env, catStyle, t)) : [];
  const gap = fontSize * PT * 0.4;
  const maxW = Math.max(0, ...catBlocks.map(b => b.w));
  const maxH = Math.max(0, ...catBlocks.map(b => b.h));
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  // the radius leaves room for the category labels, but not below a tenth of the plot area
  const R = Math.max(Math.min(plot.w, plot.h) * 0.1, Math.min(plot.w / 2 - maxW - gap, plot.h / 2 - maxH - gap));

  // the value scale, spaced like an upright axis of length R
  let lo = Infinity;
  let hi = -Infinity;
  for (const s of g.series) {
    for (const v of s.val?.pts ?? []) {
      if (v === undefined) continue;
      lo = Math.min(lo, v);
      hi = Math.max(hi, v);
    }
  }
  const lineH = labelBlock(env, valStyle, '0').h || fontSize * PT;
  const intervals = (s: { min: number; max: number; major: number }) => (s.max - s.min) / s.major;
  const input: ScaleInput = {
    lo: lo <= hi ? lo : 0, hi: lo <= hi ? hi : 1,
    fits: s => intervals(s) <= 10.0001 && R / Math.max(1, intervals(s)) >= lineH * 0.85,
  };
  for (const k of ['min', 'max', 'majorUnit'] as const) if (valM?.[k] !== undefined) input[k] = valM[k];
  const scale = autoScale(input);
  const angle = (i: number) => -Math.PI / 2 + (2 * Math.PI * i) / n;
  const rad = (v: number) => R * Math.max(0, Math.min(1, fraction(scale, v)));
  const at = (i: number, r: number): Pt => ({ x: cx + r * Math.cos(angle(i)), y: cy + r * Math.sin(angle(i)) });

  // gridlines: polygons for the value axis, spokes for the categories
  const grid = (spPr: Axis['majorGrid'], cmds: Cmd[]) => {
    if (!spPr) return;
    const stroke = strokeOf(env, spPr.spPr, themedLine(env, axisColor(env.style)));
    const prim = stroke ? pathPrim(cmds, NONE, stroke) : undefined;
    if (prim) out.push(prim);
  };
  const tickValues = ticks(scale);
  const ring = (v: number) => polyline(Array.from({ length: n }, (_, i) => at(i, rad(v))), true);
  grid(valM?.majorGrid, tickValues.filter(v => rad(v) > 0).flatMap(ring));
  grid(catM?.majorGrid, Array.from({ length: n }, (_, i) => lineCmds({ x: cx, y: cy }, at(i, R))).flat());

  // series: blanks sit at the centre (zero) or at the scale's minimum (gap)
  for (const s of g.series) {
    const vals = s.val?.pts ?? [];
    const pts: Pt[] = [];
    const marks: Mark[] = [];
    for (let i = 0; i < n; i++) {
      const p = at(i, rad(vals[i] ?? (chart.dispBlanksAs === 'gap' ? scale.min : 0)));
      pts.push(p);
      if (vals[i] !== undefined) marks.push({ p, i });
    }
    if (g.radarStyle === 'filled') {
      const st = filledStyle(env, s.index, env.maxIndex, false, [s.spPr]);
      const prim = pathPrim(polyline(pts, true), st.fill, st.line, st.effects);
      if (prim) out.push(prim);
      for (const { p, i } of marks) sites.push({ g, s, i, value: vals[i]!, x: p.x, y: p.y, style: st });
      continue;
    }
    const { line, color } = seriesLine(env, s.index, env.maxIndex, [s.spPr]);
    const prim = line ? pathPrim(polyline(pts, true), NONE, line) : undefined;
    if (prim) out.push(prim);
    const radius = drawMarkers(env, g, s, marks, g.radarStyle === 'marker' ? 'auto' : 'none', n, out);
    const style = { color, fill: solid(env, color), line };
    for (const { p, i } of marks) sites.push({ g, s, i, value: vals[i]!, x: p.x, y: p.y, radius, style });
  }

  // the value axis up the first spoke, its labels to the left
  if (valM && !valM.deleted) {
    const stroke = strokeOf(env, valM.spPr, themedLine(env, axisColor(env.style)));
    const cmds: Cmd[] = lineCmds({ x: cx, y: cy }, { x: cx, y: cy - R });
    const len = 3 * PT;
    // out is to the left, towards the labels
    const tick = valM.majorTick;
    const [t0, t1] = tick === 'cross' ? [-len, len] : tick === 'in' ? [0, len] : [-len, 0];
    for (const v of tick === 'none' ? [] : tickValues) {
      const y = cy - rad(v);
      cmds.push(...lineCmds({ x: cx + t0, y }, { x: cx + t1, y }));
    }
    const prim = stroke ? pathPrim(cmds, NONE, stroke) : undefined;
    if (prim) out.push(prim);
    if (valM.tickLblPos !== 'none') {
      const fmt = valM.numFmt && !valM.numFmt.linked ? valM.numFmt.code : g.series[0]?.val?.formatCode ?? 'General';
      for (const v of tickValues) {
        const b = labelBlock(env, valStyle, formatValue(v, fmt, env.date1904));
        texts.push(textPrim(b, placeAt(cx - 4 * PT, cy - rad(v), b.w, b.h / 2)));
      }
    }
  }
  // category labels around the outside, pushed away from the centre
  catBlocks.forEach((b, i) => {
    const p = at(i, R + gap);
    const c = Math.cos(angle(i));
    const sn = Math.sin(angle(i));
    const px = Math.abs(c) < 0.2 ? b.w / 2 : c > 0 ? 0 : b.w;
    const py = Math.abs(sn) < 0.2 ? b.h / 2 : sn > 0 ? 0 : b.h;
    texts.push(textPrim(b, placeAt(p.x, p.y, px, py)));
  });
}
