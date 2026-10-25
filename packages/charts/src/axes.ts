// packages/charts/src/axes.ts  -- cartesian axes: roles, data extents, scaling, tick labels, lines and gridlines
//
// Each chart group names two axes (c:axId). For bar, line, area and stock charts the first is the category axis and
// the second the value axis; scatter and bubble charts have two value axes, x first. Horizontal bars (c:barDir bar)
// turn the category axis upright. An axis's c:crosses says where it crosses the other axis, in that axis's units.
import type { ColorDef } from '@pptx/model';
import type { Primitive } from '@pptx/render';
import type { Axis, Chart, Group } from './model.ts';
import { dateToSerial, formatValue, serialToDate } from './numfmt.ts';
import { autoScale, fraction, ticks, type Scale, type ScaleInput } from './scale.ts';
import { line as lineCmds, pathPrim, type Pt, type Rect } from './draw.ts';
import { axisColor, minorGridColor, NONE, strokeOf, themedLine, type Env } from './style.ts';
import {
  block, placeAt, plainBody, richBody, rotatedSize, textPrim, textStyle, type Block, type TextStyle,
} from './text.ts';

export const PT = 12700;
const MAJOR_TICK = 4 * PT;
const MINOR_TICK = 2 * PT;

export interface PlacedLabel { b: Block; x: number; y: number; px: number; py: number; rot: number; }

/** An axis while the chart is laid out and drawn. */
export interface AxisRt {
  m: Axis;
  dir: 'h' | 'v';
  type: 'cat' | 'val';
  /** What the axis measures: categories, values, or a scatter chart's x or y. */
  role: 'cat' | 'val' | 'x' | 'y';
  groups: Group[];
  other?: AxisRt;
  // value axes
  lo: number;
  hi: number;
  hasData: boolean;
  percent: boolean;
  scale: Scale;
  // category axes
  count: number;
  between: boolean;
  texts: string[];
  /** Date axes: the slot (one per base time unit) of each point, and the label steps that unit allows. */
  slots?: number[];
  steps?: number[];
  /** Date axes: tick marks every this many slots, with the labels. */
  tickEvery?: number;
  /** Outer levels of multi-level category labels (level 1 outermost last), with each label's first category. */
  levels: { text: string; from: number; to: number }[][];
  format: string;
  style: TextStyle;
  // placement, in chart space
  start: number;
  end: number;
  line: number;
  side: -1 | 1;
  /** Where the labels hang: the axis line, or the plot edge for low and high labels. */
  labelLine: number;
  labels: PlacedLabel[];
  outer: PlacedLabel[];
  extent: number;
  title?: { b: Block; rot: number };
}

const SCATTERY = new Set(['scatter', 'bubble']);
const ROUND = new Set(['pie', 'doughnut', 'ofPie', 'radar']);
/** Whether a group is drawn against a pair of axes (pies, doughnuts and radar charts are not). */
export const isCartesian = (g: Group) => !ROUND.has(g.kind);
type Series = Group['series'][number];

function defaultAxis(kind: 'cat' | 'val', id: number, crossAx: number, pos: Axis['pos']): Axis {
  return {
    kind, id, crossAx, pos, deleted: false, reversed: false, crosses: 'autoZero', majorTick: 'out', minorTick: 'none',
    tickLblPos: 'nextTo', lblOffset: 100, noMultiLvlLbl: false, ...(kind === 'val' ? { majorGrid: {} } : {}),
  };
}

/** Values of a series as plotted on a value axis: c:val, or c:yVal for scatter and bubble charts. */
export const yValues = (g: Group, s: Series) => (SCATTERY.has(g.kind) ? s.y ?? s.val : s.val)?.pts ?? [];
/** X values of a scatter or bubble series; 1, 2, 3... when it has none (or only text). */
export function xValues(s: Series): (number | undefined)[] {
  if (s.x && s.x.pts.some(v => v !== undefined)) return s.x.pts;
  const n = (s.y ?? s.val)?.pts.length ?? 0;
  return Array.from({ length: n }, (_, i) => i + 1);
}

/** The number of category slots a group spreads over. */
export const categoryCount = (g: Group) =>
  Math.max(0, ...g.series.map(s => Math.max(s.val?.pts.length ?? 0, s.cat?.pts.length ?? 0)));

/** Stacked values per series and category: where each bar or area band starts and ends. */
export function stack(g: Group, n: number): { base: number[][]; top: number[][] } {
  const pos = new Array<number>(n).fill(0);
  const neg = new Array<number>(n).fill(0);
  const total = new Array<number>(n).fill(0);
  if (g.grouping === 'percentStacked') {
    for (const s of g.series) for (let i = 0; i < n; i++) total[i]! += Math.abs(s.val?.pts[i] ?? 0);
  }
  const base: number[][] = [];
  const top: number[][] = [];
  for (const s of g.series) {
    const b: number[] = [];
    const t: number[] = [];
    for (let i = 0; i < n; i++) {
      let v = s.val?.pts[i] ?? 0;
      if (g.grouping === 'percentStacked') v = total[i] ? v / total[i]! : 0;
      // lines stack on the running total; bars and areas keep positives and negatives apart
      const acc = g.kind === 'line' || v >= 0 ? pos : neg;
      if (g.kind === 'line') {
        b.push(pos[i]!);
        pos[i]! += v;
        t.push(pos[i]!);
        continue;
      }
      b.push(acc[i]!);
      acc[i]! += v;
      t.push(acc[i]!);
    }
    base.push(b);
    top.push(t);
  }
  return { base, top };
}

function extentOf(g: Group, role: 'x' | 'y'): { lo: number; hi: number } | undefined {
  let lo = Infinity;
  let hi = -Infinity;
  const see = (v: number | undefined) => {
    if (v === undefined || !Number.isFinite(v)) return;
    if (v < lo) lo = v;
    if (v > hi) hi = v;
  };
  if (role === 'x') for (const s of g.series) xValues(s).forEach(see);
  else if ((g.grouping === 'stacked' || g.grouping === 'percentStacked') && !SCATTERY.has(g.kind)) {
    const { top, base } = stack(g, categoryCount(g));
    for (const t of top) t.forEach(see);
    for (const b of base) b.forEach(see);
  } else for (const s of g.series) yValues(g, s).forEach(see);
  return lo <= hi ? { lo, hi } : undefined;
}

/** A group's two axes: categories (or x) and values (or y). */
export interface AxisPair { a: AxisRt; b: AxisRt; }

/** Build the axes of the chart's cartesian groups and pair each group with its two axes. */
export function buildAxes(chart: Chart): { axes: AxisRt[]; of: Map<Group, AxisPair> } {
  const byId = new Map<number, AxisRt>();
  const axes: AxisRt[] = [];
  const of = new Map<Group, AxisPair>();
  const models = new Map(chart.plotArea.axes.map(a => [a.id, a]));
  const rt = (m: Axis, role: AxisRt['role'], dir: 'h' | 'v'): AxisRt => {
    let r = byId.get(m.id);
    if (r) return r;
    r = {
      m, dir, type: role === 'cat' ? 'cat' : 'val', role, groups: [], lo: 0, hi: 1, hasData: false, percent: false,
      scale: { min: 0, max: 1, major: 0.2, minor: 0.04 },
      // the text style is set by prepareAxes, before anything reads it
      count: 0, between: true, texts: [], levels: [], format: 'General', style: undefined as never,
      start: 0, end: 0, line: 0, side: 1, labelLine: 0, labels: [], outer: [], extent: 0,
    };
    byId.set(m.id, r);
    axes.push(r);
    return r;
  };
  let synthetic = -1;
  for (const g of chart.plotArea.groups) {
    if (!isCartesian(g) || g.kind === 'surface') continue;
    const scatter = SCATTERY.has(g.kind);
    let m0 = models.get(g.axIds[0] ?? NaN);
    let m1 = models.get(g.axIds[1] ?? NaN);
    // the category axis comes first; a file that lists them the other way round still pairs them right
    if (!scatter && m0?.kind === 'val' && m1 && m1.kind !== 'val') [m0, m1] = [m1, m0];
    m0 ??= defaultAxis(scatter ? 'val' : 'cat', synthetic--, NaN, 'b');
    m1 ??= defaultAxis('val', synthetic--, NaN, 'l');
    const turned = g.kind === 'bar' && g.barDir === 'bar';
    const a = rt(m0, scatter ? 'x' : 'cat', turned ? 'v' : 'h');
    const b = rt(m1, scatter ? 'y' : 'val', turned ? 'h' : 'v');
    a.other ??= b;
    b.other ??= a;
    a.groups.push(g);
    b.groups.push(g);
    of.set(g, { a, b });
  }
  // data extents and category counts
  for (const [g, { a, b }] of of) {
    const ey = extentOf(g, 'y');
    if (ey) widen(b, ey);
    if (g.grouping === 'percentStacked') b.percent = true;
    if (a.type === 'val') {
      const ex = extentOf(g, 'x');
      if (ex) widen(a, ex);
    } else a.count = Math.max(a.count, categoryCount(g));
  }
  // bars sit between tick marks; lines and areas follow the value axis's c:crossBetween
  for (const a of axes) {
    if (a.type !== 'cat') continue;
    const first = a.groups[0]?.kind;
    const cb = a.other?.m.crossBetween ?? (first === 'area' || first === 'stock' ? 'midCat' : 'between');
    a.between = a.groups.some(g => g.kind === 'bar') || cb === 'between';
  }
  return { axes, of };
}

function widen(a: AxisRt, e: { lo: number; hi: number }): void {
  if (!a.hasData) {
    a.lo = e.lo;
    a.hi = e.hi;
  } else {
    a.lo = Math.min(a.lo, e.lo);
    a.hi = Math.max(a.hi, e.hi);
  }
  a.hasData = true;
}

// ------------------------------------------------------------------------------------------------------- labels

/** The number format of a value axis's labels: its own, or the source data's when linked (percent stacks: 0%). */
function axisFormat(a: AxisRt): string {
  if (a.percent && (!a.m.numFmt || a.m.numFmt.linked)) return '0%';
  if (a.m.numFmt && !a.m.numFmt.linked) return a.m.numFmt.code;
  const s = a.groups[0]?.series[0];
  const source = a.role === 'x' ? s?.x : a.role === 'y' ? s?.y ?? s?.val : s?.val;
  return source?.formatCode ?? a.m.numFmt?.code ?? 'General';
}

/** The label steps a date axis allows, in its base time unit. */
const STEPS: Record<string, number[]> = {
  days: [1, 2, 3, 4, 5, 7, 14, 28, 56, 91, 182, 364],
  months: [1, 2, 3, 4, 6, 12, 24, 60, 120],
  years: [1, 2, 5, 10, 20, 50, 100],
};
/** The serial of 31 Dec 9999, the last date spreadsheets know. */
const LAST_DATE = 2958466;

/**
 * A date axis has one slot per base time unit (c:baseTimeUnit) from the first date to the last, and each point sits
 * in the slot of its date, so uneven dates are spaced by time. Undefined when the dates do not call for it.
 */
function dateSlots(
  env: Env, a: AxisRt, nums: (number | undefined)[], fmt: string,
): { slots: number[]; texts: string[] } | undefined {
  const unit = a.m.baseTimeUnit === 'months' || a.m.baseTimeUnit === 'years' ? a.m.baseTimeUnit : 'days';
  const isDate = (n: number | undefined): n is number => n !== undefined && n >= 0 && n < LAST_DATE;
  const known = nums.filter(isDate).map(Math.floor);
  if (known.length < 2 || known.length < nums.length) return undefined;
  const min = Math.min(...known);
  const p0 = serialToDate(min, env.date1904);
  const index = (serial: number) => {
    if (unit === 'days') return serial - min;
    const p = serialToDate(serial, env.date1904);
    return unit === 'months' ? (p.y - p0.y) * 12 + (p.m - p0.m) : p.y - p0.y;
  };
  const slots = known.map(index);
  const count = Math.max(...slots) + 1;
  if (count > 5000) return undefined;
  const serialOf = (k: number) => {
    if (unit === 'days') return min + k;
    const [y, m] = unit === 'months' ? [p0.y, p0.m + k] : [p0.y + k, 1];
    return dateToSerial(y, m, 1, env.date1904);
  };
  a.steps = STEPS[unit];
  return { slots, texts: Array.from({ length: count }, (_, k) => formatValue(serialOf(k), fmt, env.date1904)) };
}

/** Prepare label texts and styles; categories come from the first series that has them. */
export function prepareAxes(env: Env, chart: Chart, axes: AxisRt[], fontSize: number): void {
  for (const a of axes) {
    a.style = textStyle(env, [a.m.txPr, chart.txPr], { size: fontSize, bold: false });
    if (a.type !== 'cat') {
      a.format = axisFormat(a);
      continue;
    }
    const source = a.groups.flatMap(g => g.series).find(s => s.cat && s.cat.pts.some(t => t !== undefined))?.cat;
    const fmt = a.m.numFmt && !a.m.numFmt.linked ? a.m.numFmt.code : source?.formatCode ?? 'General';
    a.format = fmt;
    const nums = a.m.kind === 'date' ? source?.nums : undefined;
    const dated = nums ? dateSlots(env, a, nums.slice(0, a.count), fmt) : undefined;
    if (dated) {
      a.slots = dated.slots;
      a.texts = dated.texts;
      a.count = dated.texts.length;
      a.levels = [];
      continue;
    }
    a.texts = Array.from({ length: a.count }, (_, i) => {
      const n = source?.nums?.[i];
      if (n !== undefined) return formatValue(n, fmt, env.date1904);
      const t = source?.pts[i];
      return t ?? (source ? '' : String(i + 1));
    });
    a.levels = [];
    if (source && !a.m.noMultiLvlLbl) {
      for (const lvl of source.levels) {
        const spans: { text: string; from: number; to: number }[] = [];
        for (let i = 0; i < Math.min(lvl.length, a.count); i++) {
          const t = lvl[i];
          if (t !== undefined && t !== '') spans.push({ text: t, from: i, to: i });
          else if (spans.length) spans.at(-1)!.to = i;
        }
        if (spans.length) a.levels.push(spans);
      }
    }
  }
}

// ------------------------------------------------------------------------------------------------ positioning

/** A position along an axis, t = 0 at its minimum and 1 at its maximum, in chart space. */
export const at = (a: AxisRt, t: number) => a.start + t * (a.end - a.start);
export const valueAt = (a: AxisRt, v: number) => at(a, fraction(a.scale, v));
/** Category i's centre (between) or tick (midCat) as a fraction of the axis. */
export function catT(a: AxisRt, i: number): number {
  if (a.between) return (i + 0.5) / Math.max(1, a.count);
  return a.count > 1 ? i / (a.count - 1) : 0.5;
}
export const catAt = (a: AxisRt, i: number) => at(a, catT(a, i));
/** The slot point i sits in: its own index, or its date's slot on a date axis. */
export const slotOf = (a: AxisRt, i: number) => a.slots?.[i] ?? i;
/** Where point i of a series sits along a category axis. */
export const pointAt = (a: AxisRt, i: number) => catAt(a, slotOf(a, i));
/** Width of one category slot along the axis (EMU, positive). */
export const slotSize = (a: AxisRt) => Math.abs(a.end - a.start) / Math.max(1, a.between ? a.count : a.count - 1);

/** Set an axis's ends for the inner plot rectangle; vertical axes run upwards unless reversed. */
export function setEnds(a: AxisRt, r: Rect): void {
  const [lo, hi] = a.dir === 'h' ? [r.x, r.x + r.w] : [r.y + r.h, r.y];
  [a.start, a.end] = a.m.reversed ? [hi, lo] : [lo, hi];
}

/** Where axis `a` crosses its partner, as a fraction along the partner. */
export function crossFraction(a: AxisRt): number {
  const o = a.other;
  if (!o) return 0;
  const { crossesAt, crosses } = a.m;
  if (o.type === 'cat') {
    // c:crossesAt counts categories from 1: the axis crosses at that category's edge, or at its tick (midCat)
    if (crossesAt === undefined) return crosses === 'max' ? 1 : 0;
    return Math.max(0, Math.min(1, (crossesAt - 1) / Math.max(1, o.between ? o.count : o.count - 1)));
  }
  const s = o.scale;
  const auto = s.logBase ? s.min : Math.max(s.min, Math.min(s.max, 0));
  const v = crossesAt ?? (crosses === 'max' ? s.max : crosses === 'min' ? s.min : auto);
  return fraction(s, Math.max(Math.min(s.min, s.max), Math.min(Math.max(s.min, s.max), v)));
}

/** The value where bars and areas of a group start: where the category axis crosses the value axis. */
export function baseValue(cat: AxisRt, val: AxisRt): number {
  const s = val.scale;
  if (s.logBase) return s.min;
  const m = cat.m;
  const v = m.crossesAt ?? (m.crosses === 'max' ? s.max : m.crosses === 'min' ? s.min : 0);
  return Math.max(s.min, Math.min(s.max, v));
}

/** Which side the labels go, and the line they hang from (an edge for low and high labels). */
function labelSide(a: AxisRt, r: Rect): void {
  const o = a.other;
  const p = a.m.tickLblPos;
  if (a.dir === 'h') {
    const bottom = r.y + r.h;
    const lowEdge = o?.m.reversed ? r.y : bottom;
    if (p === 'low' || p === 'high') {
      const edge = p === 'low' ? lowEdge : lowEdge === bottom ? r.y : bottom;
      a.labelLine = edge;
      a.side = edge === bottom ? 1 : -1;
    } else {
      a.side = a.m.pos === 't' ? -1 : 1;
      a.labelLine = a.line;
    }
  } else {
    const lowEdge = o?.m.reversed ? r.x + r.w : r.x;
    if (p === 'low' || p === 'high') {
      const edge = p === 'low' ? lowEdge : lowEdge === r.x ? r.x + r.w : r.x;
      a.labelLine = edge;
      a.side = edge === r.x ? -1 : 1;
    } else {
      a.side = a.m.pos === 'r' ? 1 : -1;
      a.labelLine = a.line;
    }
  }
}

function tickOut(a: AxisRt): number {
  return a.m.majorTick === 'out' || a.m.majorTick === 'cross' ? MAJOR_TICK : 0;
}

/**
 * The gap between the axis and its labels: ticks drawn outward plus an offset (scaled by c:lblOffset), smaller below
 * and above an axis, where the line's own leading already separates the text from it.
 */
const labelGap = (a: AxisRt) =>
  tickOut(a) + a.style.size * PT * (a.dir === 'h' ? 0.3 : 0.6) * (a.type === 'cat' ? a.m.lblOffset / 100 : 1);

const blocks = new WeakMap<TextStyle, Map<string, Block>>();
/** A one-line label in an axis's style, cached per style and text. */
export function labelBlock(env: Env, style: TextStyle, text: string, maxW?: number): Block {
  let m = blocks.get(style);
  if (!m) {
    m = new Map();
    blocks.set(style, m);
  }
  const key = maxW ? `${maxW}|${text}` : text;
  let b = m.get(key);
  if (!b) {
    b = block(env, plainBody(style, text), maxW);
    if (m.size < 5000) m.set(key, b);
  }
  return b;
}

/** The tick label texts of a value axis on a scale, in its display units. */
const valueTexts = (env: Env, a: AxisRt, s: Scale) =>
  ticks(s).map(v => formatValue(a.m.dispUnit ? v / a.m.dispUnit : v, a.format, env.date1904));

/**
 * Choose a value axis's scale for its current length. Office makes at most ten major intervals and keeps labels
 * apart: upright axes leave a little more than a line between labels, horizontal ones two ems beside the widest
 * label (fitted to PowerPoint's choices in the corpus thumbnails).
 */
export function scaleAxis(env: Env, a: AxisRt): void {
  if (a.type !== 'val') return;
  // a 3-D chart's value axis is foreshortened by the view: Office picks its units for the shorter axis
  const len = Math.abs(a.end - a.start) * (a.groups.some(g => g.threeD) ? 0.7 : 1);
  const lineH = labelBlock(env, a.style, '0').h || a.style.size * PT;
  const em = a.style.size * PT;
  const fits = (s: Scale) => {
    const n = Math.max(1, (s.max - s.min) / s.major);
    if (n > 10.0001) return false;
    const step = len / n;
    if (a.dir === 'v') return step >= lineH * 1.12;
    if (step < lineH) return false;
    const widest = Math.max(...valueTexts(env, a, s).map(t => labelBlock(env, a.style, t).w));
    return step >= widest + 2 * em;
  };
  const m = a.m;
  const input: ScaleInput = {
    lo: a.hasData ? a.lo : 0, hi: a.hasData ? a.hi : 1, percent: a.percent, fits, tight: a.groups.some(g => g.threeD),
  };
  for (const k of ['min', 'max', 'majorUnit', 'minorUnit', 'logBase'] as const) if (m[k] !== undefined) input[k] = m[k];
  a.scale = autoScale(input);
}

/** Lay out an axis's tick labels for its current position, and how far they reach beyond the label line. */
export function layoutLabels(env: Env, a: AxisRt, r: Rect, chartW: number, chartH: number): void {
  labelSide(a, r);
  a.labels = [];
  a.outer = [];
  a.extent = 0;
  if (a.m.deleted || a.m.tickLblPos === 'none') return;
  const gap = labelGap(a);
  const base = a.labelLine + a.side * gap;
  if (a.type === 'val') {
    const vs = ticks(a.scale);
    const texts = valueTexts(env, a, a.scale);
    const rot = a.style.rot ?? 0;
    vs.forEach((v, i) => {
      const b = labelBlock(env, a.style, texts[i]!);
      const p = valueAt(a, v);
      a.labels.push(a.dir === 'v' ? anchorV(b, base, p, a.side, rot) : anchorH(b, p, base, a.side, rot));
    });
  } else {
    layoutCategoryLabels(env, a, base, chartW, chartH);
  }
  const ext = a.labels.reduce((m, l) => Math.max(m, reach(l, a.dir, a.side)), 0);
  a.extent = ext ? gap + ext : tickOut(a);
  // outer levels of multi-level categories, one row each
  let edge = base + a.side * ext;
  for (const lvl of a.levels) {
    let rowExt = 0;
    const row: PlacedLabel[] = [];
    const gapRow = a.style.size * PT * 0.3;
    const across = edge + a.side * gapRow;
    for (const span of lvl) {
      const b = labelBlock(env, a.style, span.text);
      const p = (at(a, span.from / Math.max(1, a.count)) + at(a, (span.to + 1) / Math.max(1, a.count))) / 2;
      row.push(a.dir === 'h' ? anchorH(b, p, across, a.side, 0) : anchorV(b, across, p, a.side, 0));
      rowExt = Math.max(rowExt, (a.dir === 'h' ? b.h : b.w) + gapRow);
    }
    a.outer.push(...row);
    edge += a.side * rowExt;
    a.extent += rowExt;
  }
}

/** How far a placed label reaches beyond its anchor, across the axis on the label side. */
export function reach(l: PlacedLabel, dir: 'h' | 'v', side: number): number {
  const r = (l.rot * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  let most = 0;
  for (const [x, y] of [[0, 0], [l.b.w, 0], [0, l.b.h], [l.b.w, l.b.h]] as const) {
    const dx = (x - l.px) * c - (y - l.py) * s;
    const dy = (x - l.px) * s + (y - l.py) * c;
    most = Math.max(most, side * (dir === 'h' ? dy : dx));
  }
  return most;
}

/** A label for an upright axis: right-aligned against it on the left (side -1), left-aligned on the right. */
function anchorV(b: Block, x: number, y: number, side: number, rot: number): PlacedLabel {
  return side < 0 ? { b, x, y, px: b.w, py: b.h / 2, rot } : { b, x, y, px: 0, py: b.h / 2, rot };
}

/** A label for a horizontal axis: centred on its position, hanging below (side 1) or above; turned labels end at it. */
function anchorH(b: Block, x: number, y: number, side: number, rot: number): PlacedLabel {
  if (!rot) return { b, x, y, px: b.w / 2, py: side > 0 ? 0 : b.h, rot };
  // turned text: its end (counter-clockwise) or start (clockwise) sits at the tick, its middle line through it
  if (side > 0) return rot < 0 ? { b, x, y, px: b.w, py: b.h / 2, rot } : { b, x, y, px: 0, py: b.h / 2, rot };
  return rot < 0 ? { b, x, y, px: 0, py: b.h / 2, rot } : { b, x, y, px: b.w, py: b.h / 2, rot };
}

/**
 * The labels to measure when deciding how labels fit: all of them, or on long axes the longest texts and a regular
 * sample, so that a file with tens of thousands of categories does not lay out every label to find the widest.
 */
function sampleOf(texts: string[]): number[] {
  const n = texts.length;
  if (n <= 300) return texts.map((_, i) => i);
  const longest = texts.map((t, i) => ({ i, l: t.length })).sort((p, q) => q.l - p.l).slice(0, 40).map(x => x.i);
  const stride = Math.ceil(n / 200);
  const regular = Array.from({ length: Math.ceil(n / stride) }, (_, k) => k * stride);
  return [...new Set([...longest, ...regular])];
}

/**
 * Category labels: horizontal when they fit their slot, else wrapped to it, else turned 45 degrees (Office's
 * automatic rotation), with labels skipped when even that crowds. Upright axes skip labels that would overlap.
 */
function layoutCategoryLabels(env: Env, a: AxisRt, base: number, chartW: number, chartH: number): void {
  const n = a.count;
  if (!n) return;
  const slot = slotSize(a);
  const style = a.style;
  const lineH = labelBlock(env, style, '0').h || style.size * PT;
  const pos = (i: number) => catAt(a, i);
  const show = (skip: number) => Array.from({ length: Math.ceil(n / skip) }, (_, k) => k * skip);
  const sample = sampleOf(a.texts);
  if (a.dir === 'v') {
    const maxW = chartW * 0.3;
    const at = (i: number) => labelBlock(env, style, a.texts[i]!, maxW);
    const tallest = Math.max(lineH, ...sample.map(i => at(i).h));
    const skip = a.m.tickLblSkip ?? Math.max(1, Math.ceil(tallest / Math.max(1, slot) - 1e-9));
    for (const i of show(skip)) a.labels.push(anchorV(at(i), base, pos(i), a.side, style.rot ?? 0));
    return;
  }
  let rot = style.rot ?? 0;
  let wrapAt: number | undefined;
  const at = (i: number) => labelBlock(env, style, a.texts[i]!, wrapAt);
  let skip = a.m.tickLblSkip ?? 1;
  if (a.steps) {
    // a date axis labels every few time units: the major unit, or as many as the widest label needs
    const widestDate = Math.max(0, ...sample.map(i => rotatedSize(at(i), rot).w));
    const need = a.m.majorUnit ?? Math.ceil((widestDate + style.size * PT * 0.6) / Math.max(1, slot) - 1e-9);
    skip = a.m.tickLblSkip ?? a.steps.find(st => st >= need) ?? Math.max(1, Math.ceil(need));
    a.tickEvery = skip;
    for (const i of show(skip)) a.labels.push(anchorH(at(i), pos(i), base, a.side, rot));
    return;
  }
  const auto = style.rot === undefined;
  const widest = Math.max(0, ...sample.map(i => at(i).w));
  // each label gets an n-th of the axis, also when the labels sit on the ticks (midCat)
  const avail = (Math.abs(a.end - a.start) / n) * skip;
  if (auto && widest > avail * 0.98) {
    // wrapped onto lines no wider than the slot, while no word is wider than the slot and they stay short
    const width = Math.max(1, avail * 0.98);
    const words = Math.max(0, ...sample.flatMap(i => a.texts[i]!.split(/\s+/)).map(w => labelBlock(env, style, w).w));
    const lines = (i: number) => Math.round(labelBlock(env, style, a.texts[i]!, width).h / lineH);
    if (words <= width && Math.max(...sample.map(lines)) <= 3) wrapAt = width;
    else rot = -45;
  }
  if (rot && a.m.tickLblSkip === undefined) {
    // turned labels need room across their slope: their line height over the sine of the angle
    const need = lineH / Math.max(0.2, Math.abs(Math.sin((rot * Math.PI) / 180)));
    skip = Math.max(1, Math.ceil(need / Math.max(1, slot) - 1e-9));
  } else if (!rot && a.m.tickLblSkip === undefined) {
    const w = Math.max(0, ...sample.map(i => at(i).w));
    if (w > avail * 1.02) skip = Math.max(1, Math.ceil((w + style.size * PT * 0.3) / Math.max(1, avail) - 1e-9));
  }
  // turned labels longer than a third of the chart's height are cut down to it, as Office does
  const sin = Math.abs(Math.sin((rot * Math.PI) / 180));
  for (const i of show(skip)) {
    let b = at(i);
    if (rot && b.w * sin > chartH / 3) b = labelBlock(env, style, a.texts[i]!, chartH / 3 / sin);
    a.labels.push(anchorH(b, pos(i), base, a.side, rot));
  }
}

/** The axis title: rich text or the default "Axis Title", turned upright for vertical axes unless it says otherwise. */
export function layoutTitle(env: Env, a: AxisRt, chart: Chart, fontSize: number): void {
  const t = a.m.title;
  if (!t || a.m.deleted) return;
  const d = { size: fontSize, bold: true };
  const chain = [t.txPr, chart.txPr];
  const rich = t.rich ? richBody(env, t.rich, chain, d) : undefined;
  const style = textStyle(env, chain, d);
  const body = rich?.body ?? plainBody(style, t.text ?? 'Axis Title');
  const rotSet = rich ? rich.rot : style.rot;
  const vert = rich ? rich.vert : style.vert;
  let rot = rotSet ?? (a.dir === 'v' ? -90 : 0);
  if (vert === 'vert270' || vert === 'wordArtVertRtl') rot = -90;
  else if (vert === 'vert' || vert === 'eaVert') rot = 90;
  a.title = { b: block(env, body), rot };
}

/** Space the title takes across the axis. */
export function titleExtent(a: AxisRt): number {
  if (!a.title) return 0;
  const size = rotatedSize(a.title.b, a.title.rot);
  return (a.dir === 'h' ? size.h : size.w) + 4 * PT;
}

// ---------------------------------------------------------------------------------------------------- drawing

/** A point given by its position along an axis and across it. */
const pt = (a: AxisRt, along: number, across: number): Pt =>
  (a.dir === 'h' ? { x: along, y: across } : { x: across, y: along });

/** Gridlines of an axis across the plot rectangle (minor ones first, so major lines sit on top). */
export function drawGridlines(env: Env, a: AxisRt, r: Rect, out: Primitive[]): void {
  const [from, to] = a.dir === 'h' ? [r.y, r.y + r.h] : [r.x, r.x + r.w];
  const draw = (positions: number[], grid: Axis['majorGrid'], c: ColorDef) => {
    if (!grid) return;
    const stroke = strokeOf(env, grid.spPr, themedLine(env, c));
    if (!stroke || !positions.length) return;
    const prim = pathPrim(positions.flatMap(p => lineCmds(pt(a, p, from), pt(a, p, to))), NONE, stroke);
    if (prim) out.push(prim);
  };
  const minorColor = minorGridColor(env.style);
  if (a.type === 'val') {
    if (a.m.minorGrid) draw(minorPositions(a), a.m.minorGrid, minorColor);
    draw(ticks(a.scale).map(v => valueAt(a, v)), a.m.majorGrid, axisColor(env.style));
  } else {
    // minor gridlines of a category axis run through the categories, between the major ones
    const centres = () => Array.from({ length: a.count }, (_, i) => catAt(a, i));
    if (a.m.minorGrid && a.between) draw(centres(), a.m.minorGrid, minorColor);
    draw(categoryTicks(a), a.m.majorGrid, axisColor(env.style));
  }
}

/** Minor positions of a value axis, leaving out the major ones; on a log axis, the multiples within each decade. */
function minorPositions(a: AxisRt): number[] {
  const s = a.scale;
  const out: number[] = [];
  if (s.logBase) {
    for (const t of ticks(s)) {
      for (let k = 2; k < s.logBase && out.length < 2000; k++) if (t * k < s.max) out.push(valueAt(a, t * k));
    }
    return out;
  }
  const majors = ticks(s);
  return ticks(s, s.minor).filter(v => !majors.some(m => Math.abs(m - v) < s.minor * 1e-6)).map(v => valueAt(a, v));
}

/** Tick positions of a category axis: slot edges between categories, or the categories themselves (midCat). */
function categoryTicks(a: AxisRt): number[] {
  const skip = a.m.tickMarkSkip ?? a.tickEvery ?? 1;
  if (!a.between) return Array.from({ length: a.count }, (_, i) => i).filter(i => i % skip === 0).map(i => catAt(a, i));
  const edges = Array.from({ length: a.count + 1 }, (_, k) => k).filter(k => k % skip === 0 || k === a.count);
  return edges.map(k => at(a, k / Math.max(1, a.count)));
}

/** The axis line and its tick marks. */
export function drawAxisLine(env: Env, a: AxisRt, out: Primitive[]): void {
  if (a.m.deleted) return;
  const stroke = strokeOf(env, a.m.spPr, themedLine(env, axisColor(env.style)));
  if (!stroke) return;
  const cmds = lineCmds(pt(a, a.start, a.line), pt(a, a.end, a.line));
  const tick = (p: number, kind: Axis['majorTick'], len: number) => {
    if (kind === 'none') return;
    // out is towards the labels' side; cross marks run both ways
    const from = kind === 'cross' ? -len : 0;
    const to = kind === 'in' ? -len : len;
    cmds.push(...lineCmds(pt(a, p, a.line + a.side * from), pt(a, p, a.line + a.side * to)));
  };
  if (a.type === 'val') {
    for (const v of ticks(a.scale)) tick(valueAt(a, v), a.m.majorTick, MAJOR_TICK);
    const minor = a.m.minorTick;
    if (minor !== 'none' && !a.scale.logBase) for (const p of minorPositions(a)) tick(p, minor, MINOR_TICK);
  } else {
    for (const p of categoryTicks(a)) tick(p, a.m.majorTick, MAJOR_TICK);
  }
  const prim = pathPrim(cmds, NONE, stroke);
  if (prim) out.push(prim);
}

/** Tick labels and the axis title. */
export function drawAxisText(a: AxisRt, r: Rect, out: Primitive[]): void {
  if (a.m.deleted) return;
  for (const l of [...a.labels, ...a.outer]) out.push(textPrim(l.b, placeAt(l.x, l.y, l.px, l.py, l.rot)));
  if (a.title) {
    const t = a.title;
    const sz = rotatedSize(t.b, t.rot);
    const across = a.labelLine + a.side * (a.extent + 2 * PT) + a.side * (a.dir === 'h' ? sz.h : sz.w) / 2;
    const along = a.dir === 'h' ? r.x + r.w / 2 : r.y + r.h / 2;
    const [x, y] = a.dir === 'h' ? [along, across] : [across, along];
    out.push(textPrim(t.b, placeAt(x, y, t.b.w / 2, t.b.h / 2, t.rot)));
  }
}
