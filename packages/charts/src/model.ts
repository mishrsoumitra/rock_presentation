// packages/charts/src/model.ts  -- a c:chartSpace part as plain data (ECMA-376 Part 1, 21.2), cached values only
//
// Values come from the caches the writing application stored with each reference (c:numCache, c:strCache,
// c:multiLvlStrCache) or from literals; the embedded workbook is never opened. Formatting stays as the XML
// elements (spPr, txPr, rich text): they are resolved when drawn, against the slide's theme and colours.
import type { XmlElement } from '@pptx/xml';
import { attr, child, children, is, lname, textOf } from '@pptx/model';
import { flag, kid, kids, kidsNamed, numOf, numText, valOf } from './xml.ts';

/** Points beyond this index are ignored: a hostile ptCount must not size an array. */
export const MAX_POINTS = 32_000;

export type ChartKind =
  | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'ofPie' | 'scatter' | 'bubble' | 'radar' | 'stock' | 'surface';
export type Grouping = 'standard' | 'clustered' | 'stacked' | 'percentStacked';
export type TickMark = 'out' | 'in' | 'cross' | 'none';
type Mode = 'edge' | 'factor';

/** c:manualLayout, as fractions of the chart area. Absent values keep the automatic position or size. */
export interface Layout {
  target: 'inner' | 'outer';
  xMode: Mode;
  yMode: Mode;
  wMode: Mode;
  hMode: Mode;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

/** A title (chart or axis): its own rich text or a cached cell reference, and its formatting. */
export interface Title {
  rich?: XmlElement;
  text?: string;
  txPr?: XmlElement;
  spPr?: XmlElement;
  layout?: Layout;
  overlay: boolean;
}

/** Numbers by point index (undefined: blank), with the cache's number format and any per-point formats. */
export interface Values { pts: (number | undefined)[]; formatCode: string; formats: (string | undefined)[]; }

/** Category labels by point index; `levels` holds the outer levels of multi-level categories, innermost first. */
export interface Labels {
  pts: (string | undefined)[];
  /** The numbers behind numeric (and date) categories. */
  nums?: (number | undefined)[];
  formatCode: string;
  levels: (string | undefined)[][];
}

export interface Marker { symbol?: string; size?: number; spPr?: XmlElement; }
export interface DataPoint { spPr?: XmlElement; marker?: Marker; explosion?: number; invertIfNegative?: boolean; }

export interface Series {
  /** c:idx: the formatting index that picks the automatic colour. */
  index: number;
  order: number;
  /** The cached name, or Office's "SeriesN" when the series has none. */
  name: string;
  /** Whether the name came from the file (c:tx): only then does it become the automatic title. */
  named: boolean;
  spPr?: XmlElement;
  marker?: Marker;
  points: Map<number, DataPoint>;
  dLbls?: XmlElement;
  cat?: Labels;
  val?: Values;
  x?: Values;
  y?: Values;
  size?: Values;
  smooth: boolean;
  invertIfNegative: boolean;
  explosion: number;
  /** c15:datalabelsRange cache: label text for [CELLRANGE] fields. */
  labelRange?: (string | undefined)[];
}

export interface Group {
  kind: ChartKind;
  threeD: boolean;
  barDir: 'col' | 'bar';
  grouping: Grouping;
  varyColors: boolean;
  /** In c:order order, which is the order Office plots and lists them. */
  series: Series[];
  axIds: number[];
  gapWidth: number;
  overlap: number;
  firstSliceAng: number;
  holeSize: number;
  bubbleScale: number;
  sizeRepresents: 'area' | 'w';
  showNegBubbles: boolean;
  scatterStyle: string;
  radarStyle: string;
  showMarker: boolean;
  dLbls?: XmlElement;
  hiLowLines?: { spPr?: XmlElement };
  dropLines?: { spPr?: XmlElement };
  upDownBars?: { gapWidth: number; up?: XmlElement; down?: XmlElement };
}

export interface Axis {
  kind: 'cat' | 'val' | 'date' | 'ser';
  id: number;
  crossAx: number;
  pos: 'b' | 't' | 'l' | 'r';
  deleted: boolean;
  reversed: boolean;
  logBase?: number;
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  crosses: 'autoZero' | 'min' | 'max';
  crossesAt?: number;
  crossBetween?: 'between' | 'midCat';
  numFmt?: { code: string; linked: boolean };
  majorTick: TickMark;
  minorTick: TickMark;
  tickLblPos: 'nextTo' | 'low' | 'high' | 'none';
  majorGrid?: { spPr?: XmlElement };
  minorGrid?: { spPr?: XmlElement };
  title?: Title;
  spPr?: XmlElement;
  txPr?: XmlElement;
  tickLblSkip?: number;
  tickMarkSkip?: number;
  lblOffset: number;
  noMultiLvlLbl: boolean;
  /** Display units: tick labels are divided by this. */
  dispUnit?: number;
  baseTimeUnit?: string;
  majorTimeUnit?: string;
}

export interface Legend {
  pos: 'r' | 'l' | 't' | 'b' | 'tr';
  layout?: Layout;
  overlay: boolean;
  spPr?: XmlElement;
  txPr?: XmlElement;
  /** c:legendEntry: entries deleted, and entries with text properties of their own, by entry index. */
  deleted: Set<number>;
  entryTxPr: Map<number, XmlElement>;
}

export interface Chart {
  /** c:style 1..48 (c14:style 101..148 is the same style). */
  style: number;
  date1904: boolean;
  roundedCorners: boolean;
  clrMapOvr?: Record<string, string>;
  spPr?: XmlElement;
  txPr?: XmlElement;
  title?: Title;
  autoTitleDeleted: boolean;
  plotArea: { layout?: Layout; groups: Group[]; axes: Axis[]; spPr?: XmlElement };
  legend?: Legend;
  dispBlanksAs: 'gap' | 'zero' | 'span';
  /** c:view3D rotX: how far 3-D pies are tilted. */
  rotX?: number;
  /** Chart types the file uses that are drawn approximately (3-D as 2-D, surface as a placeholder...). */
  notes: string[];
}

/** Files written by Office 2007 mean other defaults by absent elements, as LibreOffice found. */
export interface ParseOptions { office2007: boolean; }

// ------------------------------------------------------------------------------------------------------- caches

/** n blank points. */
const blanks = <T>(n: number): (T | undefined)[] => new Array<T | undefined>(n).fill(undefined);

const pointsOf = (cache: XmlElement | undefined): { idx: number; el: XmlElement }[] => {
  const out: { idx: number; el: XmlElement }[] = [];
  children(cache, 'c:pt').forEach((pt, i) => {
    const idx = Number.parseInt(attr(pt, 'idx') ?? String(i), 10);
    if (Number.isInteger(idx) && idx >= 0 && idx < MAX_POINTS) out.push({ idx, el: pt });
  });
  return out;
};

/** How many points a cache declares, bounded by MAX_POINTS and never less than the points it holds. */
function countOf(cache: XmlElement | undefined, pts: { idx: number }[]): number {
  const declared = Number.parseInt(valOf(cache, 'c:ptCount') ?? '0', 10);
  const held = pts.reduce((m, p) => Math.max(m, p.idx + 1), 0);
  return Math.min(MAX_POINTS, Math.max(held, Number.isFinite(declared) && declared > 0 ? declared : 0));
}

function numCache(cache: XmlElement | undefined): Values | undefined {
  if (!cache) return undefined;
  const pts = pointsOf(cache);
  const formatCode = textOf(child(cache, 'c:formatCode')).trim() || 'General';
  const out: Values = { pts: blanks<number>(countOf(cache, pts)), formatCode, formats: [] };
  for (const { idx, el } of pts) {
    out.pts[idx] = numText(textOf(child(el, 'c:v')));
    const f = attr(el, 'formatCode');
    if (f) out.formats[idx] = f;
  }
  return out;
}

function strCache(cache: XmlElement | undefined): (string | undefined)[] | undefined {
  if (!cache) return undefined;
  const pts = pointsOf(cache);
  const out = blanks<string>(countOf(cache, pts));
  for (const { idx, el } of pts) out[idx] = textOf(child(el, 'c:v'));
  return out;
}

/** c:val, c:xVal, c:yVal, c:bubbleSize: numbers from a reference's cache or a literal. */
function values(el: XmlElement | undefined): Values | undefined {
  if (!el) return undefined;
  const ref = child(el, 'c:numRef');
  if (ref) return numCache(child(ref, 'c:numCache')) ?? { pts: [], formatCode: 'General', formats: [] };
  const lit = child(el, 'c:numLit');
  if (lit) return numCache(lit);
  // text where numbers belong (a category-like xVal): its numbers where they parse
  const s = labels(el);
  if (!s) return undefined;
  return { pts: s.pts.map(t => (t === undefined ? undefined : numText(t))), formatCode: 'General', formats: [] };
}

/** c:cat or c:xVal as labels: strings, numbers (kept for date axes) or multi-level strings. */
function labels(el: XmlElement | undefined): Labels | undefined {
  if (!el) return undefined;
  const multi = child(child(el, 'c:multiLvlStrRef'), 'c:multiLvlStrCache');
  if (multi) {
    const lvls = children(multi, 'c:lvl').map(pointsOf);
    const n = countOf(multi, lvls.flat());
    const levels = lvls.map(pts => {
      const out = blanks<string>(n);
      for (const { idx, el: pt } of pts) out[idx] = textOf(child(pt, 'c:v'));
      return out;
    });
    return { pts: levels[0] ?? blanks<string>(n), formatCode: 'General', levels: levels.slice(1) };
  }
  const strRef = child(el, 'c:strRef');
  const str = strRef ? strCache(child(strRef, 'c:strCache')) ?? [] : strCache(child(el, 'c:strLit'));
  if (str) return { pts: str, formatCode: 'General', levels: [] };
  const numRef = child(el, 'c:numRef');
  const num = numRef ? numCache(child(numRef, 'c:numCache')) : numCache(child(el, 'c:numLit'));
  if (!num) return undefined;
  const pts = num.pts.map(v => (v === undefined ? undefined : String(v)));
  return { pts, nums: num.pts, formatCode: num.formatCode, levels: [] };
}

/** c:tx of a series or title: the cached text. */
function cachedText(tx: XmlElement | undefined): string | undefined {
  if (!tx) return undefined;
  const ref = child(tx, 'c:strRef');
  if (ref) return (strCache(child(ref, 'c:strCache')) ?? []).filter(s => s !== undefined).join(' ');
  const v = child(tx, 'c:v');
  return v ? textOf(v) : undefined;
}

/** A string cache in any namespace (c15:dlblRangeCache holds c:pt elements). */
function strCacheAny(cache: XmlElement): (string | undefined)[] {
  const out: (string | undefined)[] = [];
  for (const pt of children(cache)) {
    if (pt.local !== 'pt') continue;
    const idx = Number.parseInt(attr(pt, 'idx') ?? '-1', 10);
    if (idx >= 0 && idx < MAX_POINTS) out[idx] = textOf(children(pt).find(e => e.local === 'v'));
  }
  return out;
}

// ------------------------------------------------------------------------------------------------------- pieces

function layoutOf(el: XmlElement | undefined): Layout | undefined {
  const m = kid(kid(el, 'c:layout'), 'c:manualLayout');
  if (!m) return undefined;
  const mode = (name: string): Mode => (valOf(m, name) === 'edge' ? 'edge' : 'factor');
  const out: Layout = {
    target: valOf(m, 'c:layoutTarget') === 'inner' ? 'inner' : 'outer',
    xMode: mode('c:xMode'), yMode: mode('c:yMode'), wMode: mode('c:wMode'), hMode: mode('c:hMode'),
  };
  for (const k of ['x', 'y', 'w', 'h'] as const) {
    const v = numOf(m, `c:${k}`);
    if (v !== undefined) out[k] = v;
  }
  return out;
}

const withSpPr = (el: XmlElement): { spPr?: XmlElement } => {
  const spPr = kid(el, 'c:spPr');
  return spPr ? { spPr } : {};
};

function titleOf(el: XmlElement | undefined): Title | undefined {
  if (!el) return undefined;
  const tx = kid(el, 'c:tx');
  const out: Title = { overlay: flag(el, 'c:overlay', false), ...withSpPr(el) };
  const rich = child(tx, 'c:rich');
  const text = rich ? undefined : cachedText(tx);
  if (rich) out.rich = rich;
  if (text !== undefined) out.text = text;
  const txPr = kid(el, 'c:txPr');
  if (txPr) out.txPr = txPr;
  const layout = layoutOf(el);
  if (layout) out.layout = layout;
  return out;
}

function markerOf(el: XmlElement | undefined): Marker | undefined {
  if (!el) return undefined;
  const out: Marker = { ...withSpPr(el) };
  const symbol = valOf(el, 'c:symbol');
  if (symbol) out.symbol = symbol;
  const size = numOf(el, 'c:size');
  if (size !== undefined) out.size = Math.max(2, Math.min(72, size));
  return out;
}

function pointOf(dPt: XmlElement): DataPoint {
  const p: DataPoint = { ...withSpPr(dPt) };
  const m = markerOf(kid(dPt, 'c:marker'));
  if (m) p.marker = m;
  const ex = numOf(dPt, 'c:explosion');
  if (ex !== undefined) p.explosion = ex;
  if (kid(dPt, 'c:invertIfNegative')) p.invertIfNegative = flag(dPt, 'c:invertIfNegative', false);
  return p;
}

function seriesOf(el: XmlElement, position: number): Series {
  const index = Number.parseInt(valOf(el, 'c:idx') ?? String(position), 10);
  const idx = Number.isFinite(index) && index >= 0 ? Math.min(index, 10_000) : position;
  const order = Number.parseInt(valOf(el, 'c:order') ?? String(idx), 10);
  const name = cachedText(kid(el, 'c:tx'));
  const s: Series = {
    index: idx, order: Number.isFinite(order) ? order : position,
    name: name ?? `Series${idx + 1}`, named: name !== undefined,
    points: new Map(), smooth: flag(el, 'c:smooth', false), invertIfNegative: flag(el, 'c:invertIfNegative', false),
    explosion: numOf(el, 'c:explosion') ?? 0, ...withSpPr(el),
  };
  const marker = markerOf(kid(el, 'c:marker'));
  if (marker) s.marker = marker;
  for (const dPt of kidsNamed(el, 'c:dPt')) {
    const i = Number.parseInt(valOf(dPt, 'c:idx') ?? '-1', 10);
    if (i >= 0 && i < MAX_POINTS) s.points.set(i, pointOf(dPt));
  }
  const dLbls = kid(el, 'c:dLbls');
  if (dLbls) s.dLbls = dLbls;
  const cat = labels(kid(el, 'c:cat'));
  if (cat) s.cat = cat;
  const val = values(kid(el, 'c:val'));
  if (val) s.val = val;
  // x values that are text do not place points: Office numbers the points 1, 2, 3... instead
  const xEl = kid(el, 'c:xVal');
  const xText = labels(xEl);
  const x = values(xEl);
  if (x && !(xText && !xText.nums && xText.pts.some(t => t !== undefined && numText(t) === undefined))) s.x = x;
  const y = values(kid(el, 'c:yVal'));
  if (y) s.y = y;
  const size = values(kid(el, 'c:bubbleSize'));
  if (size) s.size = size;
  // Office 2013's "value from cells" labels keep their text in an extension
  for (const ext of children(kid(el, 'c:extLst'), 'c:ext')) {
    for (const c of kids(ext)) {
      const cache = c.local === 'datalabelsRange' ? children(c).find(e => e.local === 'dlblRangeCache') : undefined;
      if (cache) s.labelRange = strCacheAny(cache);
    }
  }
  return s;
}

const KINDS: Record<string, { kind: ChartKind; threeD: boolean }> = {
  'c:barChart': { kind: 'bar', threeD: false }, 'c:bar3DChart': { kind: 'bar', threeD: true },
  'c:lineChart': { kind: 'line', threeD: false }, 'c:line3DChart': { kind: 'line', threeD: true },
  'c:areaChart': { kind: 'area', threeD: false }, 'c:area3DChart': { kind: 'area', threeD: true },
  'c:pieChart': { kind: 'pie', threeD: false }, 'c:pie3DChart': { kind: 'pie', threeD: true },
  'c:doughnutChart': { kind: 'doughnut', threeD: false }, 'c:ofPieChart': { kind: 'ofPie', threeD: false },
  'c:scatterChart': { kind: 'scatter', threeD: false }, 'c:bubbleChart': { kind: 'bubble', threeD: false },
  'c:radarChart': { kind: 'radar', threeD: false }, 'c:stockChart': { kind: 'stock', threeD: false },
  'c:surfaceChart': { kind: 'surface', threeD: false }, 'c:surface3DChart': { kind: 'surface', threeD: true },
};

const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));

function groupOf(el: XmlElement, kind: ChartKind, threeD: boolean, opts: ParseOptions): Group {
  const grouping = valOf(el, 'c:grouping');
  const stacked = grouping === 'stacked' || grouping === 'percentStacked';
  // c:order decides plotting order; ties keep document order
  const series = kidsNamed(el, 'c:ser').map(seriesOf).sort((a, b) => a.order - b.order);
  const g: Group = {
    kind, threeD, series,
    barDir: valOf(el, 'c:barDir') === 'bar' ? 'bar' : 'col',
    // bars side by side are "clustered"; a 3-D chart's "standard" (series one behind another) is drawn the same way
    grouping: stacked ? grouping : kind === 'bar' ? 'clustered' : 'standard',
    varyColors: flag(el, 'c:varyColors', false),
    axIds: kidsNamed(el, 'c:axId').map(a => Number.parseInt(attr(a, 'val') ?? 'NaN', 10)).filter(Number.isFinite),
    gapWidth: clamp(numOf(el, 'c:gapWidth') ?? 150, 0, 500),
    overlap: clamp(numOf(el, 'c:overlap') ?? (kind === 'bar' && stacked ? 100 : 0), -100, 100),
    firstSliceAng: numOf(el, 'c:firstSliceAng') ?? 0,
    holeSize: clamp(numOf(el, 'c:holeSize') ?? (kind === 'doughnut' ? 50 : 10), 1, 90),
    bubbleScale: clamp(numOf(el, 'c:bubbleScale') ?? 100, 0, 300),
    sizeRepresents: valOf(el, 'c:sizeRepresents') === 'w' ? 'w' : 'area',
    showNegBubbles: flag(el, 'c:showNegBubbles', false),
    scatterStyle: valOf(el, 'c:scatterStyle') ?? 'marker',
    radarStyle: valOf(el, 'c:radarStyle') ?? 'standard',
    showMarker: flag(el, 'c:marker', !opts.office2007),
  };
  const dLbls = kid(el, 'c:dLbls');
  if (dLbls) g.dLbls = dLbls;
  const hl = kid(el, 'c:hiLowLines');
  if (hl) g.hiLowLines = withSpPr(hl);
  const dl = kid(el, 'c:dropLines');
  if (dl) g.dropLines = withSpPr(dl);
  const ud = kid(el, 'c:upDownBars');
  if (ud) {
    const up = kid(kid(ud, 'c:upBars'), 'c:spPr');
    const down = kid(kid(ud, 'c:downBars'), 'c:spPr');
    g.upDownBars = { gapWidth: numOf(ud, 'c:gapWidth') ?? 150, ...(up ? { up } : {}), ...(down ? { down } : {}) };
  }
  return g;
}

const BUILT_IN_UNITS: Record<string, number> = {
  hundreds: 1e2, thousands: 1e3, tenThousands: 1e4, hundredThousands: 1e5, millions: 1e6, tenMillions: 1e7,
  hundredMillions: 1e8, billions: 1e9, trillions: 1e12,
};

const TICK_MARKS = new Set(['out', 'in', 'cross', 'none']);
/** A tick mark element: absent gives `absent`, present without a value the schema's cross. */
const tickMark = (el: XmlElement, name: string, absent: TickMark): TickMark => {
  const v = valOf(el, name);
  return v && TICK_MARKS.has(v) ? (v as TickMark) : kid(el, name) ? 'cross' : absent;
};

function axisOf(el: XmlElement, opts: ParseOptions): Axis {
  const kind = is(el, 'c:catAx') ? 'cat' : is(el, 'c:dateAx') ? 'date' : is(el, 'c:serAx') ? 'ser' : 'val';
  const scaling = kid(el, 'c:scaling');
  const pos = valOf(el, 'c:axPos');
  const crosses = valOf(el, 'c:crosses');
  const tlp = valOf(el, 'c:tickLblPos');
  const a: Axis = {
    kind,
    id: Number.parseInt(valOf(el, 'c:axId') ?? 'NaN', 10),
    crossAx: Number.parseInt(valOf(el, 'c:crossAx') ?? 'NaN', 10),
    pos: pos === 'b' || pos === 't' || pos === 'l' || pos === 'r' ? pos : 'b',
    deleted: flag(el, 'c:delete', false),
    reversed: valOf(scaling, 'c:orientation') === 'maxMin',
    crosses: crosses === 'min' || crosses === 'max' ? crosses : 'autoZero',
    majorTick: tickMark(el, 'c:majorTickMark', opts.office2007 ? 'out' : 'cross'),
    minorTick: tickMark(el, 'c:minorTickMark', opts.office2007 ? 'none' : 'cross'),
    tickLblPos: tlp === 'low' || tlp === 'high' || tlp === 'none' ? tlp : 'nextTo',
    lblOffset: clamp(numOf(el, 'c:lblOffset') ?? 100, 0, 1000),
    noMultiLvlLbl: flag(el, 'c:noMultiLvlLbl', false),
    ...withSpPr(el),
  };
  const logBase = numOf(scaling, 'c:logBase');
  if (logBase !== undefined && logBase >= 2 && logBase <= 1000) a.logBase = logBase;
  const set = <K extends keyof Axis>(k: K, v: Axis[K] | undefined) => {
    if (v !== undefined) a[k] = v;
  };
  set('min', numOf(scaling, 'c:min'));
  set('max', numOf(scaling, 'c:max'));
  const positive = (v: number | undefined) => (v !== undefined && v > 0 ? v : undefined);
  set('majorUnit', positive(numOf(el, 'c:majorUnit')));
  set('minorUnit', positive(numOf(el, 'c:minorUnit')));
  set('crossesAt', numOf(el, 'c:crossesAt'));
  const cb = valOf(el, 'c:crossBetween');
  if (cb === 'between' || cb === 'midCat') a.crossBetween = cb;
  const nf = kid(el, 'c:numFmt');
  const linked = attr(nf, 'sourceLinked');
  if (nf) a.numFmt = { code: attr(nf, 'formatCode') ?? 'General', linked: linked === '1' || linked === 'true' };
  const major = kid(el, 'c:majorGridlines');
  if (major) a.majorGrid = withSpPr(major);
  const minor = kid(el, 'c:minorGridlines');
  if (minor) a.minorGrid = withSpPr(minor);
  set('title', titleOf(kid(el, 'c:title')));
  set('txPr', kid(el, 'c:txPr'));
  const skip = numOf(el, 'c:tickLblSkip');
  if (skip !== undefined && skip >= 1) a.tickLblSkip = Math.floor(skip);
  const mskip = numOf(el, 'c:tickMarkSkip');
  if (mskip !== undefined && mskip >= 1) a.tickMarkSkip = Math.floor(mskip);
  const du = kid(el, 'c:dispUnits');
  if (du) set('dispUnit', positive(BUILT_IN_UNITS[valOf(du, 'c:builtInUnit') ?? ''] ?? numOf(du, 'c:custUnit')));
  set('baseTimeUnit', valOf(el, 'c:baseTimeUnit'));
  set('majorTimeUnit', valOf(el, 'c:majorTimeUnit'));
  return a;
}

function legendOf(el: XmlElement | undefined): Legend | undefined {
  if (!el) return undefined;
  const pos = valOf(el, 'c:legendPos');
  const out: Legend = {
    pos: pos === 'l' || pos === 't' || pos === 'b' || pos === 'tr' ? pos : 'r',
    overlay: flag(el, 'c:overlay', false), deleted: new Set(), entryTxPr: new Map(), ...withSpPr(el),
  };
  const layout = layoutOf(el);
  if (layout) out.layout = layout;
  const txPr = kid(el, 'c:txPr');
  if (txPr) out.txPr = txPr;
  for (const e of kidsNamed(el, 'c:legendEntry')) {
    const idx = Number.parseInt(valOf(e, 'c:idx') ?? '-1', 10);
    if (!(idx >= 0)) continue;
    if (flag(e, 'c:delete', false)) out.deleted.add(idx);
    const t = kid(e, 'c:txPr');
    if (t) out.entryTxPr.set(idx, t);
  }
  return out;
}

/** Read a c:chartSpace element. Unknown or broken pieces are left out rather than failing the chart. */
export function parseChart(root: XmlElement, opts: ParseOptions): Chart {
  const chart = kid(root, 'c:chart');
  const plot = kid(chart, 'c:plotArea');
  const notes: string[] = [];
  const groups: Group[] = [];
  const axes: Axis[] = [];
  for (const c of kids(plot)) {
    const k = KINDS[lname(c)];
    if (k) {
      groups.push(groupOf(c, k.kind, k.threeD, opts));
      if (k.kind === 'surface') notes.push(`${c.local} not drawn`);
      else if (k.threeD) notes.push(`${c.local} drawn as 2-D`);
      if (k.kind === 'ofPie') notes.push('ofPieChart drawn as one pie');
    } else if (is(c, 'c:catAx') || is(c, 'c:valAx') || is(c, 'c:dateAx') || is(c, 'c:serAx')) {
      axes.push(axisOf(c, opts));
    }
  }
  const style = Number.parseInt(valOf(root, 'c:style') ?? '2', 10);
  const dba = valOf(chart, 'c:dispBlanksAs');
  const out: Chart = {
    style: style > 100 ? style - 100 : style,
    date1904: flag(root, 'c:date1904', false),
    roundedCorners: flag(root, 'c:roundedCorners', false),
    autoTitleDeleted: flag(chart, 'c:autoTitleDeleted', !opts.office2007),
    plotArea: { groups, axes, ...(plot ? withSpPr(plot) : {}) },
    dispBlanksAs: dba === 'gap' || dba === 'zero' || dba === 'span' ? dba : opts.office2007 ? 'gap' : 'zero',
    notes, ...withSpPr(root),
  };
  if (!(out.style >= 1 && out.style <= 48)) out.style = 2;
  const map = kid(root, 'c:clrMapOvr');
  if (map) out.clrMapOvr = Object.fromEntries(map.attrs.filter(a => a.ns === '').map(a => [a.local, a.value]));
  const txPr = kid(root, 'c:txPr');
  if (txPr) out.txPr = txPr;
  const title = titleOf(kid(chart, 'c:title'));
  if (title) out.title = title;
  const layout = layoutOf(plot);
  if (layout) out.plotArea.layout = layout;
  const legend = legendOf(kid(chart, 'c:legend'));
  if (legend) out.legend = legend;
  const rotX = numOf(kid(chart, 'c:view3D'), 'c:rotX');
  if (rotX !== undefined) out.rotX = rotX;
  return out;
}
