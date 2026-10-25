// packages/charts/src/style.ts  -- the chart style's automatic formatting (ECMA-376 Part 1, 21.2.3.46 ST_Style)
//
// A chart element without its own spPr takes a themed fill, line or effect (the theme's subtle, moderate or intense
// style matrix entry) in a colour that depends on the chart style and, for data points, on the series' formatting
// index (c:idx). The tables below are the standard's Tables 2 to 6; where the standard and Office differ, Office's
// observed output wins (axis and gridline greys are tints of tx1, as LibreOffice also found). PowerPoint draws the
// chart and plot areas of styles 1-32 without a fill, unlike Excel.
import type { XmlElement } from '@pptx/xml';
import type { ColorDef, LineDef, PresentationDocument, ThemeDef } from '@pptx/model';
import { child, effectsIn, fillIn, parseLine } from '@pptx/model';
import {
  mergeLine, resolveEffects, resolveFill, resolveLine, themeEffects, themeFill, themeLine, withPlaceholderColor,
  type Paint, type PaintScope, type ResolvedEffects, type Stroke,
} from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';

/** Everything drawing needs about where the chart lives. */
export interface Env {
  doc: PresentationDocument;
  part: string;
  theme: ThemeDef;
  /** Colours resolve against the slide's colour map (or the chart's c:clrMapOvr) and the chart's theme. */
  scope: PaintScope;
  style: number;
  fonts: FontRegistry;
  date1904: boolean;
  /** The highest series formatting index in the chart: automatic colours fade across 0..maxIndex. */
  maxIndex: number;
}

const color = (name: string, ...mods: [string, number][]): ColorDef =>
  ({ base: { kind: 'scheme', name }, transforms: mods.map(([n, val]) => ({ name: n, val })) });
const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

/** Table 6: the repeating colour patterns. */
const PATTERNS: Record<number, ColorDef[]> = {
  1: [0.885, 0.55, 0.78, 0.925, 0.7, 0.3].map(t => color('dk1', ['tint', t])),
  2: ACCENTS.map(a => color(a)),
  3: ACCENTS.map(a => color(a, ['shade', 0.5])),
  4: [0.05, 0.55, 0.78, 0.15, 0.7, 0.3].map(t => color('dk1', ['tint', t])),
};

/** The column of the style gallery: 0 greys, 1 the accents in turn, 2..7 a single accent. */
const columnOf = (style: number) => (style - 1) % 8;

/** The pattern a style's data points cycle through, or the single accent its series fade across. */
function paletteOf(style: number): ColorDef[] {
  if (style === 41) return PATTERNS[4]!;
  const col = columnOf(style);
  if (col === 0) return PATTERNS[1]!;
  if (col === 1) return PATTERNS[2]!;
  return [color(ACCENTS[col - 2]!)];
}

/**
 * The automatic colour for formatting index `k` when indexes run to `max`: the pattern repeats, and each repetition
 * (for a one-colour fade, each series) is shaded or tinted, darkest first, between -70% and +70%.
 */
export function autoColor(style: number, k: number, max: number): ColorDef {
  const palette = paletteOf(style);
  const n = palette.length;
  const i = Math.max(0, Math.floor(k));
  const base = palette[i % n]!;
  const cycle = Math.floor(i / n);
  const maxCycle = Math.floor(Math.max(i, max) / n);
  const f = Math.round((((cycle + 1) / (maxCycle + 2)) * 1.4 - 0.7) * 100000) / 100000;
  if (f === 0) return base;
  const fade = f < 0 ? { name: 'shade', val: 1 + f } : { name: 'tint', val: 1 - f };
  return { base: base.base, transforms: [...base.transforms, fade] };
}

/** Theme style matrix index for data point fills: 1 subtle, 3 intense (Table 5). */
function themedFill(style: number, threeD: boolean): number {
  if (threeD) return style === 18 || style === 26 || style === 42 ? 3 : 1;
  return (style >= 17 && style <= 32) || style >= 41 ? 3 : 1;
}

/** Theme effect style for filled data points (Table 5): 0 none. */
function themedEffect(style: number): number {
  if (style <= 8 || (style >= 33 && style <= 40)) return 0;
  if (style <= 16) return 1;
  return style <= 24 ? 2 : 3;
}

/** The border of filled data points (Table 5), or null for none. */
function pointBorder(style: number, k: number): ColorDef | null {
  if (style >= 9 && style <= 16) return color('lt1');
  if (style === 33) return color('dk1', ['shade', 0.5]);
  if (style === 34) return PATTERNS[3]![Math.max(0, k) % 6]!;
  if (style >= 35 && style <= 40) return color(ACCENTS[style - 35]!, ['shade', 0.5]);
  return null;
}

/** Width multiplier of lines for data points (Table 5). */
const lineWidthFactor = (style: number) => (style <= 8 ? 3 : style >= 25 && style <= 32 ? 7 : 5);

const PH: ColorDef = { base: { kind: 'scheme', name: 'phClr' }, transforms: [] };

/** The theme's subtle line, as a definition whose colour is phClr. */
const subtleLine = (theme: ThemeDef): LineDef => themeLine(theme, 1) ?? { w: 9525, fill: { kind: 'solid', color: PH } };

/** The subtle line widened for data points: series lines and marker outlines. */
function dataLine(env: Env): LineDef {
  const base = subtleLine(env.theme);
  return { ...base, w: (base.w ?? 9525) * lineWidthFactor(env.style), join: { kind: 'round' } };
}

const inScope = (env: Env, c: ColorDef): PaintScope => withPlaceholderColor(env.scope, { color: c });

/** The spPr elements that apply to a data point, outermost first (the series', then the point's own). */
type SpPrs = (XmlElement | undefined)[];
/** An automatic line and the colour its phClr stands for. */
export interface AutoLine { line: LineDef; color: ColorDef; }

/**
 * An automatic line with the a:ln of each spPr laid over it in turn, piece by piece. The colours of each piece
 * resolve where they were written: the automatic line's phClr is its colour, a file's colours are the file's.
 */
function layered(env: Env, auto: LineDef | undefined, c: ColorDef | undefined, spPrs: SpPrs): Stroke | null {
  let line = auto;
  let scope = c ? inScope(env, c) : env.scope;
  for (const s of spPrs) {
    const own = parseLine(child(s, 'a:ln'));
    if (!own) continue;
    line = mergeLine(own, line);
    if (own.fill) scope = env.scope;
  }
  return line ? resolveLine(line, scope) : null;
}

// ------------------------------------------------------------------------------------------------- resolution

/** A fill from an spPr, or `auto` when the spPr has none. */
export function fillOf(env: Env, spPr: XmlElement | undefined, auto: () => Paint): Paint {
  const f = fillIn(spPr);
  return f ? resolveFill(f, env.scope) : auto();
}

/** The line of an spPr over an automatic line in a colour (or over nothing). */
export function strokeOf(env: Env, spPr: XmlElement | undefined, auto: AutoLine | null): Stroke | null {
  return layered(env, auto?.line, auto?.color, [spPr]);
}

export function effectsOf(env: Env, spPr: XmlElement | undefined, auto = (): ResolvedEffects => ({})): ResolvedEffects {
  const e = effectsIn(spPr);
  return e ? resolveEffects(e, env.scope) : auto();
}

export const solid = (env: Env, c: ColorDef): Paint => resolveFill({ kind: 'solid', color: c }, env.scope);
export const NONE: Paint = { kind: 'none' };

/** How something is filled and outlined. */
export interface BoxStyle { fill: Paint; line: Stroke | null; effects: ResolvedEffects; }

/** A box the chart draws (chart or plot area, legend, title, label) as its spPr says: by default unfilled, unlined. */
export const boxStyle = (env: Env, spPr: XmlElement | undefined, autoFill = (): Paint => NONE): BoxStyle =>
  ({ fill: fillOf(env, spPr, autoFill), line: strokeOf(env, spPr, null), effects: effectsOf(env, spPr) });

/** Formatting of one data point (a bar, slice, area or bubble): fill, border and effects, and its automatic colour. */
export interface PointStyle extends BoxStyle { color: ColorDef; }

/** A filled data point's automatic formatting overlaid with its spPr (series spPr, then the point's own). */
export function filledStyle(env: Env, k: number, max: number, threeD: boolean, spPrs: SpPrs): PointStyle {
  const c = autoColor(env.style, k, max);
  const fill = fillOf(env, spPrs.findLast(s => fillIn(s) !== undefined), () => {
    const def = themeFill(env.theme, themedFill(env.style, threeD));
    return def ? resolveFill(def, inScope(env, c)) : solid(env, c);
  });
  const border = pointBorder(env.style, k);
  const line = layered(env, border ? subtleLine(env.theme) : undefined, border ?? undefined, spPrs);
  const effects = effectsOf(env, spPrs.findLast(s => effectsIn(s) !== undefined), () => {
    const idx = themedEffect(env.style);
    return idx ? resolveEffects(themeEffects(env.theme, idx), inScope(env, c)) : {};
  });
  return { fill, line, effects, color: c };
}

/** A line series' automatic line (Lines for Data Points): the subtle theme line in the series colour, widened. */
export function seriesLine(env: Env, k: number, max: number, spPrs: SpPrs): { line: Stroke | null; color: ColorDef } {
  const c = autoColor(env.style, k, max);
  return { line: layered(env, { ...dataLine(env), cap: 'rnd' }, c, spPrs), color: c };
}

/**
 * Marker formatting: filled and outlined in the series colour unless the marker's spPr says otherwise. The outline
 * is as wide as the style's lines for data points, which is why Office's markers look larger than their size.
 */
export function markerStyle(env: Env, c: ColorDef, spPrs: SpPrs): { fill: Paint; line: Stroke | null } {
  const fill = fillOf(env, spPrs.findLast(s => fillIn(s) !== undefined), () => solid(env, c));
  return { fill, line: layered(env, dataLine(env), c, spPrs) };
}

// ------------------------------------------------------------------------------------------- the other elements

/** Table 2: axes and major gridlines. */
export function axisColor(style: number): ColorDef {
  if (style >= 41) return color('lt1', ['tint', 0.75]);
  return style >= 33 ? color('dk1', ['tint', 0.75]) : color('tx1', ['tint', 0.75]);
}

/** Table 2: minor gridlines. */
export const minorGridColor = (style: number): ColorDef => color('tx1', ['tint', style >= 41 ? 0.9 : 0.5]);

/** Table 2: other lines (high-low lines, drop lines, the borders of up and down bars). */
export function otherLineColor(style: number): ColorDef {
  if (style >= 41) return color('lt1');
  if (style >= 35) return color('dk1', ['shade', 0.25]);
  return style >= 33 ? color('dk1') : color('tx1');
}

/** Default text colour: the chart's text is tx1, light on the dark styles. */
export const textColor = (style: number): ColorDef => (style >= 41 ? color('lt1') : color('tx1'));

/** An automatic line in a colour (axes, gridlines, leader lines): the theme's subtle line. */
export const themedLine = (env: Env, c: ColorDef): AutoLine => ({ line: subtleLine(env.theme), color: c });

/** Table 3: plot area fill (2-D charts); PowerPoint leaves styles 1-32 transparent. */
export function plotAreaFill(env: Env, threeD: boolean): Paint {
  const s = env.style;
  if (threeD || s <= 32) return NONE;
  if (s <= 34) return solid(env, color('dk1', ['tint', 0.2]));
  if (s <= 40) return solid(env, color(ACCENTS[s - 35]!, ['tint', 0.2]));
  return solid(env, color('dk1', ['tint', 0.95]));
}

/** Table 3: chart area fill; only the dark styles fill it in PowerPoint. */
export const chartAreaFill = (env: Env): Paint => (env.style >= 41 ? solid(env, color('dk1')) : NONE);

/** Table 4: the up and down bars of stock charts. */
export function upDownFill(env: Env, up: boolean): Paint {
  const s = env.style;
  const col = columnOf(s);
  const accent = ACCENTS[col - 2];
  if (up) {
    if ((s >= 33 && s <= 40) || s === 42) return solid(env, color('lt1'));
    if (col === 0) return solid(env, color('dk1', ['tint', 0.25]));
    return solid(env, accent ? color(accent, ['tint', 0.25]) : color('dk1', ['tint', 0.05]));
  }
  if (s === 42) return solid(env, color('dk1'));
  if (col === 0) return solid(env, color('dk1', ['tint', 0.85]));
  return solid(env, accent ? color(accent, ['shade', 0.25]) : color('dk1', ['tint', 0.95]));
}
