// packages/charts/src/index.ts  -- charts as render primitives (DrawingML charts, ECMA-376 Part 1, 21.2)
import type { XmlElement } from '@pptx/xml';
import { childElements } from '@pptx/xml';
import { is, textOf, type PresentationDocument, type ThemeDef } from '@pptx/model';
import type { Primitive } from '@pptx/render';
import type { SlideContext } from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';
import { drawChart } from './chart.ts';
import { parseChart, type Chart } from './model.ts';
import type { Env } from './style.ts';

export { parseChart, type Chart, type Group, type Series, type Axis } from './model.ts';
export { formatValue, formatGeneral, isDateFormat, serialToDate } from './numfmt.ts';
export { autoScale, ticks, type Scale, type ScaleInput } from './scale.ts';
export { autoColor } from './style.ts';
export { drawChart } from './chart.ts';

/** Draws a chart part into a box (EMU) for buildSlide's `chart` option. */
export type ChartRenderer =
  (chartPart: string, box: { w: number; h: number }, ctx: SlideContext) => Primitive[] | undefined;

const office2007 = new WeakMap<PresentationDocument, boolean>();

/** Whether the deck was written by Office 2007, whose charts leave out elements with other defaults. */
function isOffice2007(doc: PresentationDocument): boolean {
  let v = office2007.get(doc);
  if (v !== undefined) return v;
  v = false;
  try {
    const part = doc.targets('/', 'extended-properties')[0];
    if (part) {
      const props = childElements(doc.root(part));
      const text = (name: string) => textOf(props.find(e => e.local === name));
      v = text('Application').startsWith('Microsoft') && text('AppVersion').startsWith('12.');
    }
  } catch {
    v = false;
  }
  office2007.set(doc, v);
  return v;
}

/** The slide's theme with a chart's c:themeOverride (colours, fonts, formats) laid over it. */
function themeFor(ctx: SlideContext, chartPart: string): ThemeDef {
  const override = ctx.doc.targets(chartPart, 'themeOverride')[0];
  if (!override) return ctx.theme;
  try {
    const o = ctx.doc.theme(override);
    return {
      ...ctx.theme,
      ...(o.colors.size ? { colors: o.colors } : {}),
      ...(o.minor.latin || o.major.latin ? { major: o.major, minor: o.minor } : {}),
      ...(o.fillStyles.length ? { fillStyles: o.fillStyles } : {}),
      ...(o.lineStyles.length ? { lineStyles: o.lineStyles } : {}),
      ...(o.effectStyles.length ? { effectStyles: o.effectStyles } : {}),
      ...(o.bgFillStyles.length ? { bgFillStyles: o.bgFillStyles } : {}),
    };
  } catch {
    return ctx.theme;
  }
}

/** Parse a chart part and set up what drawing it needs; undefined when the part is not a chart. */
export function loadChart(
  chartPart: string, ctx: SlideContext, fonts: FontRegistry,
): { chart: Chart; env: Env } | undefined {
  let root: XmlElement;
  try {
    root = ctx.doc.root(chartPart);
  } catch {
    return undefined;
  }
  if (!is(root, 'c:chartSpace')) return undefined;
  const chart = parseChart(root, { office2007: isOffice2007(ctx.doc) });
  const theme = themeFor(ctx, chartPart);
  const clrMap = chart.clrMapOvr && Object.keys(chart.clrMapOvr).length ? chart.clrMapOvr : ctx.clrMap;
  const maxIndex = Math.max(0, ...chart.plotArea.groups.flatMap(g => g.series.map(s => s.index)));
  const scope = { doc: ctx.doc, part: chartPart, colors: { theme: theme.colors, clrMap } };
  const { style, date1904 } = chart;
  const env: Env = { doc: ctx.doc, part: chartPart, theme, scope, style, fonts, date1904, maxIndex };
  return { chart, env };
}

/**
 * The chart hook for buildSlide: draws a chart part into a box (EMU). It returns undefined, for the slide's own
 * placeholder, only when the part cannot be read as a chart, and never throws.
 */
export function chartRenderer(fonts: FontRegistry): ChartRenderer {
  return (chartPart, box, ctx) => {
    let loaded: ReturnType<typeof loadChart>;
    try {
      loaded = loadChart(chartPart, ctx, fonts);
    } catch {
      return undefined;
    }
    if (!loaded) return undefined;
    if (!(box.w > 0) || !(box.h > 0)) return [];
    try {
      return drawChart(loaded.chart, loaded.env, box);
    } catch {
      // content this renderer cannot cope with draws as a labelled box rather than failing the slide
      return [{ kind: 'missing', transform: [1, 0, 0, 1, 0, 0], box, label: 'Chart' }];
    }
  };
}
