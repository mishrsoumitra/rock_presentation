// packages/editor-core/src/charts.ts  -- charts: inserting one, reading its data, and changing its data
import { encodeXmlText, type XmlElement } from '@pptx/xml';
import { attrNs, child, children, clone, createElement, relType, setAttr, setText, textOf, type ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { freeName } from './parts.ts';
import { topSlot, type Rect } from './shapes.ts';
import { element, esc, insertAt, nextShapeId } from './xml.ts';

const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const C_NS_STRICT = 'http://purl.oclc.org/ooxml/drawingml/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const A_NS_STRICT = 'http://purl.oclc.org/ooxml/drawingml/main';

/** The kinds of chart that can be inserted. */
export type ChartKind = 'column' | 'bar' | 'line' | 'pie';
/** A chart's data as a table: categories down the side, a series per column. A blank value is null. */
export interface ChartData { categories: string[]; series: { name: string; values: (number | null)[] }[]; }

/** The data PowerPoint puts in a new chart, so there is something to see and change. */
export const SAMPLE_DATA: ChartData = {
  categories: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
  series: [
    { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', values: [2, 2, 3, 5] },
  ],
};

const strLit = (values: string[]) => `<c:strLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strLit>`;
const numLit = (values: (number | null)[]) =>
  `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => (v === null || !Number.isFinite(v) ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numLit>`;

function seriesXml(kind: ChartKind, data: ChartData): string {
  return data.series.map((s, i) => {
    const marker = kind === 'line' ? '<c:marker><c:symbol val="none"/></c:marker>' : '';
    const fill = kind === 'pie' ? '' : `<c:spPr>${kind === 'line' ? `<a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="accent${(i % 6) + 1}"/></a:solidFill><a:round/></a:ln>` : `<a:solidFill><a:schemeClr val="accent${(i % 6) + 1}"/></a:solidFill>`}</c:spPr>`;
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${esc(s.name)}</c:v></c:tx>${fill}${marker}`
      + `<c:cat>${strLit(data.categories)}</c:cat><c:val>${numLit(data.categories.map((_, j) => s.values[j] ?? null))}</c:val>${kind === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
  }).join('');
}

/** A chart part as PowerPoint writes a simple one: a plot of one kind, its axes, and a legend at the bottom. */
function chartXml(ed: Editor, kind: ChartKind, data: ChartData): string {
  const ser = seriesXml(kind, kind === 'pie' ? { ...data, series: data.series.slice(0, 1) } : data);
  const axes = '<c:axId val="111111111"/><c:axId val="222222222"/>';
  const plot = kind === 'pie' ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
    : kind === 'line' ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/>${axes}</c:lineChart>`
      : `<c:barChart><c:barDir val="${kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="219"/><c:overlap val="-27"/>${axes}</c:barChart>`;
  const line = '<c:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></c:spPr>';
  const cat = kind === 'bar' ? 'l' : 'b', val = kind === 'bar' ? 'b' : 'l';
  const axesXml = kind === 'pie' ? ''
    : `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${cat}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${line}<c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
      + `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${val}"/><c:majorGridlines>${line}</c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111111111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<c:chartSpace xmlns:c="${ed.strict ? C_NS_STRICT : C_NS}" xmlns:a="${ed.strict ? A_NS_STRICT : A_NS}"><c:roundedCorners val="0"/>`
    + `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}${axesXml}</c:plotArea>`
    + '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
    + '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>';
}

/**
 * Insert a chart: a chart part with the data in it, related from the slide, in a graphic frame. The data lives in
 * the chart itself, not in an embedded workbook, so PowerPoint draws it but has no sheet to open for editing.
 */
export function insertChart(ed: Editor, part: string, kind: ChartKind, box: Rect, data: ChartData = SAMPLE_DATA): { command: Command; id: number } {
  if (!data.series.length || !data.categories.length) throw new Error('a chart needs at least one series and one category');
  const tree = ed.spTree(part);
  const id = nextShapeId(ed.doc.root(part));
  const chart = freeName(ed.pkg, '/ppt/charts/chart#.xml');
  const r = Math.round;
  const command = ed.edit({
    label: 'Insert Chart', parts: [part, chart], structural: true, scope: () => [tree],
    apply: () => {
      ed.pkg.addPart(chart, CHART_TYPE, encodeXmlText(chartXml(ed, kind, data)));
      const rId = ed.pkg.addRelationship(part, relType('chart', ed.strict), chart);
      const xml = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id - 1}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
        + `<p:xfrm><a:off x="${r(box.x)}" y="${r(box.y)}"/><a:ext cx="${r(box.w)}" cy="${r(box.h)}"/></p:xfrm>`
        + `<a:graphic><a:graphicData uri="${ed.strict ? C_NS_STRICT : C_NS}"><c:chart xmlns:c="${ed.strict ? C_NS_STRICT : C_NS}" r:id="${rId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
      insertAt(tree, topSlot(tree), element(tree, xml, ed.strict));
    },
  });
  return { command, id };
}

/** The chart part a graphic frame shows, when it shows a chart. */
export function chartPartOf(ed: Editor, part: string, frame: ShapeView | XmlElement): string | undefined {
  const el = 'el' in frame ? frame.el : frame;
  const ref = child(child(child(el, 'a:graphic'), 'a:graphicData'), 'c:chart');
  const rId = ref ? attrNs(ref, 'r', 'id') : undefined;
  return rId ? ed.doc.target(part, rId) : undefined;
}

/** The series of a chart: those of every chart kind in its plot area, in file order. */
function seriesOf(root: XmlElement): XmlElement[] {
  const plot = child(child(root, 'c:chart'), 'c:plotArea');
  return plot ? plot.children.flatMap(g => (g.kind === 'element' && g.local.endsWith('Chart') ? children(g, 'c:ser') : [])) : [];
}

/** Where a c:tx, c:cat or c:val keeps its values: a reference's cache, or a literal. */
function store(el: XmlElement | undefined): XmlElement | undefined {
  if (!el) return undefined;
  for (const ref of ['c:strRef', 'c:numRef', 'c:multiLvlStrRef']) {
    const cache = child(child(el, ref), ref === 'c:numRef' ? 'c:numCache' : ref === 'c:strRef' ? 'c:strCache' : 'c:multiLvlStrCache');
    if (cache) return cache;
  }
  return child(el, 'c:strLit') ?? child(el, 'c:numLit');
}

function points(cache: XmlElement | undefined): string[] {
  if (!cache) return [];
  const pts = children(cache, 'c:pt');
  const declared = Number(child(cache, 'c:ptCount')?.attrs.find(a => a.local === 'val')?.value ?? 0);
  const n = Math.max(declared || 0, ...pts.map(p => Number(p.attrs.find(a => a.local === 'idx')?.value ?? 0) + 1), 0);
  const out = Array.from({ length: n }, () => '');
  for (const p of pts) out[Number(p.attrs.find(a => a.local === 'idx')?.value ?? 0)] = textOf(child(p, 'c:v'));
  return out;
}

/** Write values into a cache or literal: its count, and a point for each value that is not blank. */
function setPoints(cache: XmlElement, values: string[]): void {
  cache.children = cache.children.filter(c => !(c.kind === 'element' && (c.local === 'pt' || c.local === 'ptCount')));
  const count = createElement(cache, 'c:ptCount', { val: values.length });
  const at = cache.children.findIndex(c => c.kind === 'element' && c.local === 'formatCode');
  cache.children.splice(at + 1, 0, count);
  values.forEach((v, i) => {
    if (v === '') return;
    const pt = createElement(cache, 'c:pt', { idx: i });
    const val = createElement(pt, 'c:v');
    setText(val, v);
    pt.children.push(val);
    pt.selfClosing = false;
    cache.children.push(pt);
  });
  // an extension list, where a cache has one, stays last
  const ext = cache.children.filter(c => c.kind === 'element' && c.local === 'extLst');
  cache.children = [...cache.children.filter(c => !ext.includes(c)), ...ext];
  cache.selfClosing = false;
}

/** A chart's data as a table, from the values the file keeps with it. Undefined when the frame shows no chart. */
export function chartData(ed: Editor, part: string, frame: ShapeView | XmlElement): ChartData | undefined {
  const chart = chartPartOf(ed, part, frame);
  if (!chart || !ed.pkg.getPart(chart)) return undefined;
  const ser = seriesOf(ed.doc.root(chart));
  const categories = points(store(child(ser[0], 'c:cat')));
  const series = ser.map((s, i) => {
    const tx = child(s, 'c:tx');
    const name = tx ? (child(tx, 'c:v') ? textOf(child(tx, 'c:v')) : points(store(tx)).join(' ')) : '';
    return { name: name || `Series ${i + 1}`, values: points(store(child(s, 'c:val'))).map(v => (v === '' || !Number.isFinite(Number(v)) ? null : Number(v))) };
  });
  const rows = Math.max(categories.length, ...series.map(s => s.values.length), 0);
  return {
    categories: Array.from({ length: rows }, (_, i) => categories[i] || `${i + 1}`),
    series: series.map(s => ({ name: s.name, values: Array.from({ length: rows }, (_, i) => s.values[i] ?? null) })),
  };
}

/**
 * Change a chart's data: its categories, and each series' name and values. Series are added (looking like the
 * last one) or removed to match. A chart that came with a workbook keeps its references to it; what changes here is
 * the copy of the values the chart is drawn from, which is all PowerPoint shows until that workbook is opened.
 */
export function setChartData(ed: Editor, part: string, frame: ShapeView | XmlElement, data: ChartData): Command {
  const chart = chartPartOf(ed, part, frame);
  if (!chart) throw new Error('this shape is not a chart');
  if (!data.series.length || !data.categories.length) throw new Error('a chart needs at least one series and one category');
  const root = ed.doc.root(chart);
  return ed.edit({
    label: 'Chart Data', parts: [part, chart], scope: () => [root],
    apply: () => {
      let ser = seriesOf(root);
      if (!ser.length) throw new Error('this chart has no series to change');
      // as many series as columns: extra ones go, missing ones are made like the last
      for (const extra of ser.slice(data.series.length)) extra.parent!.children.splice(extra.parent!.children.indexOf(extra), 1);
      const last = ser[Math.min(ser.length, data.series.length) - 1]!;
      for (let i = ser.length; i < data.series.length; i++) {
        const copy = clone(last, last.parent!);
        for (const name of ['c:idx', 'c:order']) { const e = child(copy, name); if (e) setAttr(e, 'val', i); }
        // a new series takes the next accent colour rather than repeating the last one's
        const walk = (e: XmlElement) => { if (e.local === 'schemeClr' && /^accent\d$/.test(e.attrs.find(a => a.local === 'val')?.value ?? '')) setAttr(e, 'val', `accent${(i % 6) + 1}`); for (const c of e.children) if (c.kind === 'element') walk(c); };
        const spPr = child(copy, 'c:spPr');
        if (spPr) walk(spPr);
        const group = last.parent!;
        const after = children(group, 'c:ser').at(-1)!;
        group.children.splice(group.children.indexOf(after) + 1, 0, copy);
      }
      ser = seriesOf(root);
      ser.forEach((s, i) => {
        const d = data.series[i]!;
        const tx = child(s, 'c:tx');
        const v = tx ? child(tx, 'c:v') : undefined;
        if (v) setText(v, d.name);
        else if (tx && store(tx)) setPoints(store(tx)!, [d.name]);
        const cat = store(child(s, 'c:cat'));
        if (cat) setPoints(cat, data.categories);
        const val = store(child(s, 'c:val'));
        if (val) setPoints(val, data.categories.map((_, j) => { const n = d.values[j]; return n === null || n === undefined || !Number.isFinite(n) ? '' : String(n); }));
      });
    },
  });
}
