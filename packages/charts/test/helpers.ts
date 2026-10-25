// Test helpers: fonts, and a small deck whose one slide holds a chart frame with the chart part under test.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide, type ResolvedShape, type SlideContext } from '@pptx/resolve';
import { buildSlide, type PathPrim, type Primitive, type TextPrim } from '@pptx/render';
import { FontRegistry } from '@pptx/text';
import { deckFiles } from '../../resolve/test/deck.ts';
import { zipFiles } from '../../opc/test/fixtures.ts';
import { chartRenderer, drawChart, loadChart } from '../src/index.ts';

const FONTS = join(import.meta.dirname, '../../../assets/fonts');

let registry: FontRegistry | undefined;
/** Every font in assets/fonts, loaded once. */
export function fonts(): FontRegistry {
  if (registry) return registry;
  registry = new FontRegistry();
  for (const dir of readdirSync(FONTS, { withFileTypes: true })) {
    if (!dir.isDirectory()) continue;
    for (const f of readdirSync(join(FONTS, dir.name))) {
      if (f.endsWith('.ttf')) registry.addBytes(new Uint8Array(readFileSync(join(FONTS, dir.name, f))));
    }
  }
  return registry;
}

const D = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
export const C_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** The chart frame: 6096000 x 4064000 EMU (the size PowerPoint 2007 gives a new chart), at the slide's corner. */
export const FRAME = { w: 6096000, h: 4064000 };

const SLIDE = D + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Chart 3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  `<p:xfrm><a:off x="0" y="0"/><a:ext cx="${FRAME.w}" cy="${FRAME.h}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9"/></a:graphicData></a:graphic></p:graphicFrame>' +
  '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A deck with one slide showing `chartXml` (a whole c:chartSpace document); `extra` adds or replaces parts. */
export function chartDeck(chartXml: string, extra: Record<string, string> = {}, chartRels: [string, string, string][] = []): PresentationDocument {
  const base = deckFiles();
  const types = base['[Content_Types].xml']!.replace('</Types>',
    '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>');
  const files: Record<string, string> = {
    ...base,
    '[Content_Types].xml': types,
    'ppt/slides/slide1.xml': SLIDE,
    'ppt/slides/_rels/slide1.xml.rels': D + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId9" Type="${REL}/chart" Target="../charts/chart1.xml"/></Relationships>`,
    'ppt/charts/chart1.xml': chartXml.startsWith('<?xml') ? chartXml : D + chartXml,
    ...(chartRels.length ? {
      'ppt/charts/_rels/chart1.xml.rels': D + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        chartRels.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type.includes('/') ? type : `${REL}/${type}`}" Target="${target}"/>`).join('') + '</Relationships>',
    } : {}),
    ...extra,
  };
  return new PresentationDocument(OpcPackage.open(zipFiles(files)));
}

/** The slide's context and its chart frame. */
export function chartSlide(doc: PresentationDocument): { ctx: SlideContext; frame: ResolvedShape } {
  const slide = resolveSlide(doc, doc.slideParts()[0]!);
  const frame = slide.shapes.find(s => s.graphic?.kind === 'chart');
  if (!frame) throw new Error('no chart frame');
  return { ctx: slide.context, frame };
}

/** Draw a chart document into the frame: the primitives in the frame's own space. */
export function draw(chartXml: string, box = FRAME, extra: Record<string, string> = {}, chartRels: [string, string, string][] = []): Primitive[] {
  const doc = chartDeck(chartXml, extra, chartRels);
  const { ctx } = chartSlide(doc);
  const loaded = loadChart('/ppt/charts/chart1.xml', ctx, fonts());
  if (!loaded) throw new Error('chart did not load');
  return drawChart(loaded.chart, loaded.env, box);
}

/** The same chart through buildSlide and the renderer hook, as the editor draws it. */
export function build(chartXml: string): Primitive[] {
  const doc = chartDeck(chartXml);
  const slide = resolveSlide(doc, doc.slideParts()[0]!);
  const f = fonts();
  const out = buildSlide(slide, { fonts: f, chart: chartRenderer(f) });
  return out.shapes.flatMap(s => s.prims);
}

export const paths = (prims: Primitive[]) => prims.filter((p): p is PathPrim => p.kind === 'path');
export const texts = (prims: Primitive[]) => prims.filter((p): p is TextPrim => p.kind === 'text');
/** The text of a text primitive. */
export const textOf = (p: TextPrim) => p.layout.paragraphText.join('\n');
/** A path primitive's box in chart space (it is only ever translated). */
export const boxOf = (p: PathPrim) => ({ x: p.transform[4], y: p.transform[5], w: p.box.w, h: p.box.h });
/** '#RRGGBB' of a solid fill, or undefined. */
export function hexOf(p: PathPrim): string | undefined {
  if (p.fill.kind !== 'solid') return undefined;
  const h = (v: number) => Math.round(v * 255).toString(16).padStart(2, '0').toUpperCase();
  return `#${h(p.fill.color.r)}${h(p.fill.color.g)}${h(p.fill.color.b)}`;
}

// ------------------------------------------------------------------------------------------------- chart XML

const pts = (vals: (string | number | undefined)[]) => vals.map((v, i) => (v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('');
export const strCache = (vals: (string | undefined)[]) => `<c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="${vals.length}"/>${pts(vals)}</c:strCache></c:strRef>`;
export const numCache = (vals: (number | undefined)[], fmt = 'General') => `<c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${vals.length}"/>${pts(vals)}</c:numCache></c:numRef>`;

export interface SeriesSpec { name?: string; cats?: string[]; vals: (number | undefined)[]; spPr?: string; extra?: string; idx?: number; order?: number; }

export function ser(s: SeriesSpec, i: number, valTag = 'c:val'): string {
  return `<c:ser><c:idx val="${s.idx ?? i}"/><c:order val="${s.order ?? i}"/>` +
    (s.name !== undefined ? `<c:tx>${strCache([s.name])}</c:tx>` : '') + (s.spPr ?? '') + (s.extra ?? '') +
    (s.cats ? `<c:cat>${strCache(s.cats)}</c:cat>` : '') + `<${valTag}>${numCache(s.vals)}</${valTag}></c:ser>`;
}

export const AXES = (cat = 1, val = 2, more = { cat: '', val: '' }) =>
  `<c:catAx><c:axId val="${cat}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${val}"/><c:crosses val="autoZero"/>${more.cat}</c:catAx>` +
  `<c:valAx><c:axId val="${val}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cat}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>${more.val}</c:valAx>`;

/** A whole chart space around a plot area's content, with a legend unless `legend` is ''. */
export function chartSpace(plot: string, opts: { style?: number; legend?: string; title?: string; txPr?: string; autoTitleDeleted?: boolean; extra?: string } = {}): string {
  return `<c:chartSpace ${C_NS}><c:date1904 val="0"/><c:roundedCorners val="0"/><c:style val="${opts.style ?? 2}"/>${opts.extra ?? ''}<c:chart>` +
    (opts.title ?? '') + `<c:autoTitleDeleted val="${opts.autoTitleDeleted === false ? 0 : 1}"/>` +
    `<c:plotArea><c:layout/>${plot}</c:plotArea>` + (opts.legend ?? '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>') +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    (opts.txPr ?? '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>') + '</c:chartSpace>';
}

/** A bar chart group. */
export function barChart(series: SeriesSpec[], opts: { dir?: 'col' | 'bar'; grouping?: string; gap?: number; overlap?: number; vary?: boolean; dLbls?: string } = {}): string {
  return `<c:barChart><c:barDir val="${opts.dir ?? 'col'}"/><c:grouping val="${opts.grouping ?? 'clustered'}"/><c:varyColors val="${opts.vary ? 1 : 0}"/>` +
    series.map((s, i) => ser(s, i)).join('') + (opts.dLbls ?? '') +
    `<c:gapWidth val="${opts.gap ?? 150}"/>${opts.overlap !== undefined ? `<c:overlap val="${opts.overlap}"/>` : ''}<c:axId val="1"/><c:axId val="2"/></c:barChart>`;
}
