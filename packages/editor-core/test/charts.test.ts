// Charts: inserting one of each kind, reading its data back as a table, and changing that table.
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { parseXml, encodeXmlText } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { chartData, chartPartOf, deleteShapes, duplicateSlide, insertChart, SAMPLE_DATA, setChartData, type ChartData, type ChartKind, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const BOX = { x: 1000000, y: 1000000, w: 6000000, h: 3500000 };
const charts = (ed: Editor) => ed.pkg.partNames().filter(n => n.startsWith('/ppt/charts/') && !n.endsWith('.rels'));
const frame = (ed: Editor, id: number) => ed.shape(SLIDE, id)!;
const xmlOf = (ed: Editor, id: number) => new TextDecoder().decode(OpcPackage.open(ed.save()).getPart(chartPartOf(ed, SLIDE, frame(ed, id))!)!.bytes);

function withChart(kind: ChartKind = 'column', data?: ChartData): { ed: Editor; id: number } {
  const ed = openEditor(deckBytes());
  const c = insertChart(ed, SLIDE, kind, BOX, data);
  ed.run(c.command);
  ed.log.clear();
  return { ed, id: c.id };
}

describe('inserting a chart', () => {
  test('a chart part with the data in it, related from the slide, in a frame of the size asked; undone exactly', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => insertChart(ed, SLIDE, 'column', BOX).command)).toEqual([]);
    expect(charts(ed)).toEqual([]);
    const c = insertChart(ed, SLIDE, 'column', BOX);
    ed.run(c.command);
    expect(charts(ed)).toEqual(['/ppt/charts/chart1.xml']);
    expect(ed.doc.targets(SLIDE, 'chart')).toEqual(['/ppt/charts/chart1.xml']);
    expect(frame(ed, c.id).xfrm).toMatchObject({ x: 1000000, y: 1000000, cx: 6000000, cy: 3500000 });
    expect(frame(ed, c.id).name).toBe(`Chart ${c.id - 1}`);
    const g = resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === c.id && s.layer === 'slide')!.graphic!;
    expect(g.kind).toBe('chart');
    expect(chartData(ed, SLIDE, frame(ed, c.id))).toEqual(SAMPLE_DATA);
    expect(ed.log.undoLabel).toBe('Insert Chart');
  });

  test('each kind writes its own plot: columns, bars, lines, and a pie of the first series', () => {
    const plots: Record<ChartKind, RegExp> = {
      column: /<c:barChart><c:barDir val="col"\/>/, bar: /<c:barChart><c:barDir val="bar"\/>/, line: /<c:lineChart>/, pie: /<c:pieChart>/,
    };
    for (const kind of Object.keys(plots) as ChartKind[]) {
      const { ed, id } = withChart(kind);
      const xml = xmlOf(ed, id);
      expect(xml, kind).toMatch(plots[kind]);
      expect(() => parseXml(xml)).not.toThrow();
      expect(chartData(ed, SLIDE, frame(ed, id))!.series.length, kind).toBe(kind === 'pie' ? 1 : 3);
      expect(xml.includes('<c:catAx>'), kind).toBe(kind !== 'pie');
    }
  });

  test('a saved deck opens with the chart and its data; a chart needs data to be a chart', () => {
    const { ed, id } = withChart('line', { categories: ['Q1', 'Q2 & "Q3"'], series: [{ name: 'Sales <net>', values: [1.5, null] }] });
    const back = openEditor(ed.save());
    expect(chartData(back, SLIDE, back.shape(SLIDE, id)!)).toEqual({ categories: ['Q1', 'Q2 & "Q3"'], series: [{ name: 'Sales <net>', values: [1.5, null] }] });
    expect(new PresentationDocument(OpcPackage.open(ed.save())).targets(SLIDE, 'chart').length).toBe(1);
    expect(() => insertChart(ed, SLIDE, 'column', BOX, { categories: [], series: [] })).toThrow(/at least one/);
  });
});

describe('changing a chart\'s data', () => {
  test('values, names and categories change; the chart reads back as the table given', () => {
    const { ed, id } = withChart();
    const next: ChartData = { categories: ['North', 'South'], series: [{ name: 'This year', values: [10, 20] }, { name: 'Last year', values: [8, null] }, { name: 'Plan', values: [12, 18] }] };
    expect(cycle(ed, () => setChartData(ed, SLIDE, frame(ed, id), next))).toEqual([]);
    ed.run(setChartData(ed, SLIDE, frame(ed, id), next));
    expect(chartData(ed, SLIDE, frame(ed, id))).toEqual(next);
    expect(ed.log.undoLabel).toBe('Chart Data');
    ed.undo();
    expect(chartData(ed, SLIDE, frame(ed, id))).toEqual(SAMPLE_DATA);
  });

  test('more series are added, each in the next accent colour; fewer are removed; rows grow and shrink', () => {
    const { ed, id } = withChart();
    const five: ChartData = { categories: ['a', 'b', 'c', 'd', 'e'], series: [1, 2, 3, 4, 5].map(n => ({ name: `S${n}`, values: [n, n, n, n, n] })) };
    ed.run(setChartData(ed, SLIDE, frame(ed, id), five));
    expect(chartData(ed, SLIDE, frame(ed, id))).toEqual(five);
    const xml = xmlOf(ed, id);
    expect(xml.match(/<c:ser>/g)?.length).toBe(5);
    expect([...xml.matchAll(/<c:idx val="(\d)"\/><c:order val="(\d)"\/>/g)].map(m => `${m[1]}${m[2]}`)).toEqual(['00', '11', '22', '33', '44']);
    expect([...xml.matchAll(/<c:spPr><a:solidFill><a:schemeClr val="(accent\d)"\/>/g)].map(m => m[1])).toEqual(['accent1', 'accent2', 'accent3', 'accent4', 'accent5']);
    const one: ChartData = { categories: ['only'], series: [{ name: 'Solo', values: [7] }] };
    ed.run(setChartData(ed, SLIDE, frame(ed, id), one));
    expect(chartData(ed, SLIDE, frame(ed, id))).toEqual(one);
    expect(() => setChartData(ed, SLIDE, frame(ed, id), { categories: ['x'], series: [] })).toThrow(/at least one/);
  });

  test('a chart that came with a workbook keeps its references; the values it is drawn from change', () => {
    const { ed, id } = withChart();
    // rewrite the first series as PowerPoint writes one: references into a sheet, each with a cache of its values
    const part = chartPartOf(ed, SLIDE, frame(ed, id))!;
    const xml = new TextDecoder().decode(OpcPackage.open(ed.save()).getPart(part)!.bytes)
      .replace(/<c:tx><c:v>Series 1<\/c:v><\/c:tx>/, '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx>')
      .replace(/<c:val><c:numLit>(.*?)<\/c:numLit><\/c:val>/, '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f><c:numCache>$1</c:numCache></c:numRef></c:val>');
    ed.pkg.setPartBytes(part, encodeXmlText(xml));
    ed.changed([part]);
    expect(chartData(ed, SLIDE, frame(ed, id))).toEqual(SAMPLE_DATA);
    const next: ChartData = { ...SAMPLE_DATA, series: SAMPLE_DATA.series.map((s, i) => (i ? s : { name: 'Renamed', values: [9, 8, 7, 6] })) };
    ed.run(setChartData(ed, SLIDE, frame(ed, id), next));
    expect(chartData(ed, SLIDE, frame(ed, id))).toEqual(next);
    const after = xmlOf(ed, id);
    expect(after).toContain('<c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Renamed</c:v></c:pt></c:strCache>');
    expect(after).toContain('<c:f>Sheet1!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>9</c:v></c:pt>');
  });
});

describe('charts and slides', () => {
  test('deleting the chart takes its part with it; a duplicated slide has a chart of its own', () => {
    const { ed, id } = withChart();
    const dup = duplicateSlide(ed, SLIDE);
    ed.run(dup.command);
    expect(charts(ed).length).toBe(2);
    const copy = ed.shape(dup.part!, id)!;
    ed.run(setChartData(ed, dup.part!, copy, { categories: ['x'], series: [{ name: 'copy', values: [1] }] }));
    expect(chartData(ed, SLIDE, frame(ed, id))).toEqual(SAMPLE_DATA);
    ed.run(deleteShapes(ed, SLIDE, [frame(ed, id)]));
    expect(charts(ed).length).toBe(1);
    expect(ed.doc.targets(SLIDE, 'chart')).toEqual([]);
  });
});
