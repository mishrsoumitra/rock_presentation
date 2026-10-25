import { describe, expect, test } from 'vitest';
import { parseXml } from '@pptx/xml';
import { autoColor } from '../src/style.ts';
import { MAX_POINTS, parseChart } from '../src/model.ts';
import { AXES, barChart, C_NS, chartSpace, numCache, strCache } from './helpers.ts';

const parse = (xml: string, office2007 = false) => parseChart(parseXml(xml).root, { office2007 });

describe('parsing c:chartSpace', () => {
  test('series are plotted in c:order order and keep their formatting index', () => {
    const c = parse(chartSpace(barChart([{ name: 'B', vals: [1], idx: 0, order: 1 }, { name: 'A', vals: [2], idx: 1, order: 0 }]) + AXES()));
    const g = c.plotArea.groups[0]!;
    expect(g.series.map(s => [s.name, s.index])).toEqual([['A', 1], ['B', 0]]);
    expect(c.plotArea.axes.map(a => a.kind)).toEqual(['cat', 'val']);
  });

  test('caches: sparse points, formats and literals', () => {
    const xml = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>` +
      `<c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strLit></c:cat>` +
      `<c:val><c:numRef><c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>0.5</c:v></c:pt><c:pt idx="3" formatCode="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>` +
      `</c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`;
    const s = parse(xml).plotArea.groups[0]!.series[0]!;
    expect(s.cat?.pts).toEqual(['a', undefined, 'c']);
    expect(s.val?.pts).toEqual([0.5, undefined, undefined, 1]);
    expect(s.val?.formatCode).toBe('0.0%');
    expect(s.val?.formats[3]).toBe('0');
    expect(s.name).toBe('Series1');
    expect(s.named).toBe(false);
  });

  test('multi-level categories keep their outer levels', () => {
    const xml = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="4"/>` +
      `<c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q1</c:v></c:pt><c:pt idx="3"><c:v>Q2</c:v></c:pt></c:lvl>` +
      `<c:lvl><c:pt idx="0"><c:v>2023</c:v></c:pt><c:pt idx="2"><c:v>2024</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>` +
      `<c:val>${numCache([1, 2, 3, 4])}</c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    const cat = parse(xml).plotArea.groups[0]!.series[0]!.cat!;
    expect(cat.pts).toEqual(['Q1', 'Q2', 'Q1', 'Q2']);
    expect(cat.levels).toEqual([['2023', undefined, '2024', undefined]]);
  });

  test('a hostile point count does not size anything', () => {
    const xml = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>` +
      `<c:val><c:numRef><c:numCache><c:ptCount val="999999999"/><c:pt idx="5"><c:v>1</c:v></c:pt><c:pt idx="2000000000"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>` +
      `</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    expect(parse(xml).plotArea.groups[0]!.series[0]!.val!.pts.length).toBe(MAX_POINTS);
  });

  test('CT_Boolean: an element without val means true', () => {
    const c = parse(`<c:chartSpace ${C_NS}><c:roundedCorners/><c:chart><c:autoTitleDeleted/><c:plotArea><c:pieChart><c:varyColors/></c:pieChart></c:plotArea></c:chart></c:chartSpace>`);
    expect(c.roundedCorners).toBe(true);
    expect(c.autoTitleDeleted).toBe(true);
    expect(c.plotArea.groups[0]!.varyColors).toBe(true);
  });

  test('absent elements: Office 2007 files mean other defaults', () => {
    const xml = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:lineChart/><c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
    expect(parse(xml, true)).toMatchObject({ autoTitleDeleted: false, dispBlanksAs: 'gap' });
    expect(parse(xml, false)).toMatchObject({ autoTitleDeleted: true, dispBlanksAs: 'zero' });
    expect(parse(xml, true).plotArea.axes[0]!.majorTick).toBe('out');
    expect(parse(xml, false).plotArea.axes[0]!.majorTick).toBe('cross');
  });

  test('the style comes from c14:style or its c:style fallback', () => {
    const alt = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14"><c14:style val="118"/></mc:Choice><mc:Fallback><c:style val="18"/></mc:Fallback></mc:AlternateContent>`;
    expect(parse(`<c:chartSpace ${C_NS}>${alt}<c:chart><c:plotArea/></c:chart></c:chartSpace>`).style).toBe(18);
    expect(parse(`<c:chartSpace ${C_NS}><c:chart><c:plotArea/></c:chart></c:chartSpace>`).style).toBe(2);
    expect(parse(`<c:chartSpace ${C_NS}><c:style val="99"/><c:chart><c:plotArea/></c:chart></c:chartSpace>`).style).toBe(2);
  });

  test('Strict namespaces read the same', () => {
    const xml = `<c:chartSpace xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"><c:chart><c:plotArea>` +
      `<c:areaChart><c:grouping val="stacked"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strCache(['S'])}</c:tx><c:val>${numCache([1, 2])}</c:val></c:ser><c:gapWidth val="50%"/></c:areaChart>` +
      `</c:plotArea></c:chart></c:chartSpace>`;
    const g = parse(xml).plotArea.groups[0]!;
    expect(g).toMatchObject({ kind: 'area', grouping: 'stacked', gapWidth: 50 });
    expect(g.series[0]!.name).toBe('S');
  });

  test('manual layouts, legend entries, display units and 3-D charts', () => {
    const xml = `<c:chartSpace ${C_NS}><c:chart><c:view3D><c:rotX val="30"/></c:view3D><c:plotArea>` +
      `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.6"/></c:manualLayout></c:layout>` +
      `<c:bar3DChart><c:barDir val="bar"/><c:grouping val="standard"/></c:bar3DChart>` +
      `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="maxMin"/><c:logBase val="10"/><c:max val="100"/></c:scaling><c:delete/><c:axPos val="b"/><c:crossAx val="1"/><c:crossesAt val="3"/><c:dispUnits><c:builtInUnit val="thousands"/></c:dispUnits></c:valAx>` +
      `</c:plotArea><c:legend><c:legendPos val="b"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry></c:legend></c:chart></c:chartSpace>`;
    const c = parse(xml);
    expect(c.plotArea.layout).toEqual({ target: 'inner', xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor', x: 0.1, y: 0.2, w: 0.5, h: 0.6 });
    expect(c.plotArea.groups[0]).toMatchObject({ kind: 'bar', threeD: true, barDir: 'bar', grouping: 'clustered' });
    expect(c.plotArea.axes[0]).toMatchObject({ reversed: true, logBase: 10, max: 100, deleted: true, crossesAt: 3, dispUnit: 1000 });
    expect(c.legend).toMatchObject({ pos: 'b' });
    expect([...c.legend!.deleted]).toEqual([1]);
    expect(c.rotX).toBe(30);
    expect(c.notes).toEqual(['bar3DChart drawn as 2-D']);
  });
});

describe('automatic colours (ECMA-376 21.2.3.46, Tables 5 and 6)', () => {
  const show = (c: ReturnType<typeof autoColor>) => [c.base.kind === 'scheme' ? c.base.name : '?', ...c.transforms.map(t => `${t.name} ${Math.round(t.val * 1000) / 1000}`)].join(' ');

  test('style 2 cycles the accents, then shades and tints each cycle', () => {
    expect([0, 1, 2, 3, 4, 5].map(k => show(autoColor(2, k, 5)))).toEqual(['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']);
    // seven series: two cycles, the first 23% darker and the second 23% lighter
    expect(show(autoColor(2, 0, 6))).toBe('accent1 shade 0.767');
    expect(show(autoColor(2, 6, 6))).toBe('accent1 tint 0.767');
  });

  test('style 18 uses the same colours as style 2', () => {
    expect(show(autoColor(18, 2, 2))).toBe('accent3');
  });

  test('one-colour styles fade a single accent across the series', () => {
    expect([0, 1, 2].map(k => show(autoColor(3, k, 2)))).toEqual(['accent1 shade 0.65', 'accent1', 'accent1 tint 0.65']);
    expect(show(autoColor(8, 0, 0))).toBe('accent6');
    expect(show(autoColor(12, 0, 0))).toBe('accent2');
  });

  test('grey styles use tints of dk1', () => {
    expect(show(autoColor(1, 0, 5))).toBe('dk1 tint 0.885');
    expect(show(autoColor(1, 1, 5))).toBe('dk1 tint 0.55');
    expect(show(autoColor(41, 0, 5))).toBe('dk1 tint 0.05');
  });
});
