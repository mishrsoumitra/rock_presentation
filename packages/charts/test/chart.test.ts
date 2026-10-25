import { describe, expect, test } from 'vitest';
import type { PathPrim } from '@pptx/render';
import { chartRenderer } from '../src/index.ts';
import {
  AXES, barChart, boxOf, build, C_NS, chartDeck, chartSlide, chartSpace, draw, fonts, FRAME, hexOf, numCache, paths, ser, strCache, texts, textOf,
} from './helpers.ts';

// the test deck's theme: Office 2013 colours
const ACCENT = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const CATS = ['Q1', 'Q2', 'Q3', 'Q4'];
const S3 = [
  { name: 'North', cats: CATS, vals: [4.3, 2.5, 3.5, 4.5] },
  { name: 'South', cats: CATS, vals: [2.4, 4.4, 1.8, 2.8] },
  { name: 'West', cats: CATS, vals: [2, 2, 3, 5] },
];
/** Bars (or areas, slices) in a colour, leaving out the legend's small square keys. */
const bars = (prims: ReturnType<typeof draw>, color: string) =>
  paths(prims).filter(p => hexOf(p) === color).map(boxOf).filter(b => !(Math.abs(b.w - b.h) < 2 && b.w < 150000));
const label = (prims: ReturnType<typeof draw>, text: string) => texts(prims).find(t => textOf(t) === text);
const finite = (prims: ReturnType<typeof draw>) =>
  prims.every(p => p.transform.every(Number.isFinite) && (p.kind !== 'path' || p.paths.every(s => s.commands.every(c => Object.values(c).every(v => typeof v !== 'number' || Number.isFinite(v))))));

describe('bar charts', () => {
  test('clustered columns: series side by side, gap width between categories, heights to scale', () => {
    const prims = draw(chartSpace(barChart(S3) + AXES()));
    const [a, b, c] = ACCENT.slice(0, 3).map(col => bars(prims, col));
    expect([a!.length, b!.length, c!.length]).toEqual([4, 4, 4]);
    const w = a![0]!.w;
    // three bars and a gap of 150% of a bar fill each category's slot
    const slot = a![1]!.x - a![0]!.x;
    expect(slot).toBeCloseTo(w * 4.5, -2);
    expect(b![0]!.x).toBeCloseTo(a![0]!.x + w, -1);
    expect(c![0]!.x).toBeCloseTo(a![0]!.x + 2 * w, -1);
    // all bars stand on the same baseline and their heights follow the values
    const base = a![0]!.y + a![0]!.h;
    for (const r of [...a!, ...b!, ...c!]) expect(r.y + r.h).toBeCloseTo(base, -1);
    expect(a![0]!.h / c![0]!.h).toBeCloseTo(4.3 / 2, 2);
    expect(finite(prims)).toBe(true);
  });

  test('overlap 100 puts the series on top of each other; stacked bars sit on the ones below', () => {
    const over = draw(chartSpace(barChart(S3, { overlap: 100 }) + AXES()));
    expect(bars(over, ACCENT[1]!)[0]!.x).toBeCloseTo(bars(over, ACCENT[0]!)[0]!.x, -1);
    const st = draw(chartSpace(barChart(S3, { grouping: 'stacked' }) + AXES()));
    const [a, b, c] = ACCENT.slice(0, 3).map(col => bars(st, col));
    expect(b![0]!.y + b![0]!.h).toBeCloseTo(a![0]!.y, -1);
    expect(c![0]!.y + c![0]!.h).toBeCloseTo(b![0]!.y, -1);
    expect(b![0]!.x).toBeCloseTo(a![0]!.x, -1);
  });

  test('percent stacks fill the axis, which is labelled in percent', () => {
    const prims = draw(chartSpace(barChart(S3, { grouping: 'percentStacked' }) + AXES()));
    const [a, , c] = ACCENT.slice(0, 3).map(col => bars(prims, col));
    const top = c![2]!.y;
    const bottom = a![2]!.y + a![2]!.h;
    const hundred = label(prims, '100%')!;
    const zero = label(prims, '0%')!;
    expect(hundred && zero).toBeTruthy();
    // the stack runs from the 0% tick to the 100% tick
    expect(Math.abs(top - (hundred.transform[5] + hundred.layout.height / 2))).toBeLessThan(FRAME.h * 0.01);
    expect(Math.abs(bottom - (zero.transform[5] + zero.layout.height / 2))).toBeLessThan(FRAME.h * 0.01);
  });

  test('horizontal bars: the first category at the bottom, and a legend that reads bottom series last', () => {
    const prims = draw(chartSpace(barChart(S3, { dir: 'bar' }) + AXES()));
    const a = bars(prims, ACCENT[0]!);
    expect(a[0]!.y).toBeGreaterThan(a[3]!.y);
    expect(a[0]!.w / a[2]!.w).toBeCloseTo(4.3 / 3.5, 2);
    const legend = ['North', 'South', 'West'].map(n => label(prims, n)!.transform[5]);
    expect(legend[2]!).toBeLessThan(legend[0]!);
  });

  test('negative values grow down from zero, and invertIfNegative draws them white', () => {
    const prims = draw(chartSpace(barChart([{ name: 'S', cats: CATS, vals: [3, -2, 1, -1], extra: '<c:invertIfNegative val="1"/>' }]) + AXES()));
    const pos = bars(prims, ACCENT[0]!);
    const white = paths(prims).filter(p => hexOf(p) === '#FFFFFF' && p.stroke).map(boxOf);
    expect(pos).toHaveLength(2);
    expect(white).toHaveLength(2);
    expect(white[0]!.y).toBeCloseTo(pos[0]!.y + pos[0]!.h, -1);
  });

  test('varyColors with one series colours each bar and lists the categories in the legend', () => {
    const prims = draw(chartSpace(barChart([{ name: 'S', cats: CATS, vals: [1, 2, 3, 4] }], { vary: true }) + AXES()));
    expect(ACCENT.slice(0, 4).map(c => bars(prims, c).filter(b => b.w > 100000).length)).toEqual([1, 1, 1, 1]);
    expect(CATS.every(c => label(prims, c))).toBe(true);
  });
});

describe('axes', () => {
  test('value labels use the source format; crowded category labels turn', () => {
    const many = Array.from({ length: 12 }, (_, i) => `Category number ${i + 1}`);
    const prims = draw(chartSpace(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser({ name: 'S', cats: many, vals: many.map((_, i) => i / 10) }, 0).replace('General', '0%')}<c:axId val="1"/><c:axId val="2"/></c:barChart>` + AXES()));
    expect(label(prims, '100%')).toBeTruthy();
    const cat = label(prims, 'Category number 1')!;
    // turned 45 degrees counter-clockwise: the matrix's b is -sin(45)
    expect(cat.transform[1]).toBeCloseTo(-Math.SQRT1_2, 3);
  });

  test('a reversed value axis puts its maximum at the bottom and the categories on top', () => {
    const prims = draw(chartSpace(barChart(S3) + AXES(1, 2).replace('<c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>', '<c:orientation val="maxMin"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>')));
    const zero = label(prims, '0')!;
    const six = label(prims, '6')!;
    expect(six.transform[5]).toBeGreaterThan(zero.transform[5]);
    const a = bars(prims, ACCENT[0]!);
    // bars hang down from the zero line
    expect(Math.abs(a[0]!.y - (zero.transform[5] + zero.layout.height / 2))).toBeLessThan(FRAME.h * 0.01);
  });

  test('a secondary value axis scales its own group', () => {
    const line = `<c:lineChart><c:grouping val="standard"/>${ser({ name: 'Big', cats: CATS, vals: [100, 400, 300, 200] }, 3)}<c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>`;
    const sec = `<c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="4"/></c:catAx>` +
      `<c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="3"/><c:crosses val="max"/></c:valAx>`;
    const prims = draw(chartSpace(barChart(S3) + line + AXES() + sec));
    const left = label(prims, '6')!;
    const right = label(prims, '450') ?? label(prims, '500')!;
    expect(right.transform[4]).toBeGreaterThan(left.transform[4] + FRAME.w / 2);
  });
});

describe('pies and doughnuts', () => {
  const pie = (vals: number[], extra = '', kind = 'pieChart', after = '') => chartSpace(
    `<c:${kind}><c:varyColors val="1"/>${ser({ name: 'S', cats: vals.map((_, i) => `P${i + 1}`), vals, extra }, 0)}${after}</c:${kind}>`);
  const slices = (prims: ReturnType<typeof draw>) => ACCENT.map(c => paths(prims).find(p => hexOf(p) === c)).filter((p): p is PathPrim => !!p);
  /** The angle (degrees clockwise from 12 o'clock) of a slice path's first edge, from its centre. */
  const startAngle = (p: PathPrim) => {
    const [m, l] = p.paths[0]!.commands as { x: number; y: number }[];
    return ((Math.atan2(l!.x - m!.x, -(l!.y - m!.y)) * 180) / Math.PI + 360) % 360;
  };

  test('slices run clockwise from c:firstSliceAng, in proportion to the values', () => {
    const prims = draw(pie([1, 1, 2]));
    const s = slices(prims);
    expect(s).toHaveLength(3);
    expect(startAngle(s[0]!)).toBeCloseTo(0, 1);
    expect(startAngle(s[1]!)).toBeCloseTo(90, 1);
    expect(startAngle(s[2]!)).toBeCloseTo(180, 1);
    const turned = slices(draw(pie([1, 1, 2], '', 'pieChart', '<c:firstSliceAng val="90"/>')));
    expect(startAngle(turned[0]!)).toBeCloseTo(90, 1);
  });

  test('an exploded point moves out along its middle', () => {
    const plain = slices(draw(pie([1, 1, 2])));
    const ex = slices(draw(pie([1, 1, 2], '<c:dPt><c:idx val="0"/><c:explosion val="25"/></c:dPt>')));
    const centre = (p: PathPrim) => ({ x: p.transform[4] + (p.paths[0]!.commands[0] as { x: number }).x, y: p.transform[5] + (p.paths[0]!.commands[0] as { y: number }).y });
    const moved = centre(ex[0]!);
    const others = centre(ex[1]!);
    // slice 1 spans 0..90 degrees: it moves up and right of the others' centre
    expect(moved.x).toBeGreaterThan(others.x);
    expect(moved.y).toBeLessThan(others.y);
    expect(plain[0]!.box.w).toBeGreaterThan(ex[1]!.box.w);
  });

  test('a doughnut leaves a hole of c:holeSize percent', () => {
    const prims = draw(pie([1, 1], '', 'doughnutChart', '<c:firstSliceAng val="0"/><c:holeSize val="50"/>'));
    const s = slices(prims)[0]!;
    const cmds = s.paths[0]!.commands as { op: string; x: number; y: number }[];
    const ring = Math.hypot(cmds[1]!.x - cmds[0]!.x, cmds[1]!.y - cmds[0]!.y);
    // the half ring from 12 to 6 o'clock is a diameter tall; its first edge runs from the hole to the rim, half a radius
    expect(ring / s.box.h).toBeCloseTo(0.25, 2);
  });

  test('a 3-D pie is tilted and gets a side', () => {
    const xml = chartSpace(`<c:pie3DChart><c:varyColors val="1"/>${ser({ name: 'S', cats: ['a', 'b'], vals: [3, 1] }, 0)}</c:pie3DChart>`).replace('<c:chart>', '<c:chart><c:view3D><c:rotX val="30"/></c:view3D>');
    const prims = draw(xml);
    const top = paths(prims).find(p => hexOf(p) === ACCENT[0])!;
    expect(top.box.h / top.box.w).toBeLessThan(0.75);
    expect(paths(prims).filter(p => p.fill.kind === 'linear').length).toBeGreaterThan(0);
  });

  test('percentage labels', () => {
    const lbls = '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:separator>\n</c:separator></c:dLbls>';
    const prims = draw(pie([1, 3], '', 'pieChart', lbls));
    expect(label(prims, 'P1\n25%')).toBeTruthy();
    expect(label(prims, 'P2\n75%')).toBeTruthy();
  });
});

describe('lines, areas and scatter', () => {
  test('blanks break a line ("gap"); markers sit on the points', () => {
    const xml = chartSpace(`<c:lineChart><c:grouping val="standard"/>${ser({ name: 'S', cats: CATS, vals: [1, 2, undefined, 4] }, 0)}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>` + AXES());
    const prims = draw(xml.replace(/<c:legend>.*<\/c:legend>/, ''));
    const accent = (p: PathPrim) => p.stroke?.paint.kind === 'solid' && hexOf({ ...p, fill: p.stroke.paint }) === ACCENT[0];
    const lines = paths(prims).filter(p => p.fill.kind === 'none' && accent(p));
    expect(lines).toHaveLength(1);
    expect(lines[0]!.paths[0]!.commands.filter(c => c.op === 'M')).toHaveLength(2);
    // three diamond markers in the series colour
    expect(paths(prims).filter(p => hexOf(p) === ACCENT[0] && p.paths[0]!.commands.length === 5)).toHaveLength(3);
  });

  test('stacked areas share edges', () => {
    const xml = chartSpace(`<c:areaChart><c:grouping val="stacked"/>${S3.slice(0, 2).map((s, i) => ser(s, i)).join('')}<c:axId val="1"/><c:axId val="2"/></c:areaChart>` + AXES());
    const prims = draw(xml);
    const [a, b] = ACCENT.slice(0, 2).map(c => paths(prims).find(p => hexOf(p) === c)!);
    // the upper band's lowest point is the lower band's highest
    expect(b!.transform[5] + b!.box.h).toBeGreaterThan(a!.transform[5]);
  });

  test('scatter points sit at their x values', () => {
    const xml = chartSpace(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` +
      `<c:marker><c:symbol val="square"/><c:size val="7"/></c:marker><c:xVal>${numCache([1, 2, 4])}</c:xVal><c:yVal>${numCache([5, 5, 5])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>` +
      AXES().replace('<c:catAx>', '<c:valAx>').replace('</c:catAx>', '</c:valAx>'));
    const marks = paths(draw(xml.replace(/<c:legend>.*<\/c:legend>/, ''))).filter(p => hexOf(p) === ACCENT[0]).map(boxOf).sort((p, q) => p.x - q.x);
    expect(marks).toHaveLength(3);
    const step = marks[1]!.x - marks[0]!.x;
    expect(marks[2]!.x - marks[1]!.x).toBeCloseTo(2 * step, -2);
  });

  test('bubbles scale by area', () => {
    const xml = chartSpace(`<c:bubbleChart><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal>${numCache([1, 2])}</c:xVal><c:yVal>${numCache([1, 2])}</c:yVal>` +
      `<c:bubbleSize>${numCache([1, 4])}</c:bubbleSize></c:ser><c:bubbleScale val="100"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>` + AXES().replace('<c:catAx>', '<c:valAx>').replace('</c:catAx>', '</c:valAx>'));
    const b = paths(draw(xml.replace(/<c:legend>.*<\/c:legend>/, ''))).filter(p => hexOf(p) === ACCENT[0]).map(p => p.box.w).sort((p, q) => p - q);
    expect(b[1]! / b[0]!).toBeCloseTo(2, 1);
  });
});

describe('titles, legends and text', () => {
  test('a single named series titles the chart unless the title was deleted', () => {
    const one = barChart([{ name: 'Sales', cats: CATS, vals: [1, 2, 3, 4] }]) + AXES();
    expect(label(draw(chartSpace(one, { autoTitleDeleted: false })), 'Sales')).toBeTruthy();
    const deleted = texts(draw(chartSpace(one, { autoTitleDeleted: true, legend: '' })));
    expect(deleted.some(t => textOf(t) === 'Sales')).toBe(false);
    // an empty c:title over several series shows Office's placeholder text
    expect(label(draw(chartSpace(barChart(S3) + AXES(), { title: '<c:title><c:overlay val="0"/></c:title>' })), 'Chart Title')).toBeTruthy();
  });

  test("text sizes: the chart's size for labels, 120% and bold for the title, 10 pt without a size", () => {
    const one = barChart([{ name: 'Sales', cats: CATS, vals: [1, 2, 3, 4] }]) + AXES();
    const prims = draw(chartSpace(one, { autoTitleDeleted: false, legend: '' }));
    const run = (t: string) => label(prims, t)!.layout.lines[0]!.runs[0]!.props;
    expect(run('Q1').size).toBe(18);
    expect(run('Sales').size).toBeCloseTo(21.6);
    expect(run('Sales').bold).toBe(true);
    const plain = draw(chartSpace(one, { txPr: '' }));
    expect(label(plain, 'Q1')!.layout.lines[0]!.runs[0]!.props.size).toBe(10);
  });

  test('rich text titles keep their runs and the paragraph defaults', () => {
    const title = '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US" i="1"/><a:t>Hello</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>';
    const t = label(draw(chartSpace(barChart(S3) + AXES(), { title })), 'Hello')!;
    expect(t.layout.lines[0]!.runs[0]!.props).toMatchObject({ size: 14, bold: false, italic: true });
  });

  test('deleted legend entries are left out', () => {
    const legend = '<c:legend><c:legendPos val="b"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry><c:overlay val="0"/></c:legend>';
    const prims = draw(chartSpace(barChart(S3) + AXES(), { legend }));
    expect(label(prims, 'North')).toBeTruthy();
    expect(label(prims, 'South')).toBeUndefined();
    // a bottom legend sits under the plot
    expect(label(prims, 'North')!.transform[5]).toBeGreaterThan(FRAME.h * 0.8);
  });

  test('data labels: value, number format and position above the bar', () => {
    const dLbls = '<c:dLbls><c:numFmt formatCode="0.0" sourceLinked="0"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
    const prims = draw(chartSpace(barChart([{ name: 'S', cats: CATS, vals: [1, 2, 3, 4] }], { dLbls }) + AXES()));
    const bar = bars(prims, ACCENT[0]!)[3]!;
    const t = label(prims, '4.0')!;
    expect(t.transform[5] + t.layout.height).toBeLessThanOrEqual(bar.y + 1);
  });

  test("a chart's theme override replaces the slide theme's colours", () => {
    const override = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:clrScheme name="Red">' +
      '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2><a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>' +
      '<a:accent5><a:srgbClr val="00FFFF"/></a:accent5><a:accent6><a:srgbClr val="FF00FF"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeOverride>';
    const prims = draw(chartSpace(barChart(S3) + AXES()), FRAME, { 'ppt/theme/themeOverride1.xml': override }, [['rId1', 'themeOverride', '../theme/themeOverride1.xml']]);
    expect(bars(prims, '#FF0000')).toHaveLength(4);
    expect(bars(prims, ACCENT[0]!)).toHaveLength(0);
  });
});

describe('the renderer hook', () => {
  test('buildSlide draws the chart in the frame', () => {
    const prims = build(chartSpace(barChart(S3) + AXES()));
    expect(prims.some(p => p.kind === 'missing')).toBe(false);
    expect(paths(prims).filter(p => hexOf(p) === ACCENT[0] && p.box.h > 150000)).toHaveLength(4);
  });

  test('a part that is not a chart gives undefined; broken content degrades', () => {
    const doc = chartDeck(chartSpace(barChart(S3) + AXES()));
    const { ctx } = chartSlide(doc);
    const hook = chartRenderer(fonts());
    expect(hook('/ppt/slides/slide1.xml', FRAME, ctx)).toBeUndefined();
    expect(hook('/ppt/charts/missing.xml', FRAME, ctx)).toBeUndefined();
    expect(hook('/ppt/charts/chart1.xml', { w: 0, h: 0 }, ctx)).toEqual([]);
  });

  test.each([
    ['an empty plot area', chartSpace('')],
    ['no axes', chartSpace(barChart(S3))],
    ['axes that point nowhere', chartSpace(barChart(S3).replace(/<c:axId val="\d"\/>/g, '<c:axId val="77"/>'))],
    ['values that are not numbers', chartSpace(barChart([{ name: 'S', cats: CATS, vals: [1, 2] }]).replace('<c:v>1</c:v>', '<c:v>NaN</c:v>') + AXES())],
    ['a zero major unit and a reversed min and max', chartSpace(barChart(S3) + AXES(1, 2, { cat: '', val: '<c:majorUnit val="0"/>' }).replace('<c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>', '<c:orientation val="minMax"/><c:max val="1"/><c:min val="5"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'))],
    ['a log axis over negative data', chartSpace(barChart([{ name: 'S', cats: CATS, vals: [-1, 0, 2, 3] }]) + AXES().replace('<c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>', '<c:logBase val="10"/><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'))],
    ['a tiny major unit', chartSpace(barChart(S3) + AXES(1, 2, { cat: '', val: '<c:majorUnit val="0.0000001"/>' }))],
    ['every kind of chart at once', chartSpace(['pieChart', 'doughnutChart', 'radarChart', 'stockChart', 'surfaceChart', 'ofPieChart', 'bubbleChart', 'area3DChart', 'line3DChart']
      .map(k => `<c:${k}>${ser({ name: k, cats: CATS, vals: [1, 2, 3, 4] }, 0)}<c:axId val="1"/><c:axId val="2"/></c:${k}>`).join('') + AXES())],
    ['a manual layout outside the chart', chartSpace(barChart(S3) + AXES()).replace('<c:plotArea><c:layout/>', '<c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="-3"/><c:y val="7"/><c:w val="1e300"/><c:h val="-2"/></c:manualLayout></c:layout>')],
  ])('%s', (_name, xml) => {
    const doc = chartDeck(xml);
    const { ctx } = chartSlide(doc);
    const prims = chartRenderer(fonts())('/ppt/charts/chart1.xml', FRAME, ctx);
    expect(prims).toBeDefined();
    expect(finite(prims!)).toBe(true);
  });

  test('huge charts stay quick: 20000 categories, a 3000-slice pie with labels', () => {
    const n = 20000;
    const cats = Array.from({ length: n }, (_, i) => `Category ${i}`);
    const started = performance.now();
    const bars = draw(chartSpace(barChart([{ name: 'S', cats, vals: cats.map((_, i) => i % 97) }]) + AXES()));
    const m = 3000;
    const lbls = '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
    const pie = draw(chartSpace(`<c:pieChart><c:varyColors val="1"/>${ser({ name: 'P', cats: cats.slice(0, m), vals: cats.slice(0, m).map((_, i) => 1 + (i % 5)) }, 0)}${lbls}</c:pieChart>`));
    expect(performance.now() - started).toBeLessThan(20000);
    expect(finite(bars) && finite(pie)).toBe(true);
    // only a readable share of the category labels is shown
    expect(texts(bars).filter(t => textOf(t).startsWith('Category')).length).toBeLessThan(200);
  });

  test('labels and strings that are not valid stay text', () => {
    const xml = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strCache(['&lt;b&gt;&amp;'])}</c:tx>` +
      `<c:cat>${strCache(['', '  ', 'x'])}</c:cat><c:val>${numCache([1, 2, 3])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES()}</c:plotArea>` +
      `<c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`;
    expect(label(draw(xml), '<b>&')).toBeTruthy();
  });
});
