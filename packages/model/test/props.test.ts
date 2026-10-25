import { describe, expect, test } from 'vitest';
import {
  colorIn, effectsIn, fillIn, parseBodyProps, parseColor, parseLine, parseListStyle, parseParaProps, parseRunProps, parseStyle,
  parseTable, parseTheme, paragraphText, paragraphsOf, runsOf,
} from '../src/index.ts';
import { find, NSDECL, parse } from './helpers.ts';

const a = (inner: string) => parse(`<p:x ${NSDECL}>${inner}</p:x>`);

describe('colours', () => {
  test('bases', () => {
    expect(parseColor(find(a('<a:srgbClr val="1F4E79"/>'), 'srgbClr'))).toEqual({ base: { kind: 'srgb', r: 0x1f, g: 0x4e, b: 0x79 }, transforms: [] });
    expect(parseColor(find(a('<a:schemeClr val="accent1"/>'), 'schemeClr'))!.base).toEqual({ kind: 'scheme', name: 'accent1' });
    expect(parseColor(find(a('<a:sysClr val="windowText" lastClr="000000"/>'), 'sysClr'))!.base).toEqual({ kind: 'sys', name: 'windowText', last: { r: 0, g: 0, b: 0 } });
    expect(parseColor(find(a('<a:prstClr val="black"/>'), 'prstClr'))!.base).toEqual({ kind: 'preset', name: 'black' });
    expect(parseColor(find(a('<a:hslClr hue="10800000" sat="100000" lum="50000"/>'), 'hslClr'))!.base).toEqual({ kind: 'hsl', h: 180, s: 1, l: 0.5 });
    expect(parseColor(find(a('<a:scrgbClr r="50%" g="0" b="100000"/>'), 'scrgbClr'))!.base).toEqual({ kind: 'scrgb', r: 0.5, g: 0, b: 1 });
    expect(parseColor(find(a('<a:srgbClr val="zz"/>'), 'srgbClr'))).toBeUndefined();
  });

  test('transforms keep document order, percentages as fractions, hue in degrees', () => {
    const c = colorIn(a('<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25%"/><a:hueOff val="5400000"/><a:alpha val="60000"/><a:gray/></a:schemeClr></a:solidFill>').children[0] as never);
    expect(c!.transforms).toEqual([
      { name: 'lumMod', val: 0.75 }, { name: 'lumOff', val: 0.25 }, { name: 'hueOff', val: 90 }, { name: 'alpha', val: 0.6 }, { name: 'gray', val: 0 },
    ]);
  });
});

describe('fills, lines, effects, style refs', () => {
  test('a gradient with stops out of order, a linear angle and scaling', () => {
    const f = fillIn(a('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>'));
    expect(f).toMatchObject({ kind: 'gradient', shade: { kind: 'linear', angle: 90, scaled: true }, rotWithShape: true });
    if (f?.kind !== 'gradient') throw new Error();
    expect(f.stops.map(s => s.pos)).toEqual([0, 1]);
  });

  test('a path gradient and a picture fill with crop and stretch', () => {
    expect(fillIn(a('<a:gradFill><a:gsLst/><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>')))
      .toMatchObject({ shade: { kind: 'path', path: 'circle', fillToRect: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 } } });
    const pic = fillIn(a('<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId2"><a:grayscl/><a:lum bright="20000" contrast="-10000"/>' +
      '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId3"/></a:ext></a:extLst></a:blip>' +
      '<a:srcRect l="10000" r="5000"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'));
    expect(pic).toMatchObject({
      kind: 'blip', srcRect: { l: 0.1, t: 0, r: 0.05, b: 0 }, mode: { kind: 'stretch' },
      blip: { embed: 'rId2', svgEmbed: 'rId3', effects: [{ kind: 'grayscl' }, { kind: 'lum', bright: 0.2, contrast: -0.1 }] },
    });
  });

  test('pattern, none and group fills', () => {
    expect(fillIn(a('<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>')))
      .toMatchObject({ kind: 'pattern', preset: 'dkDnDiag' });
    expect(fillIn(a('<a:noFill/>'))).toEqual({ kind: 'none' });
    expect(fillIn(a('<a:grpFill/>'))).toEqual({ kind: 'group' });
    expect(fillIn(a('<a:ln/>'))).toBeUndefined();
  });

  test('a line with width, dash, join, arrowheads and fill', () => {
    expect(parseLine(find(a('<a:ln w="38100" cap="rnd" cmpd="dbl"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="dashDot"/><a:miter lim="800000"/><a:tailEnd type="triangle" w="lg"/></a:ln>'), 'ln')))
      .toMatchObject({ w: 38100, cap: 'rnd', cmpd: 'dbl', fill: { kind: 'solid' }, dash: { kind: 'preset', name: 'dashDot' }, join: { kind: 'miter', lim: 8 }, tailEnd: { type: 'triangle', w: 'lg', len: 'med' } });
    expect(parseLine(find(a('<a:ln><a:custDash><a:ds d="300000" sp="100000"/></a:custDash><a:round/></a:ln>'), 'ln')))
      .toEqual({ dash: { kind: 'custom', segments: [{ d: 3, sp: 1 }] }, join: { kind: 'round' } });
  });

  test('effects: an empty list means none, which differs from no list', () => {
    expect(effectsIn(a('<a:effectLst/>'))).toEqual({});
    expect(effectsIn(a(''))).toBeUndefined();
    const e = effectsIn(a('<a:effectLst><a:glow rad="63500"><a:schemeClr val="accent1"/></a:glow><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw><a:softEdge rad="12700"/></a:effectLst>'));
    expect(e).toMatchObject({ glow: { rad: 63500 }, outerShdw: { blurRad: 40000, dist: 23000, dir: 90, rotWithShape: false }, softEdge: { rad: 12700 } });
  });

  test('style references', () => {
    const s = parseStyle(find(a('<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'), 'style'));
    expect(s).toMatchObject({ ln: { idx: 2 }, fill: { idx: 1 }, effect: { idx: 0 }, font: { idx: 'minor', color: { base: { kind: 'scheme', name: 'lt1' } } } });
  });
});

describe('text properties', () => {
  test('run properties', () => {
    const r = parseRunProps(find(a('<a:rPr lang="en-US" sz="2400" b="1" i="0" u="sng" strike="sngStrike" cap="small" spc="-50" baseline="30000" kern="1200">' +
      '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:highlight><a:srgbClr val="FFFF00"/></a:highlight><a:latin typeface="+mj-lt"/><a:ea typeface="MS Gothic" charset="-128"/><a:sym typeface="Wingdings" pitchFamily="2"/>' +
      '<a:hlinkClick r:id="rId4" tooltip="go"/></a:rPr>'), 'rPr'));
    expect(r).toMatchObject({
      lang: 'en-US', sz: 2400, b: true, i: false, u: 'sng', strike: 'sngStrike', cap: 'small', spc: -50, baseline: 0.3, kern: 1200,
      fill: { kind: 'solid' }, highlight: { base: { kind: 'srgb' } }, latin: { typeface: '+mj-lt' }, ea: { typeface: 'MS Gothic', charset: -128 },
      sym: { typeface: 'Wingdings', pitchFamily: 2 }, hlinkClick: { rId: 'rId4', tooltip: 'go' },
    });
    expect(parseRunProps(undefined)).toEqual({});
  });

  test('paragraph properties with spacing, bullets and tabs', () => {
    const p = parseParaProps(find(a('<a:pPr marL="342900" indent="-342900" algn="ctr" lvl="1" rtl="0"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef>' +
      '<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buSzPct val="80000"/><a:buFont typeface="Arial"/><a:buChar char="•"/><a:tabLst><a:tab pos="914400" algn="r"/></a:tabLst><a:defRPr sz="1800"/></a:pPr>'), 'pPr'));
    expect(p).toMatchObject({
      marL: 342900, indent: -342900, algn: 'ctr', lvl: 1, rtl: false, lnSpc: { kind: 'pct', val: 0.9 }, spcBef: { kind: 'pts', val: 1000 },
      buClr: { base: { kind: 'scheme' } }, buSz: { kind: 'pct', val: 0.8 }, buFont: { typeface: 'Arial' }, bullet: { kind: 'char', char: '•' },
      tabs: [{ pos: 914400, algn: 'r' }], defRPr: { sz: 1800 },
    });
    expect(parseParaProps(find(a('<a:pPr><a:buSzTx/><a:buFontTx/><a:buClrTx/><a:buAutoNum type="romanUcPeriod" startAt="3"/></a:pPr>'), 'pPr')))
      .toEqual({ buSz: 'followText', buFont: 'followText', buClr: 'followText', bullet: { kind: 'autonum', scheme: 'romanUcPeriod', startAt: 3 } });
  });

  test('a bullet is one character, as PowerPoint draws SmartArt\'s char="••"', () => {
    expect(parseParaProps(find(a('<a:pPr><a:buChar char="••"/></a:pPr>'), 'pPr')).bullet).toEqual({ kind: 'char', char: '•' });
    expect(parseParaProps(find(a('<a:pPr><a:buChar char="😀x"/></a:pPr>'), 'pPr')).bullet).toEqual({ kind: 'char', char: '😀' });
  });

  test('body properties and autofit', () => {
    expect(parseBodyProps(find(a('<a:bodyPr vert="vert270" wrap="none" lIns="0" tIns="45720" anchor="ctr" anchorCtr="1" numCol="2" rot="-5400000"><a:normAutofit fontScale="92500" lnSpcReduction="10000"/></a:bodyPr>'), 'bodyPr')))
      .toEqual({ vert: 'vert270', wrap: 'none', lIns: 0, tIns: 45720, anchor: 'ctr', anchorCtr: true, numCol: 2, rot: -90, autofit: { kind: 'norm', fontScale: 0.925, lnSpcReduction: 0.1 } });
    expect(parseBodyProps(find(a('<a:bodyPr><a:spAutoFit/></a:bodyPr>'), 'bodyPr')).autofit).toEqual({ kind: 'shape' });
    expect(parseBodyProps(find(a('<a:bodyPr><a:normAutofit/></a:bodyPr>'), 'bodyPr')).autofit).toEqual({ kind: 'norm', fontScale: 1, lnSpcReduction: 0 });
  });

  test('list styles have nine levels', () => {
    const ls = parseListStyle(find(a('<a:lstStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0"><a:defRPr sz="3200"/></a:lvl1pPr><a:lvl9pPr marL="3657600"/></a:lstStyle>'), 'lstStyle'));
    expect(ls.def).toEqual({ defRPr: { lang: 'en-US' } });
    expect(ls.levels.length).toBe(9);
    expect(ls.levels[0]).toEqual({ marL: 0, defRPr: { sz: 3200 } });
    expect(ls.levels[1]).toBeUndefined();
    expect(ls.levels[8]).toEqual({ marL: 3657600 });
  });

  test('runs, breaks and fields', () => {
    const body = find(a('<p:txBody><a:bodyPr/><a:p><a:r><a:t>Slide </a:t></a:r><a:fld id="{1}" type="slidenum"><a:rPr lang="en-US"/><a:t>3</a:t></a:fld><a:br/><a:r><a:t>next</a:t></a:r></a:p><a:p/></p:txBody>'), 'txBody');
    const [p1] = paragraphsOf(body);
    expect(runsOf(p1!).map(r => [r.kind, r.text, r.fieldType])).toEqual([['r', 'Slide ', undefined], ['fld', '3', 'slidenum'], ['br', '', undefined], ['r', 'next', undefined]]);
    expect(paragraphText(p1!)).toBe('Slide 3\u000bnext');
  });
});

describe('themes and tables', () => {
  test('a theme with colours, per-script fonts and the style matrix', () => {
    const t = parseTheme(parse(`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>` +
      '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme>' +
      '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
      '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill/><a:gradFill/></a:fillStyleLst>' +
      '<a:lnStyleLst><a:ln w="6350"/><a:ln w="12700"/><a:ln w="19050"/></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
      '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'));
    expect(t.name).toBe('Office Theme');
    expect([...t.colors.keys()]).toEqual(['dk1', 'lt1', 'accent1']);
    expect(t.major).toEqual({ latin: 'Calibri Light', ea: '', cs: '', scripts: { Jpan: '游ゴシック Light' } });
    expect(t.minor.latin).toBe('Calibri');
    expect([t.fillStyles.length, t.lineStyles.length, t.effectStyles.length, t.bgFillStyles.length]).toEqual([3, 3, 1, 1]);
  });

  test('a table with merges and style flags', () => {
    const t = parseTable(find(a('<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>' +
      '<a:tblGrid><a:gridCol w="100"/><a:gridCol w="200"/></a:tblGrid>' +
      '<a:tr h="50"><a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1"><a:tcPr/></a:tc></a:tr></a:tbl>'), 'tbl'));
    expect(t.styleId).toBe('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
    expect(t.flags).toMatchObject({ firstRow: true, bandRow: true, firstCol: false });
    expect(t.columns).toEqual([100, 200]);
    expect(t.rows[0]!.cells.map(c => [c.gridSpan, c.hMerge, c.txBody !== undefined])).toEqual([[2, false, true], [1, true, false]]);
  });
});
