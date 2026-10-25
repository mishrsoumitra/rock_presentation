// parseCustomGeometry: a:custGeom into ShapeGeometry (ECMA-376 Part 1, 20.1.9).
import { parseXml } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { evaluateGeometry, GeometryError, parseCustomGeometry } from '../src/index.ts';
import { A, custGeom } from './helpers.ts';

const STRICT = 'http://purl.oclc.org/ooxml/drawingml/main';

describe('parseCustomGeometry', () => {
  test('every part of a custom geometry, as written', () => {
    const g = custGeom(`
      <a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>
      <a:gdLst><a:gd name="x1" fmla="*/ w adj 100000"/><a:gd name="y1" fmla="+-  vc 0 10 "/></a:gdLst>
      <a:ahLst>
        <a:ahXY gdRefX="adj" minX="0" maxX="50000"><a:pos x="x1" y="t"/></a:ahXY>
        <a:ahPolar gdRefAng="adj" minAng="0" maxAng="21599999" gdRefR="adj" minR="0" maxR="100"><a:pos x="hc" y="vc"/></a:ahPolar>
      </a:ahLst>
      <a:cxnLst><a:cxn ang="cd4"><a:pos x="hc" y="b"/></a:cxn></a:cxnLst>
      <a:rect l="l" t="t" r="r" b="b"/>
      <a:pathLst>
        <a:path w="100" h="200" fill="none" stroke="false" extrusionOk="0">
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="x1" y="y1"/></a:lnTo>
          <a:arcTo wR="10" hR="20" stAng="0" swAng="cd4"/>
          <a:quadBezTo><a:pt x="1" y="2"/><a:pt x="3" y="4"/></a:quadBezTo>
          <a:cubicBezTo><a:pt x="5" y="6"/><a:pt x="7" y="8"/><a:pt x="9" y="10"/></a:cubicBezTo>
          <a:close/>
        </a:path>
        <a:path/>
      </a:pathLst>`);
    expect(g).toEqual({
      avLst: [{ name: 'adj', fmla: 'val 25000' }],
      gdLst: [{ name: 'x1', fmla: '*/ w adj 100000' }, { name: 'y1', fmla: '+-  vc 0 10 ' }],   // formulas as written
      ahLst: [
        { kind: 'xy', gdRefX: 'adj', minX: '0', maxX: '50000', pos: { x: 'x1', y: 't' } },
        { kind: 'polar', gdRefR: 'adj', minR: '0', maxR: '100', gdRefAng: 'adj', minAng: '0', maxAng: '21599999', pos: { x: 'hc', y: 'vc' } },
      ],
      cxnLst: [{ ang: 'cd4', pos: { x: 'hc', y: 'b' } }],
      rect: { l: 'l', t: 't', r: 'r', b: 'b' },
      pathLst: [
        {
          w: 100, h: 200, fill: 'none', stroke: false, extrusionOk: false,
          commands: [
            { op: 'moveTo', pt: { x: '0', y: '0' } },
            { op: 'lnTo', pt: { x: 'x1', y: 'y1' } },
            { op: 'arcTo', wR: '10', hR: '20', stAng: '0', swAng: 'cd4' },
            { op: 'quadBezTo', pts: [{ x: '1', y: '2' }, { x: '3', y: '4' }] },
            { op: 'cubicBezTo', pts: [{ x: '5', y: '6' }, { x: '7', y: '8' }, { x: '9', y: '10' }] },
            { op: 'close' },
          ],
        },
        { w: 0, h: 0, fill: 'norm', stroke: true, extrusionOk: true, commands: [] },   // the schema defaults
      ],
    });
  });

  test('an empty custGeom: no lists, no text rectangle', () => {
    expect(custGeom('')).toEqual({ avLst: [], gdLst: [], ahLst: [], cxnLst: [], rect: null, pathLst: [] });
  });

  test('Strict documents use another namespace; elements of any other namespace are ignored', () => {
    const strict = parseCustomGeometry(parseXml(
      `<a:custGeom xmlns:a="${STRICT}"><a:pathLst><a:path><a:moveTo><a:pt x="1" y="2"/></a:moveTo></a:path></a:pathLst></a:custGeom>`).root);
    expect(strict.pathLst[0]!.commands).toEqual([{ op: 'moveTo', pt: { x: '1', y: '2' } }]);
    const mixed = parseCustomGeometry(parseXml(`<a:custGeom xmlns:a="${A}" xmlns:x="urn:x">
      <x:avLst><a:gd name="no" fmla="val 1"/></x:avLst><a:pathLst><a:path><x:lnTo/><a:extLst/><a:close/></a:path></a:pathLst>
    </a:custGeom>`).root);
    expect(mixed.avLst).toEqual([]);
    expect(mixed.pathLst[0]!.commands).toEqual([{ op: 'close' }]);
  });

  test('names and coordinates are tokens: surrounding whitespace is dropped', () => {
    const g = custGeom('<a:gdLst><a:gd name=" g " fmla="val 1"/></a:gdLst><a:rect l=" l" t="t " r="r" b="g"/>');
    expect(g.gdLst[0]!.name).toBe('g');
    expect(g.rect).toEqual({ l: 'l', t: 't', r: 'r', b: 'g' });
    expect(evaluateGeometry(g, 10, 10).textRect).toEqual({ x: 0, y: 0, w: 10, h: 1 });
  });

  test.each([
    ['<a:gdLst><a:gd name="x"/></a:gdLst>', '<gd> has no fmla attribute'],
    ['<a:gdLst><a:gd fmla="val 1"/></a:gdLst>', '<gd> has no name attribute'],
    ['<a:rect l="l" t="t" r="r"/>', '<rect> has no b attribute'],
    ['<a:cxnLst><a:cxn ang="0"/></a:cxnLst>', '<cxn> has no pos'],
    ['<a:ahLst><a:ahXY gdRefX="adj"/></a:ahLst>', '<ahXY> has no pos'],
    ['<a:pathLst><a:path><a:moveTo/></a:path></a:pathLst>', '<moveTo> needs 1 pt, not 0'],
    ['<a:pathLst><a:path><a:quadBezTo><a:pt x="0" y="0"/></a:quadBezTo></a:path></a:pathLst>', '<quadBezTo> needs 2 pt, not 1'],
    ['<a:pathLst><a:path><a:lnTo><a:pt x="0"/></a:lnTo></a:path></a:pathLst>', '<pt> has no y attribute'],
    ['<a:pathLst><a:path><a:arcTo wR="1" hR="1" stAng="0"/></a:path></a:pathLst>', '<arcTo> has no swAng attribute'],
    ['<a:pathLst><a:path fill="solid"/></a:pathLst>', '<path> fill="solid" is not a path fill mode'],
    ['<a:pathLst><a:path stroke="yes"/></a:pathLst>', '<path> stroke="yes" is not a boolean'],
    ['<a:pathLst><a:path w="-5"/></a:pathLst>', '<path> w="-5" is not a positive coordinate'],
  ])('%s is refused', (xml, message) => {
    expect(() => custGeom(xml)).toThrow(GeometryError);
    expect(() => custGeom(xml)).toThrow(message);
  });

  test('a custom geometry evaluates like a preset', () => {
    // the triangle of the 20.1.9.11 example: its base on a guide at 2/3 of the height
    const g = custGeom(`<a:gdLst><a:gd name="myGuide" fmla="*/ h 2 3"/></a:gdLst>
      <a:pathLst><a:path w="1705233" h="679622">
        <a:moveTo><a:pt x="0" y="myGuide"/></a:moveTo><a:lnTo><a:pt x="1705233" y="myGuide"/></a:lnTo>
        <a:lnTo><a:pt x="852616" y="0"/></a:lnTo><a:close/>
      </a:path></a:pathLst>`);
    const path = evaluateGeometry(g, 1705233, 679622).paths[0]!;
    expect(path.commands).toEqual([
      { op: 'M', x: 0, y: (679622 * 2) / 3 }, { op: 'L', x: 1705233, y: (679622 * 2) / 3 }, { op: 'L', x: 852616, y: 0 }, { op: 'Z' },
    ]);
  });
});
