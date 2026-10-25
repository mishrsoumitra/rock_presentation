// The guide evaluator: formulas, built-in guides, document order, adjust values and errors (ECMA-376 Part 1, 20.1.9.11).
import { describe, expect, test } from 'vitest';
import { evaluateGeometry, GeometryError, type Guide, type ShapeGeometry } from '../src/index.ts';
import { custGeom, guidesOnly } from './helpers.ts';

const W = 1000, H = 500;
const guides = (gd: [string, string][], av: [string, string][] = [], adjust: Guide[] = []) =>
  evaluateGeometry(guidesOnly(gd, av), W, H, adjust).guides;
const one = (fmla: string) => guides([['x', fmla]]).get('x');

function rejects(geom: ShapeGeometry, ...fragments: string[]): void {
  let error: unknown;
  try {
    evaluateGeometry(geom, W, H);
  } catch (e) {
    error = e;
  }
  expect(error).toBeInstanceOf(GeometryError);
  for (const f of fragments) expect((error as Error).message).toContain(f);
}

describe('the 17 formulas', () => {
  test.each([
    ['*/ 6 7 2', 21],
    ['+- 6 7 2', 11],
    ['+/ 6 7 2', 6.5],
    ['?: 1 5 9', 5],
    ['?: 0 5 9', 9],          // "if x > 0": zero takes the else branch
    ['?: -1 5 9', 9],
    ['abs -7', 7],
    ['abs 7', 7],
    ['at2 1 1', 2700000],     // arctan(y / x), in 60000ths of a degree
    ['at2 -1 0', 10800000],
    ['at2 0 -1', -5400000],
    ['cat2 10 3 4', 6],       // x cos(arctan(z / y))
    ['sat2 10 3 4', 8],       // x sin(arctan(z / y))
    ['cos 10 10800000', -10],
    ['cos 2 3600000', 1],
    ['sin 10 5400000', 10],
    ['sin 2 1800000', 1],
    ['tan 10 2700000', 10],
    ['max 3 9', 9],
    ['max 9 3', 9],
    ['min 3 9', 3],
    ['min 9 3', 3],
    ['mod 3 4 12', 13],       // sqrt(x^2 + y^2 + z^2)
    ['pin 0 5 10', 5],
    ['pin 0 -5 10', 0],
    ['pin 0 15 10', 10],
    ['sqrt 16', 4],
    ['val 42', 42],
    ['val -3.5', -3.5],
  ])('%s = %d', (fmla, expected) => {
    expect(one(fmla)).toBeCloseTo(expected, 6);
  });

  test('arguments are guide names or numbers, separated by any whitespace', () => {
    expect(guides([['a', 'val 3'], ['b', '*/   a\t4  2 ']]).get('b')).toBe(6);
    expect(one('+- +5 .5 1.')).toBe(4.5);
  });

  test('a result that is not a finite number is 0: degenerate shapes divide by zero', () => {
    expect(one('*/ 1 1 0')).toBe(0);
    expect(one('+/ 1 1 0')).toBe(0);
    expect(one('sqrt -4')).toBe(0);
    expect(one(`val ${'9'.repeat(400)}`)).toBe(0);
  });
});

describe('built-in guides', () => {
  test('section 20.1.10.56, at 1000 x 500', () => {
    const g = guides([]);
    const expected: Record<string, number> = {
      l: 0, t: 0, r: 1000, b: 500, w: 1000, h: 500, hc: 500, vc: 250, ss: 500, ls: 1000,
      cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000,
      hd2: 250, hd3: 500 / 3, hd4: 125, hd5: 100, hd6: 500 / 6, hd8: 62.5,
      ssd2: 250, ssd4: 125, ssd6: 500 / 6, ssd8: 62.5, ssd16: 31.25, ssd32: 15.625,
      wd2: 500, wd3: 1000 / 3, wd4: 250, wd5: 200, wd6: 1000 / 6, wd8: 125, wd10: 100,
      hd10: 50, wd32: 31.25,   // not in that list, but Annex D uses them
    };
    for (const [name, value] of Object.entries(expected)) expect(g.get(name), name).toBeCloseTo(value, 9);
  });

  test('ss and ls follow the shorter and longer side', () => {
    const g = evaluateGeometry(guidesOnly([]), 300, 700).guides;
    expect([g.get('ss'), g.get('ls'), g.get('ssd2')]).toEqual([300, 700, 150]);
  });
});

describe('document order', () => {
  test('a guide sees the guides before it', () => {
    expect(guides([['a', 'val 2'], ['b', '*/ a 3 1'], ['c', '+- b a 0']]).get('c')).toBe(8);
  });

  test('a guide used before it is defined is an error', () => {
    rejects(guidesOnly([['b', '*/ a 3 1'], ['a', 'val 2']]), 'guide "b"', 'unknown guide "a"');
  });

  test('a name defined again hides the earlier value from then on (circularArrow redefines dxI and dyI)', () => {
    const g = guides([['a', 'val 1'], ['b', '+- a 0 0'], ['a', 'val 5'], ['c', '+- a 0 0']]);
    expect([g.get('b'), g.get('c'), g.get('a')]).toEqual([1, 5, 5]);
  });

  test('a guide can hide a built-in (textRingInside defines r)', () => {
    const g = guides([['before', '+- r 0 0'], ['r', 'val 7'], ['after', '+- r 0 0']]);
    expect([g.get('before'), g.get('after'), g.get('r')]).toEqual([1000, 7, 7]);
  });

  test('paths, the text rectangle, handles and connection sites see the final guides', () => {
    const g = custGeom(`
      <a:gdLst><a:gd name="x" fmla="val 10"/><a:gd name="x" fmla="val 20"/></a:gdLst>
      <a:ahLst><a:ahXY><a:pos x="x" y="x"/></a:ahXY></a:ahLst>
      <a:cxnLst><a:cxn ang="x"><a:pos x="x" y="x"/></a:cxn></a:cxnLst>
      <a:rect l="x" t="x" r="r" b="b"/>
      <a:pathLst><a:path><a:moveTo><a:pt x="x" y="x"/></a:moveTo></a:path></a:pathLst>`);
    const e = evaluateGeometry(g, W, H);
    expect(e.paths[0]!.commands).toEqual([{ op: 'M', x: 20, y: 20 }]);
    expect(e.textRect).toEqual({ x: 20, y: 20, w: 980, h: 480 });
    expect(e.handles).toEqual([{ x: 20, y: 20 }]);
    expect(e.connectionSites).toEqual([{ x: 20, y: 20, angle: 20 }]);
  });
});

describe('adjust values', () => {
  const av: [string, string][] = [['adj', 'val 100'], ['adj2', 'val 7']];
  const gd: [string, string][] = [['x', '*/ adj 2 1'], ['y', '+- adj2 0 0']];

  test('the geometry defaults apply when the shape gives none', () => {
    expect(guides(gd, av).get('x')).toBe(200);
  });

  test("the shape's own guides replace them by name", () => {
    const g = guides(gd, av, [{ name: 'adj', fmla: 'val 30' }]);
    expect([g.get('adj'), g.get('x'), g.get('y')]).toEqual([30, 60, 7]);
  });

  test('a replacement may be any formula over the built-ins and the adjust values before it', () => {
    expect(guides(gd, av, [{ name: 'adj', fmla: '*/ w 1 4' }]).get('x')).toBe(500);
    expect(guides(gd, av, [{ name: 'adj2', fmla: '+- adj 1 0' }]).get('y')).toBe(101);
  });

  test('names the geometry has no adjust value for are ignored, as are other guides of that name', () => {
    expect(guides(gd, av, [{ name: 'nope', fmla: 'val 1' }, { name: 'x', fmla: 'val 1' }]).get('x')).toBe(200);
  });

  test('when a name is given twice, the last one counts', () => {
    expect(guides(gd, av, [{ name: 'adj', fmla: 'val 1' }, { name: 'adj', fmla: 'val 2' }]).get('x')).toBe(4);
  });

  test('a replacement that names an unknown guide, or a later one, is an error', () => {
    expect(() => guides(gd, av, [{ name: 'adj', fmla: 'val nope' }])).toThrow(/adjust value "adj": unknown guide "nope"/);
    expect(() => guides(gd, av, [{ name: 'adj', fmla: '+- x 0 0' }])).toThrow(GeometryError);
    expect(() => guides(gd, av, [{ name: 'adj', fmla: 'oops 1' }])).toThrow(/unknown formula "oops"/);
  });
});

describe('errors', () => {
  test('unknown guide', () => rejects(guidesOnly([['x', '*/ zz 1 1']]), 'guide "x"', 'unknown guide "zz"'));
  test('unknown formula', () => rejects(guidesOnly([['x', 'foo 1 2']]), 'unknown formula "foo"'));
  test('an empty formula', () => rejects(guidesOnly([['x', '']]), 'unknown formula ""'));
  test('too few or too many arguments', () => {
    rejects(guidesOnly([['x', '*/ 1 2']]), '"*/" takes 3 arguments, not 2');
    rejects(guidesOnly([['x', '+- 1 2 3 0']]), '"+-" takes 3 arguments, not 4');
    rejects(guidesOnly([['x', 'abs']]), '"abs" takes 1 argument, not 0');
  });
  test('a universal measure is a coordinate, not a formula argument', () => {
    rejects(guidesOnly([['x', 'val 1in']]), 'unknown guide "1in"');
  });

  test.each([
    ['path', '<a:pathLst><a:path><a:moveTo><a:pt x="q" y="0"/></a:moveTo></a:path></a:pathLst>', 'path 1 moveTo: unknown guide "q"'],
    ['arc', '<a:pathLst><a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:arcTo wR="1" hR="1" stAng="q" swAng="0"/></a:path></a:pathLst>', 'path 1 arcTo'],
    ['text rectangle', '<a:rect l="q" t="t" r="r" b="b"/>', 'rect: unknown guide "q"'],
    ['connection site', '<a:cxnLst><a:cxn ang="q"><a:pos x="0" y="0"/></a:cxn></a:cxnLst>', 'cxn: unknown guide "q"'],
    ['handle limit', '<a:ahLst><a:ahXY gdRefX="adj" minX="0" maxX="q"><a:pos x="0" y="0"/></a:ahXY></a:ahLst>', 'handle 1: unknown guide "q"'],
  ])('unknown guide in a %s', (_, xml, message) => rejects(custGeom(xml), message));

  test('a broken geometry fails every time, and does not affect others', () => {
    const bad = guidesOnly([['x', 'nope']]);
    expect(() => evaluateGeometry(bad, W, H)).toThrow(GeometryError);
    expect(() => evaluateGeometry(bad, W, H)).toThrow('unknown formula "nope"');
    expect(evaluateGeometry(guidesOnly([['x', 'val 1']]), W, H).guides.get('x')).toBe(1);
  });
});

describe('coordinates', () => {
  test('literals, guide names and universal measures', () => {
    const g = custGeom(`<a:pathLst><a:path>
      <a:moveTo><a:pt x="1in" y="2.5cm"/></a:moveTo><a:lnTo><a:pt x="10mm" y="-1pt"/></a:lnTo><a:lnTo><a:pt x="1pc" y="1pi"/></a:lnTo>
      <a:lnTo><a:pt x="hc" y="+7"/></a:lnTo></a:path></a:pathLst>`);
    expect(evaluateGeometry(g, W, H).paths[0]!.commands).toEqual([
      { op: 'M', x: 914400, y: 900000 }, { op: 'L', x: 360000, y: -12700 }, { op: 'L', x: 152400, y: 152400 }, { op: 'L', x: 500, y: 7 },
    ]);
  });

  test('a size that is not a finite number counts as 0', () => {
    const e = evaluateGeometry(guidesOnly([]), Number.NaN, Infinity);
    expect([e.guides.get('w'), e.guides.get('h'), e.textRect.w]).toEqual([0, 0, 0]);
  });

  test('the text rectangle is the whole shape without a:rect, and is never inside out', () => {
    expect(evaluateGeometry(guidesOnly([]), W, H).textRect).toEqual({ x: 0, y: 0, w: W, h: H });
    const flipped = custGeom('<a:rect l="r" t="b" r="l" b="t"/><a:pathLst/>');
    expect(evaluateGeometry(flipped, W, H).textRect).toEqual({ x: 0, y: 0, w: W, h: H });
  });
});
