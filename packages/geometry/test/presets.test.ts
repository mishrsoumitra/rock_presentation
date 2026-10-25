// Every preset shape and text warp: complete, evaluable at any size and adjust value, and still what Annex D says.
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { childElements, getAttr, parseXml, type XmlElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import {
  evaluateGeometry, pathBounds, PRESET_NAMES, presetGeometry, TEXT_WARP_NAMES, textWarpGeometry,
  type AdjustHandle, type EvaluatedGeometry, type Guide, type ShapeGeometry,
} from '../src/index.ts';
import { DEFAULT_DIR, generate, OUTPUTS, presetShapes, readDefinitions, textWarps } from '../../../tools/geometry/gen-presets.ts';
import { numbersOf, valueOf } from './helpers.ts';

const ALL: [string, ShapeGeometry][] = [
  ...PRESET_NAMES.map((n): [string, ShapeGeometry] => [n, presetGeometry(n)!]),
  ...TEXT_WARP_NAMES.map((n): [string, ShapeGeometry] => [n, textWarpGeometry(n)!]),
];

/** The adjust values a handle moves, with their limits as written: [gdRef, min, max]. */
const axesOf = (h: AdjustHandle): [string | undefined, string | undefined, string | undefined][] =>
  h.kind === 'xy' ? [[h.gdRefX, h.minX, h.maxX], [h.gdRefY, h.minY, h.maxY]] : [[h.gdRefAng, h.minAng, h.maxAng], [h.gdRefR, h.minR, h.maxR]];

function finite(e: EvaluatedGeometry, what: string): void {
  expect(numbersOf(e).filter(v => !Number.isFinite(v)), what).toEqual([]);
}

describe('names', () => {
  test('187 preset shapes and 40 text warps, sorted', () => {
    expect(PRESET_NAMES).toHaveLength(187);
    expect(TEXT_WARP_NAMES).toHaveLength(40);
    expect([...PRESET_NAMES].sort()).toEqual(PRESET_NAMES);
    expect(PRESET_NAMES).toContain('upArrow');
    expect(TEXT_WARP_NAMES).not.toContain('textNoShape');
  });

  test('a name that is not a preset gives undefined, whatever Object.prototype has', () => {
    for (const name of ['nope', 'constructor', '__proto__', 'toString', 'hasOwnProperty', '', 'RECT']) {
      expect(presetGeometry(name), name).toBeUndefined();
      expect(textWarpGeometry(name), name).toBeUndefined();
    }
    expect(presetGeometry('textArchUp')).toBeUndefined();
    expect(textWarpGeometry('rect')).toBeUndefined();
  });

  test('presets are shared and frozen', () => {
    const g = presetGeometry('roundRect')!;
    expect(presetGeometry('roundRect')).toBe(g);
    expect(() => (g.avLst as Guide[]).push({ name: 'x', fmla: 'val 1' })).toThrow(TypeError);
    expect(() => { (g.pathLst[0] as { w: number }).w = 5; }).toThrow(TypeError);
  });
});

describe.each(ALL)('%s', (name, geom) => {
  // Degenerate and huge sizes divide by zero and overflow single steps; nothing may come out NaN or infinite.
  const SIZES = [[0, 0], [1, 100000], [100000, 1], [1e8, 1e8], [1600000, 1100000], [1100000, 1600000]] as const;

  test('evaluates to finite numbers at every size, with each handle at its limits', () => {
    for (const [w, h] of SIZES) {
      const base = evaluateGeometry(geom, w, h);
      finite(base, `${w} x ${h}`);
      geom.ahLst.forEach((handle, i) => {
        const limits = axesOf(handle)
          .filter((a): a is [string, string, string] => a[0] !== undefined && a[1] !== undefined && a[2] !== undefined)
          .map(([ref, min, max]) => [ref, valueOf(base, min), valueOf(base, max)] as const);
        const combos: Guide[][] = limits.flatMap(([ref, min, max]) => [[{ name: ref, fmla: `val ${min}` }], [{ name: ref, fmla: `val ${max}` }]]);
        if (limits.length === 2) {
          const [[r1, a1, b1], [r2, a2, b2]] = limits as unknown as [[string, number, number], [string, number, number]];
          for (const v1 of [a1, b1]) for (const v2 of [a2, b2]) combos.push([{ name: r1, fmla: `val ${v1}` }, { name: r2, fmla: `val ${v2}` }]);
        }
        for (const adjust of combos) finite(evaluateGeometry(geom, w, h, adjust), `${w} x ${h}, handle ${i}: ${JSON.stringify(adjust)}`);
      });
      // and adjust values no handle would produce
      for (const v of [-2147483647, 0, 2147483647]) {
        finite(evaluateGeometry(geom, w, h, geom.avLst.map(g => ({ name: g.name, fmla: `val ${v}` }))), `${w} x ${h}, every adjust value ${v}`);
      }
    }
  });

  test('draws something, each subpath starting with M', () => {
    const e = evaluateGeometry(geom, 1600000, 1100000);
    expect(e.paths.length).toBeGreaterThan(0);
    for (const p of e.paths) expect(p.commands[0]?.op).toBe('M');
    const b = pathBounds(e.paths);
    expect(b.w + b.h).toBeGreaterThan(0);
  });
});

describe('every preset shape', () => {
  test('keeps its text rectangle inside its box', () => {
    for (const [w, h] of [[1600000, 1100000], [1100000, 1600000], [1000000, 1000000]]) {
      for (const name of PRESET_NAMES) {
        const t = evaluateGeometry(presetGeometry(name)!, w!, h!).textRect;
        const inside = t.x >= -1e-6 && t.y >= -1e-6 && t.x + t.w <= w! + 1e-6 && t.y + t.h <= h! + 1e-6 && t.w > 0 && t.h > 0;
        expect(inside, `${name} at ${w} x ${h}: ${JSON.stringify(t)}`).toBe(true);
      }
    }
  });

  test('puts its connection sites on its box or its outline', () => {
    for (const [w, h] of [[1600000, 1100000], [1100000, 1600000]]) {
      for (const name of PRESET_NAMES) {
        const e = evaluateGeometry(presetGeometry(name)!, w!, h!);
        const b = pathBounds(e.paths);
        const x0 = Math.min(0, b.x), y0 = Math.min(0, b.y), x1 = Math.max(w!, b.x + b.w), y1 = Math.max(h!, b.y + b.h);
        for (const s of e.connectionSites) {
          expect(s.x >= x0 - 1e-6 && s.x <= x1 + 1e-6 && s.y >= y0 - 1e-6 && s.y <= y1 + 1e-6, `${name} at ${w} x ${h}: ${JSON.stringify(s)}`).toBe(true);
        }
      }
    }
  });
});

const ANNEX_D = existsSync(join(DEFAULT_DIR, 'presetShapeDefinitions.xml')) && existsSync(join(DEFAULT_DIR, 'presetTextWarpDefinitions.xml'));

describe.skipIf(!ANNEX_D)('the tables against Annex D', () => {
  test('the generated files are what the generator makes now', () => {
    const out = generate();
    expect(readFileSync(OUTPUTS.presets, 'utf8') === out.presets, 'run node tools/geometry/gen-presets.ts').toBe(true);
    expect(readFileSync(OUTPUTS.textWarps, 'utf8') === out.textWarps, 'run node tools/geometry/gen-presets.ts').toBe(true);
  });

  test('each decodes to the definition read from the XML, corrections applied', () => {
    const shapes = presetShapes();
    expect([...shapes.keys()].sort()).toEqual(PRESET_NAMES);
    for (const name of PRESET_NAMES) expect(presetGeometry(name), name).toEqual(shapes.get(name));
    const warps = textWarps();
    expect([...warps.keys()].sort()).toEqual(TEXT_WARP_NAMES);
    for (const name of TEXT_WARP_NAMES) expect(textWarpGeometry(name), name).toEqual(warps.get(name));
  });

  test('the corrections are needed: each error is in the file as published', () => {
    const raw = readDefinitions(join(DEFAULT_DIR, 'presetShapeDefinitions.xml'));
    expect(raw.has('upArrow')).toBe(false);
    expect(raw.get('curvedLeftArrow')!.cxnLst.map(c => c.ang)).toContain('cd3');
    expect(() => evaluateGeometry(raw.get('curvedLeftArrow')!, 100, 100)).toThrow('unknown guide "cd3"');
    for (const n of ['circularArrow', 'leftCircularArrow', 'leftRightCircularArrow']) {
      expect(() => evaluateGeometry(raw.get(n)!, 100, 100), n).toThrow('"+-" takes 3 arguments, not 4');
    }
    expect(raw.get('pie')!.rect).toEqual({ l: 'il', t: 'ir', r: 'it', b: 'ib' });
    expect(raw.get('squareTabs')!.cxnLst.some(c => c.pos.x === 'dx' && c.pos.y === 'x1')).toBe(true);
    // and nothing else changed
    const fixed = presetShapes();
    const changed = PRESET_NAMES.filter(n => JSON.stringify(raw.get(n)) !== JSON.stringify(fixed.get(n)));
    expect(changed).toEqual(['circularArrow', 'curvedLeftArrow', 'leftCircularArrow', 'leftRightCircularArrow', 'pie', 'squareTabs', 'upArrow']);
  });

  const XSD = join(DEFAULT_DIR, '../transitional/dml-main.xsd');
  test.skipIf(!existsSync(XSD))('the names are ST_ShapeType and ST_TextShapeType of the schema', () => {
    const schema = parseXml(readFileSync(XSD, 'utf8')).root;
    const enumeration = (type: string) => {
      const simple = childElements(schema).find(e => e.local === 'simpleType' && getAttr(e, 'name') === type)!;
      const values: string[] = [];
      const walk = (el: XmlElement) => {
        if (el.local === 'enumeration') values.push(getAttr(el, 'value')!);
        childElements(el).forEach(walk);
      };
      walk(simple);
      return values.sort();
    };
    expect(enumeration('ST_ShapeType')).toEqual(PRESET_NAMES);
    expect(enumeration('ST_TextShapeType').filter(n => n !== 'textNoShape')).toEqual(TEXT_WARP_NAMES);
  });
});
