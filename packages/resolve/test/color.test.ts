import { describe, expect, test } from 'vitest';
import type { ColorDef, ColorTransform } from '@pptx/model';
import { resolveColor, toHex, type ColorContext } from '../src/index.ts';

const theme = new Map<string, ColorDef>([
  ['dk1', { base: { kind: 'sys', name: 'windowText', last: { r: 0, g: 0, b: 0 } }, transforms: [] }],
  ['lt1', { base: { kind: 'sys', name: 'window', last: { r: 255, g: 255, b: 255 } }, transforms: [] }],
  ['accent1', { base: { kind: 'srgb', r: 0x44, g: 0x72, b: 0xc4 }, transforms: [] }],
]);
const ctx: ColorContext = { theme, clrMap: { bg1: 'lt1', tx1: 'dk1', accent1: 'accent1' } };
const scheme = (name: string, ...transforms: [string, number][]): ColorDef =>
  ({ base: { kind: 'scheme', name }, transforms: transforms.map(([n, v]): ColorTransform => ({ name: n, val: v })) });
const hex = (def: ColorDef, c: ColorContext = ctx) => toHex(resolveColor(def, c));

/** Equal within one step per channel: PowerPoint rounds its own way. */
function near(actual: string, expected: string) {
  const ch = (s: string, i: number) => Number.parseInt(s.slice(1 + 2 * i, 3 + 2 * i), 16);
  for (let i = 0; i < 3; i++) expect(Math.abs(ch(actual, i) - ch(expected, i)), `${actual} vs ${expected}`).toBeLessThanOrEqual(1);
}

describe('theme colour variants, against the values PowerPoint shows in its colour picker', () => {
  test.each([
    ['Accent 1, Darker 25%', [['lumMod', 0.75]], '#2F5597'],
    ['Accent 1, Darker 50%', [['lumMod', 0.5]], '#203864'],
    ['Accent 1, Lighter 40%', [['lumMod', 0.6], ['lumOff', 0.4]], '#8FAADC'],
    ['Accent 1, Lighter 60%', [['lumMod', 0.4], ['lumOff', 0.6]], '#B4C7E7'],
    ['Accent 1, Lighter 80%', [['lumMod', 0.2], ['lumOff', 0.8]], '#DAE3F3'],
  ] as [string, [string, number][], string][])('%s', (_name, t, expected) => near(hex(scheme('accent1', ...t)), expected));

  test('Text 1, Lighter 25% and Background 1, Darker 15%', () => {
    near(hex(scheme('tx1', ['lumMod', 0.75], ['lumOff', 0.25])), '#404040');
    near(hex(scheme('bg1', ['lumMod', 0.85])), '#D9D9D9');
  });
});

describe('tint and shade in linear light', () => {
  test("the default shape outline, accent 1 with a 50% shade, is PowerPoint's #2F528F", () => {
    expect(hex(scheme('accent1', ['shade', 0.5]))).toBe('#2F528F');
  });

  test('a 75% tint of black is the #898989 PowerPoint draws footer text in', () => {
    expect(hex(scheme('tx1', ['tint', 0.75]))).toBe('#898989');
  });

  test('extremes', () => {
    expect(hex(scheme('accent1', ['shade', 0]))).toBe('#000000');
    expect(hex(scheme('accent1', ['tint', 0]))).toBe('#FFFFFF');
    expect(hex(scheme('accent1', ['tint', 1]))).toBe('#4472C4');
  });
});

describe('bases and other transforms', () => {
  test('the colour map decides what bg1 and tx1 mean', () => {
    const dark: ColorContext = { ...ctx, clrMap: { bg1: 'dk1', tx1: 'lt1' } };
    expect(hex(scheme('bg1'), dark)).toBe('#000000');
    expect(hex(scheme('tx1'), dark)).toBe('#FFFFFF');
  });

  test('phClr is the colour of the style reference being applied', () => {
    expect(hex(scheme('phClr'), { ...ctx, phClr: { r: 1, g: 0, b: 0, a: 1 } })).toBe('#FF0000');
    expect(hex(scheme('phClr'))).toBe('#000000');
  });

  test('preset, system, HSL and scRGB colours', () => {
    expect(hex({ base: { kind: 'preset', name: 'dkBlue' }, transforms: [] })).toBe('#00008B');
    expect(hex({ base: { kind: 'preset', name: 'cornflowerBlue' }, transforms: [] })).toBe('#6495ED');
    expect(hex({ base: { kind: 'sys', name: 'highlight' }, transforms: [] })).toBe('#0078D7');
    expect(hex({ base: { kind: 'hsl', h: 120, s: 1, l: 0.5 }, transforms: [] })).toBe('#00FF00');
    expect(hex({ base: { kind: 'scrgb', r: 1, g: 0.5, b: 0 }, transforms: [] })).toBe('#FFBC00');
  });

  test('alpha, gray, inverse, complement, hue offset', () => {
    const c = resolveColor(scheme('accent1', ['alpha', 0.6], ['alphaMod', 0.5]), ctx);
    expect(c.a).toBeCloseTo(0.3);
    expect(hex(scheme('accent1', ['inv', 0]))).toBe('#BB8D3B');
    const g = resolveColor(scheme('accent1', ['gray', 0]), ctx);
    expect(g.r).toBeCloseTo(g.b);
    near(hex(scheme('accent1', ['comp', 0])), '#C49644');
    near(hex(scheme('accent1', ['hueOff', 180])), '#C49644');
  });

  test('a missing theme slot falls back to black instead of failing', () => {
    expect(hex(scheme('accent6'))).toBe('#000000');
  });
});
