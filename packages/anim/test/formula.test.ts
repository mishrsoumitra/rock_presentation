// The formula language of p:tav fmla and formula values: what PowerPoint's presets write, the operators and
// functions of ECMA-376 19.5.79, and input from hostile files, which must fail cleanly and quickly.
import { describe, expect, test } from 'vitest';
import { compileFormula, evaluateFormula, FormulaError, MAX_FORMULA_LENGTH, type FormulaScope } from '../src/index.ts';

const scope: FormulaScope = { ppt_x: 0.3, ppt_y: 0.6, ppt_w: 0.2, ppt_h: 0.1 };
const ev = (s: string, input?: number) => evaluateFormula(s, { ...scope, ...(input !== undefined ? { input } : {}) });

describe('values PowerPoint writes', () => {
  test.each([
    ['#ppt_x', 0.3], ['#ppt_y', 0.6], ['1+#ppt_h/2', 1.05], ['0-#ppt_w/2', -0.1], ['#ppt_y+.1', 0.7], ['2/3*#ppt_w', (2 / 3) * 0.2],
    ['#ppt_w*.05', 0.01], ['(0.5)', 0.5], ['(#ppt_y+0.4)', 1.0], ['(-#ppt_w/2)', -0.1], ['(#ppt_h/3+#ppt_w*0.1)', 0.1 / 3 + 0.02],
    // exit effects write the variables without #
    ['ppt_x', 0.3], ['1+ppt_h/2', 1.05],
  ])('%s', (src, want) => {
    expect(ev(src)).toBeCloseTo(want, 12);
  });

  test('the spiral of Spiral In ends where the shape is, whatever $ started at', () => {
    const x = '#ppt_x+(cos(-2*pi*(1-$))*-#ppt_x-sin(-2*pi*(1-$))*(1-#ppt_y))*(1-$)';
    const y = '#ppt_y+(sin(-2*pi*(1-$))*-#ppt_x+cos(-2*pi*(1-$))*(1-#ppt_y))*(1-$)';
    expect(ev(x, 1)).toBeCloseTo(0.3, 12);
    expect(ev(y, 1)).toBeCloseTo(0.6, 12);
    // at the start the shape sits at the left edge, at the bottom
    expect(ev(x, 0)).toBeCloseTo(0, 12);
    expect(ev(y, 0)).toBeCloseTo(1, 12);
  });

  test('the example of ECMA-376 19.5.79: #ppt_y-sin(pi*$)/3', () => {
    expect(ev('#ppt_y-sin(pi*$)/3', 0.5)).toBeCloseTo(0.6 - 1 / 3, 12);
    expect(ev('#ppt_y-sin(pi*$)/3', 1)).toBeCloseTo(0.6, 12);
  });
});

describe('the language', () => {
  test.each([
    ['1', 1], ['.5', 0.5], ['2.5e-3', 0.0025], ['1E2', 100], ['2+3*4', 14], ['(2+3)*4', 20], ['10%4', 2], ['7-2-1', 4], ['8/2/2', 2],
    ['2*3^2', 18], ['2^3^2', 512], ['-2^2', 4], ['2^-1', 0.5], ['3*-2', -6], ['+3', 3], ['--3', 3], ['pi', Math.PI], ['e', Math.E],
    ['abs(-3)', 3], ['acos(1)', 0], ['asin(1)', Math.PI / 2], ['atan(1)', Math.PI / 4], ['ceil(1.2)', 2], ['floor(1.8)', 1], ['cos(0)', 1],
    ['cosh(0)', 1], ['sin(pi/2)', 1], ['sinh(0)', 0], ['tan(0)', 0], ['tanh(0)', 0], ['sqrt(16)', 4], ['exp(0)', 1], ['ln(e)', 1],
    ['deg(pi)', 180], ['rad(180)', Math.PI], ['max(2,3)', 3], ['min(2,3)', 2], ['MAX( 2 , 3 )', 3], ['$*2', 0.5],
  ])('%s', (src, want) => {
    expect(ev(src, 0.25)).toBeCloseTo(want, 12);
  });

  test('rand(x) is in [0, x) and the same each time by default, so a frame always looks the same', () => {
    const a = ev('rand(10)');
    expect(a).toBeGreaterThanOrEqual(0);
    expect(a).toBeLessThan(10);
    expect(evaluateFormula('rand(1)', { ...scope, random: () => 0.25 })).toBe(0.25);
  });

  test('arithmetic follows IEEE: 1/0 is infinite, not an exception', () => {
    expect(ev('1/0')).toBe(Infinity);
    expect(ev('9^9^9^9')).toBe(Infinity);
  });
});

describe('hostile input fails cleanly', () => {
  test.each([
    '', ' ', 'constructor', '__proto__', 'this', 'process.exit(1)', 'globalThis', 'alert(1)', 'eval("1")', 'Function("x")', 'x=1', 'a;b',
    '`1`', '"1"', '{}', '[1]', '1 2', '(1', '1)', '1+', '*1', 'max(1)', 'sin(1,2)', 'foo(1)', 'ppt_z', '#', '#1', '1..2', 'max(1,2,3)',
  ])('%j', src => {
    expect(() => compileFormula(src)).toThrow(FormulaError);
    expect(evaluateFormula(src, scope)).toBeNaN();
  });

  test('length, token count and nesting are bounded', () => {
    expect(() => compileFormula('1'.repeat(MAX_FORMULA_LENGTH + 1))).toThrow(/longer than/);
    expect(() => compileFormula(Array.from({ length: 400 }, () => '1').join('+'))).toThrow(/tokens/);
    expect(() => compileFormula(`${'('.repeat(100)}1${')'.repeat(100)}`)).toThrow(/nested/);
    expect(() => compileFormula(`${'-'.repeat(100)}1`)).toThrow(/nested/);
    expect(() => compileFormula(`2${'^2'.repeat(60)}`)).toThrow(/nested/);
    // well within the limits
    expect(compileFormula(`${'('.repeat(20)}1${')'.repeat(20)}`)(scope)).toBe(1);
  });

  test('a long hostile formula is rejected in microseconds, not evaluated', () => {
    const started = performance.now();
    for (let i = 0; i < 1000; i++) evaluateFormula(`sin(${'1+'.repeat(300)}1)`, scope);
    expect(performance.now() - started).toBeLessThan(2000);
  });
});
