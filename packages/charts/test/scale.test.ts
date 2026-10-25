import { describe, expect, test } from 'vitest';
import { autoScale, ticks, type Scale } from '../src/scale.ts';

/** At most n major intervals: the shape of the length-based test the axes pass in. */
const atMost = (n: number) => (s: Scale) => (s.max - s.min) / s.major <= n + 1e-9;

describe('automatic axis scaling', () => {
  test('zero is the minimum unless the data sits far from it; 5% headroom; the smallest unit that fits', () => {
    // PowerPoint's choices for the corpus' bar and line charts of 1.2..8.2
    expect(autoScale({ lo: 1.2, hi: 8.2, fits: atMost(10) })).toMatchObject({ min: 0, max: 9, major: 1 });
    expect(autoScale({ lo: 1.2, hi: 8.2, fits: atMost(6) })).toMatchObject({ min: 0, max: 10, major: 2 });
    // 5 plus 5% is past 5: the next unit up
    expect(autoScale({ lo: 1.8, hi: 5, fits: atMost(10) })).toMatchObject({ min: 0, max: 6, major: 1 });
    expect(autoScale({ lo: 1.8, hi: 4.4, fits: atMost(10) })).toMatchObject({ min: 0, max: 5, major: 0.5 });
  });

  test('data far from zero keeps a minimum of its own', () => {
    const s = autoScale({ lo: 95, hi: 100, fits: atMost(10) });
    expect(s.min).toBeGreaterThan(80);
    expect(s.min).toBeLessThanOrEqual(95);
    expect(s.max).toBeGreaterThanOrEqual(100);
  });

  test('negative data', () => {
    expect(autoScale({ lo: -4.4, hi: 5, fits: atMost(10) })).toMatchObject({ min: -6, max: 6, major: 2 });
    expect(autoScale({ lo: -30, hi: -10, fits: atMost(10) })).toMatchObject({ max: 0 });
  });

  test('3-D charts get no headroom', () => {
    expect(autoScale({ lo: 2, hi: 5, tight: true, fits: atMost(6) })).toMatchObject({ min: 0, max: 5, major: 1 });
    expect(autoScale({ lo: 2, hi: 5, fits: atMost(6) })).toMatchObject({ min: 0, max: 6, major: 1 });
    expect(autoScale({ lo: 20, hi: 90, tight: true, fits: atMost(10) })).toMatchObject({ min: 0, max: 90, major: 10 });
  });

  test('values from the file win', () => {
    expect(autoScale({ lo: 0, hi: 8, min: 2, max: 7, fits: atMost(10) })).toMatchObject({ min: 2, max: 7 });
    expect(autoScale({ lo: 0, hi: 8, majorUnit: 4 })).toMatchObject({ min: 0, max: 12, major: 4 });
    expect(autoScale({ lo: 0, hi: 8, minorUnit: 0.5, fits: atMost(10) }).minor).toBe(0.5);
  });

  test('the minor unit is a fifth of the major', () => {
    expect(autoScale({ lo: 0, hi: 9, fits: atMost(5) }).minor).toBeCloseTo(0.4);
  });

  test('percent stacks span 0 to 100%', () => {
    expect(autoScale({ lo: 0, hi: 1, percent: true, fits: atMost(5) })).toMatchObject({ min: 0, max: 1, major: 0.2 });
    expect(autoScale({ lo: -0.3, hi: 1, percent: true, fits: atMost(10) })).toMatchObject({ min: -1, max: 1 });
  });

  test('flat and empty data', () => {
    expect(autoScale({ lo: 0, hi: 0 })).toMatchObject({ min: 0, max: 1 });
    expect(autoScale({ lo: 5, hi: 5, fits: atMost(10) })).toMatchObject({ min: 0, max: 6, major: 1 });
  });

  test('logarithmic axes run between powers of the base', () => {
    expect(autoScale({ lo: 3, hi: 4000, logBase: 10 })).toMatchObject({ min: 1, max: 10000, logBase: 10 });
    expect(ticks(autoScale({ lo: 3, hi: 4000, logBase: 10 }))).toEqual([1, 10, 100, 1000, 10000]);
  });

  test('ticks are clean decimals', () => {
    expect(ticks({ min: 0, max: 1, major: 0.1, minor: 0.02 })).toEqual([0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1]);
    expect(ticks({ min: -6, max: 6, major: 2, minor: 0.4 })).toEqual([-6, -4, -2, 0, 2, 4, 6]);
  });

  test('hostile inputs stay bounded', () => {
    const s = autoScale({ lo: 0, hi: 1e300, fits: () => false });
    expect(Number.isFinite(s.max) && Number.isFinite(s.major)).toBe(true);
    expect(ticks({ min: 0, max: 1e9, major: 1e-9, minor: 1 }).length).toBeLessThanOrEqual(1001);
    expect(ticks({ min: 0, max: 10, major: 0, minor: 0 })).toEqual([0, 10]);
  });
});
