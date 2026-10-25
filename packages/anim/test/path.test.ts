// Motion paths: the commands of p:animMotion's path, numbers as PowerPoint writes them, and steady speed along them.
import { describe, expect, test } from 'vitest';
import { filterReveal, parseMotionPath, pointAt, supportedFilter } from '../src/index.ts';

describe('parseMotionPath', () => {
  test('numbers with exponents are numbers, not commands ("5E-6")', () => {
    const p = parseMotionPath('M 5E-6 -2.22222E-6 L 0.13126 0.38472 ')!;
    expect(p.points).toHaveLength(2);
    expect(pointAt(p, 1)).toEqual({ x: 0.13126, y: 0.38472 });
  });

  test('relative commands, curves, close and end', () => {
    const p = parseMotionPath('M 0 0 l 0.1 0 l 0 0.1 Z E L 5 5')!;
    expect(p.points.map(q => [q.x, q.y])).toEqual([[0, 0], [0.1, 0], [0.1, 0.1], [0, 0]]);
    const c = parseMotionPath('M 0 0  c -0.004 -0.01066  -0.018 -0.02132  -0.023 -0.02132  ')!;
    expect(pointAt(c, 1)).toEqual({ x: -0.023, y: -0.02132 });
    expect(c.points.length).toBe(17);
  });

  test('points are spaced by distance, so the shape moves at a steady speed', () => {
    const p = parseMotionPath('M 0 0 L 0.3 0 L 0.3 0.1 E')!;
    expect(pointAt(p, 0.5)).toEqual({ x: 0.2, y: 0 });
    expect(pointAt(p, 0.875).y).toBeCloseTo(0.05, 12);
  });

  test('garbage gives what it can, or nothing', () => {
    expect(parseMotionPath('')).toBeUndefined();
    expect(parseMotionPath('Q 1 2')).toBeUndefined();
    expect(parseMotionPath('M 0 0 L 1')?.points).toEqual([{ x: 0, y: 0 }]);
    expect(parseMotionPath(`M 0 0 ${'L 1 1 '.repeat(20_000)}`)!.points.length).toBeLessThanOrEqual(10_001);
  });
});

describe('filters', () => {
  test('names from files never reach Object.prototype; unknown filters fade', () => {
    for (const name of ['constructor', '__proto__', 'toString', 'hasOwnProperty']) {
      expect(supportedFilter(name)).toBeUndefined();
      expect(filterReveal(name, 0.25)).toEqual({ opacity: 0.25 });
    }
    expect(supportedFilter('blinds(diagonal);blinds(vertical)')).toBe('blinds(vertical)');
  });

  test('an exit hides what the entrance would reveal; slide moves the content back out', () => {
    const area = (r: { x: number; y: number }[][] | undefined) => (r ?? []).reduce((a, poly) => a + poly.reduce((s, p, i) => {
      const q = poly[(i + 1) % poly.length]!;
      return s + (p.x * q.y - q.x * p.y) / 2;
    }, 0), 0);
    expect(area(filterReveal('box(in)', 0.3).region)).toBeCloseTo(1 - 0.7 * 0.7, 9);
    expect(area(filterReveal('box(in)', 0.3, 'out').region)).toBeCloseTo(0.7 * 0.7, 9);
    expect(filterReveal('slide(fromBottom)', 0.25, 'out').shift).toEqual({ x: 0, y: 0.25 });
    expect(filterReveal('fade', 0.25, 'out').opacity).toBe(0.75);
  });
});
