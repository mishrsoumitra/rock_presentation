// The ruler's marks: where they fall and which are numbered, as the zoom changes; and where a new guide goes.
import { describe, expect, test } from 'vitest';
import { newGuideAt, rulerTicks } from '../src/ui/rulers.ts';

const labels = (origin: number, ppi: number, length: number) => rulerTicks(origin, ppi, length).filter(t => t.label !== undefined).map(t => [t.at, t.label]);

describe('ruler marks', () => {
  test('an inch at 96 pixels: eighths marked, every inch numbered from the slide\'s edge', () => {
    const ticks = rulerTicks(0, 96, 300);
    expect(ticks.length).toBe(26);                                        // 0 to 3.125 inches in eighths
    expect(ticks.filter(t => t.size === 'inch').map(t => t.at)).toEqual([0, 96, 192, 288]);
    expect(ticks.filter(t => t.size === 'half').map(t => t.at)).toEqual([48, 144, 240]);
    expect(labels(0, 96, 300)).toEqual([[0, '0'], [96, '1'], [192, '2'], [288, '3']]);
  });

  test('the slide\'s edge away from the ruler\'s start: marks before it count away from the slide, unsigned', () => {
    expect(labels(100, 50, 260)).toEqual([[0, '2'], [50, '1'], [100, '0'], [150, '1'], [200, '2'], [250, '3']]);
    expect(rulerTicks(100, 50, 260).every(t => t.at >= 0 && t.at <= 260)).toBe(true);
  });

  test('zoomed out, fewer marks and fewer numbers: halves go below 20 pixels an inch, numbers thin to every 2, 5, 10', () => {
    expect(rulerTicks(0, 50, 100).map(t => t.size)).toEqual(['inch', 'half', 'inch', 'half', 'inch']);
    expect(rulerTicks(0, 16, 64).map(t => t.size)).toEqual(['inch', 'inch', 'inch', 'inch', 'inch']);
    expect(labels(0, 16, 100).map(l => l[1])).toEqual(['0', '2', '4', '6']);
    expect(labels(0, 6, 100).map(l => l[1])).toEqual(['0', '5', '10', '15']);
    expect(labels(0, 3, 100).map(l => l[1])).toEqual(['0', '10', '20', '30']);
  });

  test('nothing to show for no length or no zoom', () => {
    expect(rulerTicks(0, 0, 100)).toEqual([]);
    expect(rulerTicks(0, 96, 0)).toEqual([]);
    expect(rulerTicks(0, Number.NaN, 100)).toEqual([]);
  });
});

describe('where a new guide goes', () => {
  const W = 12192000;
  test('the middle of the slide first, then half an inch right of it, then left, and so on outwards', () => {
    expect(newGuideAt([], W)).toBe(6096000);
    expect(newGuideAt([6096000], W)).toBe(6096000 + 457200);
    expect(newGuideAt([6096000, 6096000 + 457200], W)).toBe(6096000 - 457200);
    expect(newGuideAt([1000], W)).toBe(6096000);
  });

  test('never off the slide', () => {
    const taken = [500000];
    for (let i = 0; i < 6; i++) taken.push(newGuideAt(taken, 1000000));
    expect(taken.every(g => g >= 0 && g <= 1000000)).toBe(true);
  });
});
