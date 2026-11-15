// Version history's rules: which versions make room for new ones, when one is kept by itself, and how a time reads.
import { describe, expect, test } from 'vitest';
import { AUTO_EVERY, dueForAuto, KEEP, prune, sizeText, versions, whenText } from '../src/engine/versions.ts';

const v = (id: number, at: number, auto: boolean) => ({ id, at, auto });

describe('making room', () => {
  test('nothing goes while there is room', () => {
    expect(prune([], 3)).toEqual([]);
    expect(prune([v(1, 10, true), v(2, 20, false), v(3, 30, true)], 3)).toEqual([]);
    expect(KEEP).toBe(20);
  });

  test('the oldest the editor made by itself go first, however old the named ones are', () => {
    const list = [v(1, 10, false), v(2, 20, true), v(3, 30, true), v(4, 40, false), v(5, 50, true)];
    expect(prune(list, 4)).toEqual([2]);
    expect(prune(list, 3)).toEqual([2, 3]);
    expect(prune(list, 2)).toEqual([2, 3, 5]);
  });

  test('with no automatic ones left to drop, the oldest named ones go', () => {
    const list = [v(1, 10, false), v(2, 20, true), v(3, 30, false), v(4, 40, false)];
    expect(prune(list, 2)).toEqual([2, 1]);
    expect(prune([v(1, 10, false), v(2, 20, false), v(3, 30, false)], 1)).toEqual([1, 2]);
  });

  test('order given does not matter, and versions kept at the same moment go by id', () => {
    expect(prune([v(5, 50, true), v(2, 20, true), v(3, 30, true)], 2)).toEqual([2]);
    expect(prune([v(7, 10, true), v(6, 10, true), v(8, 10, true)], 1)).toEqual([6, 7]);
  });
});

describe('keeping a version by itself', () => {
  test('ten minutes after the last one, and not before', () => {
    expect(AUTO_EVERY).toBe(600_000);
    expect(dueForAuto(1_000_000, 1_000_000 + AUTO_EVERY - 1)).toBe(false);
    expect(dueForAuto(1_000_000, 1_000_000 + AUTO_EVERY)).toBe(true);
    expect(dueForAuto(0, 5000, 5000)).toBe(true);
  });
});

describe('how a version reads in the list', () => {
  const now = new Date(2026, 10, 16, 15, 30).getTime();
  test('today, yesterday, an earlier day this year, and another year', () => {
    expect(whenText(new Date(2026, 10, 16, 9, 5).getTime(), now, 'en-US')).toBe('Today, 9:05 AM');
    expect(whenText(new Date(2026, 10, 15, 23, 59).getTime(), now, 'en-US')).toBe('Yesterday, 11:59 PM');
    expect(whenText(new Date(2026, 10, 2, 14, 0).getTime(), now, 'en-US')).toBe('Nov 2, 2:00 PM');
    expect(whenText(new Date(2025, 11, 31, 8, 0).getTime(), now, 'en-US')).toBe('Dec 31, 2025, 8:00 AM');
  });

  test('just after midnight, a minute ago was yesterday', () => {
    const midnight = new Date(2026, 10, 16, 0, 0, 30).getTime();
    expect(whenText(midnight - 60_000, midnight, 'en-US')).toBe('Yesterday, 11:59 PM');
  });

  test('sizes in the unit that fits', () => {
    expect(sizeText(512)).toBe('512 B');
    expect(sizeText(30 * 1024)).toBe('30 KB');
    expect(sizeText(2.5 * 1024 * 1024)).toBe('2.5 MB');
  });
});

describe('where the browser gives no database', () => {
  test('nothing is kept and nothing fails', async () => {
    expect(versions.available()).toBe(false);                             // these tests run outside a browser
    expect(await versions.add({ deck: 'a.pptx', label: 'x', at: 1, bytes: new Uint8Array(3), encrypted: false, slides: 1, auto: false })).toBeUndefined();
    expect(await versions.list('a.pptx')).toEqual([]);
    expect(await versions.get(1)).toBeUndefined();
    await expect(versions.remove(1)).resolves.toBeUndefined();
  });
});
