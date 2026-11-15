// The presenter view's timer text.
import { expect, test } from 'vitest';
import { clock } from '../src/ui/clock.ts';

test('time reads as minutes and seconds, with hours once there are any, and never negative', () => {
  expect(clock(0)).toBe('0:00');
  expect(clock(999)).toBe('0:00');
  expect(clock(7_000)).toBe('0:07');
  expect(clock(65_400)).toBe('1:05');
  expect(clock(12 * 60_000 + 30_000)).toBe('12:30');
  expect(clock(59 * 60_000 + 59_999)).toBe('59:59');
  expect(clock(3_600_000)).toBe('1:00:00');
  expect(clock(3_600_000 + 2 * 60_000 + 3_000)).toBe('1:02:03');
  expect(clock(-5_000)).toBe('0:00');
});
