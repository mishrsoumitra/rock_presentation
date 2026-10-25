import { describe, expect, test } from 'vitest';
import { formatGeneral, formatValue, isDateFormat, serialToDate } from '../src/numfmt.ts';

describe('General', () => {
  test.each([
    [0, '0'], [1, '1'], [-5, '-5'], [0.5, '0.5'], [0.1 + 0.2, '0.3'], [8.200000000000001, '8.2'], [1234567.891, '1234567.891'],
    [123456789012, '1.23457E+11'], [0.000000000123, '1.23E-10'], [1 / 3, '0.333333333'], [-1 / 3, '-0.33333333'],
  ])('%s -> %s', (v, s) => expect(formatGeneral(v)).toBe(s));

  test('General is the default and is case-insensitive', () => {
    expect(formatValue(2.5, '')).toBe('2.5');
    expect(formatValue(2.5, 'general')).toBe('2.5');
  });
});

describe('number formats', () => {
  test.each([
    ['0', 4.5, '5'], ['0', -4.5, '-5'], ['0', 0.4, '0'],
    ['0.0', 3, '3.0'], ['0.00', 1.005, '1.01'], ['0.0#', 2.5, '2.5'], ['0.0#', 2.567, '2.57'], ['0.0#', 2, '2.0'],
    ['#,##0', 1234567, '1,234,567'], ['#,##0.00', -1234.5, '-1,234.50'], ['#,##0,', 1234567, '1,235'], ['#,##0,,"M"', 2500000, '3M'],
    ['0%', 0.256, '26%'], ['0.00%', 0.12345, '12.35%'], ['0%', -0.5, '-50%'],
    ['$#,##0.00', 1234.5, '$1,234.50'], ['"$"#,##0', 42, '$42'], ['[$€-407] #,##0.00', 12.5, '€ 12.50'], ['#,##0 [$kr.-406]', 7, '7 kr.'],
    ['0.00E+00', 12345, '1.23E+04'], ['0.0E+0', 0.00042, '4.2E-4'], ['##0.0E+0', 12345, '12.3E+3'],
    ['000', 7, '007'], ['#.00', 0.5, '.50'], ['# ?/?', 1.5, '1 1/2'], ['?/8', 0.375, '3/8'],
    ['\\(0\\)', 3, '(3)'], ['0" units"', 4, '4 units'], ['[Red]0;[Blue]-0', -3, '-3'],
  ])('%s with %s -> %s', (code, v, s) => expect(formatValue(v as number, code as string)).toBe(s));

  test('sections: positive; negative; zero', () => {
    expect(formatValue(5, '0.0;(0.0);"zero"')).toBe('5.0');
    expect(formatValue(-5, '0.0;(0.0);"zero"')).toBe('(5.0)');
    expect(formatValue(0, '0.0;(0.0);"zero"')).toBe('zero');
    // two sections: the second is for negatives and writes its own sign
    expect(formatValue(-2, '0;0')).toBe('2');
    expect(formatValue(0, '0;-0')).toBe('0');
  });

  test("the accounting format Office writes for currency columns", () => {
    const code = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-';
    expect(formatValue(1234, code)).toBe('1,234');
    expect(formatValue(-1234, code)).toBe('-1,234');
    expect(formatValue(0, code)).toBe('-');
  });

  test('conditions pick the section', () => {
    expect(formatValue(1500, '[>=1000]#,##0,"K";0')).toBe('2K');
    expect(formatValue(500, '[>=1000]#,##0,"K";0')).toBe('500');
  });

  test('a negative number that rounds to zero has no sign', () => {
    expect(formatValue(-0.004, '0.00')).toBe('0.00');
  });
});

describe('dates', () => {
  test('serials in the 1900 system, with its phantom 29 February 1900', () => {
    expect(serialToDate(1)).toMatchObject({ y: 1900, m: 1, d: 1 });
    expect(serialToDate(59)).toMatchObject({ y: 1900, m: 2, d: 28 });
    expect(serialToDate(60)).toMatchObject({ y: 1900, m: 2, d: 29 });
    expect(serialToDate(61)).toMatchObject({ y: 1900, m: 3, d: 1 });
    expect(serialToDate(37261)).toMatchObject({ y: 2002, m: 1, d: 5, wd: 6 });
    expect(serialToDate(37261.75)).toMatchObject({ h: 18, mi: 0, s: 0 });
  });

  test('serials in the 1904 system', () => {
    expect(serialToDate(0, true)).toMatchObject({ y: 1904, m: 1, d: 1 });
    expect(formatValue(35799, 'm/d/yyyy', true)).toBe('1/5/2002');
  });

  test.each([
    ['m/d/yy', 37261, '1/5/02'], ['m/d/yyyy', 37261, '1/5/2002'], ['d-mmm-yy', 37261, '5-Jan-02'], ['mmm yyyy', 37261, 'Jan 2002'],
    ['mmmm d, yyyy', 37261, 'January 5, 2002'], ['dddd', 37261, 'Saturday'], ['ddd dd', 37261, 'Sat 05'], ['mmmmm', 37261, 'J'],
    ['yyyy-mm-dd', 37261, '2002-01-05'], ['d.m.yyyy', 37261, '5.1.2002'], ['[$-409]mmm-yy;@', 37261, 'Jan-02'],
    ['h:mm AM/PM', 0.75, '6:00 PM'], ['hh:mm:ss', 0.5 + 1 / 86400, '12:00:01'], ['h:mm', 37261.25, '6:00'], ['[h]:mm', 1.5, '36:00'],
    ['mm:ss.0', 1.5 / 86400, '00:01.5'],
  ])('%s', (code, v, s) => expect(formatValue(v, code)).toBe(s));

  test('date formats are recognised', () => {
    expect(isDateFormat('m/d/yy')).toBe(true);
    expect(isDateFormat('[$-409]mmm-yy;@')).toBe(true);
    expect(isDateFormat('h:mm')).toBe(true);
    expect(isDateFormat('#,##0.00')).toBe(false);
    expect(isDateFormat('General')).toBe(false);
    expect(isDateFormat('0.00E+00')).toBe(false);
  });
});
