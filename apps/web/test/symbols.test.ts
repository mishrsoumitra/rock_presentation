// The Symbol menu's characters: each one a single character with a name, none offered twice.
import { describe, expect, test } from 'vitest';
import { ALL_SYMBOLS, SYMBOL_GROUPS } from '../src/ui/symbols.ts';

describe('the characters the Symbol menu offers', () => {
  test('every entry is one character with a name', () => {
    expect(ALL_SYMBOLS.length).toBeGreaterThan(80);
    for (const s of ALL_SYMBOLS) {
      expect([...s.char].length, s.name).toBe(1);
      expect(s.name.trim().length, s.char).toBeGreaterThan(1);
      expect(s.char.charCodeAt(0), s.name).toBeGreaterThan(127);         // nothing a keyboard already types
    }
  });

  test('no character and no name is offered twice, and every group has a title and characters', () => {
    expect(new Set(ALL_SYMBOLS.map(s => s.char)).size).toBe(ALL_SYMBOLS.length);
    expect(new Set(ALL_SYMBOLS.map(s => s.name)).size).toBe(ALL_SYMBOLS.length);
    for (const [title, chars] of SYMBOL_GROUPS) {
      expect(title.length).toBeGreaterThan(2);
      expect(chars.length).toBeGreaterThan(5);
    }
  });
});
