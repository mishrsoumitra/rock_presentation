// Plan section 1.5: open then save with no edits, for every package in the corpus.
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { OpcError, OpcPackage } from '../src/index.ts';
import { corpusName, legacyPptFiles, opcFiles } from '../../../tools/corpus/list.ts';
import { expectedRefusals, packageDifferences, sameBytes } from './corpus-helpers.ts';

describe.skipIf(opcFiles.length === 0)('open then save with no edits', () => {
  test.each(opcFiles.map(f => [corpusName(f), f]))('%s', (name, file) => {
    const original = new Uint8Array(readFileSync(file));
    const refusal = expectedRefusals[name];
    if (refusal) {
      // a file on the reviewed list must still be refused, with the same reason
      expect(() => OpcPackage.open(original)).toThrow(OpcError);
      try {
        OpcPackage.open(original);
      } catch (e) {
        expect((e as OpcError).code).toBe(refusal.code);
      }
      return;
    }
    const saved = OpcPackage.open(original).save();
    expect(packageDifferences(original, saved)).toEqual([]);
    expect(sameBytes(OpcPackage.open(saved).save(), saved), 'saving again gives the same bytes').toBe(true);
  });

  test('every file on the refusal list is in the corpus', () => {
    const names = new Set(opcFiles.map(corpusName));
    expect(Object.keys(expectedRefusals).filter(n => !names.has(n))).toEqual([]);
  });
});

describe.skipIf(legacyPptFiles.length === 0)('legacy binary .ppt files are refused with a reason', () => {
  test.each(legacyPptFiles.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    let error: unknown;
    try {
      OpcPackage.open(new Uint8Array(readFileSync(file)));
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(OpcError);
    expect(['legacy-ppt', 'encrypted', 'unsupported-container', 'not-a-package']).toContain((error as OpcError).code);
  });
});

test.skipIf(opcFiles.length > 0)('corpus not downloaded: run `pnpm corpus:fetch` to enable the corpus tests', () => {});
