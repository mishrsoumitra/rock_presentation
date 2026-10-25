// Stricter than the plan's round trip: every XML part that parses is re-serialized, as an edited part would be,
// and the result must still read the same, part by part, for every package in the corpus.
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '../src/index.ts';
import { corpusName, opcFiles } from '../../../tools/corpus/list.ts';
import { expectedRefusals, isXmlEntry, packageDifferences } from './corpus-helpers.ts';

const openable = opcFiles.filter(f => !expectedRefusals[corpusName(f)]);

describe.skipIf(openable.length === 0)('re-serializing every XML part', () => {
  test.each(openable.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    const original = new Uint8Array(readFileSync(file));
    const pkg = OpcPackage.open(original);
    let reserialized = 0;
    for (const name of pkg.partNames()) {
      if (!isXmlEntry(name)) continue;
      try {
        pkg.xml(name);
      } catch {
        continue;            // not well-formed: left untouched, and the comparison requires identical bytes
      }
      pkg.markDirty(name);
      reserialized++;
    }
    expect(reserialized).toBeGreaterThan(0);
    const saved = pkg.save();
    expect(packageDifferences(original, saved)).toEqual([]);
    expect(OpcPackage.open(saved).partNames()).toEqual(pkg.partNames());
  });
});
