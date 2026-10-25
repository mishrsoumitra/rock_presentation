// Shared by the corpus round-trip tests.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { CORPUS_ROOT } from '../../../tools/corpus/list.ts';

export { isXmlEntry, packageDifferences, sameBytes } from '../../../tools/corpus/compare.ts';

/** Files the package layer must refuse, each with the error code it must refuse them with and why (reviewed by hand). */
export const expectedRefusals: Record<string, { code: string; why: string }> =
  JSON.parse(readFileSync(join(CORPUS_ROOT, 'expected-refusals.json'), 'utf8'));
