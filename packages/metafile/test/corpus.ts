// The metafile parts of the corpus decks, shared by the corpus and fuzz tests.
import { readFileSync } from 'node:fs';
import { readZip } from '../../opc/src/index.ts';
import { corpusFiles, corpusName } from '../../../tools/corpus/list.ts';

export interface CorpusPicture { name: string; bytes: Uint8Array; }

let cache: CorpusPicture[] | undefined;

/** Every ppt/media part named .emf or .wmf in every corpus deck that opens, in a stable order. */
export function corpusMetafiles(): CorpusPicture[] {
  if (cache) return cache;
  cache = [];
  for (const file of corpusFiles) {
    let entries;
    try {
      entries = readZip(new Uint8Array(readFileSync(file)));
    } catch {
      continue;              // the refused containers (fuzzer archives and such): the opc tests cover those
    }
    for (const e of entries) {
      if (/\.(emf|wmf)$/i.test(e.name)) cache.push({ name: `${corpusName(file)}!${e.name}`, bytes: e.data });
    }
  }
  return cache;
}

let tiffs: CorpusPicture[] | undefined;

/** Every ppt/media part that is a TIFF file, whatever its name, in a stable order. */
export function corpusTiffs(): CorpusPicture[] {
  if (tiffs) return tiffs;
  tiffs = [];
  for (const file of corpusFiles) {
    let entries;
    try {
      entries = readZip(new Uint8Array(readFileSync(file)));
    } catch {
      continue;
    }
    for (const e of entries) {
      const b = e.data;
      if (e.name.startsWith('ppt/media/') && b.length > 8 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 42) || (b[0] === 0x4d && b[1] === 0x4d && b[3] === 42))) {
        tiffs.push({ name: `${corpusName(file)}!${e.name}`, bytes: b });
      }
    }
  }
  return tiffs;
}
