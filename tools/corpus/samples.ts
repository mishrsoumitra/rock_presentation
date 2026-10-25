// tools/corpus/samples.ts  -- XML parts from the corpus, for parser tests that need real-world input
import { readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';
import { corpusFiles, corpusName } from './list.ts';

export interface XmlSample { name: string; text: string; }

/** Up to `perFile` XML parts (each under `maxBytes`) from about `files` packages spread evenly across the corpus. */
export function xmlSamples(options: { files?: number; perFile?: number; maxBytes?: number } = {}): XmlSample[] {
  const { files = 60, perFile = 8, maxBytes = 200_000 } = options;
  const step = Math.max(1, Math.floor(corpusFiles.length / files));
  const out: XmlSample[] = [];
  const decoder = new TextDecoder('utf-8', { fatal: true });
  for (const file of corpusFiles.filter((_, i) => i % step === 0)) {
    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipSync(new Uint8Array(readFileSync(file)), { filter: e => /\.(xml|rels)$/i.test(e.name) && e.originalSize < maxBytes });
    } catch {
      continue;                                   // a broken package; the round-trip test covers those
    }
    for (const [name, bytes] of Object.entries(entries).slice(0, perFile)) {
      try {
        out.push({ name: `${corpusName(file)}:${name}`, text: decoder.decode(bytes) });
      } catch {
        // not UTF-8; the encoding tests cover UTF-16
      }
    }
  }
  return out;
}
