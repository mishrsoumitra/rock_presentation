// tools/corpus/list.ts  -- the files every corpus-driven test runs over
import { existsSync, readdirSync, statSync } from 'node:fs';
import { dirname, extname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

export const CORPUS_ROOT = join(dirname(fileURLToPath(import.meta.url)), '../../corpus');

const PRESENTATION = new Set(['.pptx', '.pptm', '.potx', '.potm', '.ppsx', '.ppsm']);

function walk(dir: string): string[] {
  if (!existsSync(dir)) return [];
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    if (name.startsWith('.')) continue;                       // corpus/.cache holds the git checkouts
    const p = join(dir, name);
    if (statSync(p).isDirectory()) out.push(...walk(p));
    else out.push(p);
  }
  return out;
}

const all = walk(CORPUS_ROOT).sort();

/** Every OOXML presentation in the corpus: .pptx, .pptm, .potx, .potm, .ppsx, .ppsm. */
export const corpusFiles: string[] = all.filter(f => PRESENTATION.has(extname(f).toLowerCase()));
/** Every OPC package the round-trip test covers: presentations plus Office theme files (.thmx). */
export const opcFiles: string[] = all.filter(f => PRESENTATION.has(extname(f).toLowerCase()) || extname(f).toLowerCase() === '.thmx');
/** Legacy binary .ppt files, which the editor must refuse with a clear message. */
export const legacyPptFiles: string[] = all.filter(f => extname(f).toLowerCase() === '.ppt');

/** A short, stable name for test titles and reports, e.g. "poi/test-data~slideshow~sample.pptx". */
export const corpusName = (file: string) => relative(join(CORPUS_ROOT, 'external'), file);
