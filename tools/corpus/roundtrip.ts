// tools/corpus/roundtrip.ts  -- `pnpm corpus:roundtrip [--out out/corpus-roundtrip.json]`
//
// The phase 1 numbers for the fidelity dashboard: for every package in the corpus, whether open -> save is lossless,
// whether re-serializing every XML part is lossless, and why a file was refused. Writes a JSON report and prints a
// summary. Exits non-zero on any lossy save or on a refusal that is not on the reviewed list.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { OpcError, OpcPackage } from '@pptx/opc';
import { isXmlEntry, packageDifferences } from './compare.ts';
import { CORPUS_ROOT, corpusName, opcFiles } from './list.ts';

interface FileResult {
  file: string;
  source: string;
  bytes: number;
  status: 'lossless' | 'refused' | 'lossy';
  refusal?: { code: string; expected: boolean };
  reserialized?: 'lossless' | 'lossy';
  parts?: number;
  xmlParts?: number;
  unparseableXmlParts?: number;
  differences?: string[];
  diagnostics?: string[];
  ms: number;
}

function check(file: string, expected: Record<string, { code: string }>): FileResult {
  const started = performance.now();
  const name = corpusName(file);
  const original = new Uint8Array(readFileSync(file));
  const base = { file: name, source: name.split('/')[0] ?? '', bytes: original.length };
  let pkg: OpcPackage;
  try {
    pkg = OpcPackage.open(original);
  } catch (e) {
    if (!(e instanceof OpcError)) throw e;
    return { ...base, status: 'refused', refusal: { code: e.code, expected: expected[name]?.code === e.code }, ms: performance.now() - started };
  }
  const differences = packageDifferences(original, pkg.save());

  // the stricter variant: every XML part that parses is written from its tree, as an edited part would be
  let xmlParts = 0;
  let unparseable = 0;
  for (const part of pkg.partNames()) {
    if (!isXmlEntry(part)) continue;
    xmlParts++;
    try {
      pkg.xml(part);
      pkg.markDirty(part);
    } catch {
      unparseable++;
    }
  }
  const reserializeDifferences = packageDifferences(original, pkg.save());
  return {
    ...base,
    status: differences.length ? 'lossy' : 'lossless',
    reserialized: reserializeDifferences.length ? 'lossy' : 'lossless',
    parts: pkg.partNames().length,
    xmlParts,
    unparseableXmlParts: unparseable,
    ...(differences.length || reserializeDifferences.length ? { differences: [...differences, ...reserializeDifferences].slice(0, 20) } : {}),
    ...(pkg.diagnostics.length ? { diagnostics: [...pkg.diagnostics] } : {}),
    ms: performance.now() - started,
  };
}

export function run(outPath: string): number {
  const expected = JSON.parse(readFileSync(join(CORPUS_ROOT, 'expected-refusals.json'), 'utf8')) as Record<string, { code: string }>;
  if (opcFiles.length === 0) {
    console.error('The corpus is empty: run `pnpm corpus:fetch` first.');
    return 1;
  }
  const results = opcFiles.map(f => check(f, expected));

  const count = (pred: (r: FileResult) => boolean) => results.filter(pred).length;
  const opened = results.filter(r => r.status !== 'refused');
  const bySource: Record<string, { files: number; lossless: number; refused: number }> = {};
  for (const r of results) {
    const s = (bySource[r.source] ??= { files: 0, lossless: 0, refused: 0 });
    s.files++;
    if (r.status === 'lossless' && r.reserialized === 'lossless') s.lossless++;
    if (r.status === 'refused') s.refused++;
  }
  const refusals: Record<string, number> = {};
  for (const r of results) if (r.refusal) refusals[r.refusal.code] = (refusals[r.refusal.code] ?? 0) + 1;
  const summary = {
    generated: new Date().toISOString(),
    files: results.length,
    opened: opened.length,
    lossless: count(r => r.status === 'lossless'),
    lossy: count(r => r.status === 'lossy'),
    reserializedLossless: count(r => r.reserialized === 'lossless'),
    reserializedLossy: count(r => r.reserialized === 'lossy'),
    refused: count(r => r.status === 'refused'),
    unexpectedRefusals: count(r => r.refusal?.expected === false),
    xmlParts: opened.reduce((n, r) => n + (r.xmlParts ?? 0), 0),
    unparseableXmlParts: opened.reduce((n, r) => n + (r.unparseableXmlParts ?? 0), 0),
    withDiagnostics: count(r => (r.diagnostics?.length ?? 0) > 0),
    refusals,
    bySource,
  };
  mkdirSync(dirname(outPath), { recursive: true });
  writeFileSync(outPath, JSON.stringify({ summary, results }, null, 2) + '\n');

  const pct = (n: number, d: number) => (d === 0 ? '-' : `${((100 * n) / d).toFixed(2)}%`);
  console.log(`corpus round trip: ${summary.files} packages, ${summary.opened} opened, ${summary.refused} refused`);
  console.log(`  open -> save lossless:          ${summary.lossless}/${summary.opened} (${pct(summary.lossless, summary.opened)})`);
  console.log(`  every XML part re-serialized:   ${summary.reserializedLossless}/${summary.opened} (${pct(summary.reserializedLossless, summary.opened)})`);
  console.log(`  XML parts: ${summary.xmlParts}, of which not well-formed (kept as bytes): ${summary.unparseableXmlParts}`);
  console.log(`  refusals: ${Object.entries(refusals).map(([k, v]) => `${k} ${v}`).join(', ') || 'none'}; not on the reviewed list: ${summary.unexpectedRefusals}`);
  for (const [s, v] of Object.entries(bySource)) console.log(`  ${s.padEnd(12)} ${v.lossless}/${v.files - v.refused} lossless, ${v.refused} refused`);
  for (const r of results.filter(r => r.status === 'lossy' || r.reserialized === 'lossy' || r.refusal?.expected === false)) {
    console.log(`  PROBLEM ${r.file}: ${r.status}${r.refusal ? ` (${r.refusal.code}, not expected)` : ''} ${r.differences?.[0] ?? ''}`);
  }
  console.log(`report: ${resolve(outPath)}`);
  return summary.lossy + summary.reserializedLossy + summary.unexpectedRefusals === 0 ? 0 : 1;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const at = process.argv.indexOf('--out');
  const out = at > 0 && process.argv[at + 1] ? process.argv[at + 1]! : join(CORPUS_ROOT, '../out/corpus-roundtrip.json');
  process.exitCode = run(out);
}
