// tools/resave/resave.ts  -- milestone M1: re-save any deck through the lossless package layer, verified first.
//
//   pnpm resave <input> <output> [--force]   write the re-saved deck, once it has been checked against the input
//   pnpm resave --check <input>...           only check: open, save in memory, compare; nothing is written
//
// The check reads both files back with an independent reader (fflate's unzip) and requires every entry to survive:
// XML parts equal in canonical form, every other part byte for byte. Nothing is written unless it passes.
// Exit codes: 0 all good, 1 a save was not lossless (nothing written), 2 a file could not be opened, 64 usage error.
import { existsSync, readFileSync, realpathSync, renameSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { OpcError, OpcPackage } from '@pptx/opc';
import { packageDifferences } from '../corpus/compare.ts';

export type Outcome =
  | { status: 'ok'; saved: Uint8Array; parts: number; diagnostics: string[] }
  | { status: 'refused'; code: string; reason: string }
  | { status: 'not-lossless'; differences: string[] };

/** Open, save and verify in memory. */
export function resaveBytes(input: Uint8Array): Outcome {
  let pkg: OpcPackage;
  try {
    pkg = OpcPackage.open(input);
  } catch (e) {
    if (e instanceof OpcError) return { status: 'refused', code: e.code, reason: e.message };
    throw e;
  }
  const saved = pkg.save();
  const differences = packageDifferences(input, saved);
  if (differences.length) return { status: 'not-lossless', differences };
  return { status: 'ok', saved, parts: pkg.partNames().length, diagnostics: [...pkg.diagnostics] };
}

interface Io { log(line: string): void; error(line: string): void; }
const USAGE = 'usage: resave <input> <output> [--force]\n       resave --check <input>...';

export function main(argv: string[], io: Io = { log: l => console.log(l), error: l => console.error(l) }): number {
  const flags = new Set(argv.filter(a => a.startsWith('--')));
  const files = argv.filter(a => !a.startsWith('--'));
  const unknown = [...flags].filter(f => f !== '--check' && f !== '--force');
  if (unknown.length || files.length === 0 || (!flags.has('--check') && files.length !== 2)) {
    io.error(unknown.length ? `unknown option ${unknown.join(' ')}\n${USAGE}` : USAGE);
    return 64;
  }

  const report = (file: string, outcome: Outcome, bytesIn: number): number => {
    if (outcome.status === 'refused') {
      io.error(`refused  ${file}: ${outcome.reason} [${outcome.code}]`);
      return 2;
    }
    if (outcome.status === 'not-lossless') {
      io.error(`FAILED   ${file}: the re-saved package would differ; nothing written`);
      for (const d of outcome.differences.slice(0, 10)) io.error(`         ${d}`);
      return 1;
    }
    io.log(`ok       ${file}: ${outcome.parts} parts, ${bytesIn} -> ${outcome.saved.length} bytes`);
    for (const d of outcome.diagnostics) io.log(`         note: ${d}`);
    return 0;
  };

  if (flags.has('--check')) {
    let worst = 0;
    for (const file of files) {
      const input = new Uint8Array(readFileSync(file));
      worst = Math.max(worst, report(file, resaveBytes(input), input.length));
    }
    return worst;
  }

  const [inputPath, outputPath] = files as [string, string];
  if (existsSync(outputPath)) {
    if (realpathSync(outputPath) === realpathSync(inputPath)) {
      io.error('refusing to overwrite the input; choose another output path');
      return 64;
    }
    if (!flags.has('--force')) {
      io.error(`${outputPath} exists; pass --force to replace it`);
      return 64;
    }
  }
  const input = new Uint8Array(readFileSync(inputPath));
  const outcome = resaveBytes(input);
  const code = report(inputPath, outcome, input.length);
  if (outcome.status !== 'ok') return code;
  const temporary = `${outputPath}.partial`;                 // never leave a half-written deck under the real name
  writeFileSync(temporary, outcome.saved);
  renameSync(temporary, outputPath);
  io.log(`wrote    ${resolve(outputPath)}`);
  return 0;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) process.exitCode = main(process.argv.slice(2));
