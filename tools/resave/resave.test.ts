import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { packageDifferences } from '../corpus/compare.ts';
import { deck } from '../../packages/opc/test/fixtures.ts';
import { main, resaveBytes } from './resave.ts';

function run(args: string[]) {
  const out: string[] = [];
  const err: string[] = [];
  const code = main(args, { log: l => out.push(l), error: l => err.push(l) });
  return { code, out: out.join('\n'), err: err.join('\n') };
}

const workdir = () => mkdtempSync(join(tmpdir(), 'resave-test-'));

describe('resave', () => {
  test('writes a verified copy', () => {
    const dir = workdir();
    const input = join(dir, 'in.pptx');
    const output = join(dir, 'out.pptx');
    writeFileSync(input, deck());
    const r = run([input, output]);
    expect(r.code).toBe(0);
    expect(r.out).toContain('ok ');
    expect(packageDifferences(readFileSync(input), readFileSync(output))).toEqual([]);
    expect(existsSync(`${output}.partial`)).toBe(false);
  });

  test('refuses what it cannot open, with the reason, and writes nothing', () => {
    const dir = workdir();
    const input = join(dir, 'old.ppt');
    const output = join(dir, 'out.pptx');
    writeFileSync(input, new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...new Uint8Array(1016)]));
    const r = run([input, output]);
    expect(r.code).toBe(2);
    expect(r.err).toContain('refused');
    expect(existsSync(output)).toBe(false);
  });

  test('never overwrites the input, and needs --force to replace another file', () => {
    const dir = workdir();
    const input = join(dir, 'in.pptx');
    const other = join(dir, 'other.pptx');
    writeFileSync(input, deck());
    writeFileSync(other, 'keep me');
    expect(run([input, input]).code).toBe(64);
    expect(run([input, other]).code).toBe(64);
    expect(readFileSync(other, 'utf8')).toBe('keep me');
    expect(run([input, other, '--force']).code).toBe(0);
    expect(OpcPackage.open(readFileSync(other)).partNames()).toContain('/ppt/slides/slide1.xml');
  });

  test('--check verifies without writing, and reports the worst outcome', () => {
    const dir = workdir();
    const good = join(dir, 'good.pptx');
    const bad = join(dir, 'bad.pptx');
    writeFileSync(good, deck());
    writeFileSync(bad, 'not a deck');
    expect(run(['--check', good]).code).toBe(0);
    const r = run(['--check', good, bad]);
    expect(r.code).toBe(2);
    expect(r.out).toContain('good.pptx');
    expect(r.err).toContain('bad.pptx');
  });

  test('usage errors', () => {
    expect(run([]).code).toBe(64);
    expect(run(['a.pptx']).code).toBe(64);
    expect(run(['--bogus', 'a', 'b']).code).toBe(64);
  });

  test('resaveBytes reports a lossless save', () => {
    const outcome = resaveBytes(deck());
    expect(outcome.status).toBe('ok');
  });
});
