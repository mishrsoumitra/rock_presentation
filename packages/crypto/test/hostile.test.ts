// Hostile and malformed containers must fail cleanly and quickly, never hang. The real risk is cfb's
// get_sector_list following a FAT chain with no cycle check, so the headline case is a compound file whose FAT
// loops: assertSafeContainer must reject it before cfb ever walks it. Each decrypt runs in a worker with a
// deadline so that if anything ever did loop, the test fails instead of hanging the run.
import { readdirSync } from 'node:fs';
import { Worker } from 'node:worker_threads';
import { describe, expect, test } from 'vitest';
import { assertSafeContainer, decryptPackage, isEncryptedPackage, OfficeCryptoError } from '../src/index.ts';

interface Outcome { ok: boolean; code?: string; message?: string; length?: number; ms: number; }

function decryptInWorker(bytes: Uint8Array, password = 'x', timeoutMs = 5000): Promise<Outcome> {
  return new Promise((resolve, reject) => {
    const started = performance.now();
    const worker = new Worker(new URL('./fixtures/decrypt-worker.ts', import.meta.url), { workerData: { bytes, password } });
    const timer = setTimeout(() => { void worker.terminate(); reject(new Error(`no answer within ${timeoutMs} ms`)); }, timeoutMs);
    worker.once('message', (m: Omit<Outcome, 'ms'>) => { clearTimeout(timer); void worker.terminate(); resolve({ ...m, ms: performance.now() - started }); });
    worker.once('error', e => { clearTimeout(timer); reject(e); });
  });
}

const CFB_SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** A minimal v3 compound file header plus `sectorData` laid out as 512-byte sectors after it. */
function compoundFile(header: Partial<Record<string, number>>, difat: number[], sectors: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(512 + sectors.length * 512);
  out.set(CFB_SIG);
  const dv = new DataView(out.buffer);
  dv.setUint16(24, 0x003e, true);                 // minor version
  dv.setUint16(26, 3, true);                      // major version 3
  dv.setUint16(28, 0xfffe, true);                 // byte order
  dv.setUint16(30, 9, true);                      // sector shift (512)
  dv.setUint16(32, 6, true);                      // mini sector shift
  dv.setUint32(44, header.numFat ?? 1, true);     // number of FAT sectors
  dv.setUint32(48, header.firstDir ?? 0xffffffff, true);
  dv.setUint32(56, 0x1000, true);                 // mini stream cutoff
  dv.setUint32(60, 0xfffffffe, true);             // first mini FAT: ENDOFCHAIN
  dv.setUint32(68, header.firstDifat ?? 0xfffffffe, true);
  dv.setUint32(72, header.numDifat ?? 0, true);
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, difat[i] ?? 0xffffffff, true);
  for (let i = 0; i < sectors.length; i++) out.set(sectors[i]!.subarray(0, 512), 512 + i * 512);
  return out;
}

/** A FAT sector (128 entries) from a list of next-pointers; the rest are FREESECT. */
function fatSector(entries: Record<number, number>): Uint8Array {
  const s = new Uint8Array(512);
  const dv = new DataView(s.buffer);
  for (let i = 0; i < 128; i++) dv.setUint32(i * 4, entries[i] ?? 0xffffffff, true);
  return s;
}

describe('a FAT that loops is rejected before cfb walks it', () => {
  test('self-referential FAT chain -> malformed, fast', () => {
    // Sector 0 is the FAT; FAT[0]=0 (self-loop), FAT[2]=0 (so a stream starting at sector 2 would loop 2->0->0...).
    const fat = fatSector({ 0: 0, 1: 0xfffffffe, 2: 0 });
    const file = compoundFile({ numFat: 1, firstDir: 1 }, [0], [fat, new Uint8Array(512), new Uint8Array(512)]);
    const started = performance.now();
    expect(() => assertSafeContainer(file, { maxBytes: 1 << 20, maxSectors: 1 << 20 })).toThrow(OfficeCryptoError);
    try { assertSafeContainer(file, { maxBytes: 1 << 20, maxSectors: 1 << 20 }); } catch (e) {
      expect((e as OfficeCryptoError).code).toBe('malformed');
    }
    expect(performance.now() - started).toBeLessThan(500);
  });

  test('a looping DIFAT chain is rejected', () => {
    // DIFAT sector 0's last entry points back to sector 0.
    const difatSector = fatSector({ 127: 0 });
    const file = compoundFile({ numFat: 0, firstDifat: 0, numDifat: 2 }, [], [difatSector]);
    expect(() => assertSafeContainer(file, { maxBytes: 1 << 20, maxSectors: 1 << 20 })).toThrow(/DIFAT chain cycles/);
  });
});

describe('malformed and truncated containers fail cleanly in a worker', () => {
  test('a cyclic-FAT file never hangs the worker', async () => {
    const fat = fatSector({ 0: 0, 1: 0xfffffffe, 2: 0 });
    const file = compoundFile({ numFat: 1, firstDir: 1 }, [0], [fat, new Uint8Array(512), new Uint8Array(512)]);
    const r = await decryptInWorker(file);
    expect(r.ok).toBe(false);
    expect(['malformed', 'not-encrypted']).toContain(r.code);
    expect(r.ms).toBeLessThan(2000);
  });

  test('garbage with the CFB signature, a truncated file, and random bytes', async () => {
    const sig = new Uint8Array(2048); sig.set(CFB_SIG);
    const inputs: Uint8Array[] = [
      sig,
      new Uint8Array([...CFB_SIG, 0, 0, 0, 0]),                 // too short
      new Uint8Array(1024).fill(0xff).map((v, i) => (i < 8 ? CFB_SIG[i]! : v)), // 0xff header
      new Uint8Array(600).map((_, i) => (i * 97) & 0xff),       // no signature at all
    ];
    for (const input of inputs) {
      const r = await decryptInWorker(input);
      expect(r.ok, `message: ${r.message}`).toBe(false);
      expect(r.code, `code was ${r.code}`).not.toMatch(/^crash:/);
      expect(r.ms).toBeLessThan(2000);
    }
  });

  test('a plain (unencrypted) ZIP is reported as not-encrypted', async () => {
    const { readFixture } = await import('./fixtures.ts');
    await expect(decryptPackage(readFixture('plain/default.pptx'), 'x')).rejects.toMatchObject({ code: 'not-encrypted' });
  });
});

// If the corpus is present, the real POI clusterfuzz samples (compound files with an EncryptionInfo) must also
// fail cleanly. They are skipped when the gitignored corpus has not been fetched.
describe('corpus fuzzer samples (when present)', () => {
  const dir = new URL('../../../corpus/external/poi/', import.meta.url);
  let files: string[] = [];
  try { files = readdirSync(dir).filter(n => n.includes('clusterfuzz') && n.endsWith('.ppt')); } catch { files = []; }

  test.skipIf(files.length === 0)('each fuzzer compound file fails cleanly and fast', async () => {
    const { readFileSync } = await import('node:fs');
    for (const name of files) {
      const bytes = new Uint8Array(readFileSync(new URL(name, dir)));
      if (!isEncryptedPackage(bytes)) continue; // only the ones that look like encrypted packages
      const r = await decryptInWorker(bytes);
      expect(r.ok, `${name}: ${r.message}`).toBe(false);
      expect(r.code, `${name} crashed: ${r.message}`).not.toMatch(/^crash:/);
      expect(r.ms, `${name} was slow`).toBeLessThan(3000);
    }
  });
});
