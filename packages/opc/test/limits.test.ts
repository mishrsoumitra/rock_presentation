// Hostile containers: each must fail cleanly, quickly, and inside a worker so the host process never notices.
import { Worker } from 'node:worker_threads';
import { strToU8, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { readZip } from '../src/index.ts';
import { deckFiles, rawZip, zipFiles } from './fixtures.ts';

interface Outcome { ok: boolean; code?: string; message?: string; diagnostics?: string[]; ms: number; }

function openInWorker(bytes: Uint8Array, parse?: string, timeoutMs = 5000): Promise<Outcome> {
  return new Promise((resolve, reject) => {
    const started = performance.now();
    const worker = new Worker(new URL('./fixtures/open-worker.ts', import.meta.url), {
      workerData: { bytes, parse },
      resourceLimits: { maxOldGenerationSizeMb: 1024 },
    });
    const timer = setTimeout(() => {
      void worker.terminate();
      reject(new Error(`no answer within ${timeoutMs} ms`));
    }, timeoutMs);
    worker.once('message', (m: Omit<Outcome, 'ms'>) => {
      clearTimeout(timer);
      void worker.terminate();
      resolve({ ...m, ms: performance.now() - started });
    });
    worker.once('error', e => {
      clearTimeout(timer);
      reject(e);
    });
  });
}

const CT = deckFiles()['[Content_Types].xml'] as string;
const tiny = strToU8('<r/>');

describe('container limits, checked before anything is inflated', () => {
  test('an entry declaring 10 GB through ZIP64', async () => {
    const zip = rawZip([{ name: '[Content_Types].xml', content: strToU8(CT) }, { name: 'ppt/bomb.xml', content: tiny, zip64: true, declaredSize: 10e9 }], { zip64Locator: true });
    const r = await openInWorker(zip);
    expect(r).toMatchObject({ ok: false, code: 'limit' });
    expect(r.message).toContain('10000000000');
  });

  test('sizes of 0xFFFFFFFF without ZIP64 records', async () => {
    const zip = rawZip([{ name: '[Content_Types].xml', content: strToU8(CT) }, { name: 'ppt/x.xml', content: tiny, zip64ExtraMissing: true }]);
    expect(await openInWorker(zip)).toMatchObject({ ok: false, code: 'limit' });
  });

  test('the malformed ZIP64 entry that sent fflate 0.8.2 into an infinite loop (GHSA-px8p-9vwx-vf98)', async () => {
    const zip = rawZip([{ name: '[Content_Types].xml', content: strToU8(CT) }, { name: 'ppt/x.xml', content: tiny, zip64ExtraMissing: true }], { zip64Locator: true });
    expect(await openInWorker(zip)).toMatchObject({ ok: false, code: 'invalid-zip' });
  });

  test('entries that are each allowed but together declare more than the total limit', async () => {
    const zip = rawZip([
      { name: '[Content_Types].xml', content: strToU8(CT) },
      { name: 'a.bin', content: tiny, declaredSize: 900_000_000 },
      { name: 'b.bin', content: tiny, declaredSize: 900_000_000 },
      { name: 'c.bin', content: tiny, declaredSize: 900_000_000 },
    ]);
    const r = await openInWorker(zip);
    expect(r).toMatchObject({ ok: false, code: 'limit' });
    expect(r.ms).toBeLessThan(2000);   // nothing was allocated: the first pass refused it
  });

  test('more ZIP entries than the limit', async () => {
    const files: Record<string, Uint8Array> = { '[Content_Types].xml': strToU8(CT) };
    for (let i = 0; i < 20_000; i++) files[`ppt/e${i}.bin`] = new Uint8Array(0);
    expect(await openInWorker(zipSync(files, { level: 0 }))).toMatchObject({ ok: false, code: 'limit' });
  });

  test('an end record that claims 200,000 entries over a short central directory', async () => {
    const zip = rawZip([{ name: '[Content_Types].xml', content: strToU8(CT) }], { zip64Locator: true, claimedEntries: 200_000 });
    const r = await openInWorker(zip);
    expect(r.ok).toBe(false);
    expect(['limit', 'invalid-zip', 'invalid-package']).toContain(r.code);
  });

  test('duplicate entry names', async () => {
    const zip = rawZip([{ name: '[Content_Types].xml', content: strToU8(CT) }, { name: 'a.xml', content: tiny }, { name: 'a.xml', content: tiny }]);
    expect(await openInWorker(zip)).toMatchObject({ ok: false, code: 'invalid-package' });
  });

  test('an unsupported compression method', async () => {
    const zip = rawZip([{ name: '[Content_Types].xml', content: strToU8(CT) }, { name: 'a.xml', content: tiny, method: 14 }]);
    expect(await openInWorker(zip)).toMatchObject({ ok: false, code: 'invalid-zip' });
  });

  test('truncated and garbage archives', async () => {
    const good = zipFiles(deckFiles());
    for (const bad of [good.subarray(0, good.length - 30), good.subarray(0, 100), new Uint8Array([0x50, 0x4b, 3, 4, 0, 0, 0])]) {
      const r = await openInWorker(bad);
      expect(r.ok).toBe(false);
      expect(r.code).toBe('invalid-zip');
    }
  });
});

describe('entries whose content disagrees with their declared size', () => {
  test('an entry shorter than it declares is opened and reported', async () => {
    const zip = rawZip([{ name: '[Content_Types].xml', content: strToU8(CT) }, { name: 'ppt/x.bin', content: new Uint8Array(10), declaredSize: 1000 }]);
    const r = await openInWorker(zip);
    expect(r.ok).toBe(true);
    expect(r.diagnostics).toContain('ZIP entry ppt/x.bin holds 10 bytes but declares 1000');
  });

  test('an entry longer than it declares cannot grow past the declared size', () => {
    const entries = readZip(rawZip([{ name: 'big.bin', content: new Uint8Array(1 << 20), declaredSize: 16 }]));
    expect(entries[0]!.data.length).toBe(16);
  });
});

describe('XML parts that would cost too much to parse', () => {
  test('a part nested a million levels deep', async () => {
    const deep = '<a>'.repeat(1_000_000) + '</a>'.repeat(1_000_000);
    const zip = zipFiles({ ...deckFiles(), 'ppt/deep.xml': deep });
    const r = await openInWorker(zip, '/ppt/deep.xml');
    expect(r).toMatchObject({ ok: false, code: 'invalid-part' });
    expect(r.message).toContain('nested deeper than 256 levels');
  });

  test('a 100 MB attribute is refused by the part size limit before parsing', async () => {
    const huge = `<r a="${'x'.repeat(100_000_000)}"/>`;
    const files: Record<string, Uint8Array | [Uint8Array, { level: 1 }]> = {};
    for (const [k, v] of Object.entries(deckFiles())) files[k] = typeof v === 'string' ? strToU8(v) : v;
    files['ppt/huge.xml'] = [strToU8(huge), { level: 1 }];
    const r = await openInWorker(zipSync(files), '/ppt/huge.xml', 15_000);
    expect(r).toMatchObject({ ok: false, code: 'limit' });
  }, 30_000);
});
