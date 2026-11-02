// apps/web/precompress.ts  -- after a build, a Brotli and a gzip copy next to every file worth compressing
//
// The server (apps/server/src/web.ts) serves the editor from memory and sends these copies to the clients that
// take them. Without them it has to compress at every start, and Brotli at quality 11 takes about 19 s for the
// editor and its fonts on one core. Here the build does it once, on every core. A file gets its copies when its
// type compresses (scripts, styles, WebAssembly, fonts, text), it is bigger than 1 KB, and a copy saves at least a
// tenth: the same rules the server applies to a file that has none.
import { readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { availableParallelism } from 'node:os';
import { extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { isMainThread, parentPort, Worker, workerData } from 'node:worker_threads';
import { brotliCompressSync, constants, gzipSync } from 'node:zlib';
import type { Plugin } from 'vite';

const COMPRESSIBLE = new Set(['.html', '.js', '.css', '.json', '.wasm', '.svg', '.txt', '.ttf', '.otf', '.ico']);
const MIN_BYTES = 1024;
const WORTH_IT = 0.9;

/** Whether a built file of this name and size is worth compressing at all. */
export function compressible(name: string, size: number): boolean {
  return COMPRESSIBLE.has(extname(name).toLowerCase()) && size > MIN_BYTES && !/\.(map|br|gz)$/.test(name);
}

/** Both encodings of `body`, at their best; one that saves less than a tenth is left out. */
export function compress(body: Uint8Array): { br?: Buffer; gzip?: Buffer } {
  const br = brotliCompressSync(body, { params: { [constants.BROTLI_PARAM_QUALITY]: 11, [constants.BROTLI_PARAM_SIZE_HINT]: body.length } });
  const gzip = gzipSync(body, { level: 9 });
  return {
    ...(br.length < body.length * WORTH_IT ? { br } : {}),
    ...(gzip.length < body.length * WORTH_IT ? { gzip } : {}),
  };
}

/** `name.br` and `name.gz` next to `path`, when worth it; how many were written. */
async function writeCopies(path: string): Promise<number> {
  const { br, gzip } = compress(await readFile(path));
  const writes = [br && writeFile(`${path}.br`, br), gzip && writeFile(`${path}.gz`, gzip)].filter(Boolean);
  await Promise.all(writes);
  return writes.length;
}

/** Every file under `dir` worth compressing, biggest first. */
async function candidates(dir: string): Promise<{ path: string; size: number }[]> {
  const out: { path: string; size: number }[] = [];
  for (const entry of await readdir(dir, { recursive: true, withFileTypes: true })) {
    if (!entry.isFile() || entry.name.startsWith('.')) continue;
    const path = join(entry.parentPath, entry.name);
    const { size } = await stat(path);
    if (compressible(entry.name, size)) out.push({ path, size });
  }
  return out.sort((a, b) => b.size - a.size);
}

/**
 * Write the compressed copies of every file under `dir` worth them, spread over `threads` worker threads (one
 * per core by default; the biggest files go first, each to the thread with the least to do). Returns the number
 * of copies written.
 */
export async function precompressDir(dir: string, threads = availableParallelism()): Promise<number> {
  const files = await candidates(dir);
  if (!files.length) return 0;
  const shares: { paths: string[]; bytes: number }[] = Array.from({ length: Math.max(1, Math.min(threads, files.length)) }, () => ({ paths: [], bytes: 0 }));
  for (const f of files) {
    const share = shares.reduce((least, s) => (s.bytes < least.bytes ? s : least));
    share.paths.push(f.path);
    share.bytes += f.size;
  }
  const counts = await Promise.all(shares.map(share => new Promise<number>((resolveShare, reject) => {
    const worker = new Worker(fileURLToPath(import.meta.url), { workerData: share.paths });
    worker.once('message', (n: number) => resolveShare(n));
    worker.once('error', reject);
    worker.once('exit', code => { if (code !== 0) reject(new Error(`compression worker exited with code ${code}`)); });
  })));
  return counts.reduce((a, b) => a + b, 0);
}

/** The Vite plugin: the copies are written once the build is in its output directory. */
export function precompress(): Plugin {
  let outDir = '';
  let log = (_message: string) => {};
  return {
    name: 'precompress',
    apply: 'build',
    configResolved(config) {
      outDir = resolve(config.root, config.build.outDir);
      log = message => config.logger.info(message);
    },
    async closeBundle() {
      const started = performance.now();
      const n = await precompressDir(outDir);
      log(`precompressed ${n} copies (Brotli and gzip) in ${((performance.now() - started) / 1000).toFixed(1)} s`);
    },
  };
}

// The worker: its share of the files, one after the other.
if (!isMainThread && parentPort) {
  let n = 0;
  for (const path of workerData as string[]) n += await writeCopies(path);
  parentPort.postMessage(n);
}
