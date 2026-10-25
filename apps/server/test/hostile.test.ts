// Hostile input and worker isolation. Exit criteria: a 10 GB-declared ZIP bomb, a 200,000-entry archive and a
// malformed compound file each fail within the deadline without affecting requests running at the same time; a
// parse that exceeds the worker's memory or time fails cleanly, the thread is replaced, and the server keeps serving.
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { TaskFailure, WorkerError, WorkerPool } from '../src/workers.ts';
import { resolveConfig } from '../src/config.ts';
import { cleanup, compoundFile, deck, testServer, upload, zeroBomb, zip, type ZipEntry } from './helpers.ts';

const CT = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>';

/** One entry declaring 10 GB through ZIP64 (oracle of packages/opc's limits test, sent over HTTP). */
const bomb10GB = () => zip([{ name: '[Content_Types].xml', data: CT }, { name: 'ppt/bomb.xml', data: '<r/>', size: 10e9, zip64: true }], { zip64End: true });
/** A real bomb: 1 GiB of zeros that deflate to about 1 MB, sizes declared truthfully. */
const bomb1GiB = () => zip([{ name: '[Content_Types].xml', data: CT }, { name: 'ppt/media/zeros.bin', deflated: zeroBomb(1024), size: 1 << 30 }]);
/** 200,000 real entries (about 20 MB): ten times the entry limit. */
function entries200k(): Buffer {
  const entries: ZipEntry[] = [{ name: '[Content_Types].xml', data: CT }];
  for (let i = 0; i < 200_000; i++) entries.push({ name: `ppt/e${i}.bin` });
  return zip(entries);
}
/** An end record claiming 200,000 entries over a central directory of one. */
const claims200k = () => zip([{ name: '[Content_Types].xml', data: CT }], { zip64End: true, claimedEntries: 200_000 });
/** A compound file whose FAT loops, with the stream names of a password-protected package. */
const loopingCfb = () => compoundFile(['Root Entry', 'EncryptedPackage', 'EncryptionInfo'], { fatCycle: true });

let app: FastifyInstance;
beforeAll(async () => {
  app = await testServer({ workers: { size: 2, timeoutMs: 10_000, heapMb: 128 } });
});
afterAll(async () => {
  await app.close();
  cleanup();
});

describe('hostile archives are refused before anything is inflated', () => {
  const cases: [string, () => Uint8Array, number, string[]][] = [
    ['an entry declaring 10 GB (ZIP64)', bomb10GB, 413, ['limit']],
    ['1 GiB of zeros deflated to 1 MB', bomb1GiB, 413, ['limit']],
    ['200,000 entries', entries200k, 413, ['limit']],
    ['an end record claiming 200,000 entries', claims200k, 422, ['invalid-zip', 'invalid-package']],
    ['a compound file whose FAT loops', loopingCfb, 422, ['malformed']],
  ];
  test.each(cases)('%s', async (_name, build, status, codes) => {
    const bytes = build();
    const started = performance.now();
    const res = await upload(app, bytes);
    expect(res.statusCode).toBe(status);
    expect(codes).toContain(res.json().error);
    expect(performance.now() - started).toBeLessThan(5000);
  });

  test('all at once, while a health check and an ordinary upload go through untouched', async () => {
    const hostile = [bomb10GB(), bomb1GiB(), entries200k(), claims200k(), loopingCfb()];
    const started = performance.now();
    const [health, ordinary, ...refused] = await Promise.all([
      app.inject({ url: '/api/health' }),
      upload(app, deck(2, { note: 'ordinary' })),
      ...hostile.map(b => upload(app, b)),
    ]);
    expect(health!.statusCode).toBe(200);
    expect(ordinary!.statusCode).toBe(201);
    expect(ordinary!.json().slideCount).toBe(2);
    for (const r of refused) expect([413, 422]).toContain(r.statusCode);
    expect(performance.now() - started).toBeLessThan(10_000);
    expect(app.services.pool.stats()).toMatchObject({ timeouts: 0, outOfMemory: 0, crashes: 0 });
  });
});

describe('a parse that exceeds the worker\'s limits', () => {
  let tight: FastifyInstance;
  beforeAll(async () => {
    tight = await testServer({ workers: { size: 1, queue: 1, timeoutMs: 1500, heapMb: 48, youngMb: 8, testTasks: true } });
    expect((await upload(tight, deck(1))).statusCode).toBe(201);     // warm: the thread has loaded its modules
  });
  afterAll(async () => {
    await tight.close();
  });

  test('out of heap (test-only task): the job fails, the thread is replaced, the server keeps serving', async () => {
    const pool = tight.services.pool;
    const job = pool.run('test:alloc', null);
    const [health] = await Promise.all([tight.inject({ url: '/api/health' })]);
    expect(health.statusCode).toBe(200);
    const error = await job.catch((e: unknown) => e);
    expect(error).toBeInstanceOf(WorkerError);
    expect((error as WorkerError).code).toBe('out-of-memory');
    expect(pool.stats().outOfMemory).toBe(1);
    const after = await upload(tight, deck(2, { note: 'after oom' }));
    expect(after.statusCode).toBe(201);
  });

  test('past the deadline (test-only task): terminated on time, replaced, the server keeps serving', async () => {
    const pool = tight.services.pool;
    const started = performance.now();
    const job = pool.run('test:spin', null, { timeoutMs: 300 });
    const health = await tight.inject({ url: '/api/health' });
    expect(health.statusCode).toBe(200);
    const error = await job.catch((e: unknown) => e);
    expect((error as WorkerError).code).toBe('timeout');
    expect(performance.now() - started).toBeLessThan(3000);
    expect((await upload(tight, deck(1, { note: 'after spin' }))).statusCode).toBe(201);
  });

  test('a thread that dies (test-only task) fails only its job', async () => {
    const error = await tight.services.pool.run('test:crash', null).catch((e: unknown) => e);
    expect((error as WorkerError).code).toBe('crashed');
    expect((await upload(tight, deck(1, { note: 'after crash' }))).statusCode).toBe(201);
  });

  test('a real upload whose XML needs more heap than the worker has: 422, and the next upload works', async () => {
    // 1.4 million slide ids: 50 MB of XML, under the 64 MiB part limit, deflated to a few hundred KB
    const ids = '<p:sldId id="256" r:id="rId1"/>'.repeat(1_400_000);
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`;
    const res = await upload(tight, deck(1, { presentationXml: xml }));
    expect(res.statusCode).toBe(422);
    expect(res.json().error).toBe('too-complex');
    expect((await tight.inject({ url: '/api/health' })).statusCode).toBe(200);
    expect((await upload(tight, deck(1, { note: 'after the big one' }))).statusCode).toBe(201);
  });

  test('a full queue answers 503 at once', async () => {
    const pool = tight.services.pool;
    const running = pool.run('test:spin', null, { timeoutMs: 1000 }).catch(() => undefined);
    const queued = pool.run('test:spin', null, { timeoutMs: 1000 }).catch(() => undefined);
    const res = await upload(tight, deck(1));
    expect(res.statusCode).toBe(503);
    expect(res.json().error).toBe('busy');
    expect(res.headers['retry-after']).toBeDefined();
    await Promise.all([running, queued]);
    expect((await upload(tight, deck(1, { note: 'after the queue drained' }))).statusCode).toBe(201);
  });
});

test('the test-only tasks do not exist unless the pool is built for tests', async () => {
  const pool = new WorkerPool(resolveConfig().workers);
  try {
    const error = await pool.run('test:spin', null).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(TaskFailure);
    expect((error as TaskFailure).code).toBe('unknown-task');
  } finally {
    await pool.close();
  }
});
