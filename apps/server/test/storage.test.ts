// Storage on the filesystem and the render job queue, without HTTP.
import { existsSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, describe, expect, test } from 'vitest';
import { LocalJobQueue, QueueFullError } from '../src/jobs.ts';
import { currentVersion, DeckStore, FsBlobStore, FsMetaStore, sha256, writeAtomic, type DeckMeta, type VersionInfo } from '../src/storage.ts';
import { cleanup, tempDir } from './helpers.ts';

afterAll(() => cleanup());

const INFO: VersionInfo = { encrypted: false, slideCount: 1, slideSize: { cx: 1, cy: 1 }, contentType: 'x', macros: false };
const bytes = (s: string) => new TextEncoder().encode(s);
const exists = (meta: DeckMeta | undefined): DeckMeta => {
  if (!meta) throw new Error('missing');
  return meta;
};

describe('filesystem stores', () => {
  test('blobs are named by their SHA-256 and written once', async () => {
    const dir = tempDir();
    const blobs = new FsBlobStore(join(dir, 'blobs'), dir);
    const hash = await blobs.put(bytes('hello'));
    expect(hash).toBe(sha256(bytes('hello')));
    expect(hash).toBe('2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824');
    expect(await blobs.put(bytes('hello'))).toBe(hash);
    expect(Buffer.from((await blobs.get(hash))!).toString()).toBe('hello');
    expect(await blobs.has(hash)).toBe(true);
    await blobs.delete(hash);
    expect(await blobs.get(hash)).toBeUndefined();
    await expect(blobs.get('../../etc/passwd')).rejects.toThrow(/SHA-256/);
    expect(readdirSync(dir).filter(n => n.endsWith('.tmp'))).toEqual([]);
  });

  test('metadata writes are compare-and-set on the version count', async () => {
    const dir = tempDir();
    const metas = new FsMetaStore(join(dir, 'decks'), dir);
    const meta: DeckMeta = { id: 'a'.repeat(32), name: 'x', createdAt: '', updatedAt: '', versions: [{ version: 1, sha256: '0'.repeat(64), size: 1, createdAt: '', ...INFO }] };
    expect(await metas.put(meta, 0)).toBe(true);
    expect(await metas.put(meta, 0)).toBe(false);                 // exists already
    const next = { ...meta, versions: [...meta.versions, { ...meta.versions[0]!, version: 2 }] };
    const results = await Promise.all([metas.put(next, 1), metas.put(next, 1)]);
    expect(results.sort()).toEqual([false, true]);
    expect((await metas.get(meta.id))!.versions).toHaveLength(2);
    expect(await metas.list()).toEqual([meta.id]);
    expect(await metas.delete(meta.id)).toBe(true);
    expect(await metas.delete(meta.id)).toBe(false);
    await expect(metas.get('../x')).rejects.toThrow(/deck id/);
  });

  test('an atomic write leaves the old file or the new one, never a part', async () => {
    const dir = tempDir();
    const target = join(dir, 'f.json');
    writeFileSync(target, 'old');
    await writeAtomic(dir, target, 'new');
    expect(readdirSync(dir)).toEqual(['f.json']);
    await expect(writeAtomic(join(dir, 'missing'), target, 'x')).rejects.toThrow();
    expect(readdirSync(dir)).toEqual(['f.json']);
  });
});

describe('DeckStore', () => {
  test('references survive a restart: a shared blob goes only with its last deck', async () => {
    const dir = tempDir();
    const store = await DeckStore.openDirectory(dir);
    const a = await store.create('a', bytes('same'), INFO);
    const b = await store.create('b', bytes('same'), INFO);
    const v2 = await store.addVersion(a.id, bytes('newer'), INFO, exists);
    expect(v2.versions.map(v => v.version)).toEqual([1, 2]);
    expect(currentVersion(v2).sha256).toBe(sha256(bytes('newer')));
    writeFileSync(join(dir, 'tmp', 'interrupted.tmp'), 'x');

    const reopened = await DeckStore.openDirectory(dir);
    expect(existsSync(join(dir, 'tmp', 'interrupted.tmp'))).toBe(false);
    await reopened.delete(a.id, exists);
    expect(await reopened.blobs.has(sha256(bytes('same')))).toBe(true);     // b still has it
    expect(await reopened.blobs.has(sha256(bytes('newer')))).toBe(false);
    await reopened.delete(b.id, exists);
    expect(await reopened.blobs.has(sha256(bytes('same')))).toBe(false);
    expect(await reopened.get(b.id)).toBeUndefined();
    expect(await reopened.get('not an id')).toBeUndefined();
  });

  test('a precondition that throws leaves everything as it was, without a stray blob', async () => {
    const store = await DeckStore.openDirectory(tempDir());
    const a = await store.create('a', bytes('one'), INFO);
    await expect(store.addVersion(a.id, bytes('two'), INFO, () => { throw new Error('stale'); })).rejects.toThrow('stale');
    expect(await store.blobs.has(sha256(bytes('two')))).toBe(false);
    // the same refusal with content another deck holds keeps that deck's blob
    await expect(store.addVersion(a.id, bytes('one'), INFO, () => { throw new Error('stale'); })).rejects.toThrow('stale');
    expect(await store.blobs.has(sha256(bytes('one')))).toBe(true);
    await expect(store.delete(a.id, () => { throw new Error('stale'); })).rejects.toThrow('stale');
    expect((await store.get(a.id))!.versions).toHaveLength(1);
  });

  test('render drawings go with their content', async () => {
    const store = await DeckStore.openDirectory(tempDir());
    const a = await store.create('a', bytes('drawn'), INFO);
    const hash = currentVersion(a).sha256;
    await store.cache.put(hash, 'slide-1-w640.png', bytes('png'));
    expect(Buffer.from((await store.cache.get(hash, 'slide-1-w640.png'))!).toString()).toBe('png');
    await expect(store.cache.put(hash, '../escape', bytes('x'))).rejects.toThrow(/cache entry/);
    await expect(store.cache.put(hash, 'png/slide.png', bytes('x'))).rejects.toThrow(/cache entry/);
    await expect(store.cache.get(`../${hash.slice(3)}`, 'slide-1-w640.png')).rejects.toThrow(/SHA-256/);
    await store.delete(a.id, exists);
    expect(await store.cache.get(hash, 'slide-1-w640.png')).toBeUndefined();
  });
});

describe('LocalJobQueue', () => {
  test('one run per key, a bounded number at once, a bounded wait list', async () => {
    const q = new LocalJobQueue(1, 1);
    let runs = 0;
    let release!: () => void;
    const gate = new Promise<void>(r => { release = r; });
    const job = async () => { runs++; await gate; return runs; };
    const a = q.run('k', job);
    const b = q.run('k', job);                    // shares a's run
    const c = q.run('other', async () => 'c');    // waits for the slot
    await expect(q.run('third', async () => 'x')).rejects.toBeInstanceOf(QueueFullError);
    expect(q.stats()).toEqual({ running: 1, waiting: 1 });
    release();
    expect(await a).toBe(1);
    expect(await b).toBe(1);
    expect(await c).toBe('c');
    expect(runs).toBe(1);
    expect(q.stats()).toEqual({ running: 0, waiting: 0 });
    expect(await q.run('k', async () => 'fresh')).toBe('fresh');
  });

  test('a failed job frees its slot and its key', async () => {
    const q = new LocalJobQueue(1, 0);
    await expect(q.run('k', async () => { throw new Error('boom'); })).rejects.toThrow('boom');
    expect(await q.run('k', async () => 'ok')).toBe('ok');
  });
});
