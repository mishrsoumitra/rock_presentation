// apps/server/src/storage.ts  -- decks at rest: content-addressed blobs, a metadata record per deck, the render cache
//
// The plan's production design is S3-compatible objects keyed by content hash, a Postgres table of deck metadata and
// a job queue. BlobStore, MetaStore and RenderCache are that split; the implementations here keep everything on the
// local filesystem under DATA_DIR, for one server process:
//
//   blobs/ab/<sha256>          deck bytes, named by their SHA-256: written once, never rewritten
//   decks/<id>.json            one metadata record per deck, replaced atomically on every change
//   cache/<sha256>/<name>      slide PNGs and PDFs drawn from that content
//   tmp/                       files being written, renamed into place once complete and flushed
import { createHash, randomBytes } from 'node:crypto';
import { mkdir, open, readdir, readFile, rename, rm, stat } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';

export const HASH_RE = /^[0-9a-f]{64}$/;
export const DECK_ID_RE = /^[0-9a-f]{32}$/;
const CACHE_NAME_RE = /^[a-z0-9][a-z0-9.-]{0,63}$/;

export interface DeckVersion {
  version: number;                              // 1, 2, ... in the order they were saved
  sha256: string;                               // the blob holding this version
  size: number;
  createdAt: string;
  encrypted: boolean;                           // password-protected: stored as uploaded, only a client can open it
  slideCount: number | null;                    // null when encrypted
  slideSize: { cx: number; cy: number } | null; // EMU; null when encrypted
  contentType: string;                          // the media type downloads are sent with
  macros: boolean;                              // carries a VBA project, which is stored and never run
}

/** What the validation worker found out about an upload: everything in a version but where it was stored. */
export type VersionInfo = Omit<DeckVersion, 'version' | 'sha256' | 'size' | 'createdAt'>;

export interface DeckMeta {
  id: string;
  name: string;
  createdAt: string;
  updatedAt: string;
  versions: DeckVersion[];                      // never empty; the last one is current
}

export const currentVersion = (meta: DeckMeta): DeckVersion => meta.versions[meta.versions.length - 1]!;

/** Immutable objects keyed by the SHA-256 of their bytes (S3 in the plan). */
export interface BlobStore {
  /** Store the bytes and return their hash. An object that exists already is left as it is. */
  put(bytes: Uint8Array): Promise<string>;
  get(hash: string): Promise<Uint8Array | undefined>;
  has(hash: string): Promise<boolean>;
  delete(hash: string): Promise<void>;
}

/** Deck records (a Postgres table in the plan). */
export interface MetaStore {
  get(id: string): Promise<DeckMeta | undefined>;
  /**
   * Write the record if the stored one still has `expected` versions (0: there must be none yet), the compare-and-set
   * that keeps concurrent saves from losing each other. False when the record moved on.
   */
  put(meta: DeckMeta, expected: number): Promise<boolean>;
  list(): Promise<string[]>;
  delete(id: string): Promise<boolean>;
}

/** Drawings of a content hash: slide PNGs by (hash, slide, width), PDFs by (hash, width). */
export interface RenderCache {
  get(hash: string, name: string): Promise<Uint8Array | undefined>;
  put(hash: string, name: string, bytes: Uint8Array): Promise<void>;
  /** Forget everything drawn from this content, when its blob goes. */
  drop(hash: string): Promise<void>;
}

export const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

/**
 * A name to use as one path segment: the request's ids, hashes and cache names only ever reach the file system
 * through here. It must match its pattern and be its own basename (no separators, so no way out of the directory).
 */
function safeName(name: string, pattern: RegExp, what: string): string {
  const base = basename(name);
  if (base !== name || !pattern.test(base)) throw new Error(`not ${what}: ${name.slice(0, 80)}`);
  return base;
}

const isMissing = (e: unknown) => (e as NodeJS.ErrnoException).code === 'ENOENT';

async function readOrUndefined(path: string): Promise<Buffer | undefined> {
  try {
    return await readFile(path);
  } catch (e) {
    if (isMissing(e)) return undefined;
    throw e;
  }
}

async function exists(path: string): Promise<boolean> {
  try {
    await stat(path);
    return true;
  } catch (e) {
    if (isMissing(e)) return false;
    throw e;
  }
}

/**
 * Write a file so that a reader finds either the old content or all of the new: a temp file on the same filesystem,
 * flushed to disk, then renamed over the target (rename is atomic within a filesystem). Private to this user.
 */
export async function writeAtomic(tmpDir: string, target: string, data: Uint8Array | string): Promise<void> {
  const tmp = join(tmpDir, `${randomBytes(12).toString('hex')}.tmp`);
  const fh = await open(tmp, 'wx', 0o600);
  try {
    try {
      await fh.writeFile(data);
      await fh.sync();
    } finally {
      await fh.close();
    }
    await mkdir(dirname(target), { recursive: true, mode: 0o700 });
    await rename(tmp, target);
  } catch (e) {
    await rm(tmp, { force: true });
    throw e;
  }
}

/** Runs tasks one at a time, in the order they arrive. */
export class Mutex {
  private tail: Promise<unknown> = Promise.resolve();
  run<T>(task: () => Promise<T>): Promise<T> {
    const result = this.tail.then(task);
    this.tail = result.catch(() => undefined);
    return result;
  }
}

export class FsBlobStore implements BlobStore {
  private readonly root: string;
  private readonly tmp: string;
  constructor(root: string, tmp: string) {
    this.root = root;
    this.tmp = tmp;
  }

  private path(hash: string): string {
    return join(this.root, hash.slice(0, 2), safeName(hash, HASH_RE, 'a SHA-256'));
  }

  async put(bytes: Uint8Array): Promise<string> {
    const hash = sha256(bytes);
    const path = this.path(hash);
    if (!(await exists(path))) await writeAtomic(this.tmp, path, bytes);
    return hash;
  }

  async get(hash: string): Promise<Uint8Array | undefined> {
    return readOrUndefined(this.path(hash));
  }

  has(hash: string): Promise<boolean> {
    return exists(this.path(hash));
  }

  async delete(hash: string): Promise<void> {
    await rm(this.path(hash), { force: true });
  }
}

export class FsMetaStore implements MetaStore {
  private readonly root: string;
  private readonly tmp: string;
  private readonly lock = new Mutex();
  constructor(root: string, tmp: string) {
    this.root = root;
    this.tmp = tmp;
  }

  private path(id: string): string {
    return join(this.root, `${safeName(id, DECK_ID_RE, 'a deck id')}.json`);
  }

  async get(id: string): Promise<DeckMeta | undefined> {
    const raw = await readOrUndefined(this.path(id));
    if (!raw) return undefined;
    const meta = JSON.parse(raw.toString('utf8')) as DeckMeta;
    if (meta.id !== id || !Array.isArray(meta.versions) || meta.versions.length === 0) throw new Error(`deck record ${id} is damaged`);
    return meta;
  }

  put(meta: DeckMeta, expected: number): Promise<boolean> {
    return this.lock.run(async () => {
      const stored = await this.get(meta.id);
      if ((stored?.versions.length ?? 0) !== expected) return false;
      await writeAtomic(this.tmp, this.path(meta.id), JSON.stringify(meta, null, 2) + '\n');
      return true;
    });
  }

  async list(): Promise<string[]> {
    let names: string[];
    try {
      names = await readdir(this.root);
    } catch (e) {
      if (isMissing(e)) return [];
      throw e;
    }
    return names.filter(n => n.endsWith('.json') && DECK_ID_RE.test(n.slice(0, -5))).map(n => n.slice(0, -5)).sort();
  }

  delete(id: string): Promise<boolean> {
    return this.lock.run(async () => {
      const path = this.path(id);
      if (!(await exists(path))) return false;
      await rm(path);
      return true;
    });
  }
}

export class FsRenderCache implements RenderCache {
  private readonly root: string;
  private readonly tmp: string;
  constructor(root: string, tmp: string) {
    this.root = root;
    this.tmp = tmp;
  }

  private dir(hash: string): string {
    return join(this.root, safeName(hash, HASH_RE, 'a SHA-256'));
  }

  private path(hash: string, name: string): string {
    return join(this.dir(hash), safeName(name, CACHE_NAME_RE, 'a cache entry name'));
  }

  async get(hash: string, name: string): Promise<Uint8Array | undefined> {
    return readOrUndefined(this.path(hash, name));
  }

  async put(hash: string, name: string, bytes: Uint8Array): Promise<void> {
    await writeAtomic(this.tmp, this.path(hash, name), bytes);
  }

  async drop(hash: string): Promise<void> {
    await rm(this.dir(hash), { recursive: true, force: true });
  }
}

export interface DeckStoreParts { blobs: BlobStore; metas: MetaStore; cache: RenderCache; }

/** Checks the current record (undefined: no such deck) and returns it, or throws to refuse the change. */
export type Precondition = (meta: DeckMeta | undefined) => DeckMeta;

/**
 * Decks over the three stores. A blob can belong to several decks (two uploads of the same file), so it is deleted
 * only when the last version naming it goes: references are counted in memory, built from the records at startup.
 * Changes to records and references run one at a time; writing a blob, the slow part, happens outside that lock.
 */
export class DeckStore {
  readonly blobs: BlobStore;
  readonly metas: MetaStore;
  readonly cache: RenderCache;
  private readonly refs = new Map<string, number>();
  private readonly lock = new Mutex();

  private constructor(parts: DeckStoreParts) {
    this.blobs = parts.blobs;
    this.metas = parts.metas;
    this.cache = parts.cache;
  }

  static async open(parts: DeckStoreParts): Promise<DeckStore> {
    const store = new DeckStore(parts);
    for (const id of await parts.metas.list()) {
      const meta = await parts.metas.get(id);
      for (const v of meta?.versions ?? []) store.refs.set(v.sha256, (store.refs.get(v.sha256) ?? 0) + 1);
    }
    return store;
  }

  /** The filesystem stores under one directory (DATA_DIR). */
  static async openDirectory(dataDir: string): Promise<DeckStore> {
    const tmp = join(dataDir, 'tmp');
    for (const d of ['blobs', 'decks', 'cache', 'tmp']) await mkdir(join(dataDir, d), { recursive: true, mode: 0o700 });
    // leftovers of writes (and validator runs) interrupted by a crash; a DATA_DIR belongs to one server process
    for (const name of await readdir(tmp)) await rm(join(tmp, name), { recursive: true, force: true });
    return DeckStore.open({
      blobs: new FsBlobStore(join(dataDir, 'blobs'), tmp),
      metas: new FsMetaStore(join(dataDir, 'decks'), tmp),
      cache: new FsRenderCache(join(dataDir, 'cache'), tmp),
    });
  }

  get(id: string): Promise<DeckMeta | undefined> {
    return DECK_ID_RE.test(id) ? this.metas.get(id) : Promise.resolve(undefined);
  }

  /**
   * Under the lock: count a reference to the blob, written beforehand outside the lock (a delete in between may have
   * taken it, so it is put back), then write the record. When the write is refused (a failed precondition, a lost
   * compare-and-set) the reference is dropped again and an unreferenced blob removed: a refused save leaves nothing.
   */
  private commit(bytes: Uint8Array, hash: string, write: () => Promise<DeckMeta>): Promise<DeckMeta> {
    return this.lock.run(async () => {
      if (!(await this.blobs.has(hash))) await this.blobs.put(bytes);
      this.refs.set(hash, (this.refs.get(hash) ?? 0) + 1);
      try {
        return await write();
      } catch (e) {
        await this.forget(this.release([hash]));
        throw e;
      }
    });
  }

  async create(name: string, bytes: Uint8Array, info: VersionInfo): Promise<DeckMeta> {
    const hash = await this.blobs.put(bytes);
    return this.commit(bytes, hash, async () => {
      const now = new Date().toISOString();
      const meta: DeckMeta = {
        id: randomBytes(16).toString('hex'), name, createdAt: now, updatedAt: now,
        versions: [{ version: 1, sha256: hash, size: bytes.length, createdAt: now, ...info }],
      };
      if (!(await this.metas.put(meta, 0))) throw new Error('deck id collision');
      return meta;
    });
  }

  /**
   * Save a new version. `precondition` runs under the lock with the current record and throws to refuse (a stale
   * If-Match), so the check and the write cannot be split by another save.
   */
  async addVersion(id: string, bytes: Uint8Array, info: VersionInfo, precondition: Precondition, name?: string): Promise<DeckMeta> {
    const hash = await this.blobs.put(bytes);
    return this.commit(bytes, hash, async () => {
      const meta = precondition(await this.get(id));
      const now = new Date().toISOString();
      const next: DeckMeta = {
        ...meta, name: name ?? meta.name, updatedAt: now,
        versions: [...meta.versions, { version: meta.versions.length + 1, sha256: hash, size: bytes.length, createdAt: now, ...info }],
      };
      if (!(await this.metas.put(next, meta.versions.length))) throw new Error(`deck ${id} changed during a locked update`);
      return next;
    });
  }

  /** Delete a deck; `precondition` as for addVersion. Blobs no other deck refers to go too, with their drawings. */
  async delete(id: string, precondition: Precondition): Promise<void> {
    await this.lock.run(async () => {
      const meta = precondition(await this.get(id));
      await this.metas.delete(id);
      await this.forget(this.release(meta.versions.map(v => v.sha256)));
    });
  }

  private async forget(hashes: string[]): Promise<void> {
    for (const hash of hashes) {
      await this.blobs.delete(hash);
      await this.cache.drop(hash);
    }
  }

  /** Drop one reference per hash; returns the hashes nothing refers to any more. */
  private release(hashes: string[]): string[] {
    const gone = new Set<string>();
    for (const h of hashes) {
      const n = (this.refs.get(h) ?? 0) - 1;
      if (n > 0) {
        this.refs.set(h, n);
      } else {
        this.refs.delete(h);
        gone.add(h);
      }
    }
    return [...gone];
  }
}
