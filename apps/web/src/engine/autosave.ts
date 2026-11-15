// apps/web/src/engine/autosave.ts  -- the unsaved deck, kept in this browser so a closed tab or a crash loses nothing

/** What is kept: the deck as it would be saved (encrypted, when it has a password), its name, and when. */
export interface Autosaved { name: string; bytes: Uint8Array; at: number; encrypted: boolean; }

const DB = 'presentations', STORE = 'autosave', KEY = 'current';
/** The store of kept versions (see versions.ts): in the same database, so both are made, and cleared, together. */
export const VERSIONS = 'versions';

/** The database: version 1 had the autosaved deck alone; version 2 adds the kept versions, found by their deck. */
function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 2);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
      if (!db.objectStoreNames.contains(VERSIONS)) db.createObjectStore(VERSIONS, { keyPath: 'id', autoIncrement: true }).createIndex('deck', 'deck');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('the browser refused to open its database'));
  });
}

/** One request on one store, resolved when its transaction is through. */
export function withStore<T>(store: string, mode: IDBTransactionMode, work: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return open().then(db => new Promise<T>((resolve, reject) => {
    const tx = db.transaction(store, mode);
    const req = work(tx.objectStore(store));
    tx.oncomplete = () => { db.close(); resolve(req.result); };
    tx.onerror = tx.onabort = () => { db.close(); reject(tx.error ?? new Error('the browser could not write its database')); };
  }));
}
const run = <T>(mode: IDBTransactionMode, work: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> => withStore(STORE, mode, work);

/**
 * One deck is kept at a time: the one being edited, whenever it has changes that are not saved. Where the browser
 * has no database to give (private windows of some browsers), nothing is kept and nothing fails.
 */
export const autosave = {
  available: () => typeof indexedDB !== 'undefined',
  put: (record: Autosaved): Promise<void> => (autosave.available() ? run('readwrite', s => s.put(record, KEY)).then(() => undefined).catch(() => undefined) : Promise.resolve()),
  get: (): Promise<Autosaved | undefined> => (autosave.available() ? run<Autosaved | undefined>('readonly', s => s.get(KEY)).catch(() => undefined) : Promise.resolve(undefined)),
  clear: (): Promise<void> => (autosave.available() ? run('readwrite', s => s.delete(KEY)).then(() => undefined).catch(() => undefined) : Promise.resolve()),
};
