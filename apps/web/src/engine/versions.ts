// apps/web/src/engine/versions.ts  -- version history: earlier states of a deck, kept in this browser
import { VERSIONS, withStore } from './autosave.ts';

/** A kept version: the deck as it would be saved then (encrypted, when it has a password), and what to call it. */
export interface Version {
  id: number;
  /** The deck it is a version of: its file name. */
  deck: string;
  /** "Saved", "Autosaved", "Before restoring", or the name the user gave it. */
  label: string;
  at: number;
  bytes: Uint8Array;
  encrypted: boolean;
  slides: number;
  /** Made by the editor (a save, a timer), not named by the user: these go first when room is needed. */
  auto: boolean;
}
/** A version without its bytes, for listing. */
export type VersionInfo = Omit<Version, 'bytes'> & { size: number };

/** How many versions of one deck are kept. */
export const KEEP = 20;
/** While a deck is being changed, a version is kept this often. */
export const AUTO_EVERY = 10 * 60_000;
/** A deck larger than this is not kept as versions (as it is not autosaved). */
export const VERSION_LIMIT = 48 << 20;

/**
 * The versions to drop so that at most `keep` are left: the oldest the editor made by itself go first, and only
 * when there are none of those left to drop do the oldest named ones go. Returns their ids.
 */
export function prune(list: Pick<VersionInfo, 'id' | 'at' | 'auto'>[], keep = KEEP): number[] {
  const over = list.length - keep;
  if (over <= 0) return [];
  const oldestFirst = [...list].sort((a, b) => a.at - b.at || a.id - b.id);
  const drop = oldestFirst.filter(v => v.auto).slice(0, over);
  if (drop.length < over) drop.push(...oldestFirst.filter(v => !v.auto).slice(0, over - drop.length));
  return drop.map(v => v.id);
}

/** Is it time to keep a version by itself? `last` is when one was last kept (or when the deck was opened). */
export const dueForAuto = (last: number, now: number, every = AUTO_EVERY): boolean => now - last >= every;

/** When a version was kept, as a person says it: the time today, "Yesterday" with the time, else the date and time. */
export function whenText(at: number, now: number, locale?: string): string {
  const d = new Date(at), n = new Date(now);
  const day = (x: Date) => new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
  const days = Math.round((day(n) - day(d)) / 86_400_000);
  const time = d.toLocaleTimeString(locale, { hour: 'numeric', minute: '2-digit' });
  if (days === 0) return `Today, ${time}`;
  if (days === 1) return `Yesterday, ${time}`;
  return `${d.toLocaleDateString(locale, { day: 'numeric', month: 'short', ...(d.getFullYear() === n.getFullYear() ? {} : { year: 'numeric' }) })}, ${time}`;
}

/** A size as a person reads it. */
export function sizeText(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

const info = (v: Version): VersionInfo => { const { bytes, ...rest } = v; return { ...rest, size: bytes.length }; };

/**
 * The versions kept in this browser. They are this browser's own: another browser or computer does not have them,
 * and clearing the site's data removes them. Where the browser gives no database, nothing is kept and nothing fails.
 */
export const versions = {
  available: () => typeof indexedDB !== 'undefined',

  /** Keep a version; older ones of the same deck are dropped to make room. Returns its id, or undefined when it was not kept. */
  async add(v: Omit<Version, 'id'>): Promise<number | undefined> {
    if (!versions.available() || v.bytes.length > VERSION_LIMIT) return undefined;
    try {
      const id = Number(await withStore<IDBValidKey>(VERSIONS, 'readwrite', s => s.add(v)));
      const drop = prune(await versions.list(v.deck));
      for (const old of drop) await versions.remove(old);
      return id;
    } catch {
      return undefined;                                                   // no room, or no database: the deck itself is unharmed
    }
  },

  /** A deck's versions, newest first. */
  async list(deck: string): Promise<VersionInfo[]> {
    if (!versions.available()) return [];
    try {
      const all = await withStore<Version[]>(VERSIONS, 'readonly', s => s.index('deck').getAll(deck));
      return all.map(info).sort((a, b) => b.at - a.at || b.id - a.id);
    } catch {
      return [];
    }
  },

  get: (id: number): Promise<Version | undefined> => (versions.available() ? withStore<Version | undefined>(VERSIONS, 'readonly', s => s.get(id)).catch(() => undefined) : Promise.resolve(undefined)),
  remove: (id: number): Promise<void> => (versions.available() ? withStore(VERSIONS, 'readwrite', s => s.delete(id)).then(() => undefined).catch(() => undefined) : Promise.resolve()),
};
