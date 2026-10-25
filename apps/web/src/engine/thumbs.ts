// apps/web/src/engine/thumbs.ts  -- the main thread's side of the thumbnail worker: keeping its deck in step
import type { Editor } from '@pptx/editor-core';
import { relsNameFor } from '@pptx/opc';
import { encodeXmlText, serializeXml } from '@pptx/xml';
import type { ThumbMessage } from './thumbs.render.ts';

/**
 * Thumbnails rendered off the main thread. The worker opens its own copy of the deck; after every change the parts
 * that changed (and their relationships), the parts added and the parts removed are sent over, so the worker's
 * copy draws what the editor shows. A slide's thumbnail is cached until its scene key changes.
 */
/** A thumbnail not drawn in this time means the worker is stuck (a hostile picture): it is replaced. */
const WATCHDOG_MS = 30_000;

export class Thumbnails {
  private worker!: Worker;
  private readonly ed: Editor;
  private names: Set<string>;
  private readonly bytesOf = new Map<string, Uint8Array>();
  private next = 1;
  private readonly waiting = new Map<number, { resolve: (b: ImageBitmap | undefined) => void; timer: ReturnType<typeof setTimeout> }>();
  private readonly cache = new Map<string, { key: string; bitmap: Promise<ImageBitmap | undefined> }>();

  /** Bumped when the worker is replaced, so thumbnails drawn by the old one are asked for again. */
  generation = 0;
  private readonly restarted: () => void;

  constructor(ed: Editor, zip: Uint8Array, restarted: () => void = () => undefined) {
    this.ed = ed;
    this.restarted = restarted;
    this.names = new Set();
    this.start(zip);
  }

  /** A worker with its own copy of the deck as `zip` holds it (the file as opened, or as it is now). */
  private start(zip: Uint8Array): void {
    this.worker = new Worker(new URL('./thumbs.worker.ts', import.meta.url), { type: 'module' });
    this.worker.onmessage = (e: MessageEvent<{ type: 'thumb'; id: number; bitmap?: ImageBitmap; error?: string }>) => {
      if (e.data.error) console.warn(`thumbnail: ${e.data.error}`);
      const job = this.waiting.get(e.data.id);
      if (job) clearTimeout(job.timer);
      job?.resolve(e.data.bitmap);
      this.waiting.delete(e.data.id);
    };
    this.worker.onerror = e => console.warn(`thumbnail worker: ${e.message}`);
    const copy = zip.slice();
    this.post({ type: 'init', bytes: copy }, [copy.buffer]);
    this.names = new Set(this.ed.pkg.partNames());
    this.bytesOf.clear();
    for (const n of this.names) this.bytesOf.set(n, this.ed.pkg.getPart(n)!.bytes);
  }

  /** Replace a stuck worker with one opened from the deck as it is now; pending thumbnails are drawn again later. */
  private restart(): void {
    this.worker.terminate();
    for (const job of this.waiting.values()) {
      clearTimeout(job.timer);
      job.resolve(undefined);
    }
    this.waiting.clear();
    this.cache.clear();
    this.start(this.ed.pkg.save());
    this.generation++;
    this.restarted();
  }

  private post(m: ThumbMessage, transfer: Transferable[] = []): void { this.worker.postMessage(m, transfer); }

  /** Send the worker what changed: the parts a command touched, their relationships, and added or removed parts. */
  sync(parts: readonly string[]): void {
    const pkg = this.ed.pkg;
    const now = new Set(pkg.partNames());
    const removed = [...this.names].filter(n => !now.has(n));
    const changed = new Set<string>();
    for (const p of parts) {
      if (now.has(p)) changed.add(p);
      const rels = relsNameFor(p);
      if (now.has(rels)) changed.add(rels);
    }
    for (const n of now) if (!this.names.has(n) || this.bytesOf.get(n) !== pkg.getPart(n)!.bytes) changed.add(n);
    this.names = now;
    if (!removed.length && !changed.size) return;
    const out = [...changed].map(name => {
      const p = pkg.getPart(name)!;
      this.bytesOf.set(name, p.bytes);
      const bytes = p.xml && p.dirty ? encodeXmlText(serializeXml(p.xml), p.encoding) : p.bytes;
      return { name, contentType: p.contentType, bytes };
    });
    for (const n of removed) this.bytesOf.delete(n);
    this.post({ type: 'update', removed, parts: out });
  }

  /** A slide's thumbnail for a scene key (see SceneCache.key): rendered once per key. */
  get(part: string, key: string, width: number): Promise<ImageBitmap | undefined> {
    const hit = this.cache.get(part);
    if (hit && hit.key === `${key}@${width}`) return hit.bitmap;
    const id = this.next++;
    const bitmap = new Promise<ImageBitmap | undefined>(resolve => {
      const timer = setTimeout(() => {
        console.warn('thumbnail worker stopped answering; starting a new one');
        this.restart();
      }, WATCHDOG_MS);
      this.waiting.set(id, { resolve, timer });
    });
    this.post({ type: 'render', id, part, width });
    this.cache.set(part, { key: `${key}@${width}`, bitmap });
    return bitmap;
  }

  dispose(): void { this.worker.terminate(); }
}
