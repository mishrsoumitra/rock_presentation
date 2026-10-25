// apps/web/src/engine/images.ts  -- decoding the pictures slides show, once per picture
import { metafileImage, PICTURE_TYPES } from '@pptx/metafile';
import { metafileInWorker } from './workers.ts';
import type { OpcPackage } from '@pptx/opc';
import { BROWSER_IMAGE_TYPES } from '@pptx/render';

/**
 * Decode picture bytes for drawing: browser formats with createImageBitmap, SVG through an <img> element (where a
 * picture cannot run script or load anything), EMF, WMF and TIFF with @pptx/metafile. Undefined when the format
 * cannot be shown here.
 */
export async function decodeImage(bytes: Uint8Array, type: string, canSvg = typeof Image !== 'undefined'): Promise<CanvasImageSource | undefined> {
  if (PICTURE_TYPES.has(type)) {
    // on the page, in a worker with a deadline; inside a worker (thumbnails) the worker itself is the sandbox
    if (typeof Image !== 'undefined') return metafileInWorker(bytes);
    return (await metafileImage(bytes, (w, h) => new OffscreenCanvas(w, h))) ?? undefined;
  }
  if (!BROWSER_IMAGE_TYPES.has(type)) return undefined;
  const blob = new Blob([bytes as BlobPart], { type });
  if (type === 'image/svg+xml') {
    if (!canSvg) return undefined;
    const url = URL.createObjectURL(blob);
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      // rasterize now: an <img> of an SVG with no intrinsic size draws at 0x0 in some browsers
      const w = img.naturalWidth || 1024, h = img.naturalHeight || 1024;
      const c = new OffscreenCanvas(w, h);
      c.getContext('2d')!.drawImage(img, 0, 0, w, h);
      return c;
    } finally {
      URL.revokeObjectURL(url);
    }
  }
  return createImageBitmap(blob);
}

interface Entry { bytes: Uint8Array; image?: CanvasImageSource; done: boolean; }

/**
 * Decoded pictures by part name. An entry remembers the bytes it was decoded from, so a part whose content was
 * replaced (or a name reused after a delete) is decoded again rather than showing the old picture.
 */
export class ImageCache {
  private readonly entries = new Map<string, Entry>();
  private readonly pkg: () => OpcPackage;
  private readonly loaded: () => void;
  private timer: ReturnType<typeof setTimeout> | undefined;

  constructor(pkg: () => OpcPackage, loaded: () => void) {
    this.pkg = pkg;
    this.loaded = loaded;
  }

  get(part: string): CanvasImageSource | undefined {
    const e = this.entries.get(part);
    const p = this.pkg().getPart(part);
    return e && p && e.bytes === p.bytes ? e.image : undefined;
  }

  /** Start decoding what is not decoded yet; `loaded` fires (batched) when pictures arrive. */
  request(parts: Iterable<string>): void {
    for (const part of parts) {
      const p = this.pkg().getPart(part);
      if (!p) continue;
      const e = this.entries.get(part);
      if (e && e.bytes === p.bytes) continue;
      const entry: Entry = { bytes: p.bytes, done: false };
      this.entries.set(part, entry);
      decodeImage(p.bytes, p.contentType).then(img => {
        if (img) entry.image = img;
      }, () => undefined).finally(() => {
        entry.done = true;
        clearTimeout(this.timer);
        this.timer = setTimeout(this.loaded, 30);
      });
    }
  }

  /** Whether every picture of a set has been tried. */
  ready(parts: Iterable<string>): boolean {
    for (const part of parts) {
      const e = this.entries.get(part);
      if (!e || !e.done) return false;
    }
    return true;
  }

  clear(): void { this.entries.clear(); }
}
