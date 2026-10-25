// packages/render/src/images.ts  -- loading the pictures a slide draws, before drawing it
import type { RenderSlide } from './tree.ts';

/** Decodes picture bytes into something drawImage accepts; the host provides it (createImageBitmap in browsers). */
export type Decoder = (bytes: Uint8Array, contentType: string, part: string) => Promise<CanvasImageSource | undefined>;

/** Formats browsers decode natively. EMF, WMF and (outside Safari) TIFF are not: @pptx/metafile draws those. */
export const BROWSER_IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp', 'image/svg+xml', 'image/x-icon', 'image/avif']);

/**
 * Decode every picture of the slides not decoded yet, into `cache`. Failures and unsupported formats are left out
 * of the cache, so the renderer draws a grey box there rather than nothing.
 */
export async function loadImages(
  slides: RenderSlide[],
  read: (part: string) => { bytes: Uint8Array; contentType: string } | undefined,
  decode: Decoder,
  cache: Map<string, CanvasImageSource>,
  failed: Set<string> = new Set(),
): Promise<void> {
  const wanted = new Set<string>();
  for (const s of slides) for (const p of s.images) if (!cache.has(p) && !failed.has(p)) wanted.add(p);
  await Promise.all([...wanted].map(async part => {
    const data = read(part);
    if (!data) {
      failed.add(part);
      return;
    }
    try {
      const img = await decode(data.bytes, data.contentType, part);
      if (img) cache.set(part, img);
      else failed.add(part);
    } catch {
      failed.add(part);
    }
  }));
}
