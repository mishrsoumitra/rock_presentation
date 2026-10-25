// packages/metafile/src/index.ts  -- EMF and WMF pictures, drawn with Canvas 2D
//
// parseMetafile turns the bytes into records, drawMetafile plays them onto any 2D context, and metafileImage
// does both onto a new canvas sized from the picture, for hosts that treat pictures as images.
import { isEmf, parseEmf } from './emf.ts';
import { MetafileError } from './errors.ts';
import { dibImageKey, drawMetafile, playsPlus, type DrawOptions, type MetaCtx } from './player.ts';
import { plusImages } from './plus.ts';
import type { Metafile } from './types.ts';
import { decodeTiff } from './tiff.ts';
import { parseWmf, wmfHeaderAt } from './wmf.ts';

export { MetafileError, type MetafileErrorCode } from './errors.ts';
export { decodeDib, type DecodeOptions, type Pixels } from './dib.ts';
export { DEFAULT_MAX_WORK, dibImageKey, drawMetafile, type Dest, type DrawOptions, type MetaCtx, type ScratchCanvas } from './player.ts';
export { decodeTiff, isTiff, MAX_TIFF_PIXELS, type Raster } from './tiff.ts';
export type * from './types.ts';

/** Content types Office gives metafile parts. */
export const METAFILE_TYPES = new Set(['image/x-emf', 'image/emf', 'image/x-wmf', 'image/wmf']);

/** Content types of pictures to draw with metafileImage rather than the browser: metafiles, and TIFF (only Safari decodes it). */
export const PICTURE_TYPES = new Set([...METAFILE_TYPES, 'image/tiff']);

/** Parse an EMF or WMF (told apart by content, not by name) into its records. */
export function parseMetafile(bytes: Uint8Array): Metafile {
  if (isEmf(bytes)) return parseEmf(bytes);
  if (wmfHeaderAt(bytes) >= 0) return parseWmf(bytes, parseEmf);
  throw new MetafileError('unsupported', 'neither an EMF nor a WMF');
}

/** The longer side metafileImage draws at least, and the most pixels it draws. */
export const MIN_SIDE = 1024;
export const MAX_PIXELS = 16 * 1024 * 1024;
const MAX_SIDE = 16_384;

/**
 * The canvas size for a picture: its natural size at 96 dpi, scaled so that the longer side is at least
 * MIN_SIDE pixels (the renderer scales it into the shape box), and down to at most MAX_PIXELS pixels.
 */
export function canvasSize(mf: Pick<Metafile, 'width' | 'height' | 'frame'>): { width: number; height: number } {
  let w = mf.width, h = mf.height;
  if (!(w > 0 && h > 0) || !Number.isFinite(w * h)) {
    w = Math.abs(mf.frame.r - mf.frame.l);
    h = Math.abs(mf.frame.b - mf.frame.t);
    if (!(w > 0 && h > 0) || !Number.isFinite(w * h)) { w = 1; h = 1; }
  }
  let k = Math.max(1, MIN_SIDE / Math.max(w, h));
  k = Math.min(k, Math.sqrt(MAX_PIXELS / (w * h)), MAX_SIDE / w, MAX_SIDE / h);
  return { width: Math.max(1, Math.min(MAX_SIDE, Math.round(w * k))), height: Math.max(1, Math.min(MAX_SIDE, Math.round(h * k))) };
}

/** Whether playback would draw anything: some empty pictures (think-cell placeholders) are only a header. */
export function hasDrawing(mf: Metafile): boolean {
  return mf.records.some(r => r.type === 'plus' ? mf.plus === 'only' : DRAWS.has(r.type));
}
const DRAWS = new Set(['lineTo', 'poly', 'polyPoly', 'polyDraw', 'box', 'roundRect', 'arc', 'angleArc', 'pixel', 'path', 'gradient', 'region', 'text', 'blt']);

/** Decodes a JPEG, PNG, GIF, BMP or TIFF file (given with its media type) into something drawImage takes. */
export type ImageDecoder = (bytes: Uint8Array, type: string) => Promise<CanvasImageSource | undefined>;

/** The browser's decoder (createImageBitmap), where there is one. */
const hostDecoder: ImageDecoder | undefined = typeof createImageBitmap === 'function' && typeof Blob === 'function'
  ? async (bytes, type) => createImageBitmap(new Blob([bytes as BlobPart], { type }))
  : undefined;

const sniff = (b: Uint8Array): string | undefined =>
  b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47 ? 'image/png'
    : b[0] === 0xff && b[1] === 0xd8 ? 'image/jpeg'
      : b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 ? 'image/gif'
        : b[0] === 0x42 && b[1] === 0x4d ? 'image/bmp'
          : (b[0] === 0x49 && b[1] === 0x49) || (b[0] === 0x4d && b[1] === 0x4d) ? 'image/tiff' : undefined;

/**
 * Decode the compressed pictures a metafile carries (JPEG and PNG bitmaps in EMF records; JPEG, PNG and GIF
 * images in EMF+ objects, when `plus` picks the EMF+ records as DrawOptions.plus does) before playback, which
 * is synchronous. Undecodable ones are left out.
 */
export async function decodeImages(mf: Metafile, plus: DrawOptions['plus'] = 'auto', decode: ImageDecoder | undefined = hostDecoder):
  Promise<Map<string, CanvasImageSource>> {
  const out = new Map<string, CanvasImageSource>();
  if (!decode) return out;
  const wanted: { key: string; bytes: Uint8Array; type: string }[] = [];
  for (const r of mf.records) {
    if (r.type === 'blt' && r.dib?.encoded) wanted.push({ key: dibImageKey(r.dib.bits), bytes: r.dib.bits, type: r.dib.encoded });
  }
  if (playsPlus(mf, plus)) {
    for (const img of plusImages(mf)) {
      const type = sniff(img.bytes);
      if (type) wanted.push({ ...img, type });
    }
  }
  await Promise.all(wanted.map(async w => {
    try {
      const img = await decode(w.bytes, w.type);
      if (img) out.set(w.key, img);
    } catch {
      // a broken embedded picture is left out, like any picture that does not decode
    }
  }));
  return out;
}

async function tiffCanvas<C extends MetaCanvas>(bytes: Uint8Array, createCanvas: (w: number, h: number) => C): Promise<C | undefined> {
  const raster = await decodeTiff(bytes);
  if (!raster) return undefined;
  const canvas = createCanvas(raster.width, raster.height);
  const ctx = canvas.getContext('2d');
  if (!ctx) return undefined;
  const image = ctx.createImageData(raster.width, raster.height);
  image.data.set(raster.data);
  ctx.putImageData(image, 0, 0);
  return canvas;
}

/** A canvas this package can draw on: OffscreenCanvas, HTMLCanvasElement, or anything with a matching 2D context. */
export interface MetaCanvas {
  getContext(contextId: '2d'): (MetaCtx & ImageDataSink) | null;
}
type ImageDataSink = { createImageData(sw: number, sh: number): ImageData; putImageData(data: ImageData, dx: number, dy: number): void };

/**
 * Draw an EMF or WMF picture onto a new canvas made by `createCanvas`, sized from the picture's own frame (see
 * canvasSize), or undefined when the bytes are not a picture this package can draw. A picture with nothing in it
 * gives a 1x1 transparent canvas. Raster files saved under a metafile name (it happens) are drawn as they are, and
 * TIFF files (PICTURE_TYPES) are decoded here when the browser cannot.
 */
export async function metafileImage<C extends MetaCanvas>(bytes: Uint8Array, createCanvas: (w: number, h: number) => C): Promise<C | undefined> {
  try {
    if (!isEmf(bytes) && wmfHeaderAt(bytes) < 0) return await rasterImage(bytes, createCanvas);
    const mf = parseMetafile(bytes);
    if (!hasDrawing(mf)) return createCanvas(1, 1);
    const { width, height } = canvasSize(mf);
    const images = await decodeImages(mf);
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d');
    if (!ctx) return undefined;
    drawMetafile(ctx, mf, { x: 0, y: 0, w: width, h: height }, { createCanvas, images, snap: true });
    return canvas;
  } catch (e) {
    if (e instanceof MetafileError) return undefined;
    throw e;
  }
}

async function rasterImage<C extends MetaCanvas>(bytes: Uint8Array, createCanvas: (w: number, h: number) => C): Promise<C | undefined> {
  const type = sniff(bytes);
  if (!type) return undefined;
  let img: CanvasImageSource | undefined;
  try {
    img = hostDecoder ? await hostDecoder(bytes, type) : undefined;
  } catch {
    img = undefined;
  }
  // the browser's own decoder first; TIFF, which most do not decode, with ours
  if (!img) return type === 'image/tiff' ? tiffCanvas(bytes, createCanvas) : undefined;
  const { width, height } = img as { width: number; height: number };
  if (!(width > 0 && height > 0)) return undefined;
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d');
  ctx?.drawImage(img, 0, 0, width, height, 0, 0, width, height);
  return ctx ? canvas : undefined;
}
