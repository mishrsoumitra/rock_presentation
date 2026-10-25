// packages/metafile/src/dib.ts  -- device-independent bitmaps: headers ([MS-WMF] 2.2.2.2-2.2.2.9) and pixel decoding
import { MetafileError } from './errors.ts';
import type { Dib } from './types.ts';
import type { View } from './view.ts';

export const BI_RGB = 0;
export const BI_RLE8 = 1;
export const BI_RLE4 = 2;
export const BI_BITFIELDS = 3;
export const BI_JPEG = 4;
export const BI_PNG = 5;
export const BI_ALPHABITFIELDS = 6;

/** Largest bitmap side, and largest pixel count, the decoder accepts. */
export const MAX_DIB_SIDE = 16_384;
export const MAX_DIB_PIXELS = 64 * 1024 * 1024;

/** The header fields of a DIB and where its colour table ends, relative to the header's start. */
interface Header { dib: Omit<Dib, 'bits'>; end: number; }

/** Bytes per row of an uncompressed DIB: rows are padded to 4 bytes. */
export const dibStride = (d: Pick<Dib, 'width' | 'bitCount' | 'stride'>) => d.stride ?? (((d.width * d.bitCount + 31) >>> 5) << 2);

function checkSize(width: number, height: number): void {
  if (!(width > 0) || !(height > 0)) throw new MetafileError('malformed', `bitmap of ${width}x${height} pixels`);
  if (width > MAX_DIB_SIDE || height > MAX_DIB_SIDE || width * height > MAX_DIB_PIXELS) {
    throw new MetafileError('limit', `bitmap of ${width}x${height} pixels is over the size limit`);
  }
}

/** Read a BITMAPCOREHEADER, BITMAPINFOHEADER or V4/V5 header, its masks and its colour table, from `v` at `at`. */
function readHeader(v: View, at: number, usage: number): Header {
  const size = v.u32(at);
  let width: number, height: number, bitCount: number, compression = BI_RGB, used = 0, topDown = false;
  let masks: Dib['masks'];
  let p: number;
  if (size === 12) {
    width = v.u16(at + 4);
    height = v.u16(at + 6);
    bitCount = v.u16(at + 10);
    p = at + 12;
  } else if (size >= 16 && size <= 1024) {
    width = v.i32(at + 4);
    const h = v.i32(at + 8);
    topDown = h < 0;
    height = Math.abs(h);
    bitCount = v.u16(at + 14);
    if (size >= 20) compression = v.u32(at + 16);
    if (size >= 36) used = v.u32(at + 32);
    if (size >= 52 && (compression === BI_BITFIELDS || compression === BI_ALPHABITFIELDS)) {
      masks = [v.u32(at + 40), v.u32(at + 44), v.u32(at + 48), size >= 56 ? v.u32(at + 52) : 0];
    }
    p = at + size;
    if (size < 52 && (compression === BI_BITFIELDS || compression === BI_ALPHABITFIELDS)) {
      // the masks follow a 40-byte header
      const n = compression === BI_ALPHABITFIELDS ? 4 : 3;
      masks = [v.u32(p), v.u32(p + 4), v.u32(p + 8), n === 4 ? v.u32(p + 12) : 0];
      p += n * 4;
    }
  } else {
    throw new MetafileError('malformed', `bitmap header of ${size} bytes`);
  }
  if (compression === BI_JPEG || compression === BI_PNG) {
    // the image file carries its own size; the header's may be zero
    const encoded = compression === BI_JPEG ? 'image/jpeg' : 'image/png';
    return { dib: { width: Math.max(1, width), height: Math.max(1, height), topDown, bitCount, compression, colors: new Uint32Array(0), encoded }, end: p };
  }
  checkSize(width, height);
  if (![1, 4, 8, 16, 24, 32].includes(bitCount)) throw new MetafileError('unsupported', `${bitCount} bits per pixel`);
  if (compression === BI_RLE8 && bitCount !== 8) throw new MetafileError('malformed', 'RLE8 bitmap that is not 8 bits per pixel');
  if (compression === BI_RLE4 && bitCount !== 4) throw new MetafileError('malformed', 'RLE4 bitmap that is not 4 bits per pixel');
  if (compression > BI_ALPHABITFIELDS) throw new MetafileError('unsupported', `bitmap compression ${compression}`);
  if ((compression === BI_RLE8 || compression === BI_RLE4) && topDown) throw new MetafileError('malformed', 'top-down RLE bitmap');
  if ((compression === BI_BITFIELDS || compression === BI_ALPHABITFIELDS) && bitCount !== 16 && bitCount !== 32) {
    throw new MetafileError('malformed', `bit fields with ${bitCount} bits per pixel`);
  }
  // colour table: indexed formats always have one; others may carry an optional one (ignored, but skipped over)
  const max = bitCount <= 8 ? 1 << bitCount : 0;
  // DIB_PAL_INDICES (usage 2) has no colour table at all
  let count = usage === 2 ? 0 : bitCount <= 8 ? (used > 0 && used < max ? used : max) : Math.min(used, 256);
  const entry = usage === 1 ? 2 : size === 12 ? 3 : 4;
  // a truncated table keeps what is there (writers do declare more colours than they store)
  if (bitCount <= 8) count = Math.max(0, Math.min(count, Math.floor((v.size - p) / entry)));
  let colors = new Uint32Array(0);
  let paletteIndexes: Uint16Array | undefined;
  if (bitCount <= 8) {
    if (usage === 1) {
      paletteIndexes = new Uint16Array(count);
      for (let i = 0; i < count; i++) paletteIndexes[i] = v.u16(p + i * 2);
    } else {
      colors = new Uint32Array(count);
      for (let i = 0; i < count; i++) {
        const q = p + i * entry;
        colors[i] = (v.u8(q + 2) << 16) | (v.u8(q + 1) << 8) | v.u8(q);
      }
    }
  }
  p += count * entry;
  const dib: Omit<Dib, 'bits'> = { width, height, topDown, bitCount, compression, colors };
  if (paletteIndexes) dib.paletteIndexes = paletteIndexes;
  if (masks) dib.masks = masks;
  return { dib, end: p };
}

/** A DIB whose header (with colour table) and bits are separate, as EMF records store them. */
export function splitDib(v: View, bmiAt: number, bmiSize: number, bitsAt: number, bitsSize: number, usage: number): Dib {
  if (!v.has(bmiAt, bmiSize) || bmiSize < 12) throw new MetafileError('malformed', 'bitmap header outside its record');
  const header = readHeader(v.view(bmiAt, bmiSize), 0, usage);
  // a bits size past the record end is cut to what is there (decoding treats the rest as zero)
  const n = Math.max(0, Math.min(bitsSize, v.size - bitsAt));
  if (!v.has(bitsAt, n)) throw new MetafileError('malformed', 'bitmap bits outside their record');
  return { ...header.dib, bits: v.sub(bitsAt, n) };
}

/** A packed DIB: header, colour table and bits one after the other, as WMF records store them. */
export function packedDib(v: View, at: number, usage: number): Dib {
  const header = readHeader(v, at, usage);
  return { ...header.dib, bits: v.sub(header.end, Math.max(0, v.size - header.end)) };
}

/**
 * A WMF Bitmap16 (a device-dependent bitmap, rows padded to 2 bytes, top row first). One bit per pixel is the
 * common case; it takes the text and background colours when drawn.
 */
export function bitmap16(v: View, at: number): Dib {
  const width = v.i16(at + 2);
  const height = v.i16(at + 4);
  const widthBytes = v.i16(at + 6);
  const planes = v.u8(at + 8);
  const bitCount = v.u8(at + 9) * Math.max(1, planes);
  checkSize(width, height);
  if (![1, 4, 8, 16, 24, 32].includes(bitCount)) throw new MetafileError('unsupported', `Bitmap16 with ${bitCount} bits per pixel`);
  const stride = Math.max(widthBytes, ((width * bitCount + 15) >>> 4) << 1);
  const colors = bitCount === 4 ? VGA16 : bitCount === 8 ? GRAY256 : new Uint32Array(0);
  return { width, height, topDown: true, bitCount, compression: BI_RGB, stride, colors, bits: v.sub(at + 10, Math.max(0, v.size - at - 10)) };
}

// The 16 VGA colours, for 4-bpp device bitmaps that carry no palette.
const VGA16 = Uint32Array.from([
  0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xc0c0c0,
  0x808080, 0xff0000, 0x00ff00, 0xffff00, 0x0000ff, 0xff00ff, 0x00ffff, 0xffffff,
]);
const GRAY256 = Uint32Array.from({ length: 256 }, (_, i) => (i << 16) | (i << 8) | i);

export interface DecodeOptions {
  /** Colours for a palette-less 1-bpp bitmap (and mono brushes): 0 bits, 1 bits, as 0xRRGGBB. */
  fg?: number;
  bg?: number;
  /** A colour table to use instead of the bitmap's own (DIB_PAL_COLORS resolved through the logical palette). */
  colors?: Uint32Array;
  /** 32-bpp alpha: ignored (GDI's rule for everything but AlphaBlend), straight, or premultiplied. */
  alpha?: 'ignore' | 'straight' | 'premultiplied';
}

export interface Pixels { width: number; height: number; data: Uint8ClampedArray; }

const maskShift = (m: number) => (m === 0 ? 0 : 31 - Math.clz32(m & -m));
const maskBits = (m: number) => (m === 0 ? 0 : 32 - Math.clz32(m >>> maskShift(m)));

/** Expand a channel of `bits` bits to 0..255. */
function scale(v: number, bits: number): number {
  if (bits === 0) return 0;
  if (bits >= 8) return v >>> (bits - 8);
  return Math.round((v * 255) / ((1 << bits) - 1));
}

/**
 * Decode a DIB to RGBA, top row first. Pixels an RLE bitmap skips stay transparent; bytes missing at the end of a
 * short bitmap read as zero.
 */
export function decodeDib(d: Dib, opts: DecodeOptions = {}): Pixels {
  if (d.encoded) throw new MetafileError('unsupported', 'JPEG and PNG bitmaps are decoded by the host');
  const { width: w, height: h } = d;
  checkSize(w, h);
  // a few bytes may not ask for a huge bitmap: uncompressed pixels must be at least half there, RLE within reason
  const rle = d.compression === BI_RLE8 || d.compression === BI_RLE4;
  if (rle ? w * h > Math.max(1 << 20, d.bits.length * 1024) : d.bits.length * 2 < dibStride(d) * h) {
    throw new MetafileError('malformed', `${d.bits.length} bytes of pixels for a ${w}x${h} bitmap`);
  }
  const out = new Uint8ClampedArray(w * h * 4);
  let table = opts.colors ?? d.colors;
  if (d.bitCount === 1 && (opts.fg !== undefined || table.length === 0)) {
    table = Uint32Array.from([opts.fg ?? 0, opts.bg ?? 0xffffff]);
  }
  const bits = d.bits;
  const put = (i: number, rgb: number, a = 255) => {
    out[i] = (rgb >>> 16) & 255;
    out[i + 1] = (rgb >>> 8) & 255;
    out[i + 2] = rgb & 255;
    out[i + 3] = a;
  };
  const color = (index: number) => (index < table.length ? table[index]! : 0);

  if (d.compression === BI_RLE8 || d.compression === BI_RLE4) {
    decodeRle(d, (x, y, index) => put(((h - 1 - y) * w + x) * 4, color(index)));
    return { width: w, height: h, data: out };
  }

  const stride = dibStride(d);
  const at = (i: number) => (i < bits.length ? bits[i]! : 0);
  const masks = d.masks ?? (d.bitCount === 16 ? [0x7c00, 0x03e0, 0x001f, 0] : [0xff0000, 0xff00, 0xff, 0]);
  const shifts = masks.map(maskShift);
  const widths = masks.map(maskBits);
  const alpha = opts.alpha ?? 'ignore';
  const useMaskAlpha = d.compression === BI_ALPHABITFIELDS || (d.compression === BI_BITFIELDS && masks[3] !== 0 && alpha !== 'ignore');
  for (let y = 0; y < h; y++) {
    const row = (d.topDown ? y : h - 1 - y) * stride;
    let o = y * w * 4;
    switch (d.bitCount) {
      case 1:
        for (let x = 0; x < w; x++, o += 4) put(o, color((at(row + (x >>> 3)) >>> (7 - (x & 7))) & 1));
        break;
      case 4:
        for (let x = 0; x < w; x++, o += 4) put(o, color((at(row + (x >>> 1)) >>> (x & 1 ? 0 : 4)) & 15));
        break;
      case 8:
        for (let x = 0; x < w; x++, o += 4) put(o, color(at(row + x)));
        break;
      case 24:
        for (let x = 0; x < w; x++, o += 4) {
          const p = row + x * 3;
          out[o] = at(p + 2); out[o + 1] = at(p + 1); out[o + 2] = at(p); out[o + 3] = 255;
        }
        break;
      case 16:
      case 32: {
        const bytes = d.bitCount >>> 3;
        const plain32 = d.bitCount === 32 && d.compression === BI_RGB;
        for (let x = 0; x < w; x++, o += 4) {
          const p = row + x * bytes;
          if (plain32) {
            out[o] = at(p + 2); out[o + 1] = at(p + 1); out[o + 2] = at(p);
            out[o + 3] = alpha === 'ignore' ? 255 : at(p + 3);
            continue;
          }
          const px = bytes === 2 ? at(p) | (at(p + 1) << 8) : (at(p) | (at(p + 1) << 8) | (at(p + 2) << 16) | (at(p + 3) << 24)) >>> 0;
          for (let c = 0; c < 3; c++) out[o + c] = scale((px & masks[c]!) >>> shifts[c]!, widths[c]!);
          out[o + 3] = useMaskAlpha && masks[3] ? scale((px & masks[3]) >>> shifts[3]!, widths[3]!) : 255;
        }
        break;
      }
    }
  }
  if (alpha === 'premultiplied') {
    for (let i = 0; i < out.length; i += 4) {
      const a = out[i + 3]!;
      if (a > 0 && a < 255) {
        out[i] = Math.min(255, (out[i]! * 255) / a);
        out[i + 1] = Math.min(255, (out[i + 1]! * 255) / a);
        out[i + 2] = Math.min(255, (out[i + 2]! * 255) / a);
      }
    }
  }
  return { width: w, height: h, data: out };
}

/**
 * Walk an RLE4 or RLE8 bitmap ([MS-WMF] 3.1.6), calling `set` for each pixel it paints, with y counted from the
 * bottom row as RLE bitmaps are. Every step consumes input, so the walk ends; writes outside the bitmap are dropped.
 */
export function decodeRle(d: Dib, set: (x: number, y: number, index: number) => void): void {
  const { width: w, height: h, bits } = d;
  const four = d.compression === BI_RLE4;
  let x = 0, y = 0, i = 0;
  const plot = (index: number) => {
    if (x < w && y < h) set(x, y, index);
    x++;
  };
  while (i + 1 < bits.length && y < h) {
    const n = bits[i]!;
    const c = bits[i + 1]!;
    i += 2;
    if (n > 0) {
      // encoded run: n pixels of c (RLE4: alternating its two nibbles)
      for (let k = 0; k < n && x < w; k++) plot(four ? (k & 1 ? c & 15 : c >>> 4) : c);
      continue;
    }
    if (c === 0) { x = 0; y++; continue; }           // end of line
    if (c === 1) break;                              // end of bitmap
    if (c === 2) {                                   // delta
      if (i + 1 >= bits.length) break;
      x += bits[i]!;
      y += bits[i + 1]!;
      i += 2;
      continue;
    }
    // absolute run of c pixels, padded to a 16-bit boundary
    const bytes = four ? (c + 1) >>> 1 : c;
    for (let k = 0; k < c; k++) {
      const b = bits[i + (four ? k >>> 1 : k)];
      if (b === undefined) break;
      plot(four ? (k & 1 ? b & 15 : b >>> 4) : b);
    }
    i += bytes + (bytes & 1);
  }
}
