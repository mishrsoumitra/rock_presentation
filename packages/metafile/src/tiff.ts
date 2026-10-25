// packages/metafile/src/tiff.ts  -- TIFF pictures, which browsers other than Safari do not decode
//
// The file's first image, as straight-alpha RGBA: 1 to 16-bit samples of greys, palette colours, RGB and CMYK, in
// strips or tiles, chunky or planar, uncompressed, PackBits, LZW or Deflate, with or without the horizontal
// predictor; an alpha extra sample, premultiplied or not. Anything else (JPEG or fax compression, YCbCr, floating
// point samples, BigTIFF) and anything damaged beyond its pixel data gives undefined. Missing or short pixel data
// leaves those pixels transparent, as other readers do.

/** Decoded pixels: width x height, four bytes (RGBA, not premultiplied) per pixel, rows top to bottom. */
export interface Raster {
  width: number;
  height: number;
  data: Uint8ClampedArray;
}

/** The largest picture decoded (32 megapixels); bigger ones are refused rather than risk the memory. */
export const MAX_TIFF_PIXELS = 32 * 1024 * 1024;

/** The most one strip or tile may hold once decompressed (64 MB). */
const MAX_CHUNK = 64 * 1024 * 1024;

/** Whether bytes start like a TIFF file, in either byte order (BigTIFF is not one: its version is 43). */
export const isTiff = (b: Uint8Array): boolean =>
  b.length >= 8 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 42 && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 42));

/** Decode the first image of a TIFF file, or undefined when it is not one this reader handles. */
export async function decodeTiff(bytes: Uint8Array): Promise<Raster | undefined> {
  try {
    return await decode(bytes);
  } catch {
    return undefined;            // a field or table outside the file: DataView throws, and the file is damaged
  }
}

const TAG = {
  width: 256, height: 257, bits: 258, compression: 259, photometric: 262, fillOrder: 266, stripOffsets: 273, samples: 277,
  rowsPerStrip: 278, stripCounts: 279, planar: 284, predictor: 317, colorMap: 320, tileWidth: 322, tileLength: 323,
  tileOffsets: 324, tileCounts: 325, inkSet: 332, extraSamples: 338, sampleFormat: 339,
} as const;

const NONE = 1, LZW = 5, DEFLATE = 8, ADOBE_DEFLATE = 32946, PACKBITS = 32773;
const WHITE_IS_ZERO = 0, BLACK_IS_ZERO = 1, RGB = 2, PALETTE = 3, CMYK = 5;

async function decode(bytes: Uint8Array): Promise<Raster | undefined> {
  if (!isTiff(bytes)) return undefined;
  const le = bytes[0] === 0x49;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tags = readIfd(bytes, view, le, view.getUint32(4, le));
  const one = (tag: number, fallback: number): number => tags.get(tag)?.[0] ?? fallback;

  const width = one(TAG.width, 0);
  const height = one(TAG.height, 0);
  if (!width || !height || width * height > MAX_TIFF_PIXELS) return undefined;
  const spp = one(TAG.samples, 1);
  const bitList = tags.get(TAG.bits) ?? [1];
  const bits = bitList[0]!;
  if (spp < 1 || spp > 8 || bitList.some(b => b !== bits) || ![1, 2, 4, 8, 16].includes(bits)) return undefined;
  if (width * height * spp > 4 * MAX_TIFF_PIXELS) return undefined;
  // unsigned integers (1), or "undefined" (4), which writers use for the same
  if ((tags.get(TAG.sampleFormat) ?? [1]).some(f => f !== 1 && f !== 4)) return undefined;

  const compression = one(TAG.compression, NONE);
  if (![NONE, LZW, DEFLATE, ADOBE_DEFLATE, PACKBITS].includes(compression)) return undefined;
  const photometric = one(TAG.photometric, spp >= 3 ? RGB : BLACK_IS_ZERO);
  const colours = photometric === RGB ? 3 : photometric === CMYK ? 4 : 1;
  if (![WHITE_IS_ZERO, BLACK_IS_ZERO, RGB, PALETTE, CMYK].includes(photometric) || spp < colours) return undefined;
  if (photometric === CMYK && one(TAG.inkSet, 1) !== 1) return undefined;
  if (photometric === PALETTE && bits > 8) return undefined;
  if ((photometric === RGB || photometric === CMYK) && bits < 8) return undefined;
  const predictor = one(TAG.predictor, 1);
  if (predictor !== 1 && !(predictor === 2 && bits >= 8)) return undefined;
  // the first extra sample is alpha when it says so: 1 premultiplied, 2 straight
  const extra = tags.get(TAG.extraSamples)?.[0] ?? 0;
  const alpha = spp > colours && (extra === 1 || extra === 2) ? extra : 0;
  const palette = photometric === PALETTE ? paletteOf(tags.get(TAG.colorMap), bits) : undefined;
  if (photometric === PALETTE && !palette) return undefined;

  // the pixel data, in chunks: strips of whole rows, or tiles (padded to whole tiles at the right and bottom)
  const tiled = tags.has(TAG.tileWidth);
  const cw = tiled ? one(TAG.tileWidth, 0) : width;
  const ch = tiled ? one(TAG.tileLength, 0) : Math.min(one(TAG.rowsPerStrip, height) || height, height);
  const offsets = tags.get(tiled ? TAG.tileOffsets : TAG.stripOffsets);
  const counts = tags.get(tiled ? TAG.tileCounts : TAG.stripCounts);
  if (!offsets || !cw || !ch || cw > 65536 || ch > 65536) return undefined;
  const planar = one(TAG.planar, 1) === 2 && spp > 1;
  const across = Math.ceil(width / cw);
  const perPlane = across * Math.ceil(height / ch);
  const spc = planar ? 1 : spp;                                    // samples per pixel in a chunk
  const rowBytes = Math.ceil((cw * spc * bits) / 8);
  if (rowBytes * ch > MAX_CHUNK) return undefined;                // a tile far bigger than any picture
  const reverse = one(TAG.fillOrder, 1) === 2;

  const samples = new Uint8Array(width * height * spp);         // 8-bit values (raw values below 8 bits)
  const covered = new Uint8Array(width * height);               // pixels some chunk gave data for
  const row = new Uint16Array(cw * spc);
  const mask = bits === 16 ? 0xffff : 0xff;
  for (let plane = 0; plane < (planar ? spp : 1); plane++) {
    for (let c = 0; c < perPlane; c++) {
      const i = plane * perPlane + c;
      const at = offsets[i];
      if (at === undefined || at >= bytes.length) continue;
      const x0 = (c % across) * cw;
      const y0 = Math.floor(c / across) * ch;
      const rows = tiled ? ch : Math.min(ch, height - y0);
      const size = rowBytes * rows;
      const count = counts?.[i] ?? size;
      const chunk = await decompress(bytes.subarray(at, Math.min(bytes.length, at + count)), compression, size);
      if (reverse) for (let k = 0; k < chunk.length; k++) chunk[k] = REVERSED[chunk[k]!]!;
      const available = Math.floor(chunk.length / rowBytes);
      for (let y = 0; y < Math.min(rows, available, height - y0); y++) {
        readRow(chunk, y * rowBytes, bits, le, row);
        if (predictor === 2) for (let k = spc; k < row.length; k++) row[k] = (row[k]! + row[k - spc]!) & mask;
        for (let x = 0; x < Math.min(cw, width - x0); x++) {
          const pixel = (y0 + y) * width + x0 + x;
          covered[pixel] = 1;
          for (let s = 0; s < spc; s++) {
            const v = row[x * spc + s]!;
            samples[pixel * spp + (planar ? plane : s)] = bits === 16 ? v >> 8 : v;
          }
        }
      }
    }
  }

  const out = new Uint8ClampedArray(width * height * 4);
  const max = bits >= 8 ? 255 : (1 << bits) - 1;
  for (let p = 0; p < width * height; p++) {
    if (!covered[p]) continue;
    const s = p * spp;
    let r: number, g: number, b: number;
    if (photometric === RGB) {
      r = samples[s]!; g = samples[s + 1]!; b = samples[s + 2]!;
    } else if (photometric === CMYK) {
      const k = 255 - samples[s + 3]!;
      r = ((255 - samples[s]!) * k) / 255; g = ((255 - samples[s + 1]!) * k) / 255; b = ((255 - samples[s + 2]!) * k) / 255;
    } else if (palette) {
      const at = samples[s]! * 3;
      r = palette[at] ?? 0; g = palette[at + 1] ?? 0; b = palette[at + 2] ?? 0;
    } else {
      const v = (samples[s]! * 255) / max;
      r = g = b = photometric === WHITE_IS_ZERO ? 255 - v : v;
    }
    let a = 255;
    if (alpha) {
      a = (samples[s + colours]! * 255) / max;
      if (alpha === 1 && a > 0) {
        r = (r * 255) / a; g = (g * 255) / a; b = (b * 255) / a;
      }
    }
    const o = p * 4;
    out[o] = r; out[o + 1] = g; out[o + 2] = b; out[o + 3] = a;
  }
  return { width, height, data: out };
}

/** The integer fields of an image file directory (the rest, text and fractions, are not needed here). */
function readIfd(bytes: Uint8Array, view: DataView, le: boolean, at: number): Map<number, number[]> {
  const tags = new Map<number, number[]>();
  const n = view.getUint16(at, le);
  for (let i = 0; i < n; i++) {
    const e = at + 2 + i * 12;
    const tag = view.getUint16(e, le);
    const type = view.getUint16(e + 2, le);
    const count = view.getUint32(e + 4, le);
    const size = type === 1 || type === 7 ? 1 : type === 3 ? 2 : type === 4 ? 4 : 0;
    if (!size || !count || count > 1 << 20) continue;
    const from = size * count <= 4 ? e + 8 : view.getUint32(e + 8, le);
    if (from + size * count > bytes.length) continue;
    const values = new Array<number>(count);
    for (let k = 0; k < count; k++) {
      values[k] = size === 1 ? bytes[from + k]! : size === 2 ? view.getUint16(from + 2 * k, le) : view.getUint32(from + 4 * k, le);
    }
    tags.set(tag, values);
  }
  return tags;
}

/** A palette image's colours, as RGB bytes per index (the file gives all reds, then greens, then blues, 16-bit). */
function paletteOf(map: number[] | undefined, bits: number): Uint8Array | undefined {
  const n = 1 << bits;
  if (!map || map.length < 3 * n) return undefined;
  const out = new Uint8Array(3 * n);
  for (let i = 0; i < n; i++) {
    out[i * 3] = map[i]! >> 8;
    out[i * 3 + 1] = map[n + i]! >> 8;
    out[i * 3 + 2] = map[2 * n + i]! >> 8;
  }
  return out;
}

/** One row's sample values, unpacked from `bits` bits each (most significant bits first). */
function readRow(chunk: Uint8Array, at: number, bits: number, le: boolean, row: Uint16Array): void {
  if (bits === 8) {
    for (let k = 0; k < row.length; k++) row[k] = chunk[at + k] ?? 0;
  } else if (bits === 16) {
    for (let k = 0; k < row.length; k++) {
      const a = chunk[at + 2 * k] ?? 0;
      const b = chunk[at + 2 * k + 1] ?? 0;
      row[k] = le ? a | (b << 8) : (a << 8) | b;
    }
  } else {
    const m = (1 << bits) - 1;
    for (let k = 0; k < row.length; k++) {
      const bit = k * bits;
      row[k] = ((chunk[at + (bit >> 3)] ?? 0) >> (8 - bits - (bit & 7))) & m;
    }
  }
}

/** A chunk's bytes, decompressed to at most `size` (what it should hold): a chunk cannot make more than that. */
async function decompress(data: Uint8Array, compression: number, size: number): Promise<Uint8Array> {
  if (compression === NONE) return data.slice(0, size);
  if (compression === PACKBITS) return unpackBits(data, size);
  if (compression === LZW) return unLzw(data, size);
  return inflate(data, size);
}

function unpackBits(src: Uint8Array, size: number): Uint8Array {
  const out = new Uint8Array(size);
  let i = 0;
  let o = 0;
  while (i < src.length && o < size) {
    const n = (src[i++]! << 24) >> 24;
    if (n >= 0) {
      for (let k = 0; k <= n && i < src.length && o < size; k++) out[o++] = src[i++]!;
    } else if (n !== -128) {
      const v = src[i++] ?? 0;
      for (let k = 0; k < 1 - n && o < size; k++) out[o++] = v;
    }
  }
  return out.subarray(0, o);
}

/** TIFF's LZW: codes of 9 to 12 bits, most significant bit first, widening one code early. */
function unLzw(src: Uint8Array, size: number): Uint8Array {
  const out = new Uint8Array(size);
  const prefix = new Int16Array(4096);
  const last = new Uint8Array(4096);
  const first = new Uint8Array(4096);
  const length = new Uint16Array(4096);
  for (let i = 0; i < 256; i++) {
    prefix[i] = -1;
    last[i] = i;
    first[i] = i;
    length[i] = 1;
  }
  let next = 258;
  let width = 9;
  let old = -1;
  let o = 0;
  let bit = 0;
  const end = src.length * 8;
  // a string is written from its end, following the codes it was made from
  const emit = (code: number): void => {
    for (let c = code, at = o + length[code]! - 1; c >= 0; c = prefix[c]!, at--) if (at < size) out[at] = last[c]!;
    o += length[code]!;
  };
  const add = (code: number, byte: number): void => {
    if (next >= 4096) return;
    prefix[next] = code;
    last[next] = byte;
    first[next] = first[code]!;
    length[next] = length[code]! + 1;
    next++;
  };
  while (o < size && bit + width <= end) {
    const byte = bit >> 3;
    const word = ((src[byte] ?? 0) << 16) | ((src[byte + 1] ?? 0) << 8) | (src[byte + 2] ?? 0);
    const code = (word >> (24 - (bit & 7) - width)) & ((1 << width) - 1);
    bit += width;
    if (code === 257) break;
    if (code === 256) {
      next = 258;
      width = 9;
      old = -1;
      continue;
    }
    if (old < 0) {
      if (code > 255) break;
    } else if (code < next) {
      add(old, first[code]!);
    } else if (code === next) {
      add(old, first[old]!);
    } else {
      break;                                                     // a code not made yet: the data is damaged
    }
    emit(code);
    old = code;
    if (next >= (1 << width) - 1 && width < 12) width++;
  }
  return out.subarray(0, Math.min(o, size));
}

/** Deflate (zlib) data, through the platform's decompressor, stopping at `size` bytes. */
async function inflate(src: Uint8Array, size: number): Promise<Uint8Array> {
  const out = new Uint8Array(size);
  let o = 0;
  const reader = new Blob([src as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate')).getReader();
  try {
    while (o < size) {
      const { done, value } = await reader.read();
      if (done) break;
      out.set(value.subarray(0, size - o), o);
      o += value.length;
    }
  } catch {
    // damaged or cut short: the rows that came out
  } finally {
    reader.cancel().catch(() => undefined);
  }
  return out.subarray(0, Math.min(o, size));
}

const REVERSED = Uint8Array.from({ length: 256 }, (_, i) => {
  let r = 0;
  for (let b = 0; b < 8; b++) if (i & (1 << b)) r |= 0x80 >> b;
  return r;
});
