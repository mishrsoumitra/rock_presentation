import { describe, expect, test } from 'vitest';
import { decodeTiff, metafileImage, type Raster } from '../src/index.ts';
import { corpusTiffs } from './corpus.ts';
import { MockCanvas } from './mock.ts';

interface Spec {
  width: number;
  height: number;
  photometric: number;
  bits?: number;
  spp?: number;
  compression?: number;
  predictor?: number;
  planar?: number;
  extra?: number;
  colorMap?: number[];
  fillOrder?: number;
  rowsPerStrip?: number;
  tile?: [number, number];
  be?: boolean;
  extraTags?: [number, number][];
  chunks: Uint8Array[];
}

/** A TIFF file: header, the chunks, then one image file directory (arrays after it). */
function tiff(s: Spec): Uint8Array {
  const le = !s.be;
  const entries: [number, number, number[]][] = [];
  const short = (tag: number, ...v: number[]) => entries.push([tag, 3, v]);
  const long = (tag: number, ...v: number[]) => entries.push([tag, 4, v]);
  long(256, s.width);
  long(257, s.height);
  short(258, ...new Array<number>(s.spp ?? 1).fill(s.bits ?? 8));
  short(259, s.compression ?? 1);
  short(262, s.photometric);
  if (s.fillOrder) short(266, s.fillOrder);
  short(277, s.spp ?? 1);
  if (s.planar) short(284, s.planar);
  if (s.predictor) short(317, s.predictor);
  if (s.colorMap) short(320, ...s.colorMap);
  if (s.extra !== undefined) short(338, s.extra);
  for (const [tag, v] of s.extraTags ?? []) short(tag, v);
  let at = 8;
  const offsets = s.chunks.map(c => { const o = at; at += c.length; return o; });
  const counts = s.chunks.map(c => c.length);
  if (s.tile) {
    long(322, s.tile[0]);
    long(323, s.tile[1]);
    long(324, ...offsets);
    long(325, ...counts);
  } else {
    long(273, ...offsets);
    long(278, s.rowsPerStrip ?? s.height);
    long(279, ...counts);
  }
  entries.sort((a, b) => a[0] - b[0]);
  const ifd = at + (at & 1);
  let more = ifd + 2 + entries.length * 12 + 4;
  const size = entries.reduce((n, [, type, v]) => n + ((type === 3 ? 2 : 4) * v.length > 4 ? (type === 3 ? 2 : 4) * v.length : 0), more);
  const buf = new Uint8Array(size);
  const view = new DataView(buf.buffer);
  buf.set(le ? [0x49, 0x49] : [0x4d, 0x4d], 0);
  view.setUint16(2, 42, le);
  view.setUint32(4, ifd, le);
  s.chunks.forEach((c, i) => buf.set(c, offsets[i]!));
  view.setUint16(ifd, entries.length, le);
  entries.forEach(([tag, type, v], i) => {
    const e = ifd + 2 + i * 12;
    const w = type === 3 ? 2 : 4;
    view.setUint16(e, tag, le);
    view.setUint16(e + 2, type, le);
    view.setUint32(e + 4, v.length, le);
    let to = e + 8;
    if (w * v.length > 4) {
      view.setUint32(e + 8, more, le);
      to = more;
      more += w * v.length;
    }
    v.forEach((x, k) => (w === 2 ? view.setUint16(to + 2 * k, x, le) : view.setUint32(to + 4 * k, x, le)));
  });
  return buf;
}

/** LZW as libtiff writes it: MSB first, the code width growing one code early, a clear code when the table fills. */
function lzw(data: Uint8Array): Uint8Array {
  const out: number[] = [];
  let acc = 0;
  let nbits = 0;
  const put = (code: number, width: number) => {
    acc = (acc << width) | code;
    nbits += width;
    while (nbits >= 8) {
      out.push((acc >> (nbits - 8)) & 0xff);
      nbits -= 8;
    }
    acc &= (1 << nbits) - 1;
  };
  let dict = new Map<string, number>();
  let next = 258;
  let width = 9;
  const reset = () => {
    dict = new Map();
    for (let i = 0; i < 256; i++) dict.set(String.fromCharCode(i), i);
    next = 258;
    width = 9;
  };
  const grow = () => {
    next++;
    if (next === 4094) {
      put(256, width);
      reset();
    } else if (next > (1 << width) - 1) {
      width++;
    }
  };
  reset();
  put(256, 9);
  let w = '';
  for (const byte of data) {
    const c = String.fromCharCode(byte);
    if (dict.has(w + c)) {
      w += c;
      continue;
    }
    put(dict.get(w)!, width);
    dict.set(w + c, next);
    grow();
    w = c;
  }
  if (w) {
    put(dict.get(w)!, width);
    grow();
  }
  put(257, width);
  if (nbits) out.push((acc << (8 - nbits)) & 0xff);
  return Uint8Array.from(out);
}

async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new CompressionStream('deflate'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Bytes of a pseudo-random picture, the same on every run. */
function noise(n: number, seed = 7): Uint8Array {
  const out = new Uint8Array(n);
  let x = seed;
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0;
    out[i] = (x >>> 16) & 0xff;
  }
  return out;
}

/** The horizontal predictor, as a writer applies it to 8-bit rows of `spp` samples per pixel. */
function predict(data: Uint8Array, width: number, spp: number): Uint8Array {
  const out = data.slice();
  const row = width * spp;
  for (let y = 0; y < data.length / row; y++) {
    for (let k = row - 1; k >= spp; k--) out[y * row + k] = (data[y * row + k]! - data[y * row + k - spp]! + 256) & 0xff;
  }
  return out;
}

const pixels = (r: Raster) => Array.from({ length: r.width * r.height }, (_, i) => Array.from(r.data.subarray(i * 4, i * 4 + 4)));
const rgbOf = (r: Raster) => Uint8Array.from(pixels(r).flatMap(p => p.slice(0, 3)));

describe('TIFF', () => {
  test('uncompressed 8-bit RGB in two strips, little-endian', async () => {
    const t = tiff({ width: 2, height: 2, photometric: 2, spp: 3, rowsPerStrip: 1, chunks: [Uint8Array.of(255, 0, 0, 0, 255, 0), Uint8Array.of(0, 0, 255, 9, 8, 7)] });
    expect(pixels((await decodeTiff(t))!)).toEqual([[255, 0, 0, 255], [0, 255, 0, 255], [0, 0, 255, 255], [9, 8, 7, 255]]);
  });

  test('16-bit grey, big-endian, with the horizontal predictor', async () => {
    // 0x1000, 0x2000, 0xF000 written as differences
    const t = tiff({ width: 3, height: 1, photometric: 1, bits: 16, predictor: 2, be: true, chunks: [Uint8Array.of(0x10, 0, 0x10, 0, 0xd0, 0)] });
    expect(pixels((await decodeTiff(t))!).map(p => p[0])).toEqual([0x10, 0x20, 0xf0]);
  });

  test('PackBits, 4-bit palette', async () => {
    const colorMap = new Array<number>(48).fill(0);
    colorMap[1] = 65535;                              // index 1: red
    colorMap[32 + 2] = 65535;                         // index 2: blue
    const t = tiff({ width: 4, height: 1, photometric: 3, bits: 4, compression: 32773, colorMap, chunks: [Uint8Array.of(0xff, 0x12)] });
    expect(pixels((await decodeTiff(t))!)).toEqual([[255, 0, 0, 255], [0, 0, 255, 255], [255, 0, 0, 255], [0, 0, 255, 255]]);
  });

  test('LZW with the predictor, across code widths and table resets', async () => {
    for (const [width, height] of [[16, 8], [96, 64]] as const) {
      const rgb = noise(width * height * 3, width);
      // smooth rows compress into long strings; noise fills the table and makes the writer clear it
      for (let i = 0; i < rgb.length / 2; i++) rgb[i] = (i >> 4) & 0xff;
      const t = tiff({ width, height, photometric: 2, spp: 3, compression: 5, predictor: 2, chunks: [lzw(predict(rgb, width, 3))] });
      expect(rgbOf((await decodeTiff(t))!)).toEqual(rgb);
    }
  });

  test('Deflate RGBA with premultiplied alpha comes out straight', async () => {
    const t = tiff({ width: 2, height: 1, photometric: 2, spp: 4, extra: 1, compression: 8, chunks: [await deflate(Uint8Array.of(100, 50, 0, 128, 10, 20, 30, 0))] });
    expect(pixels((await decodeTiff(t))!)).toEqual([[199, 100, 0, 128], [10, 20, 30, 0]]);
    // straight alpha (2) is kept as it is; an unspecified extra sample (0) is not alpha
    const straight = tiff({ width: 1, height: 1, photometric: 2, spp: 4, extra: 2, chunks: [Uint8Array.of(100, 50, 0, 128)] });
    expect(pixels((await decodeTiff(straight))!)).toEqual([[100, 50, 0, 128]]);
    const unspecified = tiff({ width: 1, height: 1, photometric: 2, spp: 4, extra: 0, chunks: [Uint8Array.of(100, 50, 0, 128)] });
    expect(pixels((await decodeTiff(unspecified))!)).toEqual([[100, 50, 0, 255]]);
  });

  test('tiles, padded past the right and bottom edges', async () => {
    // a 3x3 grey picture in four 2x2 tiles: value = 10 * row + column
    const tile = (x0: number, y0: number) => Uint8Array.from([0, 1, 2, 3].map(i => 10 * (y0 + (i >> 1)) + x0 + (i & 1)));
    const t = tiff({ width: 3, height: 3, photometric: 1, tile: [2, 2], chunks: [tile(0, 0), tile(2, 0), tile(0, 2), tile(2, 2)] });
    expect(pixels((await decodeTiff(t))!).map(p => p[0])).toEqual([0, 1, 2, 10, 11, 12, 20, 21, 22]);
  });

  test('planar RGB: a strip per sample', async () => {
    const t = tiff({ width: 2, height: 1, photometric: 2, spp: 3, planar: 2, chunks: [Uint8Array.of(1, 2), Uint8Array.of(3, 4), Uint8Array.of(5, 6)] });
    expect(pixels((await decodeTiff(t))!)).toEqual([[1, 3, 5, 255], [2, 4, 6, 255]]);
  });

  test('1-bit, white is zero, bits stored least significant first', async () => {
    // pixels 1 0 1 1 (black white black black) are 0b1011_0000, stored reversed
    const t = tiff({ width: 4, height: 1, photometric: 0, bits: 1, fillOrder: 2, chunks: [Uint8Array.of(0b0000_1101)] });
    expect(pixels((await decodeTiff(t))!).map(p => p[0])).toEqual([0, 255, 0, 0]);
  });

  test('CMYK', async () => {
    const t = tiff({ width: 2, height: 1, photometric: 5, spp: 4, chunks: [Uint8Array.of(255, 0, 0, 0, 0, 0, 0, 255)] });
    expect(pixels((await decodeTiff(t))!)).toEqual([[0, 255, 255, 255], [0, 0, 0, 255]]);
  });

  test('missing or short pixel data leaves those pixels transparent', async () => {
    const t = tiff({ width: 2, height: 2, photometric: 1, rowsPerStrip: 1, chunks: [Uint8Array.of(7, 8), Uint8Array.of(9)] });
    expect(pixels((await decodeTiff(t))!)).toEqual([[7, 7, 7, 255], [8, 8, 8, 255], [0, 0, 0, 0], [0, 0, 0, 0]]);
  });

  test('what it does not read gives undefined', async () => {
    const rgb = { width: 1, height: 1, photometric: 2, spp: 3, chunks: [Uint8Array.of(1, 2, 3)] };
    expect(await decodeTiff(tiff({ ...rgb, compression: 7 }))).toBeUndefined();             // JPEG
    expect(await decodeTiff(tiff({ ...rgb, photometric: 6 }))).toBeUndefined();             // YCbCr
    expect(await decodeTiff(tiff({ ...rgb, extraTags: [[339, 3]] }))).toBeUndefined();      // floating point
    const big = tiff(rgb);
    big[2] = 43;                                                                            // BigTIFF
    expect(await decodeTiff(big)).toBeUndefined();
    expect(await decodeTiff(tiff({ ...rgb, width: 10_000, height: 10_000 }))).toBeUndefined();
    expect(await decodeTiff(Uint8Array.of(0x89, 0x50, 0x4e, 0x47))).toBeUndefined();
  });

  test('damaged files never throw', async () => {
    const base = tiff({ width: 8, height: 8, photometric: 2, spp: 3, compression: 5, predictor: 2, rowsPerStrip: 3,
      chunks: [0, 1, 2].map(i => lzw(noise(8 * 3 * 3, i))) });
    for (let n = 0; n < base.length; n += 3) await decodeTiff(base.subarray(0, n));
    const rand = noise(4000, 99);
    for (let i = 0; i < 1000; i++) {
      const b = base.slice();
      b[rand[i * 2]! % b.length] = rand[i * 2 + 1]!;
      b[(rand[i * 2 + 1]! * 7) % b.length] = rand[i * 2]!;
      const r = await decodeTiff(b);
      if (r) expect(r.data.length).toBe(r.width * r.height * 4);
    }
  });

  test('metafileImage draws a TIFF the host cannot decode', async () => {
    const t = tiff({ width: 2, height: 1, photometric: 2, spp: 3, chunks: [Uint8Array.of(1, 2, 3, 4, 5, 6)] });
    const canvas = (await metafileImage(t, (w, h) => new MockCanvas(w, h)))!;
    expect([canvas.width, canvas.height]).toEqual([2, 1]);
    expect(Array.from(canvas.ctx.images[0]!.data)).toEqual([1, 2, 3, 255, 4, 5, 6, 255]);
  });

  test('the corpus TIFF pictures decode, with their alpha', async () => {
    const all = corpusTiffs();
    for (const { name, bytes } of all) {
      const r = await decodeTiff(bytes);
      expect(r, name).toBeDefined();
      const opaque = pixels(r!).filter(p => p[3]! > 0).length;
      expect(opaque, name).toBeGreaterThan(0);
    }
  });
});
