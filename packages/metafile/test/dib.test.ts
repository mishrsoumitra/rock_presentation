// DIB decoding: every header kind, bit depth and compression on hand-made bitmaps.
import { describe, expect, test } from 'vitest';
import { decodeDib, MetafileError } from '../src/index.ts';
import { decodeRle, packedDib as readPacked, splitDib } from '../src/dib.ts';
import { View } from '../src/view.ts';
import { dibParts, packedDib, type DibSpec } from './build.ts';

/** Decode a spec through the split (EMF) path and return the pixels as 0xRRGGBBAA numbers, rows top first. */
function decode(s: DibSpec, opts: Parameters<typeof decodeDib>[1] = {}): number[][] {
  const { header, bits } = dibParts(s);
  const all = new Uint8Array(header.length + bits.length);
  all.set(header);
  all.set(bits, header.length);
  const dib = splitDib(new View(all), 0, header.length, header.length, bits.length, 0);
  const px = decodeDib(dib, opts);
  const rows: number[][] = [];
  for (let y = 0; y < px.height; y++) {
    const row: number[] = [];
    for (let x = 0; x < px.width; x++) {
      const i = (y * px.width + x) * 4;
      row.push(((px.data[i]! << 24) | (px.data[i + 1]! << 16) | (px.data[i + 2]! << 8) | px.data[i + 3]!) >>> 0);
    }
    rows.push(row);
  }
  return rows;
}

const RED = 0xff0000ff, GREEN = 0x00ff00ff, BLUE = 0x0000ffff, BLACK = 0x000000ff, WHITE = 0xffffffff;

describe('bit depths', () => {
  test('1 bpp with a colour table, bottom-up', () => {
    expect(decode({ width: 3, height: 2, bitCount: 1, colors: [0xff0000, 0x0000ff], rows: [[1, 0, 1], [0, 0, 1]] }))
      .toEqual([[BLUE, RED, BLUE], [RED, RED, BLUE]]);
  });

  test('1 bpp takes the text and background colours when asked (mono brushes)', () => {
    expect(decode({ width: 2, height: 1, bitCount: 1, colors: [0, 0xffffff], rows: [[0, 1]] }, { fg: 0x00ff00, bg: 0xff0000 }))
      .toEqual([[GREEN, RED]]);
  });

  test('4 bpp, two pixels a byte, high nibble first', () => {
    const colors = Array.from({ length: 16 }, (_, i) => (i === 3 ? 0x00ff00 : i === 12 ? 0x0000ff : 0));
    expect(decode({ width: 3, height: 1, bitCount: 4, colors, rows: [[3, 12, 3]] })).toEqual([[GREEN, BLUE, GREEN]]);
  });

  test('8 bpp, a short colour table: indexes past it are black', () => {
    expect(decode({ width: 3, height: 1, bitCount: 8, colors: [0xff0000, 0x00ff00], rows: [[1, 0, 200]] })).toEqual([[GREEN, RED, BLACK]]);
  });

  test('16 bpp defaults to 5-5-5', () => {
    expect(decode({ width: 2, height: 1, bitCount: 16, rows: [[0x7c00, 0x001f]] })).toEqual([[RED, BLUE]]);
  });

  test('16 bpp with 5-6-5 bit fields', () => {
    expect(decode({ width: 2, height: 1, bitCount: 16, compression: 3, masks: [0xf800, 0x07e0, 0x001f], rows: [[0x07e0, 0xf800]] }))
      .toEqual([[GREEN, RED]]);
  });

  test('24 bpp is stored blue, green, red; rows are padded to 4 bytes', () => {
    expect(decode({ width: 3, height: 2, bitCount: 24, rows: [[0xff0000, 0x00ff00, 0x0000ff], [0xffffff, 0, 0x808080]] }))
      .toEqual([[RED, GREEN, BLUE], [WHITE, BLACK, 0x808080ff]]);
  });

  test('32 bpp: the fourth byte is ignored unless alpha is asked for', () => {
    const s: DibSpec = { width: 2, height: 1, bitCount: 32, rows: [[0x40ff0000, 0x80008000]] };
    expect(decode(s)).toEqual([[RED, 0x008000ff]]);
    expect(decode(s, { alpha: 'straight' })).toEqual([[0xff000040, 0x00800080]]);
    // premultiplied: the colour is divided back out
    expect(decode(s, { alpha: 'premultiplied' })[0]![1]).toBe(0x00ff0080);
  });

  test('32 bpp bit fields with an alpha mask (V5 header)', () => {
    const s: DibSpec = { width: 1, height: 1, bitCount: 32, compression: 3, headerSize: 124, masks: [0x00ff0000, 0x0000ff00, 0x000000ff, 0xff000000], rows: [[0x80123456]] };
    expect(decode(s, { alpha: 'straight' })).toEqual([[0x12345680]]);
    expect(decode(s)).toEqual([[0x123456ff]]);
  });

  test('top-down bitmaps keep their row order', () => {
    expect(decode({ width: 1, height: 2, bitCount: 24, topDown: true, rows: [[0xff0000], [0x0000ff]] })).toEqual([[RED], [BLUE]]);
  });

  test('BITMAPCOREHEADER with RGBTRIPLE colours', () => {
    expect(decode({ width: 2, height: 1, bitCount: 8, core: true, colors: [0x0000ff, 0x00ff00], rows: [[1, 0]] })).toEqual([[GREEN, BLUE]]);
  });

  test('V4 header with bit fields inside it', () => {
    expect(decode({ width: 1, height: 1, bitCount: 16, compression: 3, headerSize: 108, masks: [0xf800, 0x07e0, 0x001f, 0], rows: [[0x001f]] }))
      .toEqual([[BLUE]]);
  });

  test('a bitmap shorter than it says reads the missing bytes as zero', () => {
    const { header, bits } = dibParts({ width: 2, height: 2, bitCount: 24, rows: [[0xffffff, 0xffffff], [0xffffff, 0xffffff]] });
    const all = new Uint8Array(header.length + 8);
    all.set(header);
    all.set(bits.subarray(0, 8), header.length);
    const px = decodeDib(splitDib(new View(all), 0, header.length, header.length, 8, 0));
    // bottom-up: the bottom row is first in the file and survives; the top row is missing
    expect(Array.from(px.data.subarray(8, 16))).toEqual([255, 255, 255, 255, 255, 255, 255, 255]);
    expect(Array.from(px.data.subarray(0, 4))).toEqual([0, 0, 0, 255]);
  });
});

describe('RLE', () => {
  test('RLE8: encoded runs, an absolute run, end of line, delta, end of bitmap; skipped pixels stay transparent', () => {
    const colors = [0x000000, 0xff0000, 0x00ff00, 0x0000ff];
    // bottom row: 2 x colour 1, then absolute [2, 3, 2], end of line; next row: delta (1, 0), 1 x colour 3, end
    const bits = [2, 1, 0, 3, 2, 3, 2, 0, 0, 0, 0, 2, 1, 0, 1, 3, 0, 1];
    const rows = decode({ width: 5, height: 2, bitCount: 8, compression: 1, colors, bits, rows: [] });
    expect(rows[1]).toEqual([RED, RED, GREEN, BLUE, GREEN]);
    expect(rows[0]).toEqual([0, BLUE, 0, 0, 0]);
  });

  test('RLE4: runs alternate the two nibbles; absolute runs are padded', () => {
    const colors = [0x000000, 0xff0000, 0x00ff00, 0x0000ff];
    // 3 pixels alternating 1 and 2, then absolute 3 pixels [3, 1, 2] (2 bytes, padded to an even count), end
    const bits = [3, 0x12, 0, 3, 0x31, 0x20, 0, 1];
    expect(decode({ width: 6, height: 1, bitCount: 4, compression: 2, colors, bits, rows: [] })[0])
      .toEqual([RED, GREEN, RED, BLUE, RED, GREEN]);
  });

  test('a delta past the bitmap and runs past its width are dropped, and the walk ends', () => {
    const d = { width: 2, height: 2, topDown: false, bitCount: 8, compression: 1, colors: new Uint32Array(0), bits: Uint8Array.from([200, 1, 0, 2, 50, 50, 5, 1]) };
    const seen: [number, number][] = [];
    decodeRle(d, (x, y) => seen.push([x, y]));
    expect(seen).toEqual([[0, 0], [1, 0]]);
  });
});

describe('headers and limits', () => {
  test('a packed DIB finds its bits after the colour table', () => {
    const bytes = packedDib({ width: 2, height: 1, bitCount: 8, colors: [0xff0000, 0x00ff00], rows: [[1, 0]] });
    const px = decodeDib(readPacked(new View(bytes), 0, 0));
    expect(Array.from(px.data)).toEqual([0, 255, 0, 255, 255, 0, 0, 255]);
  });

  test('DIB_PAL_COLORS keeps palette indexes for playback to resolve', () => {
    const bytes = packedDib({ width: 1, height: 1, bitCount: 8, rows: [[0]] });
    // rewrite as two 16-bit palette indexes in place of the (empty) colour table
    const header = new Uint8Array(bytes.subarray(0, 40));
    new DataView(header.buffer).setUint32(32, 2, true);
    const all = Uint8Array.from([...header, 5, 0, 7, 0, 1, 0, 0, 0]);
    const dib = readPacked(new View(all), 0, 1);
    expect(Array.from(dib.paletteIndexes!)).toEqual([5, 7]);
    expect(Array.from(decodeDib(dib, { colors: Uint32Array.from([0x112233, 0x445566]) }).data)).toEqual([0x44, 0x55, 0x66, 255]);
  });

  test('JPEG and PNG bitmaps are marked for the host to decode', () => {
    const { header } = dibParts({ width: 10, height: 10, bitCount: 0, compression: 5, rows: [] });
    const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
    const all = Uint8Array.from([...header, ...png]);
    const dib = splitDib(new View(all), 0, header.length, header.length, png.length, 0);
    expect(dib.encoded).toBe('image/png');
    expect(Array.from(dib.bits)).toEqual(Array.from(png));
    expect(() => decodeDib(dib)).toThrow(MetafileError);
  });

  test('sizes over the limits are refused with a limit error, impossible ones as malformed', () => {
    const big = dibParts({ width: 20_000, height: 1, bitCount: 24, rows: [] }).header;
    expect(() => splitDib(new View(big), 0, big.length, 0, 0, 0)).toThrow(expect.objectContaining({ code: 'limit' }));
    const huge = dibParts({ width: 16_000, height: 16_000, bitCount: 1, rows: [] }).header;
    expect(() => splitDib(new View(huge), 0, huge.length, 0, 0, 0)).toThrow(expect.objectContaining({ code: 'limit' }));
    const zero = dibParts({ width: 0, height: 5, bitCount: 24, rows: [] }).header;
    expect(() => splitDib(new View(zero), 0, zero.length, 0, 0, 0)).toThrow(expect.objectContaining({ code: 'malformed' }));
    const odd = dibParts({ width: 1, height: 1, bitCount: 7, rows: [] }).header;
    expect(() => splitDib(new View(odd), 0, odd.length, 0, 0, 0)).toThrow(expect.objectContaining({ code: 'unsupported' }));
    const header = dibParts({ width: 1, height: 1, bitCount: 24, rows: [] }).header;
    expect(() => splitDib(new View(header), 0, header.length + 4, 0, 0, 0)).toThrow(expect.objectContaining({ code: 'malformed' }));
  });
});
