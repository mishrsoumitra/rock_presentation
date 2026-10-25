// Hand-made pictures that exercise what the corpus lacks: EMF+ only content, embedded EMF in a WMF, RLE bitmaps,
// regions. Used by the EMF+ tests and as fuzzing seeds.
import { Bytes, dibParts, E, emf, packedDib, plus, plusComment, plusHeader, W, wmf } from './build.ts';

const OBJECT = 0x4008;

/** EmfPlusObject records: a solid brush, a pen, a path, a font and a 2x2 32-bpp bitmap image. */
export function plusObjects(): Uint8Array[] {
  const solid = (id: number, argb: number) => plus(OBJECT, (1 << 8) | id, b => b.u32(0xdbc01002, 0, argb));
  const pen = plus(OBJECT, (2 << 8) | 1, b => b.u32(0xdbc01002, 0, 0x20, 0).f32(4).u32(1).u32(0xdbc01002, 0, 0xff0000ff));
  // a closed triangle: start, line, line|close
  const path = plus(OBJECT, (3 << 8) | 2, b => b.u32(0xdbc01002, 3, 0).f32(0, 0, 50, 0, 0, 50).u8(0x00, 0x01, 0x81).pad(4));
  const font = plus(OBJECT, (6 << 8) | 3, b => b.u32(0xdbc01002).f32(12).u32(2, 1, 0, 5).utf16('Arial').pad(4));
  const image = plus(OBJECT, (5 << 8) | 4, b => b.u32(0xdbc01002, 1).i32(2, 2, 8).u32(0x26200a, 0).u32(0xffff0000, 0xff00ff00, 0xff0000ff, 0x80ffffff));
  const linear = plus(OBJECT, (1 << 8) | 5, b => b.u32(0xdbc01002, 4, 0, 0).f32(0, 0, 100, 10).u32(0xffff0000, 0xff0000ff, 0, 0));
  const region = plus(OBJECT, (4 << 8) | 6, b => b.u32(0xdbc01002, 2).u32(1).u32(0x10000000).f32(0, 0, 50, 50).u32(0x10000000).f32(25, 25, 50, 50));
  return [solid(0, 0xff00ff00), pen, path, font, image, linear, region];
}

/** An EMF+-only picture: objects, then fills, outlines, an image, a string, transforms and a clip. */
export function plusOnly(extra: Uint8Array[] = []): Uint8Array {
  const draw = [
    plusHeader(false),
    plus(0x4030, 2, b => b.f32(1)),                                      // SetPageTransform: pixels
    ...plusObjects(),
    plus(0x400a, 0x8000, b => b.u32(0xffff0000, 1).f32(10, 10, 30, 20)),  // FillRects, solid red
    plus(0x400b, 1, b => b.u32(1).f32(5, 5, 10, 10)),                    // DrawRects with pen 1
    plus(0x4014, 2, b => b.u32(0)),                                       // FillPath 2 with brush 0
    plus(0x4015, 2, b => b.u32(1)),                                       // DrawPath 2 with pen 1
    plus(0x400d, 0x4000 | 0x2000 | 1, b => b.u32(3).i16(0, 0, 10, 0, 10, 10)), // DrawLines, 16-bit, closed
    plus(0x400e, 0x8000, b => b.u32(0xff0000ff).f32(60, 60, 20, 10)),     // FillEllipse
    plus(0x4010, 0x8000, b => b.u32(0xff00ff00).f32(0, 90).f32(0, 0, 40, 40)), // FillPie
    plus(0x402a, 0, b => b.f32(1, 0, 0, 1, 5, 5)),                       // SetWorldTransform
    plus(0x401a, 4, b => b.u32(0xffffffff, 2).f32(0, 0, 2, 2).f32(70, 0, 20, 20)), // DrawImage
    plus(0x402b),                                                         // ResetWorldTransform
    plus(0x4032, 1 << 8, b => b.f32(0, 0, 60, 60)),                       // SetClipRect, intersect
    plus(0x401c, 0x8000 | 3, b => b.u32(0xff000000, 0xffffffff, 2).f32(0, 70, 100, 20).utf16('Hi').pad(4)), // DrawString
    plus(0x4031),                                                         // ResetClip
    plus(0x4025, 0, b => b.u32(7)),                                       // Save
    plus(0x402d, 0, b => b.f32(10, 10)),                                  // TranslateWorldTransform
    plus(0x400c, 0x8000, b => b.u32(0xff123456, 3).f32(0, 0, 10, 0, 0, 10)), // FillPolygon
    plus(0x4026, 0, b => b.u32(7)),                                       // Restore
    plus(0x4034, (1 << 8) | 6),                                           // SetClipRegion (intersect) with region 6
    plus(0x400a, 5, b => b.u32(5, 1).f32(0, 0, 100, 10)),                // FillRects with the linear gradient brush
    ...extra,
    plus(0x4002),
  ];
  return emf([plusComment(draw)]);
}

/** A WMF whose escapes carry a whole EMF, split in two pieces (or only the first, `complete` false). */
export function wmfWithEmf(complete = true): Uint8Array {
  const inner = emf([E.createBrush(1, 0, 0x0000ff), E.select(1), E.rectangle(10, 10, 50, 50)]);
  const half = Math.ceil(inner.length / 2);
  const piece = (from: number, to: number, remaining: number) => new Bytes()
    .u32(0x43464d57, 1, 0x10000).u16(0).u32(0, 2, to - from, remaining, inner.length).bytes(inner.subarray(from, to)).done();
  return wmf([
    W.windowExt(100, 100), W.rectangle(0, 0, 10, 10),
    W.escape(0x000f, piece(0, half, inner.length - half)),
    ...(complete ? [W.escape(0x000f, piece(half, inner.length, 0))] : []),
  ]);
}

/** An EMF of bitmaps in several encodings, regions and a gradient. */
export function bitmaps(): Uint8Array {
  const rle8 = dibParts({ width: 4, height: 2, bitCount: 8, compression: 1, colors: [0, 0xff0000, 0x00ff00], bits: [4, 1, 0, 0, 2, 2, 0, 1], rows: [] });
  const mono = dibParts({ width: 8, height: 2, bitCount: 1, colors: [0, 0xffffff], rows: [[1, 0, 1, 0, 1, 0, 1, 0], [0, 1, 0, 1, 0, 1, 0, 1]] });
  const rgb = dibParts({ width: 2, height: 2, bitCount: 24, rows: [[0xff0000, 0x00ff00], [0x0000ff, 0xffffff]] });
  return emf([
    E.stretchDIBits(0, 0, 40, 20, [0, 0, 4, 2], rle8),
    E.bitBlt(0, 30, 8, 2, 0x8800c6, mono),
    E.bitBlt(0, 30, 2, 2, 0xee0086, rgb),
    E.bitBlt(20, 30, 2, 2, 0xe20746, mono),
    E.createBrush(1, 2, 0x00ff00, 5), E.select(1),
    E.bitBlt(50, 50, 30, 30, 0xf00021),
    E.clipRegion(1, [[0, 0, 60, 60], [70, 70, 90, 90]]),
    E.patternBrush(94, 2, mono), E.select(2),
    E.ellipse(0, 0, 90, 90),
  ]);
}

/** A WMF with a placeable header, text in two character sets, clipping and a DIB. */
export function textAndClip(): Uint8Array {
  return wmf([
    W.mapMode(8), W.windowOrg(0, 0), W.windowExt(1000, 500),
    W.createFont(-80, 'Times New Roman', { weight: 700 }), W.select(0),
    W.textAlign(24 | 6), W.textOut(500, 250, 'Hello'),
    W.createFont(-60, 'Arial', { charSet: 204 }), W.select(1),
    W.intersectClip(0, 0, 600, 400),
    W.extTextOut(10, 400, [0xcf, 0xf0, 0xe8], { options: 4, rect: [0, 300, 400, 450], dx: [30, 30, 30] }),
    W.dibBitBlt(800, 100, 2, 2, 0, 0, 0xcc0020, packedDib({ width: 2, height: 2, bitCount: 4, colors: Array.from({ length: 16 }, (_, i) => i * 0x111111), rows: [[1, 2], [3, 15]] })),
    W.patBlt(0, 0, 100, 100, 0x5a0049),
  ], { l: 0, t: 0, r: 1000, b: 500, inch: 1000 });
}
