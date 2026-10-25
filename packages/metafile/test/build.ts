// Builds small EMF, WMF and DIB files byte by byte, for tests.

/** A little-endian byte writer. */
export class Bytes {
  private out: number[] = [];
  u8(...v: number[]): this { for (const x of v) this.out.push(x & 0xff); return this; }
  u16(...v: number[]): this { for (const x of v) this.out.push(x & 0xff, (x >>> 8) & 0xff); return this; }
  i16(...v: number[]): this { return this.u16(...v.map(x => x & 0xffff)); }
  u32(...v: number[]): this { for (const x of v) this.out.push(x & 0xff, (x >>> 8) & 0xff, (x >>> 16) & 0xff, (x >>> 24) & 0xff); return this; }
  i32(...v: number[]): this { return this.u32(...v.map(x => x >>> 0)); }
  f32(...v: number[]): this {
    const b = new Uint8Array(4);
    for (const x of v) { new DataView(b.buffer).setFloat32(0, x, true); this.out.push(...b); }
    return this;
  }
  bytes(b: ArrayLike<number>): this { for (let i = 0; i < b.length; i++) this.out.push(b[i]! & 0xff); return this; }
  utf16(s: string): this { for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i)); return this; }
  pad(to: number): this { while (this.out.length % to) this.out.push(0); return this; }
  get length(): number { return this.out.length; }
  done(): Uint8Array { return Uint8Array.from(this.out); }
}

// ---------------------------------------------------------------------------------------------------------- EMF

/** An EMF record: type, size, then the body (padded to 4 bytes). */
export function emr(type: number, body: Bytes | ((b: Bytes) => Bytes) = new Bytes()): Uint8Array {
  const b = typeof body === 'function' ? body(new Bytes()) : body;
  const data = b.pad(4).done();
  return new Bytes().u32(type, 8 + data.length).bytes(data).done();
}

export interface EmfOptions {
  bounds?: [number, number, number, number];
  /** Frame in 0.01 mm. */
  frame?: [number, number, number, number];
  device?: [number, number];
  mm?: [number, number];
  handles?: number;
}

/** A whole EMF: header, the records, EOF. By default a 100x100-pixel reference device at 100 pixels per inch... */
export function emf(records: Uint8Array[], o: EmfOptions = {}): Uint8Array {
  const device = o.device ?? [1000, 1000];
  const mm = o.mm ?? [254, 254];                           // 100 dpi: one pixel is 0.254 mm
  const frame = o.frame ?? [0, 0, 2540, 2540];             // 25.4 mm = 100 device pixels
  const bounds = o.bounds ?? [0, 0, 99, 99];
  const body = [...records, emr(14, b => b.u32(0, 0, 20))];
  const size = 108 + body.reduce((n, r) => n + r.length, 0);
  const header = new Bytes()
    .u32(1, 108).i32(...bounds).i32(...frame).u32(0x464d4520, 0x10000, size, body.length + 1).u16(o.handles ?? 16, 0)
    .u32(0, 0, 0).i32(...device).i32(...mm).u32(0, 0, 0).i32(mm[0] * 1000, mm[1] * 1000);
  const out = new Bytes().bytes(header.done());
  for (const r of body) out.bytes(r);
  return out.done();
}

export const rectL = (b: Bytes, l: number, t: number, r: number, bb: number) => b.i32(l, t, r, bb);

// common EMF records
export const E = {
  mapMode: (m: number) => emr(17, b => b.u32(m)),
  windowExt: (x: number, y: number) => emr(9, b => b.i32(x, y)),
  windowOrg: (x: number, y: number) => emr(10, b => b.i32(x, y)),
  viewportExt: (x: number, y: number) => emr(11, b => b.i32(x, y)),
  viewportOrg: (x: number, y: number) => emr(12, b => b.i32(x, y)),
  saveDC: () => emr(33),
  restoreDC: (n: number) => emr(34, b => b.i32(n)),
  worldTransform: (m: number[]) => emr(35, b => b.f32(...m)),
  modifyWorldTransform: (m: number[], mode: number) => emr(36, b => b.f32(...m).u32(mode)),
  select: (i: number) => emr(37, b => b.u32(i)),
  deleteObject: (i: number) => emr(40, b => b.u32(i)),
  createPen: (i: number, style: number, width: number, color: number) => emr(38, b => b.u32(i, style).i32(width, 0).u32(color)),
  extCreatePen: (i: number, style: number, width: number, color: number, dashes: number[] = []) =>
    emr(95, b => b.u32(i, 0, 0, 0, 0, style, width, 0, color, 0, dashes.length).u32(...dashes)),
  createBrush: (i: number, style: number, color: number, hatch = 0) => emr(39, b => b.u32(i, style, color, hatch)),
  createFont: (i: number, height: number, face: string, o: { weight?: number; italic?: boolean; escapement?: number; charSet?: number } = {}) =>
    emr(82, b => {
      b.u32(i).i32(height, 0, o.escapement ?? 0, o.escapement ?? 0, o.weight ?? 400).u8(o.italic ? 1 : 0, 0, 0, o.charSet ?? 0, 0, 0, 0, 0);
      const name = new Bytes().utf16(face.slice(0, 31)).done();
      b.bytes(name);
      for (let k = name.length; k < 64; k++) b.u8(0);
      return b;
    }),
  bkMode: (m: number) => emr(18, b => b.u32(m)),
  rop2: (r: number) => emr(20, b => b.u32(r)),
  /** CREATEMONOBRUSH (93) or CREATEDIBPATTERNBRUSHPT (94) with the DIB right after the fixed part. */
  patternBrush: (type: 93 | 94, i: number, dibBytes: { header: Uint8Array; bits: Uint8Array }) =>
    emr(type, b => b.u32(i, 0, 32, dibBytes.header.length, 32 + dibBytes.header.length, dibBytes.bits.length).bytes(dibBytes.header).bytes(dibBytes.bits)),
  bkColor: (c: number) => emr(25, b => b.u32(c)),
  textColor: (c: number) => emr(24, b => b.u32(c)),
  textAlign: (a: number) => emr(22, b => b.u32(a)),
  polyFillMode: (m: number) => emr(19, b => b.u32(m)),
  arcDirection: (d: number) => emr(57, b => b.u32(d)),
  moveTo: (x: number, y: number) => emr(27, b => b.i32(x, y)),
  lineTo: (x: number, y: number) => emr(54, b => b.i32(x, y)),
  rectangle: (l: number, t: number, r: number, bb: number) => emr(43, b => b.i32(l, t, r, bb)),
  ellipse: (l: number, t: number, r: number, bb: number) => emr(42, b => b.i32(l, t, r, bb)),
  arc: (box: number[], start: number[], end: number[]) => emr(45, b => b.i32(...box, ...start, ...end)),
  polygon: (pts: number[]) => emr(3, b => b.i32(0, 0, 0, 0).u32(pts.length / 2).i32(...pts)),
  polyline16: (pts: number[]) => emr(87, b => b.i32(0, 0, 0, 0).u32(pts.length / 2).i16(...pts)),
  polyBezierTo: (pts: number[]) => emr(5, b => b.i32(0, 0, 0, 0).u32(pts.length / 2).i32(...pts)),
  polyPolygon: (polys: number[][]) => emr(8, b => {
    b.i32(0, 0, 0, 0).u32(polys.length, polys.reduce((n, p) => n + p.length / 2, 0));
    for (const p of polys) b.u32(p.length / 2);
    for (const p of polys) b.i32(...p);
    return b;
  }),
  beginPath: () => emr(59),
  endPath: () => emr(60),
  closeFigure: () => emr(61),
  fillPath: () => emr(62, b => b.i32(0, 0, 0, 0)),
  strokePath: () => emr(64, b => b.i32(0, 0, 0, 0)),
  selectClipPath: (mode: number) => emr(67, b => b.u32(mode)),
  intersectClip: (l: number, t: number, r: number, bb: number) => emr(30, b => b.i32(l, t, r, bb)),
  excludeClip: (l: number, t: number, r: number, bb: number) => emr(29, b => b.i32(l, t, r, bb)),
  clipRegion: (mode: number, rects: number[][] | null) => emr(75, b => {
    if (!rects) return b.u32(0, mode);
    b.u32(32 + rects.length * 16, mode).u32(32, 1, rects.length, rects.length * 16).i32(0, 0, 0, 0);
    for (const r of rects) b.i32(...r);
    return b;
  }),
  extTextOutW: (x: number, y: number, text: string, o: { options?: number; dx?: number[]; rect?: number[]; mode?: number } = {}) =>
    emr(84, b => {
      const n = text.length;
      const offString = 76;
      const offDx = o.dx ? offString + ((n * 2 + 3) & ~3) : 0;
      b.i32(0, 0, 0, 0).u32(o.mode ?? 1).f32(1, 1).i32(x, y).u32(n, offString, o.options ?? 0).i32(...(o.rect ?? [0, 0, 0, 0])).u32(offDx);
      b.utf16(text).pad(4);
      if (o.dx) b.i32(...o.dx);
      return b;
    }),
  bitBlt: (x: number, y: number, w: number, h: number, rop: number, dibBytes?: { header: Uint8Array; bits: Uint8Array }) => emr(76, b => {
    b.i32(0, 0, 0, 0).i32(x, y, w, h).u32(rop).i32(0, 0).f32(1, 0, 0, 1, 0, 0).u32(0, 0);
    if (!dibBytes) return b.u32(0, 0, 0, 0);
    const offBmi = 100, offBits = offBmi + dibBytes.header.length;
    return b.u32(offBmi, dibBytes.header.length, offBits, dibBytes.bits.length).bytes(dibBytes.header).bytes(dibBytes.bits);
  }),
  stretchDIBits: (x: number, y: number, w: number, h: number, src: number[], dibBytes: { header: Uint8Array; bits: Uint8Array }, rop = 0xcc0020) => emr(81, b => {
    const offBmi = 80, offBits = offBmi + dibBytes.header.length;
    return b.i32(0, 0, 0, 0).i32(x, y, ...src).u32(offBmi, dibBytes.header.length, offBits, dibBytes.bits.length, 0, rop).i32(w, h)
      .bytes(dibBytes.header).bytes(dibBytes.bits);
  }),
  comment: (data: Uint8Array) => emr(70, b => b.u32(data.length).bytes(data)),
  setDIBitsToDevice: (x: number, y: number, w: number, h: number, dibBytes: { header: Uint8Array; bits: Uint8Array }) => emr(80, b =>
    b.i32(0, 0, 0, 0).i32(x, y, 0, 0, w, h).u32(76, dibBytes.header.length, 76 + dibBytes.header.length, dibBytes.bits.length, 0, 0, h)
      .bytes(dibBytes.header).bytes(dibBytes.bits)),
  /** ALPHABLEND (114) or TRANSPARENTBLT (116): `extra` is the BLENDFUNCTION or the transparent colour. */
  blend: (type: 114 | 116, x: number, y: number, w: number, h: number, extra: number, dibBytes: { header: Uint8Array; bits: Uint8Array }) => emr(type, b => {
    b.i32(0, 0, 0, 0).i32(x, y, w, h).u32(extra).i32(0, 0).f32(1, 0, 0, 1, 0, 0).u32(0, 0);
    return b.u32(108, dibBytes.header.length, 108 + dibBytes.header.length, dibBytes.bits.length).i32(w, h).bytes(dibBytes.header).bytes(dibBytes.bits);
  }),
  /** MASKBLT (78) with a source and a 1-bpp mask. */
  maskBlt: (x: number, y: number, w: number, h: number, src: { header: Uint8Array; bits: Uint8Array }, mask: { header: Uint8Array; bits: Uint8Array }) => emr(78, b => {
    const s0 = 128, s1 = s0 + src.header.length, m0 = s1 + src.bits.length, m1 = m0 + mask.header.length;
    b.i32(0, 0, 0, 0).i32(x, y, w, h).u32(0xaacc0020).i32(0, 0).f32(1, 0, 0, 1, 0, 0).u32(0, 0);
    b.u32(s0, src.header.length, s1, src.bits.length).i32(0, 0).u32(0, m0, mask.header.length, m1, mask.bits.length);
    return b.bytes(src.header).bytes(src.bits).bytes(mask.header).bytes(mask.bits);
  }),
  /** PLGBLT (79) onto the parallelogram of three corners. */
  plgBlt: (corners: number[], w: number, h: number, src: { header: Uint8Array; bits: Uint8Array }) => emr(79, b => {
    const s0 = 140, s1 = s0 + src.header.length;
    b.i32(0, 0, 0, 0).i32(...corners).i32(0, 0, w, h).f32(1, 0, 0, 1, 0, 0).u32(0, 0);
    return b.u32(s0, src.header.length, s1, src.bits.length).i32(0, 0).u32(0, 0, 0, 0, 0).bytes(src.header).bytes(src.bits);
  }),
};

// ---------------------------------------------------------------------------------------------------------- WMF

/** A WMF record: size in words, function, parameters (16-bit words). */
export function wmr(fn: number, body: Bytes | ((b: Bytes) => Bytes) = new Bytes()): Uint8Array {
  const b = typeof body === 'function' ? body(new Bytes()) : body;
  const data = b.pad(2).done();
  return new Bytes().u32((6 + data.length) / 2).u16(fn).bytes(data).done();
}

/** A whole WMF, with a placeable header when `box` is given. */
export function wmf(records: Uint8Array[], box?: { l: number; t: number; r: number; b: number; inch: number }): Uint8Array {
  const body = [...records, wmr(0)];
  const words = 9 + body.reduce((n, r) => n + r.length / 2, 0);
  const out = new Bytes();
  if (box) out.u32(0x9ac6cdd7).u16(0).i16(box.l, box.t, box.r, box.b).u16(box.inch).u32(0).u16(0);
  out.u16(1, 9, 0x300).u32(words).u16(16).u32(64).u16(0);
  for (const r of body) out.bytes(r);
  return out.done();
}

export const W = {
  mapMode: (m: number) => wmr(0x0103, b => b.u16(m)),
  windowOrg: (x: number, y: number) => wmr(0x020b, b => b.i16(y, x)),
  windowExt: (x: number, y: number) => wmr(0x020c, b => b.i16(y, x)),
  viewportOrg: (x: number, y: number) => wmr(0x020d, b => b.i16(y, x)),
  viewportExt: (x: number, y: number) => wmr(0x020e, b => b.i16(y, x)),
  saveDC: () => wmr(0x001e),
  restoreDC: (n: number) => wmr(0x0127, b => b.i16(n)),
  createPen: (style: number, width: number, color: number) => wmr(0x02fa, b => b.u16(style).i16(width, 0).u32(color)),
  createBrush: (style: number, color: number, hatch = 0) => wmr(0x02fc, b => b.u16(style).u32(color).u16(hatch)),
  createFont: (height: number, face: string, o: { weight?: number; charSet?: number; escapement?: number } = {}) => wmr(0x02fb, b => {
    b.i16(height, 0, o.escapement ?? 0, o.escapement ?? 0, o.weight ?? 400).u8(0, 0, 0, o.charSet ?? 0, 0, 0, 0, 0);
    for (let k = 0; k < 32; k++) b.u8(k < face.length ? face.charCodeAt(k) : 0);
    return b;
  }),
  createPalette: (colors: number[]) => wmr(0x00f7, b => { b.u16(0x300, colors.length); for (const c of colors) b.u8(c & 255, (c >> 8) & 255, (c >> 16) & 255, 0); return b; }),
  select: (i: number) => wmr(0x012d, b => b.u16(i)),
  deleteObject: (i: number) => wmr(0x01f0, b => b.u16(i)),
  rectangle: (l: number, t: number, r: number, bb: number) => wmr(0x041b, b => b.i16(bb, r, t, l)),
  ellipse: (l: number, t: number, r: number, bb: number) => wmr(0x0418, b => b.i16(bb, r, t, l)),
  polygon: (pts: number[]) => wmr(0x0324, b => b.i16(pts.length / 2, ...pts)),
  polyline: (pts: number[]) => wmr(0x0325, b => b.i16(pts.length / 2, ...pts)),
  moveTo: (x: number, y: number) => wmr(0x0214, b => b.i16(y, x)),
  lineTo: (x: number, y: number) => wmr(0x0213, b => b.i16(y, x)),
  textOut: (x: number, y: number, s: string) => wmr(0x0521, b => b.i16(s.length).bytes(Array.from(s, c => c.charCodeAt(0))).pad(2).i16(y, x)),
  extTextOut: (x: number, y: number, bytes: number[], o: { options?: number; rect?: number[]; dx?: number[] } = {}) => wmr(0x0a32, b => {
    b.i16(y, x, bytes.length).u16(o.options ?? 0);
    if (o.rect) b.i16(...o.rect);
    b.bytes(bytes).pad(2);
    if (o.dx) b.i16(...o.dx);
    return b;
  }),
  textAlign: (a: number) => wmr(0x012e, b => b.u16(a)),
  textColor: (c: number) => wmr(0x0209, b => b.u32(c)),
  bkMode: (m: number) => wmr(0x0102, b => b.u16(m)),
  patBlt: (x: number, y: number, w: number, h: number, rop: number) => wmr(0x061d, b => b.u32(rop).i16(h, w, y, x)),
  dibBitBlt: (x: number, y: number, w: number, h: number, sx: number, sy: number, rop: number, packed: Uint8Array) =>
    wmr(0x0940, b => b.u32(rop).i16(sy, sx, h, w, y, x).bytes(packed)),
  intersectClip: (l: number, t: number, r: number, bb: number) => wmr(0x0416, b => b.i16(bb, r, t, l)),
  escape: (fn: number, data: Uint8Array) => wmr(0x0626, b => b.u16(fn, data.length).bytes(data)),
};

// ---------------------------------------------------------------------------------------------------------- DIB

export interface DibSpec {
  width: number;
  height: number;
  bitCount: number;
  /** Rows of pixel values (palette indexes, or 0xRRGGBB / 0xAARRGGBB), top row first. */
  rows: number[][];
  colors?: number[];
  topDown?: boolean;
  compression?: number;
  masks?: number[];
  /** Raw pixel bytes (for RLE), instead of `rows`. */
  bits?: number[];
  core?: boolean;
  headerSize?: number;
}

/** A DIB's header (with masks and colour table) and its bits, as separate byte arrays. */
export function dibParts(s: DibSpec): { header: Uint8Array; bits: Uint8Array } {
  const h = new Bytes();
  const colors = s.colors ?? [];
  if (s.core) {
    h.u32(12).u16(s.width, s.height, 1, s.bitCount);
    for (const c of colors) h.u8(c & 255, (c >> 8) & 255, (c >> 16) & 255);
  } else {
    const size = s.headerSize ?? 40;
    h.u32(size).i32(s.width, s.topDown ? -s.height : s.height).u16(1, s.bitCount).u32(s.compression ?? 0, 0, 2835, 2835, colors.length, 0);
    if (size >= 108) {
      h.u32(...(s.masks ?? [0, 0, 0, 0])).u32(0x73524742);
      for (let k = 0; k < 48; k++) h.u8(0);
      if (size >= 124) h.u32(4, 0, 0, 0);
    } else if (s.masks) {
      h.u32(...s.masks);
    }
    for (const c of colors) h.u8(c & 255, (c >> 8) & 255, (c >> 16) & 255, 0);
  }
  if (s.bits) return { header: h.done(), bits: Uint8Array.from(s.bits) };
  const stride = ((s.width * s.bitCount + 31) >>> 5) << 2;
  const bits = new Uint8Array(stride * s.height);
  s.rows.forEach((row, y) => {
    const r = (s.topDown ? y : s.height - 1 - y) * stride;
    row.forEach((v, x) => {
      switch (s.bitCount) {
        case 1: if (v) bits[r + (x >> 3)]! |= 0x80 >> (x & 7); break;
        case 4: bits[r + (x >> 1)]! |= x & 1 ? v & 15 : (v & 15) << 4; break;
        case 8: bits[r + x] = v; break;
        case 16: bits[r + x * 2] = v & 255; bits[r + x * 2 + 1] = (v >> 8) & 255; break;
        case 24: bits[r + x * 3] = v & 255; bits[r + x * 3 + 1] = (v >> 8) & 255; bits[r + x * 3 + 2] = (v >> 16) & 255; break;
        case 32: bits[r + x * 4] = v & 255; bits[r + x * 4 + 1] = (v >> 8) & 255; bits[r + x * 4 + 2] = (v >> 16) & 255; bits[r + x * 4 + 3] = (v >>> 24) & 255; break;
      }
    });
  });
  return { header: h.done(), bits };
}

/** A packed DIB: header, colour table and bits together, as WMF records carry them. */
export function packedDib(s: DibSpec): Uint8Array {
  const { header, bits } = dibParts(s);
  return new Bytes().bytes(header).bytes(bits).done();
}

// ---------------------------------------------------------------------------------------------------------- EMF+

/** An EMF+ record. */
export function plus(type: number, flags = 0, body: Bytes | ((b: Bytes) => Bytes) = new Bytes()): Uint8Array {
  const b = typeof body === 'function' ? body(new Bytes()) : body;
  const data = b.pad(4).done();
  return new Bytes().u16(type, flags).u32(12 + data.length, data.length).bytes(data).done();
}

/** EMF+ records wrapped in an EMR_COMMENT. */
export function plusComment(records: Uint8Array[]): Uint8Array {
  const b = new Bytes().u32(0x2b464d45);
  for (const r of records) b.bytes(r);
  return E.comment(b.done());
}

/** The EmfPlusHeader record: dual or EMF+ only, at `dpi`. */
export const plusHeader = (dual: boolean, dpi = 100) => plus(0x4001, dual ? 1 : 0, b => b.u32(0xdbc01002, 1, dpi, dpi));
