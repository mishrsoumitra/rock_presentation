// packages/metafile/src/plus.ts  -- EMF+ records ([MS-EMFPLUS]) for EMF+-only pictures
//
// The subset Office writes: objects (brushes, pens, paths, regions, images, fonts, string formats), fills and
// outlines of rectangles, polygons, lines, ellipses, arcs, pies, paths, curves and Beziers, images, strings,
// world and page transforms, saved states and containers, and clipping. Dual files are drawn from their EMF
// records instead; EMF+ is only played when nothing else describes the picture.
import { combinePath, combineRects, RGN_COPY, translateClip, type Clip } from './clip.ts';
import { decodeDib, packedDib, type Pixels } from './dib.ts';
import { MetafileError, stop, stops } from './errors.ts';
import { apply, arcCurves, axisAligned, compose, det, IDENTITY, Path, usable } from './geom.ts';
import type { DrawOptions, MetaCtx } from './player.ts';
import type { Dib, FontObject, Matrix, Metafile, PlusRecord } from './types.ts';
import { View } from './view.ts';
import { cssFont } from './text.ts';

/** What the EMF+ player draws through: the EMF player's context, output mapping, clip and bitmaps. */
export interface PlusHost {
  readonly ctx: MetaCtx;
  readonly d2o: Matrix;
  readonly pixel: number;
  readonly mf: Metafile;
  readonly opts: DrawOptions;
  spend(n: number): void;
  setClip(clip: Clip): void;
  image(px: Pixels): CanvasImageSource | undefined;
  /** Draw a metafile embedded in an image object, one level deep at most. */
  nested?(bytes: Uint8Array, corners: [number, number, number, number, number, number]): void;
}

type Brush =
  | { type: 'solid'; argb: number }
  | { type: 'hatch'; style: number; fore: number; back: number }
  | { type: 'linear'; rect: [number, number, number, number]; stops: [number, number][]; transform: Matrix }
  | { type: 'radial'; center: [number, number]; inner: number; outer: number; bounds: [number, number, number, number]; transform: Matrix }
  | { type: 'texture'; image: Image | undefined; transform: Matrix };

interface Pen { width: number; unit: number; brush: Brush; style: number; dashes: number[] | undefined; cap: number; join: number; miterLimit: number; }
interface PlusPath { points: Float64Array; types: Uint8Array; }
type Region =
  | { type: 'rect'; rect: [number, number, number, number] }
  | { type: 'path'; path: PlusPath }
  | { type: 'empty' }
  | { type: 'infinite' }
  | { type: 'op'; op: number; left?: Region; right?: Region };
/** A bitmap holds its pixels, a DIB (a compressed BMP), or compressed bytes the host decodes (by `key`). */
type Image =
  | { kind: 'bitmap'; key: string; pixels?: Pixels; dib?: Dib; encoded?: Uint8Array }
  | { kind: 'metafile'; bytes: Uint8Array };
interface Font { emSize: number; unit: number; style: number; family: string; }
interface Format { flags: number; align: number; lineAlign: number; }

type PlusObject =
  | { kind: 'brush'; brush: Brush }
  | { kind: 'pen'; pen: Pen }
  | { kind: 'path'; path: PlusPath }
  | { kind: 'region'; region: Region }
  | { kind: 'image'; image: Image }
  | { kind: 'font'; font: Font }
  | { kind: 'format'; format: Format }
  | { kind: 'other' };

interface Saved { world: Matrix; unit: number; scale: number; clip: Clip; }

export interface PlusState {
  objects: Map<number, PlusObject>;
  /** An object arriving in pieces (the continue flag): its slot, type, declared size and the pieces so far. */
  pending: { id: number; type: number; total: number; key: string; parts: Uint8Array[]; size: number } | undefined;
  world: Matrix;
  unit: number;
  scale: number;
  dpiX: number;
  dpiY: number;
  video: boolean;
  clip: Clip;
  saved: Map<number, Saved>;
}

export function plusState(mf: Metafile): PlusState {
  const dpi = mf.dpmmX * 25.4;
  return { objects: new Map(), pending: undefined, world: IDENTITY, unit: 1, scale: 1, dpiX: dpi, dpiY: mf.dpmmY * 25.4, video: true, clip: [], saved: new Map() };
}

/** Most bytes one object may gather from continued records. */
const MAX_OBJECT_BYTES = 64 * 1024 * 1024;
/** Most points one path, and most nodes one region, may have. */
const MAX_POINTS = 1_000_000;
const MAX_NODES = 10_000;

const OBJECT = 0x4008;

// ---------------------------------------------------------------------------------------------------------- objects

/**
 * Feed an EmfPlusObject record; returns the complete object data (one record's, or the pieces of a continued
 * object joined) with a key that names it the same way on every pass, or undefined while pieces are missing.
 */
export function assembleObject(state: Pick<PlusState, 'pending'>, r: PlusRecord): { id: number; type: number; data: View; key: string } | undefined {
  const id = r.flags & 0xff;
  const type = (r.flags >> 8) & 0x7f;
  const v = r.data;
  if (r.flags & 0x8000) {
    // continued: the object's total size, then this piece (DataSize covers both, as GDI+ writes them)
    if (!v.has(0, 4)) return undefined;
    const total = v.u32(0);
    const piece = v.sub(4, v.size - 4);
    let p = state.pending;
    if (!p || p.id !== id || p.type !== type) {
      if (total > MAX_OBJECT_BYTES) throw new MetafileError('malformed', `EMF+ object of ${total} bytes`);
      p = state.pending = { id, type, total, key: `plus:${v.start}`, parts: [], size: 0 };
    }
    p.parts.push(piece);
    p.size += piece.length;
    if (p.size < p.total) return undefined;
    state.pending = undefined;
    const joined = new Uint8Array(p.size);
    let at = 0;
    for (const part of p.parts) { joined.set(part, at); at += part.length; }
    return { id, type, data: new View(joined), key: p.key };
  }
  const p = state.pending;
  if (p && p.id === id && p.type === type) {
    // the last piece of a continued object comes without the continue flag, with or without the total size
    state.pending = undefined;
    const skip = v.size === p.total - p.size + 4 && v.has(0, 4) && v.u32(0) === p.total ? 4 : 0;
    const last = v.sub(skip, v.size - skip);
    const joined = new Uint8Array(p.size + last.length);
    let at = 0;
    for (const part of p.parts) { joined.set(part, at); at += part.length; }
    joined.set(last, at);
    return { id, type, data: new View(joined), key: p.key };
  }
  return { id, type, data: v, key: `plus:${v.start}` };
}

const argb = (v: View, at: number) => v.u32(at);

function matrixAt(v: View, at: number): Matrix {
  const m: Matrix = [v.f32(at), v.f32(at + 4), v.f32(at + 8), v.f32(at + 12), v.f32(at + 16), v.f32(at + 20)];
  return m.every(Number.isFinite) ? m : IDENTITY;
}

/** EmfPlusInteger7 or EmfPlusInteger15 at `at`: its value and size. */
function intR(v: View, at: number): [number, number] {
  const b = v.u8(at);
  if (b & 0x80) {
    const n = ((b & 0x7f) << 8) | v.u8(at + 1);
    return [n & 0x4000 ? n - 0x8000 : n, 2];
  }
  return [b & 0x40 ? b - 0x80 : b, 1];
}

/** `count` points at `at`: relative (P flag), 16-bit (C flag) or float; and where they end. */
function readPoints(v: View, at: number, count: number, compressed: boolean, relative: boolean): [Float64Array, number] {
  if (count > MAX_POINTS || count > v.size) throw new MetafileError('malformed', `${count} EMF+ points`);
  const out = new Float64Array(count * 2);
  let p = at;
  if (relative) {
    let x = 0, y = 0;
    for (let i = 0; i < count; i++) {
      const [dx, a] = intR(v, p); p += a;
      const [dy, b] = intR(v, p); p += b;
      x += dx; y += dy;
      out[i * 2] = x; out[i * 2 + 1] = y;
    }
  } else if (compressed) {
    for (let i = 0; i < count * 2; i++) out[i] = v.i16(at + i * 2);
    p = at + count * 4;
  } else {
    for (let i = 0; i < count * 2; i++) out[i] = v.f32(at + i * 4);
    p = at + count * 8;
  }
  return [out, p];
}

function readRect(v: View, at: number, compressed: boolean): [number, number, number, number] {
  return compressed ? [v.i16(at), v.i16(at + 2), v.i16(at + 4), v.i16(at + 6)] : [v.f32(at), v.f32(at + 4), v.f32(at + 8), v.f32(at + 12)];
}

function readPath(v: View, at: number): [PlusPath, number] {
  const count = v.u32(at + 4);
  const flags = v.u32(at + 8);
  const relative = (flags & 0x800) !== 0;
  const [points, end] = readPoints(v, at + 12, count, (flags & 0x4000) !== 0, relative);
  const types = new Uint8Array(count);
  let p = end;
  if (relative) {
    // run-length encoded types: a 16-bit run of (count, type) pairs
    for (let i = 0; i < count;) {
      const w = v.u16(p); p += 2;
      const run = Math.max(1, (w >> 8) & 0x3f);
      for (let k = 0; k < run && i < count; k++) types[i++] = w & 0xff;
    }
  } else {
    types.set(v.sub(p, count));
    p += count;
  }
  return [{ points, types }, p];
}

function readBrush(v: View, at: number, key: string): Brush {
  const type = v.u32(at + 4);
  const p = at + 8;
  switch (type) {
    case 0: return { type: 'solid', argb: argb(v, p) };
    case 1: return { type: 'hatch', style: v.u32(p), fore: argb(v, p + 4), back: argb(v, p + 8) };
    case 4: {
      const flags = v.u32(p);
      const rect: [number, number, number, number] = [v.f32(p + 8), v.f32(p + 12), v.f32(p + 16), v.f32(p + 20)];
      const start = argb(v, p + 24), end = argb(v, p + 28);
      let q = p + 40;
      let transform = IDENTITY;
      if (flags & 0x2) { transform = matrixAt(v, q); q += 24; }
      let stops: [number, number][] = [[0, start], [1, end]];
      if (flags & 0x4) {
        const n = v.u32(q);
        if (n > 0 && n <= 1024) stops = Array.from({ length: n }, (_, i) => [v.f32(q + 4 + i * 4), argb(v, q + 4 + n * 4 + i * 4)] as [number, number]);
      } else if (flags & 0x8) {
        const n = v.u32(q);
        if (n > 0 && n <= 1024) {
          stops = Array.from({ length: n }, (_, i) => [v.f32(q + 4 + i * 4), mix(start, end, v.f32(q + 4 + n * 4 + i * 4))] as [number, number]);
        }
      }
      return { type: 'linear', rect, stops, transform };
    }
    case 3: {
      const flags = v.u32(p);
      const inner = argb(v, p + 8);
      const center: [number, number] = [v.f32(p + 12), v.f32(p + 16)];
      const n = v.u32(p + 20);
      if (n > 4096 || !v.has(p + 24, n * 4)) throw new MetafileError('malformed', `${n} surrounding colours`);
      const outer = n > 0 ? argb(v, p + 24) : inner;
      let q = p + 24 + n * 4;
      let bounds: [number, number, number, number] = [center[0], center[1], 0, 0];
      // the boundary: a path or a list of points; its bounding box sets the gradient's extent
      let pts: Float64Array | undefined;
      if (flags & 1) { const size = v.u32(q); pts = readPath(v, q + 4)[0].points; q += 4 + size; }
      else { const c = v.u32(q); [pts] = readPoints(v, q + 4, c, false, false); q += 4 + c * 8; }
      if (pts.length >= 2) {
        let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
        for (let i = 0; i + 1 < pts.length; i += 2) {
          l = Math.min(l, pts[i]!); r = Math.max(r, pts[i]!); t = Math.min(t, pts[i + 1]!); b = Math.max(b, pts[i + 1]!);
        }
        bounds = [l, t, r - l, b - t];
      }
      const transform = flags & 2 && v.has(q, 24) ? matrixAt(v, q) : IDENTITY;
      return { type: 'radial', center, inner, outer, bounds, transform };
    }
    case 2: {
      const flags = v.u32(p);
      let q = p + 8;
      let transform = IDENTITY;
      if (flags & 0x2) { transform = matrixAt(v, q); q += 24; }
      const image = v.has(q, 8) ? readImage(v, q, `${key}+texture`) : undefined;
      return { type: 'texture', image, transform };
    }
    default: return { type: 'solid', argb: 0xff808080 };
  }
}

function mix(a: number, b: number, f: number): number {
  // BlendFactors: 0 is the start colour, 1 the end colour
  const t = Math.min(1, Math.max(0, f));
  let out = 0;
  for (const s of [24, 16, 8, 0]) out |= Math.round(((a >>> s) & 255) * (1 - t) + ((b >>> s) & 255) * t) << s;
  return out >>> 0;
}

function readPen(v: View, at: number, key: string): Pen {
  const flags = v.u32(at + 8);
  const unit = v.u32(at + 12);
  const width = v.f32(at + 16);
  let p = at + 20;
  const pen: Pen = {
    width: Number.isFinite(width) ? Math.abs(width) : 1, unit, brush: { type: 'solid', argb: 0xff000000 }, style: 0, dashes: undefined,
    cap: 0, join: 0, miterLimit: 10,
  };
  if (flags & 0x1) p += 24;
  if (flags & 0x2) { pen.cap = v.u32(p); p += 4; }
  if (flags & 0x4) { if (!(flags & 0x2)) pen.cap = v.u32(p); p += 4; }
  if (flags & 0x8) { pen.join = v.u32(p); p += 4; }
  if (flags & 0x10) { pen.miterLimit = v.f32(p); p += 4; }
  if (flags & 0x20) { pen.style = v.u32(p); p += 4; }
  if (flags & 0x40) p += 4;
  if (flags & 0x80) p += 4;
  if (flags & 0x100) {
    const n = v.u32(p);
    if (n > 64) throw new MetafileError('malformed', `${n} dashes`);
    pen.dashes = Array.from({ length: n }, (_, i) => v.f32(p + 4 + i * 4));
    p += 4 + n * 4;
  }
  if (flags & 0x200) p += 4;
  if (flags & 0x400) { const n = v.u32(p); if (n > 64) throw new MetafileError('malformed', `${n} compound lines`); p += 4 + n * 4; }
  if (flags & 0x800) p += 4 + v.u32(p);
  if (flags & 0x1000) p += 4 + v.u32(p);
  if (v.has(p, 8)) pen.brush = readBrush(v, p, `${key}+brush`);
  return pen;
}

function readRegion(v: View, at: number): Region {
  // the nodes are stored in pre-order; an explicit stack holds the operations waiting for their operands
  const count = v.u32(at + 4);
  if (count > MAX_NODES) throw new MetafileError('limit', `${count} region nodes`);
  let p = at + 8;
  const open: Extract<Region, { type: 'op' }>[] = [];
  let root: Region | undefined;
  for (let n = 0; n <= count; n++) {
    const type = v.u32(p);
    p += 4;
    let node: Region;
    if (type === 0x10000000) { node = { type: 'rect', rect: [v.f32(p), v.f32(p + 4), v.f32(p + 8), v.f32(p + 12)] }; p += 16; }
    else if (type === 0x10000001) { const size = v.u32(p); node = { type: 'path', path: readPath(v, p + 4)[0] }; p += 4 + size; }
    else if (type === 0x10000002) node = { type: 'empty' };
    else if (type === 0x10000003) node = { type: 'infinite' };
    else if (type >= 1 && type <= 5) node = { type: 'op', op: type };
    else throw new MetafileError('malformed', `region node type ${type}`);
    const parent = open.at(-1);
    if (parent) { if (!parent.left) parent.left = node; else parent.right = node; }
    else if (!root) root = node;
    else throw new MetafileError('malformed', 'region nodes after the tree');
    if (node.type === 'op') open.push(node);
    while (open.length && open.at(-1)!.right) open.pop();
    if (!open.length && root) break;
  }
  if (!root || open.length) throw new MetafileError('malformed', 'incomplete region');
  return root;
}

function readImage(v: View, at: number, key: string): Image {
  const type = v.u32(at + 4);
  const p = at + 8;
  if (type === 2) {
    const size = v.u32(p + 4);
    return { kind: 'metafile', bytes: v.sub(p + 8, Math.min(size, v.size - p - 8)) };
  }
  if (type !== 1) throw new MetafileError('unsupported', `EMF+ image type ${type}`);
  const width = v.i32(p), height = v.i32(p + 4), stride = v.i32(p + 8), format = v.u32(p + 12), dataType = v.u32(p + 16);
  const q = p + 20;
  if (dataType === 1) {
    const encoded = v.sub(q, v.size - q);
    // a BMP needs no host decoder
    if (encoded[0] === 0x42 && encoded[1] === 0x4d && encoded.length > 14) return { kind: 'bitmap', key, dib: packedDib(new View(encoded, 14), 0, 0) };
    return { kind: 'bitmap', key, encoded };
  }
  return { kind: 'bitmap', key, pixels: decodeBitmap(v, q, width, height, stride, format) };
}

/** Uncompressed EMF+ bitmap pixels ([MS-EMFPLUS] 2.2.2.2): indexed, 16, 24, 32 and 64 bits, top row first. */
function decodeBitmap(v: View, at: number, width: number, height: number, stride: number, format: number): Pixels {
  if (!(width > 0 && height > 0) || width > 16_384 || height > 16_384 || width * height > 64 * 1024 * 1024) {
    throw new MetafileError(width > 0 && height > 0 ? 'limit' : 'malformed', `EMF+ bitmap of ${width}x${height}`);
  }
  const bpp = (format >> 8) & 0xff;
  let p = at;
  let palette: number[] = [];
  if (format & 0x10000) {
    const n = v.u32(p + 4);
    if (n > 256) throw new MetafileError('malformed', `${n} palette entries`);
    palette = Array.from({ length: n }, (_, i) => v.u32(p + 8 + i * 4));
    p += 8 + n * 4;
  }
  const rowBytes = Math.abs(stride) || ((width * bpp + 31) >> 5) * 4;
  const bytes = v.sub(p, Math.max(0, v.size - p));
  if (bytes.length * 2 < rowBytes * height) throw new MetafileError('malformed', `${bytes.length} bytes of pixels for a ${width}x${height} bitmap`);
  const data = new Uint8ClampedArray(width * height * 4);
  const at8 = (i: number) => (i < bytes.length ? bytes[i]! : 0);
  const put = (o: number, c: number, premultiplied = false) => {
    const a = c >>> 24;
    const k = premultiplied && a > 0 && a < 255 ? 255 / a : 1;
    data[o] = ((c >>> 16) & 255) * k; data[o + 1] = ((c >>> 8) & 255) * k; data[o + 2] = (c & 255) * k; data[o + 3] = a;
  };
  for (let y = 0; y < height; y++) {
    const row = (stride < 0 ? height - 1 - y : y) * rowBytes;           // a negative stride: bottom row first
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4;
      switch (format) {
        case 0x30101: put(o, palette[(at8(row + (x >> 3)) >> (7 - (x & 7))) & 1] ?? 0xff000000); break;
        case 0x30402: put(o, palette[(at8(row + (x >> 1)) >> (x & 1 ? 0 : 4)) & 15] ?? 0xff000000); break;
        case 0x30803: put(o, palette[at8(row + x)] ?? 0xff000000); break;
        case 0x21005: case 0x61007: {
          const w = at8(row + x * 2) | (at8(row + x * 2 + 1) << 8);
          const a = format === 0x61007 && !(w & 0x8000) ? 0 : 255;
          put(o, ((a << 24) | (((w >> 10) & 31) * 255 / 31) << 16 | (((w >> 5) & 31) * 255 / 31) << 8 | ((w & 31) * 255 / 31)) >>> 0);
          break;
        }
        case 0x21006: {
          const w = at8(row + x * 2) | (at8(row + x * 2 + 1) << 8);
          const r = Math.round(((w >> 11) & 31) * 255 / 31), g = Math.round(((w >> 5) & 63) * 255 / 63), b = Math.round((w & 31) * 255 / 31);
          put(o, (0xff000000 | (r << 16) | (g << 8) | b) >>> 0);
          break;
        }
        case 0x101004: { const g = at8(row + x * 2 + 1); put(o, (0xff000000 | (g << 16) | (g << 8) | g) >>> 0); break; }
        case 0x21808: put(o, (0xff000000 | (at8(row + x * 3 + 2) << 16) | (at8(row + x * 3 + 1) << 8) | at8(row + x * 3)) >>> 0); break;
        case 0x22009: put(o, (0xff000000 | (at8(row + x * 4 + 2) << 16) | (at8(row + x * 4 + 1) << 8) | at8(row + x * 4)) >>> 0); break;
        case 0x26200a: case 0xe200b:
          put(o, ((at8(row + x * 4 + 3) << 24) | (at8(row + x * 4 + 2) << 16) | (at8(row + x * 4 + 1) << 8) | at8(row + x * 4)) >>> 0, format === 0xe200b);
          break;
        case 0x10300c: put(o, (0xff000000 | (at8(row + x * 6 + 5) << 16) | (at8(row + x * 6 + 3) << 8) | at8(row + x * 6 + 1)) >>> 0); break;
        case 0x34400d: case 0x1a400e:
          put(o, ((at8(row + x * 8 + 7) << 24) | (at8(row + x * 8 + 5) << 16) | (at8(row + x * 8 + 3) << 8) | at8(row + x * 8 + 1)) >>> 0, format === 0x1a400e);
          break;
        default: throw new MetafileError('unsupported', `EMF+ pixel format ${format.toString(16)}`);
      }
    }
  }
  return { width, height, data };
}

function readObject(type: number, v: View, key: string): PlusObject {
  switch (type) {
    case 1: return { kind: 'brush', brush: readBrush(v, 0, key) };
    case 2: return { kind: 'pen', pen: readPen(v, 0, key) };
    case 3: return { kind: 'path', path: readPath(v, 0)[0] };
    case 4: return { kind: 'region', region: readRegion(v, 0) };
    case 5: return { kind: 'image', image: readImage(v, 0, key) };
    case 6: {
      const n = v.u32(20);
      if (n > 256) throw new MetafileError('malformed', 'font name too long');
      let family = '';
      for (let i = 0; i < n; i++) family += String.fromCharCode(v.u16(24 + i * 2));
      return { kind: 'font', font: { emSize: v.f32(4), unit: v.u32(8), style: v.u32(12), family: family.replace(/\0.*$/, '') } };
    }
    case 7: return { kind: 'format', format: { flags: v.u32(4), align: v.u32(12), lineAlign: v.u32(16) } };
    default: return { kind: 'other' };
  }
}

/** The compressed (JPEG, PNG, GIF, TIFF) images among the EMF+ records, for decoding before playback. */
export function plusImages(mf: Metafile): { key: string; bytes: Uint8Array }[] {
  const out: { key: string; bytes: Uint8Array }[] = [];
  const state: Pick<PlusState, 'pending'> = { pending: undefined };
  for (const rec of mf.records) {
    if (rec.type !== 'plus') continue;
    for (const r of rec.records) {
      if (r.type !== OBJECT) continue;
      try {
        const o = assembleObject(state, r);
        if (!o || (o.type !== 1 && o.type !== 5)) continue;
        const obj = readObject(o.type, o.data, o.key);
        const img = obj.kind === 'image' ? obj.image : obj.kind === 'brush' && obj.brush.type === 'texture' ? obj.brush.image : undefined;
        if (img?.kind === 'bitmap' && img.encoded) out.push({ key: img.key, bytes: img.encoded });
      } catch (e) {
        if (!(e instanceof MetafileError)) throw e;
      }
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------------------------- playback

const css = (c: number) => `rgba(${(c >>> 16) & 255},${(c >>> 8) & 255},${c & 255},${+((c >>> 24) / 255).toFixed(4)})`;

/** Device pixels per unit ([MS-EMFPLUS] 2.1.1.32). */
function unitScale(s: PlusState, unit: number, y = false): number {
  const dpi = y ? s.dpiY : s.dpiX;
  switch (unit) {
    case 1: return s.video ? 1 : dpi / 100;
    case 3: return dpi / 72;
    case 4: return dpi;
    case 5: return dpi / 300;
    case 6: return dpi / 25.4;
    default: return 1;
  }
}

/** World units to output units. */
function toOutput(host: PlusHost, s: PlusState): Matrix {
  const page: Matrix = [s.scale * unitScale(s, s.unit), 0, 0, s.scale * unitScale(s, s.unit, true), 0, 0];
  const m = compose(host.d2o, compose(page, s.world));
  return usable(m) ? m : host.d2o;
}

/** A path object's figures in output units. */
function plusPath(p: PlusPath, m: Matrix): Path {
  const out = new Path();
  const n = p.types.length;
  for (let i = 0; i < n; i++) {
    const t = p.types[i]!;
    const [x, y] = apply(m, p.points[i * 2]!, p.points[i * 2 + 1]!);
    const kind = t & 0x07;
    if (kind === 0 || i === 0) out.moveTo(x, y);
    else if (kind === 3 && i + 2 < n) {
      const [x2, y2] = apply(m, p.points[i * 2 + 2]!, p.points[i * 2 + 3]!);
      const [x3, y3] = apply(m, p.points[i * 2 + 4]!, p.points[i * 2 + 5]!);
      out.curveTo(x, y, x2, y2, x3, y3);
      i += 2;
      if (p.types[i]! & 0x80) out.close();
      continue;
    } else out.lineTo(x, y);
    if (t & 0x80) out.close();
  }
  return out;
}

function rectPath(m: Matrix, r: [number, number, number, number]): Path {
  const [x, y, w, h] = r;
  const p = new Path();
  p.polygon([...apply(m, x, y), ...apply(m, x + w, y), ...apply(m, x + w, y + h), ...apply(m, x, y + h)]);
  return p;
}

/** An ellipse arc in the box, in degrees clockwise from the x axis (EMF+'s convention), as a path. */
function arcPath(m: Matrix, r: [number, number, number, number], start: number, sweep: number, pie: boolean, closed: boolean): Path {
  const [x, y, w, h] = r;
  const cx = x + w / 2, cy = y + h / 2, rx = w / 2, ry = h / 2;
  const s = Math.max(-360, Math.min(360, sweep));
  // the angles are of the circle the ellipse is stretched from
  const t0 = Math.atan2(Math.sin((start * Math.PI) / 180) * (rx || 1), Math.cos((start * Math.PI) / 180) * (ry || 1));
  const t1 = Math.atan2(Math.sin(((start + s) * Math.PI) / 180) * (rx || 1), Math.cos(((start + s) * Math.PI) / 180) * (ry || 1));
  let d = t1 - t0;
  if (s > 0) { while (d <= 0) d += Math.PI * 2; } else if (s < 0) { while (d >= 0) d -= Math.PI * 2; }
  if (Math.abs(s) >= 360) d = Math.sign(s) * Math.PI * 2;
  const p = new Path();
  if (pie) p.moveTo(...apply(m, cx, cy));
  const start0 = apply(m, cx + rx * Math.cos(t0), cy + ry * Math.sin(t0));
  if (pie) p.lineTo(...start0); else p.moveTo(...start0);
  arcCurves(cx, cy, rx, ry, t0, d, (a, b, c, dd, e, f) => p.curveTo(...apply(m, a, b), ...apply(m, c, dd), ...apply(m, e, f)));
  if (closed) p.close();
  return p;
}

/** A cardinal spline through the points as Beziers ([MS-EMFPLUS] DrawCurve); closed joins the last to the first. */
function curvePath(m: Matrix, pts: Float64Array, tension: number, closed: boolean, from = 0, segments = Infinity): Path {
  const n = pts.length >> 1;
  const p = new Path();
  if (n < 2) return p;
  const t = (Number.isFinite(tension) ? tension : 0.5) / 3;
  const at = (i: number): [number, number] => {
    const k = closed ? ((i % n) + n) % n : Math.max(0, Math.min(n - 1, i));
    return [pts[k * 2]!, pts[k * 2 + 1]!];
  };
  const last = closed ? n : Math.min(n - 1, from + segments);
  p.moveTo(...apply(m, ...at(from)));
  for (let i = from; i < last; i++) {
    const [x0, y0] = at(i - 1), [x1, y1] = at(i), [x2, y2] = at(i + 1), [x3, y3] = at(i + 2);
    p.curveTo(...apply(m, x1 + (x2 - x0) * t, y1 + (y2 - y0) * t), ...apply(m, x2 - (x3 - x1) * t, y2 - (y3 - y1) * t), ...apply(m, x2, y2));
  }
  if (closed) p.close();
  return p;
}

// The GDI+ hatch styles that matter, as 8x8 bits; the rest are drawn as a blend of their two colours.
const HATCH: Record<number, number[]> = {
  0: [0, 0, 0, 0, 0, 0, 0, 0xff], 1: [0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
  2: [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01], 3: [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80],
  4: [0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0xff], 5: [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81],
};
const HATCH_COVER: Record<number, number> = { 6: 0.05, 7: 0.1, 8: 0.2, 9: 0.25, 10: 0.3, 11: 0.4, 12: 0.5, 13: 0.6, 14: 0.7, 15: 0.75, 16: 0.8, 17: 0.9 };

class PlusPlayer {
  readonly host: PlusHost;
  readonly s: PlusState;
  private m: Matrix;

  constructor(host: PlusHost, s: PlusState) {
    this.host = host;
    this.s = s;
    this.m = toOutput(host, s);
  }

  private refresh(): void { this.m = toOutput(this.host, this.s); }

  private object(id: number): PlusObject | undefined { return this.s.objects.get(id & 0xff); }

  private brushFor(flags: number, id: number): Brush | undefined {
    if (flags & 0x8000) return { type: 'solid', argb: id >>> 0 };
    const o = this.object(id);
    return o?.kind === 'brush' ? o.brush : undefined;
  }

  /** A brush as a canvas fill style, in output coordinates. */
  private paint(b: Brush): string | CanvasGradient | CanvasPattern | undefined {
    const ctx = this.host.ctx;
    switch (b.type) {
      case 'solid': return b.argb >>> 24 ? css(b.argb) : undefined;
      case 'hatch': {
        const bits = HATCH[b.style];
        if (!bits) return css(mix(b.back, b.fore, HATCH_COVER[b.style] ?? 0.5));
        const data = new Uint8ClampedArray(256);
        for (let y = 0; y < 8; y++) {
          for (let x = 0; x < 8; x++) {
            const c = (bits[y]! >> (7 - x)) & 1 ? b.fore : b.back;
            data.set([(c >>> 16) & 255, (c >>> 8) & 255, c & 255, c >>> 24], (y * 8 + x) * 4);
          }
        }
        const img = this.host.image({ width: 8, height: 8, data });
        const p = img ? ctx.createPattern(img, 'repeat') : null;
        if (!p) return css(mix(b.back, b.fore, 0.5));
        p.setTransform?.({ a: this.host.pixel, b: 0, c: 0, d: this.host.pixel, e: this.host.d2o[4], f: this.host.d2o[5] });
        return p;
      }
      case 'linear': {
        const m = compose(this.m, b.transform);
        const [x, y, w] = b.rect;
        const [x0, y0] = apply(m, x, y);
        const [x1, y1] = apply(m, x + w, y);
        if (!(Number.isFinite(x0 + y0 + x1 + y1)) || (x0 === x1 && y0 === y1)) return css(b.stops[0]?.[1] ?? 0xff000000);
        const g = ctx.createLinearGradient(x0, y0, x1, y1);
        for (const [pos, c] of b.stops) if (Number.isFinite(pos)) g.addColorStop(Math.min(1, Math.max(0, pos)), css(c));
        return g;
      }
      case 'radial': {
        const m = compose(this.m, b.transform);
        const [cx, cy] = apply(m, b.center[0], b.center[1]);
        const [bx, by, bw, bh] = b.bounds;
        const corners = [apply(m, bx, by), apply(m, bx + bw, by), apply(m, bx, by + bh), apply(m, bx + bw, by + bh)];
        const r = Math.max(...corners.map(([x, y]) => Math.hypot(x - cx, y - cy)));
        if (!(r > 0) || !Number.isFinite(r)) return css(b.inner);
        const g = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
        g.addColorStop(0, css(b.inner));
        g.addColorStop(1, css(b.outer));
        return g;
      }
      case 'texture': {
        const img = b.image ? this.imageSource(b.image) : undefined;
        if (!img) return 'rgba(128,128,128,1)';
        const p = ctx.createPattern(img.source, 'repeat');
        if (!p) return 'rgba(128,128,128,1)';
        const m = compose(this.m, b.transform);
        p.setTransform?.({ a: m[0], b: m[1], c: m[2], d: m[3], e: m[4], f: m[5] });
        return p;
      }
    }
  }

  private fill(path: Path, b: Brush | undefined, rule: CanvasFillRule = 'nonzero'): void {
    if (!b) return;
    const style = this.paint(b);
    if (!style) return;
    const ctx = this.host.ctx;
    this.host.spend(path.size);
    this.host.setClip(this.s.clip);
    ctx.globalCompositeOperation = 'source-over';
    ctx.fillStyle = style;
    ctx.beginPath();
    path.trace(ctx);
    ctx.fill(rule);
  }

  private stroke(path: Path, penId: number): void {
    const o = this.object(penId);
    if (o?.kind !== 'pen') return;
    const pen = o.pen;
    const style = this.paint(pen.brush);
    if (!style) return;
    const host = this.host, ctx = host.ctx;
    // a world-unit pen goes through the world and page transforms like the coordinates; a pixel pen is a device
    // width; other units are physical sizes, scaled by the world transform and the page scale (as GDI+ draws them)
    const device = Math.hypot(host.d2o[0], host.d2o[1]);
    const w = pen.unit === 0 ? pen.width * Math.sqrt(Math.abs(det(this.m)))
      : pen.unit === 2 ? pen.width * device
      : pen.width * unitScale(this.s, pen.unit) * this.s.scale * Math.sqrt(Math.abs(det(this.s.world))) * device;
    const width = Math.max(w, host.pixel);
    host.spend(path.size);
    host.setClip(this.s.clip);
    ctx.globalCompositeOperation = 'source-over';
    ctx.strokeStyle = style;
    ctx.lineWidth = width;
    ctx.lineCap = pen.cap === 1 ? 'square' : pen.cap === 2 ? 'round' : 'butt';
    ctx.lineJoin = pen.join === 1 ? 'bevel' : pen.join === 2 ? 'round' : 'miter';
    ctx.miterLimit = Math.max(1, Number.isFinite(pen.miterLimit) ? pen.miterLimit : 10);
    const dashes = pen.style === 1 ? [3, 1] : pen.style === 2 ? [1, 1] : pen.style === 3 ? [3, 1, 1, 1] : pen.style === 4 ? [3, 1, 1, 1, 1, 1]
      : pen.style === 5 ? pen.dashes ?? [] : [];
    ctx.setLineDash(dashes.every(d => Number.isFinite(d) && d >= 0) && dashes.some(d => d > 0) ? dashes.map(d => d * width) : []);
    ctx.beginPath();
    path.trace(ctx);
    ctx.stroke();
    ctx.setLineDash([]);
  }

  private imageSource(img: Image): { source: CanvasImageSource; width: number; height: number } | undefined {
    if (img.kind !== 'bitmap') return undefined;
    let px = img.pixels;
    if (!px && img.dib) {
      this.host.spend(img.dib.width * img.dib.height);
      px = decodeDib(img.dib, {});
    }
    if (px) {
      const source = this.host.image(px);
      return source ? { source, width: px.width, height: px.height } : undefined;
    }
    const decoded = this.host.opts.images?.get(img.key);
    if (!decoded) return undefined;
    const { width, height } = decoded as { width: number; height: number };
    return { source: decoded, width, height };
  }

  private drawImage(id: number, src: [number, number, number, number], corners: [number, number, number, number, number, number]): void {
    const o = this.object(id);
    if (o?.kind !== 'image') return;
    const host = this.host, ctx = host.ctx;
    const m = this.m;
    const [x0, y0] = apply(m, corners[0], corners[1]);
    const [x1, y1] = apply(m, corners[2], corners[3]);
    const [x2, y2] = apply(m, corners[4], corners[5]);
    if (o.image.kind === 'metafile') {
      host.setClip(this.s.clip);
      host.nested?.(o.image.bytes, [x0, y0, x1, y1, x2, y2]);
      return;
    }
    const img = this.imageSource(o.image);
    if (!img) return;
    let [sx, sy, sw, sh] = src;
    if (!(sw > 0 && sh > 0) || !Number.isFinite(sx + sy + sw + sh)) { sx = 0; sy = 0; sw = img.width; sh = img.height; }
    host.setClip(this.s.clip);
    ctx.save();
    try {
      ctx.globalCompositeOperation = 'source-over';
      ctx.imageSmoothingEnabled = true;
      ctx.transform((x1 - x0) / sw, (y1 - y0) / sw, (x2 - x0) / sh, (y2 - y0) / sh, x0, y0);
      ctx.drawImage(img.source, sx, sy, sw, sh, 0, 0, sw, sh);
    } finally {
      ctx.restore();
    }
  }

  /** A font object as a LOGFONT-like description, and its em size in world units. */
  private font(id: number): { f: FontObject; em: number; underline: boolean; strike: boolean } | undefined {
    const o = this.object(id);
    if (o?.kind !== 'font') return undefined;
    const ft = o.font;
    // a size in physical units is converted to page units; world units stay as they are
    const em = ft.unit === 0 ? ft.emSize : (ft.emSize * unitScale(this.s, ft.unit, true)) / (this.s.scale * unitScale(this.s, this.s.unit, true));
    if (!(em > 0) || !Number.isFinite(em)) return undefined;
    const f: FontObject = {
      kind: 'font', height: -em, width: 0, escapement: 0, orientation: 0, weight: ft.style & 1 ? 700 : 400, italic: (ft.style & 2) !== 0,
      underline: false, strikeOut: false, charSet: 1, quality: 0, pitchAndFamily: 0, face: ft.family,
    };
    return { f, em, underline: (ft.style & 4) !== 0, strike: (ft.style & 8) !== 0 };
  }

  /**
   * Set the context up for text in world space: returns the layout scale (world units per layout unit) and
   * the font size in layout units, with the transform applied (the caller restores).
   */
  private textFrame(em: number, f: FontObject): { unit: number; size: number } | undefined {
    const host = this.host, ctx = host.ctx;
    const m = this.m;
    const gy = Math.hypot(m[2], m[3]);
    const out = em * gy;
    if (!(out > 0) || !Number.isFinite(out)) return undefined;
    const size = Math.min(2000, Math.max(1, out / host.pixel));
    const unit = em / size;
    ctx.transform(m[0] * unit, m[1] * unit, m[2] * unit, m[3] * unit, m[4], m[5]);
    ctx.font = cssFont(f, size);
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    ctx.globalCompositeOperation = 'source-over';
    return { unit, size };
  }

  private drawString(flags: number, v: View): void {
    const brush = this.brushFor(flags, v.u32(0));
    const font = this.font(flags & 0xff);
    if (!brush || !font) return;
    const fmt = this.object(v.u32(4));
    const format = fmt?.kind === 'format' ? fmt.format : { flags: 0, align: 0, lineAlign: 0 };
    const n = v.u32(8);
    const [x, y, w, h] = readRect(v, 12, false);
    if (n > v.size) throw new MetafileError('malformed', 'string longer than its record');
    let text = '';
    for (let i = 0; i < n; i++) text += String.fromCharCode(v.u16(28 + i * 2));
    this.host.spend(n);
    const style = this.paint(brush);
    if (!style || !text) return;
    const host = this.host, ctx = host.ctx;
    host.setClip(this.s.clip);
    ctx.save();
    try {
      const frame = this.textFrame(font.em, font.f);
      if (frame) this.layoutString(text, frame, format, font, style, x, y, w, h);
    } finally {
      ctx.restore();
    }
  }

  /** Lay a DrawString string out in its rectangle (world units): hard breaks, greedy wrapping, the format's alignment. */
  private layoutString(text: string, frame: { unit: number; size: number }, format: Format, font: { underline: boolean; strike: boolean },
    style: string | CanvasGradient | CanvasPattern, x: number, y: number, w: number, h: number): void {
    const ctx = this.host.ctx;
    const { unit, size } = frame;
    const m = ctx.measureText('Hg');
    const ascent = m.fontBoundingBoxAscent ?? size * 0.9, descent = m.fontBoundingBoxDescent ?? size * 0.21;
    const lineHeight = ascent + descent;
    // lines: hard breaks, then (unless NoWrap) greedy word wrap in the layout width
    const width = w / unit;
    const lines: string[] = [];
    for (const para of text.split(/\r\n|\r|\n/)) {
      if (!(width > 0) || format.flags & 0x1000) { lines.push(para); continue; }
      let line = '';
      for (const word of para.split(/(?<=\s)/)) {
        if (line && ctx.measureText(line + word).width > width) { lines.push(line.trimEnd()); line = word.trimStart(); }
        else line += word;
      }
      lines.push(line);
    }
    const total = lines.length * lineHeight;
    const top = y / unit + (format.lineAlign === 1 ? (h / unit - total) / 2 : format.lineAlign === 2 ? h / unit - total : 0);
    ctx.fillStyle = style;
    lines.forEach((line, i) => {
      const lw = ctx.measureText(line).width;
      // near, centre or far in the layout width; around the point when the rectangle has no width
      const room = width > 0 ? width - lw : -lw;
      const lx = x / unit + (format.align === 1 ? room / 2 : format.align === 2 ? room : 0);
      const ly = top + i * lineHeight + ascent;
      ctx.fillText(line, lx, ly);
      if (font.underline) ctx.fillRect(lx, ly + size * 0.1, lw, size / 14);
      if (font.strike) ctx.fillRect(lx, ly - size * 0.28, lw, size / 14);
    });
  }

  private drawDriverString(flags: number, v: View): void {
    const brush = this.brushFor(flags, v.u32(0));
    const font = this.font(flags & 0xff);
    const options = v.u32(4);
    const n = v.u32(12);
    if (!brush || !font || !(options & 1)) return;           // glyph indexes cannot be drawn without the font file
    if (n > v.size / 10) throw new MetafileError('malformed', 'glyphs outside their record');
    const chars = Array.from({ length: n }, (_, i) => String.fromCharCode(v.u16(16 + i * 2)));
    // one position per glyph (with DriverStringOptionsRealizedAdvance only the first counts); the optional
    // per-glyph transform is not applied
    const positions = readPoints(v, 16 + n * 2, n, false, false)[0];
    this.host.spend(n);
    const style = this.paint(brush);
    if (!style) return;
    const host = this.host, ctx = host.ctx;
    host.setClip(this.s.clip);
    ctx.save();
    try {
      const frame = this.textFrame(font.em, font.f);
      if (!frame) return;
      ctx.fillStyle = style;
      const { unit } = frame;
      let x = positions[0]! / unit, y = positions[1]! / unit;
      for (let i = 0; i < n; i++) {
        if (!(options & 4)) { x = positions[i * 2]! / unit; y = positions[i * 2 + 1]! / unit; }
        ctx.fillText(chars[i]!, x, y);
        if (options & 4) x += ctx.measureText(chars[i]!).width;
      }
    } finally {
      ctx.restore();
    }
  }

  /** A region object as a clip, through the current world transform (rectangles stay exact). */
  private regionClip(region: Region): Clip {
    // post-order over the tree with an explicit stack: the tree's depth comes from the file
    const done = new Map<Region, Clip>();
    const stack: Region[] = [region];
    while (stack.length) {
      const node = stack.at(-1)!;
      if (node.type === 'op' && node.left && node.right && (!done.has(node.left) || !done.has(node.right))) {
        if (!done.has(node.left)) stack.push(node.left);
        if (!done.has(node.right)) stack.push(node.right);
        continue;
      }
      stack.pop();
      done.set(node, this.regionLeaf(node, done));
    }
    return done.get(region) ?? [];
  }

  private regionLeaf(node: Region, done: Map<Region, Clip>): Clip {
    switch (node.type) {
      case 'rect': return this.combine([], RGN_COPY, node.rect);
      case 'path': return combinePath([], RGN_COPY, plusPath(node.path, this.m), 'evenodd');
      case 'empty': return [{ kind: 'rects', rects: [] }];
      case 'infinite': return [];
      case 'op': {
        const a = (node.left && done.get(node.left)) ?? [], b = (node.right && done.get(node.right)) ?? [];
        return combineClips(a, node.op, b);
      }
    }
  }

  /** Combine a rectangle (world units) with a clip. */
  private combine(clip: Clip, mode: number, r: [number, number, number, number]): Clip {
    const path = rectPath(this.m, r);
    if (axisAligned(this.m)) {
      const b = path.bounds();
      return combineRects(clip, mode, b ? [b] : []);
    }
    return combinePath(clip, mode, path, 'nonzero');
  }

  /** Play one record; returns true for EmfPlusGetDC (EMF records follow). */
  play(r: PlusRecord): boolean {
    const v = r.data, f = r.flags, s = this.s;
    const compressed = (f & 0x4000) !== 0, relative = (f & 0x0800) !== 0;
    this.host.spend(1);
    switch (r.type) {
      case 0x4001:                                        // Header
        s.video = (v.u32(4) & 1) !== 0;
        if (v.u32(8) > 0 && v.u32(8) < 10000) s.dpiX = v.u32(8);
        if (v.u32(12) > 0 && v.u32(12) < 10000) s.dpiY = v.u32(12);
        this.refresh();
        return false;
      case 0x4004: return true;                           // GetDC
      case OBJECT: {
        const o = assembleObject(s, r);
        if (o) {
          let obj: PlusObject;
          try { obj = readObject(o.type, o.data, o.key); } catch (e) {
            if (!(e instanceof MetafileError) || stops(e)) throw e;
            obj = { kind: 'other' };
          }
          s.objects.set(o.id, obj);
        }
        return false;
      }
      case 0x4009: {                                      // Clear
        const host = this.host;
        host.setClip(s.clip);
        host.ctx.globalCompositeOperation = 'source-over';
        host.ctx.fillStyle = css(v.u32(0));
        const fr = host.mf.frame;
        const [x0, y0] = apply(host.d2o, fr.l, fr.t), [x1, y1] = apply(host.d2o, fr.r, fr.b);
        host.ctx.fillRect(Math.min(x0, x1), Math.min(y0, y1), Math.abs(x1 - x0), Math.abs(y1 - y0));
        return false;
      }
      case 0x400a: case 0x400b: {                         // FillRects, DrawRects
        const fill = r.type === 0x400a;
        const at = fill ? 4 : 0;
        const n = v.u32(at);
        if (n > v.size / 8) throw new MetafileError('malformed', 'rectangles outside their record');
        const path = new Path();
        for (let i = 0; i < n; i++) path.append(rectPath(this.m, readRect(v, at + 4 + i * (compressed ? 8 : 16), compressed)));
        if (fill) this.fill(path, this.brushFor(f, v.u32(0)));
        else this.stroke(path, f & 0xff);
        return false;
      }
      case 0x400c: {                                      // FillPolygon
        const [pts] = readPoints(v, 8, v.u32(4), compressed, relative);
        const path = new Path();
        const xy: number[] = [];
        for (let i = 0; i + 1 < pts.length; i += 2) xy.push(...apply(this.m, pts[i]!, pts[i + 1]!));
        path.polygon(xy);
        this.fill(path, this.brushFor(f, v.u32(0)), 'evenodd');
        return false;
      }
      case 0x400d: {                                      // DrawLines
        const [pts] = readPoints(v, 4, v.u32(0), compressed, relative);
        const xy: number[] = [];
        for (let i = 0; i + 1 < pts.length; i += 2) xy.push(...apply(this.m, pts[i]!, pts[i + 1]!));
        const path = new Path();
        path.polygon(xy, (f & 0x2000) !== 0);
        this.stroke(path, f & 0xff);
        return false;
      }
      case 0x400e: this.fill(arcPath(this.m, readRect(v, 4, compressed), 0, 360, false, true), this.brushFor(f, v.u32(0))); return false;
      case 0x400f: this.stroke(arcPath(this.m, readRect(v, 0, compressed), 0, 360, false, true), f & 0xff); return false;
      case 0x4010: this.fill(arcPath(this.m, readRect(v, 12, compressed), v.f32(4), v.f32(8), true, true), this.brushFor(f, v.u32(0))); return false;
      case 0x4011: this.stroke(arcPath(this.m, readRect(v, 8, compressed), v.f32(0), v.f32(4), true, true), f & 0xff); return false;
      case 0x4012: this.stroke(arcPath(this.m, readRect(v, 8, compressed), v.f32(0), v.f32(4), false, false), f & 0xff); return false;
      case 0x4013: {                                      // FillRegion
        const o = this.object(f & 0xff);
        if (o?.kind !== 'region') return false;
        const clip = this.regionClip(o.region);
        const host = this.host;
        // filled as a clip over the frame
        host.setClip([...s.clip, ...clip]);
        const fr = host.mf.frame;
        const p = new Path();
        p.polygon([...apply(host.d2o, fr.l, fr.t), ...apply(host.d2o, fr.r, fr.t), ...apply(host.d2o, fr.r, fr.b), ...apply(host.d2o, fr.l, fr.b)]);
        const style = this.brushFor(f, v.u32(0));
        const paint = style && this.paint(style);
        if (paint) {
          host.ctx.globalCompositeOperation = 'source-over';
          host.ctx.fillStyle = paint;
          host.ctx.beginPath(); p.trace(host.ctx); host.ctx.fill();
        }
        return false;
      }
      case 0x4014: case 0x4015: {                         // FillPath, DrawPath
        const o = this.object(f & 0xff);
        if (o?.kind !== 'path') return false;
        const path = plusPath(o.path, this.m);
        if (r.type === 0x4014) this.fill(path, this.brushFor(f, v.u32(0)), 'evenodd');
        else this.stroke(path, v.u32(0));
        return false;
      }
      case 0x4016: {                                      // FillClosedCurve
        const [pts] = readPoints(v, 12, v.u32(8), compressed, relative);
        this.fill(curvePath(this.m, pts, v.f32(4), true), this.brushFor(f, v.u32(0)), f & 0x2000 ? 'nonzero' : 'evenodd');
        return false;
      }
      case 0x4017: {                                      // DrawClosedCurve
        const [pts] = readPoints(v, 8, v.u32(4), compressed, relative);
        this.stroke(curvePath(this.m, pts, v.f32(0), true), f & 0xff);
        return false;
      }
      case 0x4018: {                                      // DrawCurve
        const [pts] = readPoints(v, 16, v.u32(12), compressed, false);
        this.stroke(curvePath(this.m, pts, v.f32(0), false, v.u32(4), v.u32(8)), f & 0xff);
        return false;
      }
      case 0x4019: {                                      // DrawBeziers
        const [pts] = readPoints(v, 4, v.u32(0), compressed, relative);
        const path = new Path();
        const n = pts.length >> 1;
        if (n >= 1) path.moveTo(...apply(this.m, pts[0]!, pts[1]!));
        const at = (k: number) => apply(this.m, pts[k * 2]!, pts[k * 2 + 1]!);
        for (let i = 1; i + 2 < n; i += 3) path.curveTo(...at(i), ...at(i + 1), ...at(i + 2));
        this.stroke(path, f & 0xff);
        return false;
      }
      case 0x401a: {                                      // DrawImage
        const src = readRect(v, 8, false);
        const [x, y, w, h] = readRect(v, 24, compressed);
        this.drawImage(f & 0xff, src, [x, y, x + w, y, x, y + h]);
        return false;
      }
      case 0x401b: {                                      // DrawImagePoints
        const src = readRect(v, 8, false);
        const [pts] = readPoints(v, 28, Math.min(3, v.u32(24)), compressed, relative);
        if (pts.length < 6) return false;
        this.drawImage(f & 0xff, src, [pts[0]!, pts[1]!, pts[2]!, pts[3]!, pts[4]!, pts[5]!]);
        return false;
      }
      case 0x401c: this.drawString(f, v); return false;
      case 0x4036: this.drawDriverString(f, v); return false;
      case 0x4025: case 0x4027: case 0x4028: {             // Save, BeginContainer(NoParams)
        const index = v.u32(r.type === 0x4027 ? 32 : 0);
        if (s.saved.size >= 4096) throw stop('too many saved EMF+ states');
        s.saved.set(index, { world: s.world, unit: s.unit, scale: s.scale, clip: s.clip });
        if (r.type === 0x4027) {
          // the container maps its source rectangle (in the flags' unit) onto the destination rectangle (in world
          // units); the page transform stays
          const [dx, dy, dw, dh] = readRect(v, 0, false);
          const ux = unitScale(s, f & 0xff), uy = unitScale(s, f & 0xff, true);
          const [sx, sy, sw, sh] = readRect(v, 16, false).map((n, i) => n * (i % 2 ? uy : ux)) as [number, number, number, number];
          if (sw && sh) {
            const m = compose(s.world, [dw / sw, 0, 0, dh / sh, dx - (sx * dw) / sw, dy - (sy * dh) / sh]);
            if (usable(m)) s.world = m;
          }
        }
        this.refresh();
        return false;
      }
      case 0x4026: case 0x4029: {                         // Restore, EndContainer
        const saved = s.saved.get(v.u32(0));
        if (!saved) return false;
        s.saved.delete(v.u32(0));
        Object.assign(s, saved);
        this.refresh();
        return false;
      }
      case 0x402a: { const m = matrixAt(v, 0); if (usable(m)) s.world = m; this.refresh(); return false; }
      case 0x402b: s.world = IDENTITY; this.refresh(); return false;
      case 0x402c: case 0x402d: case 0x402e: case 0x402f: {
        let m: Matrix;
        if (r.type === 0x402c) m = matrixAt(v, 0);
        else if (r.type === 0x402d) m = [1, 0, 0, 1, v.f32(0), v.f32(4)];
        else if (r.type === 0x402e) m = [v.f32(0), 0, 0, v.f32(4), 0, 0];
        else { const a = (v.f32(0) * Math.PI) / 180; m = [Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0]; }
        // the A flag appends (applies after the current transform); otherwise the matrix applies first
        const next = f & 0x2000 ? compose(m, s.world) : compose(s.world, m);
        if (usable(next)) s.world = next;
        this.refresh();
        return false;
      }
      case 0x4030: {                                      // SetPageTransform
        const scale = v.f32(0);
        s.unit = f & 0xff;
        s.scale = Number.isFinite(scale) && scale > 0 ? scale : 1;
        this.refresh();
        return false;
      }
      case 0x4031: s.clip = []; return false;             // ResetClip
      case 0x4032: s.clip = this.combine(s.clip, combineMode(f), readRect(v, 0, false)); return false;
      case 0x4033: {                                      // SetClipPath
        const o = this.object(f & 0xff);
        if (o?.kind === 'path') s.clip = combinePath(s.clip, combineMode(f), plusPath(o.path, this.m), 'evenodd');
        return false;
      }
      case 0x4034: {                                      // SetClipRegion
        const o = this.object(f & 0xff);
        if (o?.kind === 'region') s.clip = combineClips(s.clip, combineMode(f), this.regionClip(o.region));
        return false;
      }
      case 0x4035: {                                      // OffsetClip
        const dx = v.f32(0), dy = v.f32(4);
        if (Number.isFinite(dx + dy)) s.clip = translateClip(s.clip, this.m[0] * dx + this.m[2] * dy, this.m[1] * dx + this.m[3] * dy);
        return false;
      }
      default: return false;
    }
  }
}

/** EMF+ CombineMode to the GDI region mode. */
function combineMode(flags: number): number {
  switch ((flags >> 8) & 0xf) {
    case 1: return 1;          // intersect: RGN_AND
    case 2: return 2;          // union: RGN_OR
    case 3: return 3;          // XOR
    case 4: return 4;          // exclude: RGN_DIFF
    default: return RGN_COPY;  // replace, and complement (approximated)
  }
}

/** Combine two clips: exact for rectangles, by intersection or difference for paths, approximate otherwise. */
function combineClips(a: Clip, mode: number, b: Clip): Clip {
  // an infinite region: intersecting keeps the clip, excluding empties it, anything else is unclipped
  if (b.length === 0) return mode === 1 ? a : mode === 4 ? [{ kind: 'rects', rects: [] }] : [];
  if (b.length === 1 && b[0]!.kind === 'rects') return combineRects(a, mode, b[0]!.rects);
  if (b.length === 1 && b[0]!.kind === 'path') return combinePath(a, mode, b[0]!.path, b[0]!.rule);
  return mode === 1 ? [...a, ...b] : mode === RGN_COPY ? b : a;
}

/** Play EMF+ records; returns true when the last one was EmfPlusGetDC, so that EMF records are played next. */
export function playPlus(host: PlusHost, state: PlusState, records: PlusRecord[]): boolean {
  const p = new PlusPlayer(host, state);
  let gdi = false;
  for (const r of records) {
    try {
      gdi = p.play(r);
    } catch (e) {
      // one bad record is skipped, as with EMF records; running out of budget stops everything
      if (!(e instanceof MetafileError) || stops(e)) throw e;
      gdi = false;
    }
  }
  return gdi;
}

