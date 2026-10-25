// packages/metafile/src/emf.ts  -- Enhanced Metafiles ([MS-EMF]): header and records, as GDI operations
import { splitDib } from './dib.ts';
import { MetafileError, stop, stops } from './errors.ts';
import type { BltRecord, BrushObject, Dib, FontObject, GdiObject, Matrix, Metafile, MetaRecord, PenObject, PlusRecord, Rect, TextRecord } from './types.ts';
import { View } from './view.ts';
import { MAX_RECORDS } from './wmf.ts';

export const EMF_SIGNATURE = 0x464d4520;
const EMFPLUS_SIGNATURE = 0x2b464d45;
/** Highest object table index a record may use. */
export const MAX_OBJECTS = 65_536;
/** Stock objects are selected by index with the high bit set. */
export const STOCK = 0x80000000;

const ETO_NO_RECT = 0x100;
const ETO_SMALL_CHARS = 0x200;
const ETO_PDY = 0x2000;

/** True when the bytes start with an EMF header record. */
export function isEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 88) return false;
  const v = new View(bytes);
  return v.u32(0) === 1 && v.u32(40) === EMF_SIGNATURE;
}

const rectAt = (v: View, at: number): Rect => ({ l: v.i32(at), t: v.i32(at + 4), r: v.i32(at + 8), b: v.i32(at + 12) });

function matrixAt(v: View, at: number): Matrix {
  const m: Matrix = [v.f32(at), v.f32(at + 4), v.f32(at + 8), v.f32(at + 12), v.f32(at + 16), v.f32(at + 20)];
  if (!m.every(Number.isFinite)) throw new MetafileError('malformed', 'transform with non-finite values');
  return m;
}

/** `count` points of 32-bit (or, `small`, 16-bit) coordinates at `at`. */
function pointsAt(v: View, at: number, count: number, small: boolean): Int32Array {
  const size = small ? 4 : 8;
  if (count > (v.size - at) / size || !v.has(at, count * size)) throw new MetafileError('malformed', `${count} points do not fit their record`);
  const out = new Int32Array(count * 2);
  for (let i = 0; i < count * 2; i++) out[i] = small ? v.i16(at + i * 2) : v.i32(at + i * 4);
  return out;
}

/** The rectangles of a RegionData object at `at`, `size` bytes long. */
function regionAt(v: View, at: number, size: number): Rect[] {
  if (size < 32 || !v.has(at, size)) throw new MetafileError('malformed', 'region data outside its record');
  const count = v.u32(at + 8);
  const first = at + 32;                        // after the RegionDataHeader
  if (count > (size - 32) / 16 || !v.has(first, count * 16)) throw new MetafileError('malformed', 'region rectangles outside their record');
  const rects: Rect[] = [];
  for (let i = 0; i < count; i++) rects.push(rectAt(v, first + i * 16));
  return rects;
}

/** A bitmap in an EMF record, from its offset and size fields; undefined when the record has none. */
function dibAt(v: View, usage: number, offBmi: number, cbBmi: number, offBits: number, cbBits: number): Dib | undefined {
  if (cbBmi === 0 || offBmi === 0) return undefined;
  return splitDib(v, offBmi, cbBmi, offBits, cbBits, usage);
}

function readFontW(v: View, at: number): FontObject {
  let face = '';
  for (let i = 0; i < 32 && v.has(at + 28 + i * 2, 2); i++) {
    const c = v.u16(at + 28 + i * 2);
    if (c === 0) break;
    face += String.fromCharCode(c);
  }
  return {
    kind: 'font', height: v.i32(at), width: v.i32(at + 4), escapement: v.i32(at + 8), orientation: v.i32(at + 12),
    weight: v.i32(at + 16), italic: v.u8(at + 20) !== 0, underline: v.u8(at + 21) !== 0, strikeOut: v.u8(at + 22) !== 0,
    charSet: v.u8(at + 23), quality: v.u8(at + 26), pitchAndFamily: v.u8(at + 27), face: face.trim(),
  };
}

/** An EmrText object at `at` (offsets inside it are from the record start). */
function emrText(v: View, at: number, wide: boolean): TextRecord {
  const x = v.i32(at), y = v.i32(at + 4);
  const chars = v.u32(at + 8);
  const offString = v.u32(at + 12);
  const options = v.u32(at + 16);
  const box = options & ETO_NO_RECT ? undefined : rectAt(v, at + 20);
  const offDx = v.u32(at + 36);
  if (chars > v.size) throw new MetafileError('malformed', 'string longer than its record');
  const rec: TextRecord = { type: 'text', x, y, options };
  if (box && (options & 6)) rec.box = box;
  if (wide) {
    if (!v.has(offString, chars * 2)) throw new MetafileError('malformed', 'string outside its record');
    let s = '';
    for (let i = 0; i < chars; i++) s += String.fromCharCode(v.u16(offString + i * 2));
    rec.text = s;
  } else {
    rec.bytes = v.sub(offString, chars);
  }
  const pdy = (options & ETO_PDY) !== 0;
  if (offDx > 0 && chars > 0 && v.has(offDx, chars * (pdy ? 8 : 4))) {
    const dx = new Int32Array(chars);
    const dy = pdy ? new Int32Array(chars) : undefined;
    for (let i = 0; i < chars; i++) {
      dx[i] = v.i32(offDx + i * (pdy ? 8 : 4));
      if (dy) dy[i] = v.i32(offDx + i * 8 + 4);
    }
    rec.dx = dx;
    if (dy) rec.dy = dy;
  }
  return rec;
}

function stretchBlt(v: View, stretch: boolean): BltRecord {
  const rec: BltRecord = {
    type: 'blt', x: v.i32(24), y: v.i32(28), w: v.i32(32), h: v.i32(36), rop: v.u32(40),
    sx: v.i32(44), sy: v.i32(48), sw: stretch ? v.i32(100) : v.i32(32), sh: stretch ? v.i32(104) : v.i32(36),
  };
  const dib = dibAt(v, v.u32(80), v.u32(84), v.u32(88), v.u32(92), v.u32(96));
  if (!dib) return rec;
  rec.dib = dib;
  // the source DC's transform maps the source rectangle to bitmap pixels (identity in practice; zero when unused)
  const m = [v.f32(52), v.f32(56), v.f32(60), v.f32(64), v.f32(68), v.f32(72)];
  if (m.every(Number.isFinite) && m[0]! > 0 && m[3]! > 0 && (m[0] !== 1 || m[3] !== 1 || m[4] !== 0 || m[5] !== 0)) {
    rec.sx = rec.sx * m[0]! + m[4]!;
    rec.sy = rec.sy * m[3]! + m[5]!;
    rec.sw *= m[0]!;
    rec.sh *= m[3]!;
  }
  return rec;
}

/** EMF+ records inside an EMR_COMMENT ([MS-EMFPLUS] 2.3): their headers checked, their data left as views. */
function plusRecords(v: View, at: number, size: number): PlusRecord[] {
  const out: PlusRecord[] = [];
  let p = at;
  const end = at + size;
  while (p + 12 <= end) {
    const type = v.u16(p), flags = v.u16(p + 2), recSize = v.u32(p + 4), dataSize = v.u32(p + 8);
    if (recSize < 12 || p + recSize > end || dataSize > recSize - 12) break;
    out.push({ type, flags, data: v.view(p + 12, recSize - 12) });
    p += recSize;
  }
  return out;
}

const objIndex = (n: number) => {
  if (n >= MAX_OBJECTS) throw new MetafileError('malformed', `object index ${n} is over ${MAX_OBJECTS}`);
  return n;
};

/** One record (`v` covers it, header included) as GDI operations; undefined for records that change nothing drawn. */
function record(type: number, v: View): MetaRecord | MetaRecord[] | undefined {
  switch (type) {
    case 2: case 3: case 4: case 5: case 6: case 85: case 86: case 87: case 88: case 89: {
      const small = type >= 85;
      const shape = (['bezier', 'polygon', 'polyline', 'bezierTo', 'polylineTo'] as const)[(small ? type - 85 : type - 2)]!;
      return { type: 'poly', shape, points: pointsAt(v, 28, v.u32(24), small) };
    }
    case 7: case 8: case 90: case 91: {
      const small = type >= 90;
      const polys = v.u32(24);
      const total = v.u32(28);
      if (polys > (v.size - 32) / 4) throw new MetafileError('malformed', 'polygon counts outside their record');
      const counts = new Uint32Array(polys);
      let sum = 0;
      for (let i = 0; i < polys; i++) sum += counts[i] = v.u32(32 + i * 4);
      if (sum !== total) throw new MetafileError('malformed', 'polygon counts do not add up');
      return { type: 'polyPoly', shape: type === 7 || type === 90 ? 'polyline' : 'polygon', counts, points: pointsAt(v, 32 + polys * 4, total, small) };
    }
    case 56: case 92: {
      const small = type === 92;
      const count = v.u32(24);
      const points = pointsAt(v, 28, count, small);
      return { type: 'polyDraw', points, kinds: v.sub(28 + count * (small ? 4 : 8), count) };
    }
    case 9: return { type: 'windowExt', x: v.i32(8), y: v.i32(12) };
    case 10: return { type: 'windowOrg', x: v.i32(8), y: v.i32(12) };
    case 11: return { type: 'viewportExt', x: v.i32(8), y: v.i32(12) };
    case 12: return { type: 'viewportOrg', x: v.i32(8), y: v.i32(12) };
    case 13: return { type: 'brushOrg', x: v.i32(8), y: v.i32(12) };
    case 15: return { type: 'pixel', x: v.i32(8), y: v.i32(12), color: v.u32(16) };
    case 17: return { type: 'mapMode', mode: v.u32(8) };
    case 18: return { type: 'bkMode', value: v.u32(8) };
    case 19: return { type: 'polyFillMode', value: v.u32(8) };
    case 20: return { type: 'rop2', value: v.u32(8) };
    case 21: return { type: 'stretchMode', value: v.u32(8) };
    case 22: return { type: 'textAlign', value: v.u32(8) };
    case 24: return { type: 'textColor', color: v.u32(8) };
    case 25: return { type: 'bkColor', color: v.u32(8) };
    case 26: return { type: 'offsetClip', x: v.i32(8), y: v.i32(12) };
    case 27: return { type: 'moveTo', x: v.i32(8), y: v.i32(12) };
    case 29: return { type: 'clipRect', op: 'exclude', box: rectAt(v, 8) };
    case 30: return { type: 'clipRect', op: 'intersect', box: rectAt(v, 8) };
    case 31: return { type: 'scaleViewportExt', xNum: v.i32(8), xDenom: v.i32(12), yNum: v.i32(16), yDenom: v.i32(20) };
    case 32: return { type: 'scaleWindowExt', xNum: v.i32(8), xDenom: v.i32(12), yNum: v.i32(16), yDenom: v.i32(20) };
    case 33: return { type: 'saveDC' };
    case 34: return { type: 'restoreDC', n: v.i32(8) };
    case 35: return { type: 'worldTransform', m: matrixAt(v, 8) };
    case 36: return { type: 'modifyWorldTransform', m: matrixAt(v, 8), mode: v.u32(32) };
    case 37: case 48: return { type: 'select', index: v.u32(8) };
    case 40: return { type: 'delete', index: v.u32(8) };
    case 38: {
      // LogPen: the width is in logical units; zero is one device pixel
      const width = v.i32(16);
      const pen: PenObject = { kind: 'pen', style: v.u32(12), width: Math.abs(width), color: v.u32(24), cosmetic: width === 0 };
      return { type: 'create', index: objIndex(v.u32(8)), object: pen };
    }
    case 95: {
      // LogPenEx after ihPen and the bitmap fields (a pattern pen's bitmap is ignored: it draws in its colour)
      const style = v.u32(28);
      const width = v.u32(32);
      const geometric = (style & 0x000f0000) === 0x00010000;
      const pen: PenObject = { kind: 'pen', style, width, color: v.u32(40), cosmetic: !geometric || width === 0 };
      if (v.u32(36) === 1) pen.style = (style & ~0xf) | 5;            // BS_NULL brush: no line
      const n = v.u32(48);
      if ((style & 0xf) === 7 && n > 0 && n <= 16 && v.has(52, n * 4)) {
        pen.dashes = Array.from({ length: n }, (_, i) => v.u32(52 + i * 4));
      }
      return { type: 'create', index: objIndex(v.u32(8)), object: pen };
    }
    case 39: {
      const brush: BrushObject = { kind: 'brush', style: v.u32(12), color: v.u32(16), hatch: v.u32(20) };
      return { type: 'create', index: objIndex(v.u32(8)), object: brush };
    }
    case 93: case 94: {
      const usage = v.u32(12);
      const dib = dibAt(v, usage, v.u32(16), v.u32(20), v.u32(24), v.u32(28));
      const brush: BrushObject = dib
        ? { kind: 'brush', style: 6, color: 0, hatch: 0, dib }
        : { kind: 'brush', style: 0, color: 0x808080, hatch: 0 };
      if (type === 93) brush.mono = true;
      return { type: 'create', index: objIndex(v.u32(8)), object: brush };
    }
    case 82: return { type: 'create', index: objIndex(v.u32(8)), object: readFontW(v, 12) };
    case 49: {
      const count = v.u16(14);
      const colors = new Uint32Array(count);
      for (let i = 0; i < count; i++) colors[i] = (v.u8(16 + i * 4) << 16) | (v.u8(17 + i * 4) << 8) | v.u8(18 + i * 4);
      return { type: 'create', index: objIndex(v.u32(8)), object: { kind: 'palette', colors } };
    }
    case 99: case 122: return { type: 'create', index: objIndex(v.u32(8)), object: { kind: 'other' } };
    case 41: return { type: 'angleArc', x: v.i32(8), y: v.i32(12), r: v.u32(16), start: v.f32(20), sweep: v.f32(24) };
    case 42: return { type: 'box', shape: 'ellipse', box: rectAt(v, 8) };
    case 43: return { type: 'box', shape: 'rectangle', box: rectAt(v, 8) };
    case 44: return { type: 'roundRect', box: rectAt(v, 8), w: v.i32(24), h: v.i32(28) };
    case 45: case 46: case 47: case 55:
      return {
        type: 'arc', shape: type === 45 ? 'arc' : type === 46 ? 'chord' : type === 47 ? 'pie' : 'arcTo',
        box: rectAt(v, 8), start: [v.i32(24), v.i32(28)], end: [v.i32(32), v.i32(36)],
      };
    case 54: return { type: 'lineTo', x: v.i32(8), y: v.i32(12) };
    case 57: return { type: 'arcDirection', value: v.u32(8) };
    case 58: {
      // the spec says an integer, but writers store a float: a huge integer is a float's bits
      const raw = v.u32(8);
      return { type: 'miterLimit', value: raw > 0xffff ? v.f32(8) : raw };
    }
    case 59: return { type: 'path', op: 'begin' };
    case 60: return { type: 'path', op: 'end' };
    case 61: return { type: 'path', op: 'close' };
    case 62: return { type: 'path', op: 'fill' };
    case 63: return { type: 'path', op: 'strokeAndFill' };
    case 64: return { type: 'path', op: 'stroke' };
    case 65: return { type: 'path', op: 'flatten' };
    case 66: return { type: 'path', op: 'widen' };
    case 67: return { type: 'clipPath', mode: v.u32(8) };
    case 68: return { type: 'path', op: 'abort' };
    case 70: {
      const size = v.u32(8);
      if (size >= 4 && v.has(12, size) && v.u32(12) === EMFPLUS_SIGNATURE) return { type: 'plus', records: plusRecords(v, 16, size - 4) };
      return undefined;
    }
    case 71: return { type: 'region', op: 'fill', brush: v.u32(28), rects: regionAt(v, 32, v.u32(24)) };
    case 72: return { type: 'region', op: 'frame', brush: v.u32(28), w: v.i32(32), h: v.i32(36), rects: regionAt(v, 40, v.u32(24)) };
    case 73: return { type: 'region', op: 'invert', rects: regionAt(v, 28, v.u32(24)) };
    case 74: return { type: 'region', op: 'paint', rects: regionAt(v, 28, v.u32(24)) };
    case 75: {
      const size = v.u32(8);
      const mode = v.u32(12);
      return { type: 'clipRegion', mode, rects: size === 0 && mode === 5 ? null : regionAt(v, 16, size) };
    }
    case 76: return stretchBlt(v, false);
    case 77: return stretchBlt(v, true);
    case 114: case 116: {
      const rec = stretchBlt(v, true);
      if (type === 114) {
        rec.alpha = { constant: v.u8(42), perPixel: (v.u8(43) & 1) !== 0 };
        rec.rop = 0xcc0020;
      } else {
        rec.transparent = v.u32(40);
        rec.rop = 0xcc0020;
      }
      return rec;
    }
    case 78: {
      const rec = stretchBlt(v, false);
      // MAKEROP4(fore, back): the foreground operation sits in the usual place
      rec.rop = v.u32(40) & 0x00ffffff;
      const mask = dibAt(v, v.u32(108), v.u32(112), v.u32(116), v.u32(120), v.u32(124));
      if (mask) rec.mask = { dib: mask, x: v.i32(100), y: v.i32(104) };
      return rec;
    }
    case 79: {
      const corners: BltRecord['corners'] = [v.i32(24), v.i32(28), v.i32(32), v.i32(36), v.i32(40), v.i32(44)];
      const rec: BltRecord = { type: 'blt', x: 0, y: 0, w: 0, h: 0, rop: 0xcc0020, sx: v.i32(48), sy: v.i32(52), sw: v.i32(56), sh: v.i32(60), corners };
      const dib = dibAt(v, v.u32(92), v.u32(96), v.u32(100), v.u32(104), v.u32(108));
      if (dib) rec.dib = dib;
      const mask = dibAt(v, v.u32(120), v.u32(124), v.u32(128), v.u32(132), v.u32(136));
      if (mask) rec.mask = { dib: mask, x: v.i32(112), y: v.i32(116) };
      return rec;
    }
    case 80: {
      const dib = dibAt(v, v.u32(64), v.u32(48), v.u32(52), v.u32(56), v.u32(60));
      if (!dib) return undefined;
      const sw = v.i32(40), sh = v.i32(44), sx = v.i32(32), sy = v.i32(36);
      // the source origin is the lower-left corner of a bottom-up DIB; only the scan lines given are stored
      const start = v.u32(68), scans = v.u32(72);
      if (!dib.topDown && !dib.encoded && scans > 0 && scans < dib.height) dib.height = scans;
      const top = dib.topDown || dib.encoded ? sy : dib.height + start - sy - sh;
      return { type: 'blt', x: v.i32(24), y: v.i32(28), w: sw, h: sh, sx, sy: top, sw, sh, rop: 0xcc0020, dib, device: true };
    }
    case 81: {
      const dib = dibAt(v, v.u32(64), v.u32(48), v.u32(52), v.u32(56), v.u32(60));
      const sx = v.i32(32), sy = v.i32(36), sw = v.i32(40), sh = v.i32(44);
      const rec: BltRecord = { type: 'blt', x: v.i32(24), y: v.i32(28), w: v.i32(72), h: v.i32(76), rop: v.u32(68), sx, sy, sw, sh };
      if (dib) {
        rec.dib = dib;
        // a bottom-up DIB's source rectangle is measured from its lower-left corner
        if (!dib.topDown && !dib.encoded) rec.sy = dib.height - sy - Math.abs(sh);
      }
      return rec;
    }
    case 83: case 84: {
      const rec = emrText(v, 36, type === 84);
      if (v.u32(24) === 2) rec.advanced = true;
      return rec;
    }
    case 96: case 97: {
      const n = v.u32(36);
      if (n > (v.size - 40) / 40) throw new MetafileError('malformed', 'strings outside their record');
      const advanced = v.u32(24) === 2;
      return Array.from({ length: n }, (_, i) => ({ ...emrText(v, 40 + i * 40, type === 97), ...(advanced ? { advanced } : {}) }));
    }
    case 108: {
      const options = v.u32(20);
      const chars = v.u32(16);
      const at = options & ETO_NO_RECT ? 36 : 52;
      const rec: TextRecord = { type: 'text', x: v.i32(8), y: v.i32(12), options: options & ~ETO_SMALL_CHARS };
      if (v.u32(24) === 2) rec.advanced = true;
      if (!(options & ETO_NO_RECT) && (options & 6)) rec.box = rectAt(v, 36);
      if (chars > v.size) throw new MetafileError('malformed', 'string longer than its record');
      if (options & ETO_SMALL_CHARS) {
        // the low bytes of UTF-16 code units, so Latin-1
        let s = '';
        for (const b of v.sub(at, chars)) s += String.fromCharCode(b);
        rec.text = s;
      } else {
        let s = '';
        if (!v.has(at, chars * 2)) throw new MetafileError('malformed', 'string outside its record');
        for (let i = 0; i < chars; i++) s += String.fromCharCode(v.u16(at + i * 2));
        rec.text = s;
      }
      return rec;
    }
    case 115: return { type: 'layout', value: v.u32(8) };
    case 118: {
      const nVer = v.u32(24), nTri = v.u32(28), mode = v.u32(32);
      if (nVer > (v.size - 36) / 16 || mode > 2) throw new MetafileError('malformed', 'gradient vertices outside their record');
      const vertices = new Int32Array(nVer * 2);
      const colors = new Uint32Array(nVer);
      for (let i = 0; i < nVer; i++) {
        const p = 36 + i * 16;
        vertices[i * 2] = v.i32(p);
        vertices[i * 2 + 1] = v.i32(p + 4);
        colors[i] = ((v.u16(p + 8) >>> 8) << 16) | ((v.u16(p + 10) >>> 8) << 8) | (v.u16(p + 12) >>> 8);
      }
      const per = mode === 2 ? 3 : 2;
      const at = 36 + nVer * 16;
      if (nTri > (v.size - at) / (per * 4)) throw new MetafileError('malformed', 'gradient indexes outside their record');
      const indexes = new Uint32Array(nTri * per);
      for (let i = 0; i < indexes.length; i++) {
        indexes[i] = v.u32(at + i * 4);
        if (indexes[i]! >= nVer) throw new MetafileError('malformed', 'gradient index out of range');
      }
      return { type: 'gradient', mode, vertices, colors, indexes };
    }
    default:
      // palettes, colour management, OpenGL, escapes, flood fills, the meta region
      return undefined;
  }
}

/** Records that create an object must still take their slot when they cannot be read. */
const CREATES = new Set([38, 39, 49, 82, 93, 94, 95, 99, 122]);

export interface EmfHeader {
  bounds: Rect;
  frame: Rect;
  /** Reference device pixels per millimetre. */
  dpmmX: number;
  dpmmY: number;
}

export function readEmfHeader(v: View): EmfHeader {
  const size = v.u32(4);
  if (v.u32(0) !== 1 || size < 88 || size > v.size || v.u32(40) !== EMF_SIGNATURE) throw new MetafileError('malformed', 'not an EMF header');
  const bounds = rectAt(v, 8);
  const frame = rectAt(v, 24);
  const devX = v.i32(72), devY = v.i32(76);
  let mmX = v.i32(80), mmY = v.i32(84);
  let umX = 0, umY = 0;
  if (size >= 108) { umX = v.i32(100); umY = v.i32(104); }
  const ok = (d: number, mm: number) => d > 0 && mm > 0 && d / mm > 0.5 && d / mm < 2000;
  // micrometres are the precise figure when present; otherwise millimetres; otherwise 96 dpi
  let dpmmX = 96 / 25.4, dpmmY = 96 / 25.4;
  if (umX > 0 && umY > 0 && ok(devX, umX / 1000) && ok(devY, umY / 1000)) { mmX = umX / 1000; mmY = umY / 1000; }
  if (ok(devX, mmX) && ok(devY, mmY)) { dpmmX = devX / mmX; dpmmY = devY / mmY; }
  return { bounds, frame, dpmmX, dpmmY };
}

/** Parse an EMF: the header, then records until EMR_EOF or a record that does not fit. */
export function parseEmf(bytes: Uint8Array): Metafile {
  const v = new View(bytes);
  if (!isEmf(bytes)) throw new MetafileError('malformed', 'not an EMF');
  const { bounds, frame: frame01mm, dpmmX, dpmmY } = readEmfHeader(v);
  const records: MetaRecord[] = [];
  let plus: Metafile['plus'] = 'none';
  let skipped = 0;
  let p = v.u32(4);
  let count = 0;
  while (p + 8 <= bytes.length) {
    if (++count > MAX_RECORDS) throw stop(`more than ${MAX_RECORDS} records`);
    const type = v.u32(p);
    const size = v.u32(p + 4);
    if (size < 8 || p + size > bytes.length) break;
    if (type === 14) { records.push({ type: 'eof' }); break; }
    let out: MetaRecord | MetaRecord[] | undefined;
    try {
      out = record(type, v.view(p, size));
    } catch (e) {
      if (!(e instanceof MetafileError) || stops(e)) throw e;
      skipped++;
      const slot = CREATES.has(type) && v.has(p + 8, 4) ? v.u32(p + 8) : MAX_OBJECTS;
      out = slot < MAX_OBJECTS ? { type: 'create', index: slot, object: { kind: 'other' } as GdiObject } : undefined;
    }
    for (const rec of out === undefined ? [] : Array.isArray(out) ? out : [out]) {
      if (rec.type === 'plus' && plus === 'none') {
        const header = rec.records.find(r => r.type === 0x4001);
        if (header) plus = header.flags & 1 ? 'dual' : 'only';
      }
      records.push(rec);
    }
    p += size;
  }
  // the frame maps onto the destination; when it is empty, the bounds (device units) stand in
  let frame: Rect = {
    l: (frame01mm.l * dpmmX) / 100, t: (frame01mm.t * dpmmY) / 100, r: (frame01mm.r * dpmmX) / 100, b: (frame01mm.b * dpmmY) / 100,
  };
  if (!(frame.r - frame.l > 0 && frame.b - frame.t > 0)) {
    frame = bounds.r > bounds.l && bounds.b > bounds.t ? { l: bounds.l, t: bounds.t, r: bounds.r + 1, b: bounds.b + 1 } : { l: 0, t: 0, r: 1, b: 1 };
  }
  const width = ((frame.r - frame.l) / dpmmX / 25.4) * 96;
  const height = ((frame.b - frame.t) / dpmmY / 25.4) * 96;
  return { format: 'emf', frame, width, height, dpmmX, dpmmY, records, plus, skipped };
}
