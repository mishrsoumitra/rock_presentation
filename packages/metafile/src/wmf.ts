// packages/metafile/src/wmf.ts  -- Windows Metafiles ([MS-WMF]): header, placeable header and records, as GDI operations
import { bitmap16, packedDib } from './dib.ts';
import { MetafileError, stop, stops } from './errors.ts';
import { decodeAnsi } from './text.ts';
import type { BrushObject, Dib, FontObject, GdiObject, Metafile, MetaRecord, Rect } from './types.ts';
import { View } from './view.ts';

export const PLACEABLE_KEY = 0x9ac6cdd7;
/** Most records one file may hold. */
export const MAX_RECORDS = 1_000_000;

const BS_PATTERN = 3;
const BS_DIBPATTERNPT = 6;

/** The offset of the standard header (after any placeable header), or -1 when the bytes are not a WMF. */
export function wmfHeaderAt(bytes: Uint8Array): number {
  const v = new View(bytes);
  const at = v.has(0, 4) && v.u32(0) === PLACEABLE_KEY ? 22 : 0;
  if (!v.has(at, 18)) return -1;
  const type = v.u16(at);
  const words = v.u16(at + 2);
  const version = v.u16(at + 4);
  return (type === 1 || type === 2) && words === 9 && (version === 0x100 || version === 0x300) ? at : -1;
}

/**
 * STRETCHDIB and SETDIBTODEV record StretchDIBits and SetDIBitsToDevice calls, whose source rectangle is measured
 * from a bottom-up DIB's lower-left corner: as rows from the top. (The BITBLT family counts from the top, as the
 * device bitmap the DIB stands for does; files depend on it.)
 */
const fromTop = (dib: Dib, sy: number, sh: number) => (dib.topDown ? sy : dib.height - sy - Math.abs(sh));

/** A Rect stored as four 16-bit values in the order the record keeps them. */
const rect = (l: number, t: number, r: number, b: number): Rect => ({ l, t, r, b });

/** The "WMFC" comments that carry a whole EMF ([MS-WMF] 2.3.6.5), concatenated; undefined when incomplete. */
class EmbeddedEmf {
  private parts: Uint8Array[] = [];
  private size = 0;
  private total = -1;
  broken = false;

  add(v: View): void {
    if (this.broken) return;
    const total = v.u32(34);
    const n = v.u32(26);
    if ((this.total >= 0 && total !== this.total) || total > 64 * 1024 * 1024 || !v.has(38, n)) {
      this.broken = true;
      return;
    }
    this.total = total;
    this.parts.push(v.sub(38, n));
    this.size += n;
  }

  bytes(): Uint8Array | undefined {
    if (this.broken || this.total <= 0 || this.size !== this.total) return undefined;
    const out = new Uint8Array(this.size);
    let p = 0;
    for (const part of this.parts) { out.set(part, p); p += part.length; }
    return out;
  }
}

function readFont(v: View, at: number): FontObject {
  // the face name is ANSI text in the font's own character set, up to 32 bytes
  let end = at + 18;
  while (end < at + 50 && v.has(end, 1) && v.u8(end) !== 0) end++;
  const charSet = v.u8(at + 13);
  return {
    kind: 'font', height: v.i16(at), width: v.i16(at + 2), escapement: v.i16(at + 4), orientation: v.i16(at + 6),
    weight: v.i16(at + 8), italic: v.u8(at + 10) !== 0, underline: v.u8(at + 11) !== 0, strikeOut: v.u8(at + 12) !== 0,
    charSet, quality: v.u8(at + 16), pitchAndFamily: v.u8(at + 17), face: decodeAnsi(v.sub(at + 18, end - at - 18), charSet).trim(),
  };
}

/** A WMF Region object: its scans as rectangles, or its bounding box when the scans do not add up. */
function readRegion(v: View, at: number): Rect[] {
  const scanCount = v.i16(at + 12);
  const bounds = rect(v.i16(at + 16), v.i16(at + 18), v.i16(at + 20), v.i16(at + 22));
  const rects: Rect[] = [];
  let p = at + 24;
  for (let s = 0; s < scanCount; s++) {
    if (!v.has(p, 6)) return [bounds];
    const count = v.u16(p);
    const top = v.i16(p + 2);
    const bottom = v.i16(p + 4);
    if (count % 2 || !v.has(p + 6, count * 2 + 2)) return [bounds];
    for (let k = 0; k < count; k += 2) rects.push(rect(v.i16(p + 6 + k * 2), top, v.i16(p + 8 + k * 2), bottom));
    p += 6 + count * 2 + 2;
  }
  return rects.length || scanCount === 0 ? rects : [bounds];
}

function points(v: View, at: number, n: number): Int32Array {
  if (n < 0 || !v.has(at, n * 4)) throw new MetafileError('malformed', `${n} points do not fit their record`);
  const out = new Int32Array(n * 2);
  for (let i = 0; i < n * 2; i++) out[i] = v.i16(at + i * 2);
  return out;
}

/** One record's parameters (after the 6-byte header) as GDI operations; undefined for records that draw nothing. */
function record(fn: number, v: View, embedded: EmbeddedEmf): MetaRecord | MetaRecord[] | undefined {
  const n = v.size;
  switch (fn) {
    case 0x0102: return { type: 'bkMode', value: v.u16(0) };
    case 0x0103: return { type: 'mapMode', mode: v.u16(0) };
    case 0x0104: return { type: 'rop2', value: v.u16(0) };
    case 0x0106: return { type: 'polyFillMode', value: v.u16(0) };
    case 0x0107: return { type: 'stretchMode', value: v.u16(0) };
    case 0x0108: return { type: 'charExtra', value: v.i16(0) };
    case 0x012e: return { type: 'textAlign', value: v.u16(0) };
    case 0x0149: return { type: 'layout', value: v.u32(0) };
    case 0x0201: return { type: 'bkColor', color: v.u32(0) };
    case 0x0209: return { type: 'textColor', color: v.u32(0) };
    case 0x020b: return { type: 'windowOrg', x: v.i16(2), y: v.i16(0) };
    case 0x020c: return { type: 'windowExt', x: v.i16(2), y: v.i16(0) };
    case 0x020d: return { type: 'viewportOrg', x: v.i16(2), y: v.i16(0) };
    case 0x020e: return { type: 'viewportExt', x: v.i16(2), y: v.i16(0) };
    case 0x020f: return { type: 'offsetWindowOrg', x: v.i16(2), y: v.i16(0) };
    case 0x0211: return { type: 'offsetViewportOrg', x: v.i16(2), y: v.i16(0) };
    case 0x0410: return { type: 'scaleWindowExt', yDenom: v.i16(0), yNum: v.i16(2), xDenom: v.i16(4), xNum: v.i16(6) };
    case 0x0412: return { type: 'scaleViewportExt', yDenom: v.i16(0), yNum: v.i16(2), xDenom: v.i16(4), xNum: v.i16(6) };
    case 0x001e: return { type: 'saveDC' };
    case 0x0127: return { type: 'restoreDC', n: v.i16(0) };
    case 0x0213: return { type: 'lineTo', x: v.i16(2), y: v.i16(0) };
    case 0x0214: return { type: 'moveTo', x: v.i16(2), y: v.i16(0) };
    case 0x0415: return { type: 'clipRect', op: 'exclude', box: rect(v.i16(6), v.i16(4), v.i16(2), v.i16(0)) };
    case 0x0416: return { type: 'clipRect', op: 'intersect', box: rect(v.i16(6), v.i16(4), v.i16(2), v.i16(0)) };
    case 0x0220: return { type: 'offsetClip', x: v.i16(2), y: v.i16(0) };
    case 0x0418: return { type: 'box', shape: 'ellipse', box: rect(v.i16(6), v.i16(4), v.i16(2), v.i16(0)) };
    case 0x041b: return { type: 'box', shape: 'rectangle', box: rect(v.i16(6), v.i16(4), v.i16(2), v.i16(0)) };
    case 0x061c: return { type: 'roundRect', box: rect(v.i16(10), v.i16(8), v.i16(6), v.i16(4)), w: v.i16(2), h: v.i16(0) };
    case 0x041f: return { type: 'pixel', x: v.i16(6), y: v.i16(4), color: v.u32(0) };
    case 0x0817: case 0x081a: case 0x0830:
      return {
        type: 'arc', shape: fn === 0x0817 ? 'arc' : fn === 0x081a ? 'pie' : 'chord',
        box: rect(v.i16(14), v.i16(12), v.i16(10), v.i16(8)), start: [v.i16(6), v.i16(4)], end: [v.i16(2), v.i16(0)],
      };
    case 0x0324: case 0x0325:
      return { type: 'poly', shape: fn === 0x0324 ? 'polygon' : 'polyline', points: points(v, 2, v.i16(0)) };
    case 0x0538: {
      const polys = v.u16(0);
      const counts = new Uint32Array(polys);
      let total = 0;
      for (let i = 0; i < polys; i++) total += counts[i] = v.u16(2 + i * 2);
      return { type: 'polyPoly', shape: 'polygon', counts, points: points(v, 2 + polys * 2, total) };
    }
    case 0x0521: {
      const len = v.i16(0);
      if (len < 0) throw new MetafileError('malformed', 'negative string length');
      const at = 2 + len + (len & 1);
      return { type: 'text', bytes: v.sub(2, len), x: v.i16(at + 2), y: v.i16(at), options: 0 };
    }
    case 0x0a32: {
      const len = v.i16(4);
      const options = v.u16(6);
      if (len < 0) throw new MetafileError('malformed', 'negative string length');
      // the rectangle is there when ETO_OPAQUE or ETO_CLIPPED asks for it
      const box = options & 6 ? rect(v.i16(8), v.i16(10), v.i16(12), v.i16(14)) : undefined;
      const at = box ? 16 : 8;
      const bytes = v.sub(at, Math.min(len, n - at));
      const rec: MetaRecord = { type: 'text', bytes, x: v.i16(2), y: v.i16(0), options };
      if (box) rec.box = box;
      const dxAt = at + len + (len & 1);
      if (len > 0 && v.has(dxAt, len * 2)) {
        const dx = new Int32Array(len);
        for (let i = 0; i < len; i++) dx[i] = v.i16(dxAt + i * 2);
        rec.dx = dx;
      }
      return rec;
    }
    case 0x061d: // PATBLT
      return { type: 'blt', rop: v.u32(0), h: v.i16(4), w: v.i16(6), y: v.i16(8), x: v.i16(10), sx: 0, sy: 0, sw: 0, sh: 0 };
    case 0x0940: case 0x0922: { // DIBBITBLT, BITBLT
      const rop = v.u32(0);
      // without a bitmap, the record has one more (reserved) word: RecordSize == (RecordFunction >> 8) + 3
      if (n === ((fn >> 8) + 3) * 2 - 6) {
        return { type: 'blt', rop, sy: 0, sx: 0, h: v.i16(10), w: v.i16(12), y: v.i16(14), x: v.i16(16), sw: 0, sh: 0 };
      }
      const h = v.i16(8), w = v.i16(10);
      const dib = fn === 0x0940 ? packedDib(v, 16, 0) : bitmap16(v, 16);
      return { type: 'blt', rop, sy: v.i16(4), sx: v.i16(6), sh: h, sw: w, h, w, y: v.i16(12), x: v.i16(14), dib };
    }
    case 0x0b41: case 0x0b23: { // DIBSTRETCHBLT, STRETCHBLT
      const rop = v.u32(0);
      if (n === ((fn >> 8) + 3) * 2 - 6) {
        return { type: 'blt', rop, sh: v.i16(4), sw: v.i16(6), sy: v.i16(8), sx: v.i16(10), h: v.i16(14), w: v.i16(16), y: v.i16(18), x: v.i16(20) };
      }
      const dib = fn === 0x0b41 ? packedDib(v, 20, 0) : bitmap16(v, 20);
      return { type: 'blt', rop, sh: v.i16(4), sw: v.i16(6), sy: v.i16(8), sx: v.i16(10), h: v.i16(12), w: v.i16(14), y: v.i16(16), x: v.i16(18), dib };
    }
    case 0x0f43: { // STRETCHDIB: the source rectangle's origin is the DIB's lower left
      const dib = packedDib(v, 22, v.u16(4));
      const sh = v.i16(6), sw = v.i16(8), sy = v.i16(10), sx = v.i16(12);
      return { type: 'blt', rop: v.u32(0), sh, sw, sy: fromTop(dib, sy, sh), sx, h: v.i16(14), w: v.i16(16), y: v.i16(18), x: v.i16(20), dib };
    }
    case 0x0d33: { // SETDIBTODEV
      const dib = packedDib(v, 18, v.u16(0));
      const h = v.i16(10), w = v.i16(12), sy = v.i16(6), sx = v.i16(8);
      return { type: 'blt', rop: 0xcc0020, h, w, sh: h, sw: w, sy: fromTop(dib, sy, h), sx, y: v.i16(14), x: v.i16(16), dib, device: true };
    }
    // objects
    case 0x02fa: {
      const width = v.i16(2);
      return { type: 'create', index: -1, object: { kind: 'pen', style: v.u16(0), width: Math.abs(width), color: v.u32(6), cosmetic: width === 0 } };
    }
    case 0x02fc:
      return { type: 'create', index: -1, object: { kind: 'brush', style: v.u16(0), color: v.u32(2), hatch: v.u16(6) } };
    case 0x02fb: return { type: 'create', index: -1, object: readFont(v, 0) };
    case 0x00f7: {
      const count = v.u16(2);
      const colors = new Uint32Array(count);
      for (let i = 0; i < count; i++) colors[i] = (v.u8(4 + i * 4) << 16) | (v.u8(5 + i * 4) << 8) | v.u8(6 + i * 4);
      return { type: 'create', index: -1, object: { kind: 'palette', colors } };
    }
    case 0x0142: { // DIBCREATEPATTERNBRUSH
      const style = v.u16(0);
      let dib: Dib;
      let mono = false;
      if (style === BS_PATTERN && v.u32(4) !== 40 && v.u32(4) !== 12 && v.u32(4) !== 108 && v.u32(4) !== 124) {
        dib = bitmap16(v, 4);
        mono = dib.bitCount === 1;
      } else {
        dib = packedDib(v, 4, style === BS_PATTERN ? 0 : v.u16(2));
      }
      const brush: BrushObject = { kind: 'brush', style: BS_DIBPATTERNPT, color: 0, hatch: 0, dib };
      if (mono) brush.mono = true;
      return { type: 'create', index: -1, object: brush };
    }
    case 0x01f9: { // CREATEPATTERNBRUSH: a Bitmap16 header, 18 reserved bytes, then the bits
      const head = bitmap16(v, 0);
      const bits = v.sub(32, Math.max(0, n - 32));
      const dib: Dib = { ...head, bits };
      return { type: 'create', index: -1, object: { kind: 'brush', style: BS_DIBPATTERNPT, color: 0, hatch: 0, dib, mono: dib.bitCount === 1 } };
    }
    case 0x06ff: return { type: 'create', index: -1, object: { kind: 'region', rects: readRegion(v, 0) } };
    case 0x012d: case 0x0234: return { type: 'select', index: v.u16(0) };
    case 0x01f0: return { type: 'delete', index: v.u16(0) };
    case 0x012c: return { type: 'selectClipRegion', index: v.u16(0) };
    case 0x0228: return { type: 'region', op: 'fill', region: v.u16(0), brush: v.u16(2) };
    case 0x012b: return { type: 'region', op: 'paint', region: v.u16(0) };
    case 0x012a: return { type: 'region', op: 'invert', region: v.u16(0) };
    case 0x0429: return { type: 'region', op: 'frame', region: v.u16(0), brush: v.u16(2), h: v.i16(4), w: v.i16(6) };
    case 0x0626: // ESCAPE: only the embedded EMF matters
      if (v.has(0, 38) && v.u16(0) === 0x000f && v.u32(4) === 0x43464d57 && v.u32(8) === 1) embedded.add(v);
      return undefined;
    default:
      // palettes (realize, entries, resize, animate), relabs, mapper flags, text justification, flood fills
      return undefined;
  }
}

/** Objects a record creates take a slot, so a record that cannot be read still has to create something. */
const placeholder = (fn: number): GdiObject | undefined =>
  fn === 0x02fa || fn === 0x02fc || fn === 0x02fb || fn === 0x00f7 || fn === 0x0142 || fn === 0x01f9 || fn === 0x06ff ? { kind: 'other' } : undefined;

/** Bounds of every coordinate the drawing records use, for a WMF with neither a placeable header nor a window. */
function drawingBounds(records: MetaRecord[]): Rect | undefined {
  let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
  const add = (x: number, y: number) => { l = Math.min(l, x); t = Math.min(t, y); r = Math.max(r, x); b = Math.max(b, y); };
  for (const rec of records) {
    switch (rec.type) {
      case 'moveTo': case 'lineTo': case 'text': add(rec.x, rec.y); break;
      case 'poly': case 'polyPoly': for (let i = 0; i + 1 < rec.points.length; i += 2) add(rec.points[i]!, rec.points[i + 1]!); break;
      case 'box': case 'roundRect': case 'arc': add(rec.box.l, rec.box.t); add(rec.box.r, rec.box.b); break;
      case 'blt': add(rec.x, rec.y); add(rec.x + rec.w, rec.y + rec.h); break;
      default: break;
    }
  }
  return l < r && t < b ? { l, t, r, b } : undefined;
}

/**
 * Parse a WMF. A WMF that carries a complete embedded EMF (the comments GDI writes when converting an EMF) is
 * returned through `emf` instead, since its records are the original picture.
 */
export function parseWmf(bytes: Uint8Array, emf?: (bytes: Uint8Array) => Metafile | undefined): Metafile {
  const at = wmfHeaderAt(bytes);
  if (at < 0) throw new MetafileError('malformed', 'not a WMF');
  const v = new View(bytes);
  let placeable: Rect | undefined;
  let inch = 0;
  if (at === 22) {
    placeable = rect(v.i16(6), v.i16(8), v.i16(10), v.i16(12));
    inch = v.u16(14);
  }
  const records: MetaRecord[] = [];
  const embedded = new EmbeddedEmf();
  let skipped = 0;
  let p = at + v.u16(at + 2) * 2;
  // the window set before anything is drawn is the picture's frame when there is no placeable header
  let org: [number, number] = [0, 0];
  let ext: [number, number] | undefined;
  let drawn = false;
  while (p + 6 <= bytes.length) {
    if (records.length >= MAX_RECORDS) throw stop(`more than ${MAX_RECORDS} records`);
    const size = v.u32(p) * 2;
    const fn = v.u16(p + 4);
    if (size < 6 || p + size > bytes.length) break;        // a record that does not fit ends the file
    if (fn === 0) { records.push({ type: 'eof' }); break; }
    let out: MetaRecord | MetaRecord[] | undefined;
    try {
      out = record(fn, v.view(p + 6, size - 6), embedded);
    } catch (e) {
      if (!(e instanceof MetafileError) || stops(e)) throw e;
      skipped++;
      const object = placeholder(fn);
      out = object ? { type: 'create', index: -1, object } : undefined;
    }
    for (const rec of out === undefined ? [] : Array.isArray(out) ? out : [out]) {
      if (!drawn) {
        if (rec.type === 'windowOrg') org = [rec.x, rec.y];
        else if (rec.type === 'windowExt' && rec.x !== 0 && rec.y !== 0) ext = [rec.x, rec.y];
        else if (DRAWING.has(rec.type)) drawn = true;
      }
      records.push(rec);
    }
    p += size;
  }
  const inner = embedded.bytes();
  if (inner && emf) {
    try {
      const m = emf(inner);
      if (m) return m;
    } catch (e) {
      if (!(e instanceof MetafileError) || stops(e)) throw e;
    }
  }
  if (records.length === 0) throw new MetafileError('malformed', 'WMF without records');
  const box = placeable && placeable.r !== placeable.l && placeable.b !== placeable.t
    ? rect(Math.min(placeable.l, placeable.r), Math.min(placeable.t, placeable.b), Math.max(placeable.l, placeable.r), Math.max(placeable.t, placeable.b))
    : undefined;
  const frame = box ?? (ext ? rect(org[0], org[1], org[0] + ext[0], org[1] + ext[1]) : drawingBounds(records)) ?? rect(0, 0, 1000, 1000);
  const fw = Math.abs(frame.r - frame.l);
  const fh = Math.abs(frame.b - frame.t);
  // units per inch: the placeable header says; without one, assume a large window is in twips (1440 per inch)
  // and a small one in 96-dpi pixels, which is what writers of such files use
  const units = box && inch > 0 ? inch : Math.max(fw, fh) > 4000 ? 1440 : 96;
  return { format: 'wmf', frame, width: (fw / units) * 96, height: (fh / units) * 96, dpmmX: units / 25.4, dpmmY: units / 25.4, records, plus: 'none', skipped };
}

const DRAWING = new Set<MetaRecord['type']>(['poly', 'polyPoly', 'box', 'roundRect', 'arc', 'text', 'blt', 'lineTo', 'pixel', 'region']);
