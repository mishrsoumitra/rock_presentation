// packages/metafile/src/player.ts  -- plays a parsed metafile's GDI records onto a Canvas 2D context
//
// The playback model of [MS-EMF] 3.1 and [MS-WMF] 3.1: a device context of transforms, selected objects and modes,
// saved and restored as a whole, an object table, a path bracket and a clipping region. Everything is computed in
// output coordinates (the destination rectangle's space) and drawn with the context's transform left as the
// caller set it, so a metafile can be drawn straight into a slide.
import { combinePath, combineRects, RGN_AND, RGN_COPY, RGN_DIFF, translateClip, type Clip } from './clip.ts';
import { decodeDib, type DecodeOptions, type Pixels } from './dib.ts';
import { isEmf, parseEmf, STOCK } from './emf.ts';
import { MetafileError, stop, stops } from './errors.ts';
import { apply, arcCurves, axisAligned, compose, det, IDENTITY, normalize, Path, usable } from './geom.ts';
import { playPlus, plusState, type PlusState } from './plus.ts';
import { cssFont, decodeAnsiUnits, symbolTable, symbolText, SYMBOL_CHARSET } from './text.ts';
import type { BltRecord, BrushObject, Dib, FontObject, GdiObject, Matrix, Metafile, MetaRecord, PaletteObject, PenObject, Rect, TextRecord } from './types.ts';
import { parseWmf } from './wmf.ts';

/** The subset of CanvasRenderingContext2D (and OffscreenCanvasRenderingContext2D) the players use. */
export interface MetaCtx {
  save(): void;
  restore(): void;
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  beginPath(): void;
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  bezierCurveTo(cp1x: number, cp1y: number, cp2x: number, cp2y: number, x: number, y: number): void;
  closePath(): void;
  rect(x: number, y: number, w: number, h: number): void;
  fill(fillRule?: CanvasFillRule): void;
  stroke(): void;
  clip(fillRule?: CanvasFillRule): void;
  fillRect(x: number, y: number, w: number, h: number): void;
  fillText(text: string, x: number, y: number): void;
  measureText(text: string): { width: number; fontBoundingBoxAscent?: number; fontBoundingBoxDescent?: number };
  drawImage(image: CanvasImageSource, sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number): void;
  createPattern(image: CanvasImageSource, repetition: string | null): CanvasPattern | null;
  createLinearGradient(x0: number, y0: number, x1: number, y1: number): CanvasGradient;
  createRadialGradient(x0: number, y0: number, r0: number, x1: number, y1: number, r1: number): CanvasGradient;
  setLineDash(segments: number[]): void;
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  lineCap: CanvasLineCap;
  lineJoin: CanvasLineJoin;
  miterLimit: number;
  globalAlpha: number;
  globalCompositeOperation: GlobalCompositeOperation;
  imageSmoothingEnabled: boolean;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
}

/** A canvas the players can make bitmaps in: the host's canvas type (OffscreenCanvas in browsers and workers). */
export interface ScratchCanvas {
  getContext(contextId: '2d'): { createImageData(sw: number, sh: number): ImageData; putImageData(data: ImageData, dx: number, dy: number): void } | null;
}

export interface DrawOptions {
  /** Makes canvases for bitmaps and patterns; without it, bitmaps are skipped and patterns fill with their colour. */
  createCanvas?: (w: number, h: number) => ScratchCanvas;
  /** JPEG and PNG pictures inside the metafile, decoded beforehand by `decodeImages`, by key (see `dibImageKey`). */
  images?: ReadonlyMap<string, CanvasImageSource>;
  /**
   * Destination units per pixel of the picture at its natural size: how wide a hairline is, and how large a hatch
   * cell. Defaults to the destination width over the natural width.
   */
  pixel?: number;
  /** Which records to play in an EMF that has EMF+ records: by default EMF for dual files and EMF+ for EMF+ only. */
  plus?: 'auto' | 'emf' | 'emf+';
  /** Work limit (records, points, characters and bitmap pixels together); playback stops with a 'limit' error. */
  maxWork?: number;
  /**
   * The destination's units are device pixels: put the edges of axis-aligned fills and bitmaps on whole pixels,
   * as GDI does, so that neighbouring fills meet without antialiased seams.
   */
  snap?: boolean;
}

export interface Dest { x: number; y: number; w: number; h: number; }

/** Whether the EMF+ records are the ones played (see DrawOptions.plus); EMF records then play only after GetDC. */
export const playsPlus = (mf: Pick<Metafile, 'plus'>, mode: DrawOptions['plus'] = 'auto'): boolean =>
  mf.plus !== 'none' && (mode === 'emf+' || (mode === 'auto' && mf.plus === 'only'));

export const DEFAULT_MAX_WORK = 200_000_000;

/** How a JPEG or PNG bitmap inside an EMF record is named in `DrawOptions.images`. */
export const dibImageKey = (bits: Uint8Array) => `dib:${bits.byteOffset}:${bits.length}`;
const MAX_SAVE_DEPTH = 4096;
const MAX_OBJECTS = 65_536;

const MM_TEXT = 1, MM_ISOTROPIC = 7, MM_ANISOTROPIC = 8;
const OPAQUE = 2;
const WINDING = 2;
const PS_DASH = 1, PS_DOT = 2, PS_DASHDOT = 3, PS_DASHDOTDOT = 4, PS_NULL = 5, PS_INSIDEFRAME = 6, PS_USERSTYLE = 7, PS_ALTERNATE = 8;
const PS_GEOMETRIC = 0x10000;
const BS_SOLID = 0, BS_NULL = 1, BS_HATCHED = 2;
const TA_UPDATECP = 1, TA_RIGHT = 2, TA_CENTER = 6, TA_BOTTOM = 8, TA_BASELINE = 24;
const ETO_OPAQUE = 2, ETO_CLIPPED = 4, ETO_GLYPH_INDEX = 0x10;
const AD_CLOCKWISE = 2;

// ---------------------------------------------------------------------------------------------------------- state

const brush = (color: number, style = BS_SOLID): BrushObject => ({ kind: 'brush', style, color, hatch: 0 });
const pen = (color: number, style = 0): PenObject => ({ kind: 'pen', style, width: 0, color, cosmetic: true });
const font = (face: string, height: number, weight = 400, pitchAndFamily = 0): FontObject => ({
  kind: 'font', height, width: 0, escapement: 0, orientation: 0, weight, italic: false, underline: false, strikeOut: false,
  charSet: 0, quality: 0, pitchAndFamily, face,
});

/** GetStockObject's objects, by index ([MS-EMF] 2.1.31). */
const STOCK_OBJECTS: (GdiObject | undefined)[] = [
  brush(0xffffff), brush(0xc0c0c0), brush(0x808080), brush(0x404040), brush(0), brush(0, BS_NULL),
  pen(0xffffff), pen(0), pen(0, PS_NULL), undefined,
  font('Courier New', 16, 400, 0x31), font('Courier New', 16, 400, 0x31), font('Arial', 16), font('System', 16, 700),
  font('System', 16, 700), { kind: 'palette', colors: new Uint32Array(0) }, font('Courier New', 16, 400, 0x31),
  font('Microsoft Sans Serif', -11), brush(0xffffff), pen(0),
];

/** The playback device context ([MS-EMF] 3.1.1.2): copied whole by SaveDC. */
interface DC {
  mapMode: number;
  wox: number; woy: number; wex: number; wey: number;
  vox: number; voy: number; vex: number; vey: number;
  world: Matrix;
  pen: PenObject;
  brush: BrushObject;
  font: FontObject;
  palette: PaletteObject | undefined;
  bkMode: number;
  bkColor: number;
  textColor: number;
  textAlign: number;
  fillMode: number;
  rop2: number;
  miterLimit: number;
  arcDir: number;
  charExtra: number;
  clip: Clip;
  /** Current position, logical units. */
  cx: number;
  cy: number;
  /** Brush origin, device units. */
  bx: number;
  by: number;
}

/** The object table: EMF records name their slots, WMF objects take the lowest free one ([MS-WMF] 3.1.4.1). */
class ObjectTable {
  private slots = new Map<number, GdiObject>();
  private free: number[] = [];          // min-heap of freed slots below `next`
  private next = 0;

  create(index: number, o: GdiObject): void {
    if (index < 0) {
      index = this.next;
      while (this.free.length) {
        const i = heapPop(this.free);
        if (!this.slots.has(i)) { index = i; break; }
      }
    }
    if (index >= MAX_OBJECTS) throw stop(`more than ${MAX_OBJECTS} objects`);
    if (index >= this.next) this.next = index + 1;
    this.slots.set(index, o);
  }

  get(index: number): GdiObject | undefined { return this.slots.get(index); }

  delete(index: number): void {
    if (this.slots.delete(index)) heapPush(this.free, index);
  }
}

function heapPush(h: number[], v: number): void {
  h.push(v);
  for (let i = h.length - 1; i > 0;) {
    const p = (i - 1) >> 1;
    if (h[p]! <= h[i]!) break;
    [h[p], h[i]] = [h[i]!, h[p]!];
    i = p;
  }
}

function heapPop(h: number[]): number {
  const top = h[0]!;
  const last = h.pop()!;
  if (h.length) {
    h[0] = last;
    for (let i = 0; ;) {
      const l = i * 2 + 1, r = l + 1;
      let m = i;
      if (l < h.length && h[l]! < h[m]!) m = l;
      if (r < h.length && h[r]! < h[m]!) m = r;
      if (m === i) break;
      [h[m], h[i]] = [h[i]!, h[m]!];
      i = m;
    }
  }
  return top;
}

// ---------------------------------------------------------------------------------------------------------- colours

/** 0xRRGGBB as a CSS colour. */
const css = (rgb: number, alpha = 1) => alpha >= 1 ? `rgb(${(rgb >>> 16) & 255},${(rgb >>> 8) & 255},${rgb & 255})`
  : `rgba(${(rgb >>> 16) & 255},${(rgb >>> 8) & 255},${rgb & 255},${+alpha.toFixed(4)})`;

/** Hatch patterns, 8 rows of 8 bits each, most significant bit leftmost (as GDI draws them). */
const HATCHES = [
  [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0x00], // HS_HORIZONTAL
  [0x08, 0x08, 0x08, 0x08, 0x08, 0x08, 0x08, 0x08], // HS_VERTICAL
  [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01], // HS_FDIAGONAL
  [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80], // HS_BDIAGONAL
  [0x08, 0x08, 0x08, 0x08, 0x08, 0x08, 0xff, 0x08], // HS_CROSS
  [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81], // HS_DIAGCROSS
];

/** Truth table of a ternary raster operation, for one destination bit d, as 8-bit masks of p and s ([MS-WMF] 2.1.1.31). */
function ropBits(rop: number, d: number, p: number, s: number): number {
  const t = (ps: number) => ((rop >> ((ps << 1) | d)) & 1 ? 0xff : 0);
  const np = ~p & 0xff, ns = ~s & 0xff;
  return (np & ns & t(0)) | (np & s & t(1)) | (p & ns & t(2)) | (p & s & t(3));
}

/**
 * How a raster operation with a source bitmap is drawn: as a transform of the source pixels and a compositing
 * operation. The common operations are exact on black-and-white operands (the mask idioms); anything else is
 * approximated per pixel by evaluating the operation against black and white destinations and blending.
 */
function ropPlan(rop: number): { pixels?: 'invert' | 'and-pattern' | 'generic'; op: GlobalCompositeOperation } {
  switch (rop) {
    case 0xcc: return { op: 'source-over' };                       // SRCCOPY
    case 0x33: return { pixels: 'invert', op: 'source-over' };     // NOTSRCCOPY
    case 0x88: return { op: 'multiply' };                          // SRCAND
    case 0xee: return { op: 'lighten' };                           // SRCPAINT
    case 0x66: return { op: 'difference' };                        // SRCINVERT
    case 0xbb: return { pixels: 'invert', op: 'lighten' };         // MERGEPAINT
    case 0x22: return { pixels: 'invert', op: 'multiply' };        // ~S & D
    case 0xc0: return { pixels: 'and-pattern', op: 'source-over' }; // MERGECOPY
    case 0x44: case 0x11: return { op: 'source-over' };            // SRCERASE, NOTSRCERASE (approximated)
    default: return { pixels: 'generic', op: 'source-over' };
  }
}

// ---------------------------------------------------------------------------------------------------------- player

const isDrawing = (r: MetaRecord) => r.type !== 'plus' && r.type !== 'eof';

export class Player {
  readonly ctx: MetaCtx;
  readonly mf: Metafile;
  readonly opts: DrawOptions;
  readonly dest: Dest;
  /** Device units to output units. */
  readonly d2o: Matrix;
  /** Output units per pixel of the picture at natural size. */
  readonly pixel: number;
  dc: DC;
  private saved: DC[] = [];
  private objects = new ObjectTable();
  private mat: Matrix = IDENTITY;
  private dirty = true;
  /** The open path bracket, then the closed path waiting for FILLPATH and friends. */
  private bracket: Path | undefined;
  private path: Path | undefined;
  private pathText: (() => void)[] = [];
  private applied: Clip | undefined;
  private work = 0;
  private readonly maxWork: number;
  private patterns = new Map<string, { pattern?: CanvasPattern; average?: number }>();
  private bitmaps = new WeakMap<Dib, Map<string, CanvasImageSource | null>>();
  private fontRatios = new Map<string, number>();
  private plus: PlusState | undefined;
  /** Metafiles inside EMF+ images are drawn one level deep, never deeper. */
  private readonly depth: number;

  constructor(ctx: MetaCtx, mf: Metafile, dest: Dest, opts: DrawOptions = {}, depth = 0) {
    this.ctx = ctx;
    this.depth = depth;
    this.mf = mf;
    this.opts = opts;
    this.dest = dest;
    const f = mf.frame;
    const kx = dest.w / (f.r - f.l || 1), ky = dest.h / (f.b - f.t || 1);
    this.d2o = [kx, 0, 0, ky, dest.x - f.l * kx, dest.y - f.t * ky];
    const natural = opts.pixel ?? (mf.width > 0 ? Math.abs(dest.w) / mf.width : 1);
    // a hairline never takes more than a 64th of the picture, whatever the header claims
    this.pixel = Math.max(1e-6, Math.min(natural, Math.max(Math.abs(dest.w), Math.abs(dest.h)) / 64));
    this.maxWork = opts.maxWork ?? DEFAULT_MAX_WORK;
    const wmf = mf.format === 'wmf';
    // a WMF's logical units start out as its frame's (see Metafile.frame); an EMF starts in MM_TEXT
    const w = f.r - f.l || 1, h = f.b - f.t || 1;
    this.dc = {
      mapMode: wmf ? MM_ANISOTROPIC : MM_TEXT,
      wox: wmf ? f.l : 0, woy: wmf ? f.t : 0, wex: wmf ? w : 1, wey: wmf ? h : 1,
      vox: wmf ? f.l : 0, voy: wmf ? f.t : 0, vex: wmf ? w : 1, vey: wmf ? h : 1,
      world: IDENTITY, pen: pen(0), brush: brush(0xffffff), font: STOCK_OBJECTS[13] as FontObject, palette: undefined,
      bkMode: OPAQUE, bkColor: 0xffffff, textColor: 0, textAlign: 0, fillMode: 1, rop2: 13,
      miterLimit: 10, arcDir: 1, charExtra: 0, clip: [], cx: 0, cy: 0, bx: 0, by: 0,
    };
  }

  /** Work done so far, in the units of `DrawOptions.maxWork`. */
  get used(): number { return this.work; }

  /** An EMF+ image that is itself a metafile: drawn into the parallelogram (output units), one level deep. */
  nested(bytes: Uint8Array, c: [number, number, number, number, number, number]): void {
    if (this.depth > 0) return;
    let mf: Metafile;
    try {
      mf = isEmf(bytes) ? parseEmf(bytes) : parseWmf(bytes);
    } catch (e) {
      if (e instanceof MetafileError && !stops(e)) return;
      throw e;
    }
    const w = Math.hypot(c[2] - c[0], c[3] - c[1]);
    if (!(w > 0) || !(Math.hypot(c[4] - c[0], c[5] - c[1]) > 0)) return;
    const ctx = this.ctx;
    ctx.save();
    ctx.transform(c[2] - c[0], c[3] - c[1], c[4] - c[0], c[5] - c[1], c[0], c[1]);
    // the inner picture is drawn in a unit square: its units are not pixels, so nothing is snapped to them
    const opts = { ...this.opts, pixel: this.pixel / w, maxWork: this.maxWork - this.work, snap: false };
    const inner = new Player(ctx, mf, { x: 0, y: 0, w: 1, h: 1 }, opts, this.depth + 1);
    try {
      inner.run();
    } finally {
      ctx.restore();
      this.spend(inner.used);
    }
  }

  /** Count work, and stop a hostile file that asks for too much. */
  spend(n: number): void {
    this.work += n;
    if (this.work > this.maxWork) throw stop(`more than ${this.maxWork} units of drawing work`);
  }

  run(): void {
    const { ctx, dest } = this;
    const plusOnly = playsPlus(this.mf, this.opts.plus);
    let gdi = !plusOnly;
    ctx.save();
    ctx.beginPath();
    ctx.rect(dest.x, dest.y, dest.w, dest.h);
    ctx.clip();
    ctx.save();               // the level the clipping region is set in
    try {
      const recs = this.mf.records;
      for (let i = 0; i < recs.length; i++) {
        const rec = recs[i]!;
        this.spend(1);
        if (rec.type === 'eof') break;
        if (rec.type === 'plus') {
          if (!plusOnly) continue;
          this.plus ??= plusState(this.mf);
          gdi = playPlus(this, this.plus, rec.records);
          continue;
        }
        if (!gdi && isDrawing(rec)) continue;
        if (rec.type === 'blt' && rec.rop === 0x660046 && this.xorMask(recs, i)) { i += 2; continue; }
        this.play(rec);
      }
    } finally {
      ctx.restore();
      ctx.restore();
    }
  }

  // ------------------------------------------------------------------------------------------------ transforms

  /** Logical units to output units, for the current mapping mode, window, viewport and world transform. */
  matrix(): Matrix {
    if (!this.dirty) return this.mat;
    const dc = this.dc;
    let sx: number, sy: number;
    if (dc.mapMode === MM_ISOTROPIC || dc.mapMode === MM_ANISOTROPIC) {
      sx = dc.vex / dc.wex;
      sy = dc.vey / dc.wey;
      if (dc.mapMode === MM_ISOTROPIC) {
        const s = Math.min(Math.abs(sx), Math.abs(sy));
        sx = Math.sign(sx) * s;
        sy = Math.sign(sy) * s;
      }
    } else {
      [sx, sy] = this.fixedScale(dc.mapMode);
    }
    const page: Matrix = [sx, 0, 0, sy, dc.vox - dc.wox * sx, dc.voy - dc.woy * sy];
    const m = compose(this.d2o, compose(page, dc.world));
    this.mat = usable(m) ? m : this.d2o;
    this.dirty = false;
    return this.mat;
  }

  /** Device units per logical unit in the fixed mapping modes (y grows upward in the metric ones). */
  private fixedScale(mode: number): [number, number] {
    const { dpmmX: x, dpmmY: y } = this.mf;
    switch (mode) {
      case 2: return [x / 10, -y / 10];                 // MM_LOMETRIC, 0.1 mm
      case 3: return [x / 100, -y / 100];               // MM_HIMETRIC, 0.01 mm
      case 4: return [x * 0.254, -y * 0.254];           // MM_LOENGLISH, 0.01 inch
      case 5: return [x * 0.0254, -y * 0.0254];         // MM_HIENGLISH, 0.001 inch
      case 6: return [(x * 25.4) / 1440, (-y * 25.4) / 1440]; // MM_TWIPS
      default: return [1, 1];                            // MM_TEXT
    }
  }

  private setMapMode(mode: number): void {
    const dc = this.dc;
    if (mode < 1 || mode > 8 || mode === dc.mapMode) return;
    if ((mode === MM_ISOTROPIC || mode === MM_ANISOTROPIC) && dc.mapMode !== MM_ISOTROPIC && dc.mapMode !== MM_ANISOTROPIC) {
      // switching from a fixed mode keeps its scale as the starting extents
      const [sx, sy] = this.fixedScale(dc.mapMode);
      dc.wex = 1; dc.wey = 1; dc.vex = sx; dc.vey = sy;
    }
    dc.mapMode = mode;
    this.dirty = true;
  }

  private get scalable(): boolean { return this.dc.mapMode === MM_ISOTROPIC || this.dc.mapMode === MM_ANISOTROPIC; }

  /** A logical point in output units. */
  pt(x: number, y: number): [number, number] { return apply(this.matrix(), x, y); }

  /** Output units per logical unit along x: how pen widths and dash lengths scale. */
  private scaleX(): number {
    const m = this.matrix();
    return Math.hypot(m[0], m[1]);
  }

  // ------------------------------------------------------------------------------------------------ records

  private play(rec: MetaRecord): void {
    const dc = this.dc;
    switch (rec.type) {
      case 'mapMode': this.setMapMode(rec.mode); return;
      case 'windowOrg': dc.wox = rec.x; dc.woy = rec.y; this.dirty = true; return;
      case 'viewportOrg': dc.vox = rec.x; dc.voy = rec.y; this.dirty = true; return;
      case 'offsetWindowOrg': dc.wox += rec.x; dc.woy += rec.y; this.dirty = true; return;
      case 'offsetViewportOrg': dc.vox += rec.x; dc.voy += rec.y; this.dirty = true; return;
      case 'windowExt':
        if (this.scalable && rec.x && rec.y) { dc.wex = rec.x; dc.wey = rec.y; this.dirty = true; }
        return;
      case 'viewportExt':
        if (this.scalable && rec.x && rec.y) { dc.vex = rec.x; dc.vey = rec.y; this.dirty = true; }
        return;
      case 'scaleWindowExt': case 'scaleViewportExt': {
        if (!this.scalable || !rec.xDenom || !rec.yDenom) return;
        const w = rec.type === 'scaleWindowExt';
        const x = (w ? dc.wex : dc.vex) * rec.xNum / rec.xDenom, y = (w ? dc.wey : dc.vey) * rec.yNum / rec.yDenom;
        if (!x || !y) return;
        if (w) { dc.wex = x; dc.wey = y; } else { dc.vex = x; dc.vey = y; }
        this.dirty = true;
        return;
      }
      case 'worldTransform': if (usable(rec.m)) { dc.world = rec.m; this.dirty = true; } return;
      case 'modifyWorldTransform': {
        // MWT_IDENTITY, MWT_LEFTMULTIPLY (the record's transform first), MWT_RIGHTMULTIPLY, MWT_SET
        const m = rec.mode === 1 ? IDENTITY : rec.mode === 2 ? compose(dc.world, rec.m) : rec.mode === 3 ? compose(rec.m, dc.world)
          : rec.mode === 4 ? rec.m : undefined;
        if (m && usable(m)) { dc.world = m; this.dirty = true; }
        return;
      }
      case 'saveDC':
        if (this.saved.length >= MAX_SAVE_DEPTH) throw stop(`SaveDC nested more than ${MAX_SAVE_DEPTH} deep`);
        this.saved.push({ ...dc });
        return;
      case 'restoreDC': {
        const target = rec.n < 0 ? this.saved.length + rec.n : rec.n - 1;
        if (target < 0 || target >= this.saved.length) return;
        this.dc = this.saved[target]!;
        this.saved.length = target;
        this.dirty = true;
        return;
      }
      case 'bkMode': dc.bkMode = rec.value; return;
      case 'textAlign': dc.textAlign = rec.value; return;
      case 'polyFillMode': dc.fillMode = rec.value; return;
      case 'rop2': dc.rop2 = rec.value; return;
      case 'arcDirection': dc.arcDir = rec.value; return;
      case 'miterLimit': if (rec.value >= 1) dc.miterLimit = rec.value; return;
      case 'charExtra': dc.charExtra = rec.value; return;
      case 'layout': return;
      case 'bkColor': dc.bkColor = rec.color; return;
      case 'textColor': dc.textColor = rec.color; return;
      case 'brushOrg': dc.bx = rec.x; dc.by = rec.y; return;
      case 'create': this.objects.create(rec.index, rec.object); return;
      case 'select': this.select(rec.index); return;
      case 'delete': this.objects.delete(rec.index); return;
      case 'moveTo':
        dc.cx = rec.x; dc.cy = rec.y;
        if (this.bracket) this.bracket.moveTo(...this.pt(rec.x, rec.y));
        return;
      case 'lineTo': {
        const p = this.figure();
        p.lineTo(...this.pt(rec.x, rec.y));
        dc.cx = rec.x; dc.cy = rec.y;
        if (!this.bracket) this.emit(p, false);
        return;
      }
      case 'poly': this.poly(rec.shape, rec.points); return;
      case 'polyPoly': {
        const path = new Path();
        const xy = rec.points;
        let at = 0;
        for (const n of rec.counts) {
          const pts: number[] = [];
          for (let k = 0; k < n; k++) pts.push(...this.pt(xy[(at + k) * 2]!, xy[(at + k) * 2 + 1]!));
          path.polygon(pts, rec.shape === 'polygon');
          at += n;
        }
        this.emit(path, rec.shape === 'polygon');
        return;
      }
      case 'polyDraw': this.polyDraw(rec.points, rec.kinds); return;
      case 'box': this.box(rec.shape, rec.box); return;
      case 'roundRect': this.roundRect(rec.box, rec.w, rec.h); return;
      case 'arc': this.arc(rec.shape, rec.box, rec.start, rec.end); return;
      case 'angleArc': this.angleArc(rec.x, rec.y, rec.r, rec.start, rec.sweep); return;
      case 'pixel': {
        const [x, y] = this.pt(rec.x, rec.y);
        this.syncClip();
        this.ctx.globalCompositeOperation = 'source-over';
        this.ctx.fillStyle = css(this.rgb(rec.color));
        this.ctx.fillRect(x, y, this.pixel, this.pixel);
        return;
      }
      case 'path': this.pathOp(rec.op); return;
      case 'gradient': this.gradient(rec.mode, rec.vertices, rec.colors, rec.indexes); return;
      case 'region': this.region(rec); return;
      case 'text': this.text(rec); return;
      case 'blt': this.blt(rec); return;
      case 'clipRect': this.clipRect(rec.op, rec.box); return;
      case 'clipRegion':
        dc.clip = combineRects(dc.clip, rec.mode, rec.rects && rec.rects.map(r => this.devRect(r)));
        return;
      case 'selectClipRegion': {
        const o = this.objects.get(rec.index);
        dc.clip = o?.kind === 'region' ? combineRects([], RGN_COPY, o.rects.map(r => this.logicalRect(r))) : [];
        return;
      }
      case 'clipPath': {
        const path = this.path;
        this.path = undefined;
        if (path) dc.clip = combinePath(dc.clip, rec.mode, path, dc.fillMode === WINDING ? 'nonzero' : 'evenodd');
        return;
      }
      case 'offsetClip': {
        const m = this.matrix();
        dc.clip = translateClip(dc.clip, m[0] * rec.x + m[2] * rec.y, m[1] * rec.x + m[3] * rec.y);
        return;
      }
      default: return;
    }
  }

  private select(index: number): void {
    const dc = this.dc;
    const o = this.object(index);
    if (!o) return;
    switch (o.kind) {
      case 'pen': dc.pen = o; return;
      case 'brush': dc.brush = o; return;
      case 'font': dc.font = o; return;
      case 'palette': dc.palette = o; return;
      // selecting a region selects it as the clipping region (logical units here: WMF regions)
      case 'region': dc.clip = combineRects([], RGN_COPY, o.rects.map(r => this.logicalRect(r))); return;
      default: return;
    }
  }

  // ------------------------------------------------------------------------------------------------ shapes

  /** The path that a "-To" drawing continues: the open bracket, or a new path starting at the current position. */
  private figure(): Path {
    const at = this.pt(this.dc.cx, this.dc.cy);
    if (this.bracket) {
      if (!this.bracket.open || this.bracket.x !== at[0] || this.bracket.y !== at[1]) this.bracket.moveTo(...at);
      return this.bracket;
    }
    const p = new Path();
    p.moveTo(...at);
    return p;
  }

  private poly(shape: 'polyline' | 'polylineTo' | 'polygon' | 'bezier' | 'bezierTo', xy: Int32Array): void {
    const n = xy.length >> 1;
    this.spend(n);
    const to = shape === 'polylineTo' || shape === 'bezierTo';
    if (n === 0 || (!to && n < 2)) return;
    const path = to ? this.figure() : new Path();
    const bezier = shape === 'bezier' || shape === 'bezierTo';
    let i = 0;
    if (!to) { path.moveTo(...this.pt(xy[0]!, xy[1]!)); i = 1; }
    if (bezier) {
      for (; i + 2 < n; i += 3) {
        const [a, b] = this.pt(xy[i * 2]!, xy[i * 2 + 1]!);
        const [c, d] = this.pt(xy[i * 2 + 2]!, xy[i * 2 + 3]!);
        const [e, f] = this.pt(xy[i * 2 + 4]!, xy[i * 2 + 5]!);
        path.curveTo(a, b, c, d, e, f);
      }
    } else {
      for (; i < n; i++) path.lineTo(...this.pt(xy[i * 2]!, xy[i * 2 + 1]!));
    }
    if (shape === 'polygon') path.close();
    if (to) {
      const last = bezier ? Math.min(n, Math.floor(n / 3) * 3) : n;
      if (last > 0) { this.dc.cx = xy[last * 2 - 2]!; this.dc.cy = xy[last * 2 - 1]!; }
      if (this.bracket) return;
    }
    this.emit(path, shape === 'polygon');
  }

  private polyDraw(xy: Int32Array, kinds: Uint8Array): void {
    const n = Math.min(xy.length >> 1, kinds.length);
    this.spend(n);
    const path = this.bracket ?? new Path();
    if (!this.bracket) path.moveTo(...this.pt(this.dc.cx, this.dc.cy));
    for (let i = 0; i < n; i++) {
      const k = kinds[i]! & ~1;
      const p = this.pt(xy[i * 2]!, xy[i * 2 + 1]!);
      if (k === 6) path.moveTo(...p);
      else if (k === 4 && i + 2 < n) {
        const q = this.pt(xy[i * 2 + 2]!, xy[i * 2 + 3]!);
        const r = this.pt(xy[i * 2 + 4]!, xy[i * 2 + 5]!);
        path.curveTo(p[0], p[1], q[0], q[1], r[0], r[1]);
        i += 2;
      } else path.lineTo(...p);
      if (kinds[i]! & 1) path.close();
      this.dc.cx = xy[i * 2]!; this.dc.cy = xy[i * 2 + 1]!;
    }
    if (!this.bracket) this.emit(path, false);
  }

  /** A bounding box, shrunk by half the pen for PS_INSIDEFRAME. */
  private frameBox(box: Rect): Rect {
    const b = normalize(box);
    const p = this.dc.pen;
    if ((p.style & 0xf) !== PS_INSIDEFRAME || p.cosmetic || p.width <= 1) return b;
    const d = Math.min(p.width / 2, (b.r - b.l) / 2, (b.b - b.t) / 2);
    return { l: b.l + d, t: b.t + d, r: b.r - d, b: b.b - d };
  }

  private box(shape: 'rectangle' | 'ellipse', box: Rect): void {
    const b = this.frameBox(box);
    const path = new Path();
    if (shape === 'rectangle') {
      path.polygon([...this.pt(b.l, b.t), ...this.pt(b.r, b.t), ...this.pt(b.r, b.b), ...this.pt(b.l, b.b)]);
    } else {
      this.ellipseArc(path, b, 0, Math.PI * 2, true);
      path.close();
    }
    this.emit(path, true);
  }

  private roundRect(box: Rect, w: number, h: number): void {
    const b = this.frameBox(box);
    const rx = Math.min(Math.abs(w) / 2, (b.r - b.l) / 2), ry = Math.min(Math.abs(h) / 2, (b.b - b.t) / 2);
    const path = new Path();
    const corner = (cx: number, cy: number, t0: number) =>
      arcCurves(cx, cy, rx, ry, t0, Math.PI / 2, (a, bb, c, d, e, f) => path.curveTo(...this.pt(a, bb), ...this.pt(c, d), ...this.pt(e, f)));
    path.moveTo(...this.pt(b.l + rx, b.t));
    path.lineTo(...this.pt(b.r - rx, b.t));
    corner(b.r - rx, b.t + ry, -Math.PI / 2);
    path.lineTo(...this.pt(b.r, b.b - ry));
    corner(b.r - rx, b.b - ry, 0);
    path.lineTo(...this.pt(b.l + rx, b.b));
    corner(b.l + rx, b.b - ry, Math.PI / 2);
    path.lineTo(...this.pt(b.l, b.t + ry));
    corner(b.l + rx, b.t + ry, Math.PI);
    path.close();
    this.emit(path, true);
  }

  /** Add an arc of the ellipse in `b` (logical) to `path`, from parameter angle t0 through `sweep`. */
  private ellipseArc(path: Path, b: Rect, t0: number, sweep: number, start: boolean): void {
    const cx = (b.l + b.r) / 2, cy = (b.t + b.b) / 2, rx = (b.r - b.l) / 2, ry = (b.b - b.t) / 2;
    const p0 = this.pt(cx + rx * Math.cos(t0), cy + ry * Math.sin(t0));
    if (start) path.moveTo(...p0); else path.lineTo(...p0);
    arcCurves(cx, cy, rx, ry, t0, sweep, (a, bb, c, d, e, f) => path.curveTo(...this.pt(a, bb), ...this.pt(c, d), ...this.pt(e, f)));
  }

  /**
   * Arc, ArcTo, Chord and Pie: the ellipse in the box, from where the radial through `start` meets it to where the
   * radial through `end` does, counter-clockwise on the device unless SetArcDirection said otherwise.
   */
  private arc(shape: 'arc' | 'arcTo' | 'chord' | 'pie', box: Rect, start: [number, number], end: [number, number]): void {
    const b = shape === 'arc' || shape === 'arcTo' ? normalize(box) : this.frameBox(box);
    const cx = (b.l + b.r) / 2, cy = (b.t + b.b) / 2, rx = (b.r - b.l) / 2 || 1e-9, ry = (b.b - b.t) / 2 || 1e-9;
    const t0 = Math.atan2((start[1] - cy) / ry, (start[0] - cx) / rx);
    const t1 = Math.atan2((end[1] - cy) / ry, (end[0] - cx) / rx);
    // the angles are logical: a mapping that flips an axis (y up) turns the device's direction round
    const clockwise = (this.dc.arcDir === AD_CLOCKWISE) !== (det(this.matrix()) < 0);
    let sweep = t1 - t0;
    if (clockwise) { while (sweep <= 0) sweep += Math.PI * 2; } else { while (sweep >= 0) sweep -= Math.PI * 2; }
    if (Math.abs(Math.abs(sweep) - Math.PI * 2) < 1e-12 && start[0] === end[0] && start[1] === end[1]) sweep = clockwise ? Math.PI * 2 : -Math.PI * 2;
    if (shape === 'arcTo') {
      const path = this.figure();
      this.ellipseArc(path, b, t0, sweep, false);
      this.dc.cx = cx + rx * Math.cos(t0 + sweep);
      this.dc.cy = cy + ry * Math.sin(t0 + sweep);
      if (!this.bracket) this.emit(path, false);
      return;
    }
    const path = new Path();
    if (shape === 'pie') {
      path.moveTo(...this.pt(cx, cy));
      this.ellipseArc(path, b, t0, sweep, false);
      path.close();
    } else {
      this.ellipseArc(path, b, t0, sweep, true);
      if (shape === 'chord') path.close();
    }
    this.emit(path, shape !== 'arc');
  }

  private angleArc(x: number, y: number, r: number, start: number, sweep: number): void {
    if (!Number.isFinite(start) || !Number.isFinite(sweep)) return;
    // angles count counter-clockwise from the x axis, which on a y-down page is decreasing parameter angle
    const t0 = (-start * Math.PI) / 180;
    const s = Math.max(-Math.PI * 8, Math.min(Math.PI * 8, (-sweep * Math.PI) / 180));
    const path = this.figure();
    this.ellipseArc(path, { l: x - r, t: y - r, r: x + r, b: y + r }, t0, s, false);
    this.dc.cx = x + r * Math.cos(t0 + s);
    this.dc.cy = y + r * Math.sin(t0 + s);
    if (!this.bracket) this.emit(path, false);
  }

  /** A shape: into the open path bracket, or filled with the brush (closed shapes) and outlined with the pen. */
  private emit(path: Path, fill: boolean): void {
    if (this.bracket) {
      this.bracket.append(path);
      return;
    }
    this.spend(path.size);
    this.syncClip();
    if (this.opts.snap) path.snap();
    if (fill) this.fillPath(path, this.dc.brush, this.dc.fillMode === WINDING ? 'nonzero' : 'evenodd');
    this.strokePath(path, this.dc.pen);
  }

  /** Three corners of an axis-aligned destination on whole pixels (with `snap`), keeping at least one pixel. */
  private snapCorners(p0: [number, number], p1: [number, number], p2: [number, number]): [[number, number], [number, number], [number, number]] {
    if (!this.opts.snap || p0[1] !== p1[1] || p0[0] !== p2[0]) return [p0, p1, p2];
    const r = (a: number, b: number) => { const ra = Math.round(a); let rb = Math.round(b); if (rb === ra && b !== a) rb = ra + Math.sign(b - a); return [ra, rb]; };
    const [x0, x1] = r(p0[0], p1[0]), [y0, y2] = r(p0[1], p2[1]);
    return [[x0!, y0!], [x1!, y0!], [x0!, y2!]];
  }

  // ------------------------------------------------------------------------------------------------ paths

  private pathOp(op: 'begin' | 'end' | 'close' | 'abort' | 'fill' | 'stroke' | 'strokeAndFill' | 'flatten' | 'widen'): void {
    switch (op) {
      case 'begin': this.bracket = new Path(); this.path = undefined; this.pathText = []; return;
      case 'end': if (this.bracket) { this.path = this.bracket; this.bracket = undefined; } return;
      case 'close': this.bracket?.close(); return;
      case 'abort': this.bracket = undefined; this.path = undefined; this.pathText = []; return;
      case 'flatten': return;
      case 'widen': if (this.path) widened.add(this.path); return;
      default: {
        const path = this.path;
        this.path = undefined;
        if (!path) return;
        this.spend(path.size);
        this.syncClip();
        const rule = this.dc.fillMode === WINDING ? 'nonzero' : 'evenodd';
        if (op !== 'stroke') {
          if (widened.has(path)) this.strokePath(path, { ...this.dc.pen, color: this.dc.brush.color, style: this.dc.pen.style & ~0xf });
          else this.fillPath(closedCopy(path), this.dc.brush, rule);
          // text in a path bracket adds its outlines: approximated by drawing it in the brush's colour
          const text = this.pathText;
          this.pathText = [];
          for (const draw of text) draw();
        }
        if (op !== 'fill') this.strokePath(path, this.dc.pen);
        return;
      }
    }
  }

  // ------------------------------------------------------------------------------------------------ paint

  /** A COLORREF as 0xRRGGBB, through the selected palette for PALETTEINDEX colours. */
  rgb(c: number): number {
    if (c >>> 24 === 1) {
      const pal = this.dc.palette?.colors;
      const i = c & 0xffff;
      return pal && i < pal.length ? pal[i]! : 0;
    }
    return ((c & 0xff) << 16) | (c & 0xff00) | ((c >>> 16) & 0xff);
  }

  /** The binary raster operation's colour and compositing for a pen or brush of `rgb`; undefined draws nothing. */
  private rop2(rgb: number): { rgb: number; op: GlobalCompositeOperation } | undefined {
    switch (this.dc.rop2) {
      case 1: return { rgb: 0, op: 'source-over' };                     // R2_BLACK
      case 16: return { rgb: 0xffffff, op: 'source-over' };             // R2_WHITE
      case 11: return undefined;                                         // R2_NOP
      case 4: return { rgb: ~rgb & 0xffffff, op: 'source-over' };       // R2_NOTCOPYPEN
      case 6: return { rgb: 0xffffff, op: 'difference' };              // R2_NOT
      case 7: return { rgb, op: 'difference' };                         // R2_XORPEN
      case 10: return { rgb: ~rgb & 0xffffff, op: 'difference' };      // R2_NOTXORPEN
      case 9: return { rgb, op: 'multiply' };                           // R2_MASKPEN
      case 3: return { rgb: ~rgb & 0xffffff, op: 'multiply' };         // R2_MASKNOTPEN
      case 15: return { rgb, op: 'lighten' };                           // R2_MERGEPEN
      case 12: return { rgb: ~rgb & 0xffffff, op: 'lighten' };         // R2_MERGENOTPEN
      default: return { rgb, op: 'source-over' };                       // R2_COPYPEN and the rest
    }
  }

  /** What a brush paints with: a colour, a pattern, or nothing (BS_NULL). */
  private brushPaint(b: BrushObject): { style: string | CanvasPattern; rgb: number } | undefined {
    if (b.style === BS_NULL) return undefined;
    const rgb = this.rgb(b.color);
    if (b.style === BS_HATCHED) {
      const bg = this.dc.bkMode === OPAQUE ? this.rgb(this.dc.bkColor) : -1;
      const hatch = HATCHES[b.hatch];
      if (!hatch) return { style: css(rgb), rgb };
      const { pattern } = this.pattern(`h${b.hatch}|${rgb}|${bg}`, () => {
        const data = new Uint8ClampedArray(8 * 8 * 4);
        for (let y = 0; y < 8; y++) {
          for (let x = 0; x < 8; x++) {
            const on = (hatch[y]! >> (7 - x)) & 1;
            const c = on ? rgb : bg;
            if (c < 0) continue;
            data.set([(c >>> 16) & 255, (c >>> 8) & 255, c & 255, 255], (y * 8 + x) * 4);
          }
        }
        return { width: 8, height: 8, data };
      });
      return { style: pattern ?? css(rgb), rgb };
    }
    if (b.dib && b.style !== BS_SOLID) {
      const fg = this.rgb(this.dc.textColor), bg = this.rgb(this.dc.bkColor);
      const opts: DecodeOptions = b.mono ? { fg, bg } : {};
      const colors = this.dibColors(b.dib);
      if (colors) opts.colors = colors;
      const { pattern, average } = this.pattern(`d${dibKey(b.dib)}|${b.mono ? `${fg}|${bg}` : ''}`, () => this.decode(b.dib!, opts));
      return { style: pattern ?? css(average ?? rgb), rgb: average ?? rgb };
    }
    return { style: css(rgb), rgb };
  }

  /**
   * A repeating pattern of device pixels (a hatch or a pattern brush), cached by `key`, with the tile's average
   * colour, which stands in where a pattern cannot be made or a raster operation needs one colour.
   */
  private pattern(key: string, make: () => Pixels | undefined): { pattern?: CanvasPattern; average?: number } {
    let hit = this.patterns.get(key);
    if (!hit) {
      const px = make();
      const img = px ? this.image(px) : undefined;
      const p = img ? this.ctx.createPattern(img, 'repeat') : null;
      if (p) {
        // a pattern's pixels are pixels of the device the picture is played on, aligned to the brush origin
        const [ox, oy] = apply(this.d2o, this.dc.bx, this.dc.by);
        p.setTransform?.({ a: this.pixel, b: 0, c: 0, d: this.pixel, e: ox, f: oy });
      }
      hit = {};
      if (p) hit.pattern = p;
      if (px) hit.average = averageColor(px);
      this.patterns.set(key, hit);
    }
    return hit;
  }

  fillPath(path: Path, b: BrushObject, rule: CanvasFillRule): void {
    const paint = this.brushPaint(b);
    if (!paint) return;
    const r = this.rop2(paint.rgb);
    if (!r) return;
    const ctx = this.ctx;
    ctx.beginPath();
    path.trace(ctx);
    ctx.globalCompositeOperation = r.op;
    // a raster operation that changes the colour fills with that colour, patterns included
    ctx.fillStyle = r.rgb === paint.rgb ? paint.style : css(r.rgb);
    ctx.fill(rule);
    ctx.globalCompositeOperation = 'source-over';
  }

  /** Line width in output units: a cosmetic pen is one device pixel, a geometric one is never thinner. */
  private penWidth(p: PenObject): number {
    return p.cosmetic ? this.pixel : Math.max(p.width * this.scaleX(), this.pixel);
  }

  private dashes(p: PenObject, width: number): number[] {
    let style = p.style & 0xf;
    const geometric = (p.style & PS_GEOMETRIC) !== 0 && !p.cosmetic;
    // CreatePen: a pen wider than one pixel draws solid whatever its style
    if (!geometric && !p.cosmetic && width > this.pixel * 1.5 && style >= PS_DASH && style <= PS_DASHDOTDOT) style = 0;
    // cosmetic dashes are counted in device pixels, geometric ones in pen widths
    const u = geometric ? width : this.pixel;
    const cos = !geometric;
    switch (style) {
      case PS_DASH: return cos ? [18 * u, 6 * u] : [3 * u, u];
      case PS_DOT: return cos ? [3 * u, 3 * u] : [u, u];
      case PS_DASHDOT: return cos ? [9 * u, 6 * u, 3 * u, 6 * u] : [3 * u, u, u, u];
      case PS_DASHDOTDOT: return cos ? [9 * u, 3 * u, 3 * u, 3 * u, 3 * u, 3 * u] : [3 * u, u, u, u, u, u];
      case PS_ALTERNATE: return [this.pixel, this.pixel];
      case PS_USERSTYLE: {
        const k = p.cosmetic ? this.pixel : this.scaleX();
        const d = (p.dashes ?? []).map(v => Math.max(0, v * k));
        return d.some(v => v > 0) ? d : [];
      }
      default: return [];
    }
  }

  strokePath(path: Path, p: PenObject): void {
    const style = p.style & 0xf;
    if (style === PS_NULL) return;
    const r = this.rop2(this.rgb(p.color));
    if (!r) return;
    const ctx = this.ctx;
    const width = this.penWidth(p);
    const cap = p.style & 0xf00, join = p.style & 0xf000;
    const dash = this.dashes(p, width);
    ctx.lineWidth = width;
    ctx.lineCap = cap === 0x100 ? 'square' : cap === 0x200 || (dash.length && cap === 0) ? 'butt' : 'round';
    ctx.lineJoin = join === 0x1000 ? 'bevel' : join === 0x2000 ? 'miter' : 'round';
    ctx.miterLimit = Math.max(1, this.dc.miterLimit);
    ctx.beginPath();
    path.trace(ctx);
    ctx.globalCompositeOperation = r.op;
    if (dash.length && this.dc.bkMode === OPAQUE) {
      // the gaps of a styled line are drawn in the background colour
      ctx.setLineDash([]);
      ctx.strokeStyle = css(this.rgb(this.dc.bkColor));
      ctx.stroke();
    }
    ctx.setLineDash(dash);
    ctx.strokeStyle = css(r.rgb);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.globalCompositeOperation = 'source-over';
  }

  // ------------------------------------------------------------------------------------------------ clipping

  /** Make the canvas clip the current region (the EMF+ player sets its own through `setClip`). */
  syncClip(): void { this.setClip(this.dc.clip); }

  setClip(clip: Clip): void {
    if (this.applied === clip) return;
    const ctx = this.ctx;
    ctx.restore();
    ctx.save();
    for (const t of clip) {
      ctx.beginPath();
      if (t.kind === 'rects') {
        for (const r of t.rects) ctx.rect(r.l, r.t, r.r - r.l, r.b - r.t);
        ctx.clip('nonzero');
      } else if (t.kind === 'path') {
        t.path.trace(ctx);
        ctx.clip(t.rule);
      } else {
        ctx.rect(-1e9, -1e9, 2e9, 2e9);
        t.path.trace(ctx);
        ctx.clip('evenodd');
      }
    }
    this.applied = clip;
  }

  /** A device-unit rectangle in output units. */
  private devRect(r: Rect): Rect {
    const [l, t] = apply(this.d2o, r.l, r.t);
    const [rr, b] = apply(this.d2o, r.r, r.b);
    return normalize({ l, t, r: rr, b });
  }

  /** A logical rectangle in output units (its bounding box when the transform rotates). */
  private logicalRect(r: Rect): Rect {
    const pts = [this.pt(r.l, r.t), this.pt(r.r, r.t), this.pt(r.r, r.b), this.pt(r.l, r.b)];
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    return { l: Math.min(...xs), t: Math.min(...ys), r: Math.max(...xs), b: Math.max(...ys) };
  }

  private clipRect(op: 'intersect' | 'exclude', box: Rect): void {
    const dc = this.dc;
    const b = normalize(box);
    if (axisAligned(this.matrix())) {
      dc.clip = combineRects(dc.clip, op === 'intersect' ? RGN_AND : RGN_DIFF, [this.logicalRect(b)]);
      return;
    }
    const path = new Path();
    path.polygon([...this.pt(b.l, b.t), ...this.pt(b.r, b.t), ...this.pt(b.r, b.b), ...this.pt(b.l, b.b)]);
    dc.clip = combinePath(dc.clip, op === 'intersect' ? RGN_AND : RGN_DIFF, path, 'nonzero');
  }

  // ------------------------------------------------------------------------------------------------ regions, gradients

  /** An object by table index or stock index. */
  private object(i: number): GdiObject | undefined {
    return i >= STOCK ? STOCK_OBJECTS[i - STOCK] : this.objects.get(i);
  }

  /** FillRgn, PaintRgn, FrameRgn, InvertRgn: rectangles from the record (EMF) or a region object (WMF), logical. */
  private region(rec: Extract<MetaRecord, { type: 'region' }>): void {
    const ro = rec.region === undefined ? undefined : this.object(rec.region);
    const rects = rec.rects ?? (ro?.kind === 'region' ? ro.rects : undefined);
    if (!rects?.length) return;
    this.spend(rects.length);
    const path = new Path();
    for (const r of rects) path.polygon([...this.pt(r.l, r.t), ...this.pt(r.r, r.t), ...this.pt(r.r, r.b), ...this.pt(r.l, r.b)]);
    this.syncClip();
    const bo = rec.brush === undefined ? this.dc.brush : this.object(rec.brush);
    const b = bo?.kind === 'brush' ? bo : undefined;
    switch (rec.op) {
      case 'paint': this.fillPath(path, this.dc.brush, 'nonzero'); return;
      case 'fill': if (b) this.fillPath(path, b, 'nonzero'); return;
      case 'invert': {
        const ctx = this.ctx;
        ctx.beginPath(); path.trace(ctx);
        ctx.globalCompositeOperation = 'difference';
        ctx.fillStyle = '#fff';
        ctx.fill();
        ctx.globalCompositeOperation = 'source-over';
        return;
      }
      case 'frame': {
        if (!b) return;
        // each rectangle's border, w wide and h high, inside it
        const w = Math.abs(rec.w ?? 1), h = Math.abs(rec.h ?? 1);
        const frame = new Path();
        const quad = (l: number, t: number, r: number, b: number) => [...this.pt(l, t), ...this.pt(r, t), ...this.pt(r, b), ...this.pt(l, b)];
        for (const r0 of rects) {
          const r = normalize(r0);
          frame.polygon(quad(r.l, r.t, r.r, r.b));
          if (r.r - r.l > 2 * w && r.b - r.t > 2 * h) frame.polygon(quad(r.l + w, r.t + h, r.r - w, r.b - h));
        }
        this.fillPath(frame, b, 'evenodd');
        return;
      }
    }
  }

  private gradient(mode: number, v: Int32Array, colors: Uint32Array, idx: Uint32Array): void {
    const ctx = this.ctx;
    this.syncClip();
    const per = mode === 2 ? 3 : 2;
    this.spend(idx.length);
    for (let i = 0; i + per <= idx.length; i += per) {
      const a = idx[i]!, b = idx[i + 1]!;
      const path = new Path();
      if (mode === 2) {
        const c = idx[i + 2]!;
        path.polygon([...this.pt(v[a * 2]!, v[a * 2 + 1]!), ...this.pt(v[b * 2]!, v[b * 2 + 1]!), ...this.pt(v[c * 2]!, v[c * 2 + 1]!)]);
        // a triangle's three-colour blend is approximated by its average colour
        const avg = [16, 8, 0].map(s => Math.round((((colors[a]! >>> s) & 255) + ((colors[b]! >>> s) & 255) + ((colors[c]! >>> s) & 255)) / 3));
        ctx.fillStyle = css((avg[0]! << 16) | (avg[1]! << 8) | avg[2]!);
      } else {
        const x0 = v[a * 2]!, y0 = v[a * 2 + 1]!, x1 = v[b * 2]!, y1 = v[b * 2 + 1]!;
        path.polygon([...this.pt(x0, y0), ...this.pt(x1, y0), ...this.pt(x1, y1), ...this.pt(x0, y1)]);
        const [gx0, gy0] = this.pt(x0, y0);
        const [gx1, gy1] = mode === 0 ? this.pt(x1, y0) : this.pt(x0, y1);
        const g = ctx.createLinearGradient(gx0, gy0, gx1, gy1);
        g.addColorStop(0, css(colors[a]!));
        g.addColorStop(1, css(colors[b]!));
        ctx.fillStyle = g;
      }
      ctx.globalCompositeOperation = 'source-over';
      ctx.beginPath();
      path.trace(ctx);
      ctx.fill();
    }
  }

  // ------------------------------------------------------------------------------------------------ text

  /** Em height over cell height (ascent + descent) for a font, measured once; LOGFONT heights above 0 are cells. */
  private emPerCell(f: FontObject): number {
    const key = cssFont(f, 100);
    let r = this.fontRatios.get(key);
    if (r === undefined) {
      this.ctx.font = key;
      const m = this.ctx.measureText('Hg');
      const cell = (m.fontBoundingBoxAscent ?? 0) + (m.fontBoundingBoxDescent ?? 0);
      r = cell > 0 ? Math.min(1.2, Math.max(0.5, 100 / cell)) : 0.87;
      this.fontRatios.set(key, r);
    }
    return r;
  }

  private text(rec: TextRecord, colorOverride?: number): void {
    const dc = this.dc;
    const f = dc.font;
    const table = symbolTable(f);
    let s: string;
    let starts: number[] | undefined;
    if (rec.text !== undefined) s = rec.text;
    else {
      const d = decodeAnsiUnits(rec.bytes ?? new Uint8Array(0), table ? SYMBOL_CHARSET : f.charSet, f.face);
      s = d.text;
      starts = d.starts;
    }
    this.spend(s.length);
    if (table) s = symbolText(s, table);
    const m = this.matrix();
    const update = (dc.textAlign & TA_UPDATECP) !== 0;
    const rx = update ? dc.cx : rec.x, ry = update ? dc.cy : rec.y;

    if (this.bracket && colorOverride === undefined) {
      // text in a path bracket adds its outlines to the path: drawn when the path is filled (see pathOp)
      const snapshot = { ...dc };
      this.pathText.push(() => {
        const keep = this.dc;
        this.dc = snapshot;
        this.dirty = true;
        this.text(rec, keep.brush.color);
        this.dc = keep;
        this.dirty = true;
      });
      return;
    }

    this.syncClip();
    const ctx = this.ctx;
    if (rec.box && rec.options & ETO_OPAQUE) {
      const b = normalize(rec.box);
      const p = new Path();
      p.polygon([...this.pt(b.l, b.t), ...this.pt(b.r, b.t), ...this.pt(b.r, b.b), ...this.pt(b.l, b.b)]);
      ctx.globalCompositeOperation = 'source-over';
      ctx.fillStyle = css(this.rgb(dc.bkColor));
      ctx.beginPath(); p.trace(ctx); ctx.fill();
    }
    // glyph indexes need the font's own glyph table, which a browser does not expose: such text is not drawn
    if (!s.length || rec.options & ETO_GLYPH_INDEX) return;

    // the glyph frame: x along the baseline, y down; (a, b) and (c, d) in output units per logical unit
    const esc = ((f.escapement || 0) / 10) * (Math.PI / 180);
    const ce = Math.cos(esc), se = Math.sin(esc);
    const sx = Math.hypot(m[0], m[1]), sy = Math.hypot(m[2], m[3]);
    let a: number, b: number, c: number, d: number, advance: number;
    if (!axisAligned(m) || (rec.advanced && Math.abs(sx - sy) > 1e-9 * (sx + sy))) {
      // GM_ADVANCED: the world transform applies to the glyphs too (kept upright when it mirrors)
      const la = m[0], lb = m[1];
      let lc = m[2], ld = m[3];
      if (det(m) < 0) { lc = -lc; ld = -ld; }
      a = la * ce - lc * se; b = lb * ce - ld * se; c = la * se + lc * ce; d = lb * se + ld * ce;
      advance = Math.hypot(c, d);
    } else {
      // GM_COMPATIBLE: upright glyphs scaled by the height's scale, turned by the escapement
      a = sy * ce; b = -sy * se; c = sy * se; d = sy * ce;
      advance = sx;
    }
    const gy = Math.hypot(c, d);
    if (!(gy > 0) || !Number.isFinite(gy)) return;
    const emLogical = f.height < 0 ? -f.height : f.height > 0 ? f.height * this.emPerCell(f) : (16 * this.pixel) / gy;
    const em = emLogical * gy;
    // laid out in units of `unit` output units, so that the font size is a sane number of pixels whatever the
    // output units are (canvas pixels, or EMUs when drawing into a slide)
    const size = Math.min(2000, Math.max(1, em / this.pixel));
    const unit = em / size;
    if (!(unit > 0) || !Number.isFinite(unit)) return;
    const k = unit / gy;
    const adv = advance / unit;
    const [px, py] = this.pt(rx, ry);
    ctx.save();
    if (rec.box && rec.options & ETO_CLIPPED) {
      const bx = normalize(rec.box);
      const clip = new Path();
      clip.polygon([...this.pt(bx.l, bx.t), ...this.pt(bx.r, bx.t), ...this.pt(bx.r, bx.b), ...this.pt(bx.l, bx.b)]);
      ctx.beginPath();
      clip.trace(ctx);
      ctx.clip();
    }
    ctx.transform(a * k, b * k, c * k, d * k, px, py);
    ctx.font = cssFont(f, size);
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    ctx.globalCompositeOperation = 'source-over';
    const metrics = ctx.measureText(s);
    const ascent = metrics.fontBoundingBoxAscent ?? size * 0.9;
    const descent = metrics.fontBoundingBoxDescent ?? size * 0.21;

    // advances: from the record (per string unit, summed per character), or measured
    const chars = Array.from(s);
    let xs: number[] | undefined;
    let ys: number[] | undefined;
    let total: number;
    if (rec.dx) {
      xs = [];
      ys = rec.dy ? [] : undefined;
      let x = 0, y = 0, at = 0;
      for (const ch of chars) {
        xs.push(x);
        ys?.push(y);
        const from = starts ? starts[at]! : at;
        const to = starts ? (at + ch.length < starts.length ? starts[at + ch.length]! : rec.dx.length) : at + ch.length;
        for (let i = from; i < to && i < rec.dx.length; i++) {
          x += rec.dx[i]! * adv;
          if (rec.dy) y -= (rec.dy[i] ?? 0) * adv;
        }
        at += ch.length;
      }
      total = x;
    } else if (dc.charExtra) {
      xs = [];
      let x = 0;
      for (const ch of chars) { xs.push(x); x += ctx.measureText(ch).width + dc.charExtra * adv; }
      total = x;
    } else {
      total = metrics.width;
    }
    const h = dc.textAlign & TA_CENTER;
    const x0 = h === TA_CENTER ? -total / 2 : h === TA_RIGHT ? -total : 0;
    const v = dc.textAlign & TA_BASELINE;
    const y0 = v === TA_BASELINE ? 0 : v === TA_BOTTOM ? -descent : ascent;
    if (dc.bkMode === OPAQUE && !(rec.options & ETO_OPAQUE) && colorOverride === undefined) {
      ctx.fillStyle = css(this.rgb(dc.bkColor));
      ctx.fillRect(x0, y0 - ascent, total, ascent + descent);
    }
    ctx.fillStyle = css(this.rgb(colorOverride ?? dc.textColor));
    if (xs && !COMPLEX.test(s)) {
      // each character goes where the advances put it; a stand-in font wider than the one the advances were
      // measured with is narrowed so that its glyphs do not run into each other
      const squeeze = rec.dx && total > 0 && metrics.width > total * 1.01 ? total / metrics.width : 1;
      if (squeeze < 1) ctx.transform(squeeze, 0, 0, 1, 0, 0);
      for (let i = 0; i < chars.length; i++) ctx.fillText(chars[i]!, (x0 + xs[i]!) / squeeze, y0 + (ys?.[i] ?? 0));
      if (squeeze < 1) ctx.transform(1 / squeeze, 0, 0, 1, 0, 0);
    } else if (rec.dx && total > 0 && metrics.width > total * 1.01) {
      // complex scripts are shaped as a whole, fitted to the advances' total
      ctx.transform(total / metrics.width, 0, 0, 1, 0, 0);
      ctx.fillText(s, (x0 * metrics.width) / total, y0);
      ctx.transform(metrics.width / total, 0, 0, 1, 0, 0);
    } else {
      ctx.fillText(s, x0, y0);
    }
    if (f.underline || f.strikeOut) {
      const t = Math.max(size / 14, this.pixel / unit);
      if (f.underline) ctx.fillRect(x0, y0 + size * 0.1, total, t);
      if (f.strikeOut) ctx.fillRect(x0, y0 - size * 0.28, total, t);
    }
    ctx.restore();
    if (update) {
      dc.cx = rx + (total / adv) * ce;
      dc.cy = ry - (total / adv) * se;
    }
  }

  // ------------------------------------------------------------------------------------------------ bitmaps

  /** Pixels to something drawImage takes, through a scratch canvas. */
  image(px: Pixels): CanvasImageSource | undefined {
    const make = this.opts.createCanvas;
    if (!make) return undefined;
    const canvas = make(px.width, px.height);
    const c = canvas.getContext('2d');
    if (!c) return undefined;
    const data = c.createImageData(px.width, px.height);
    data.data.set(px.data);
    c.putImageData(data, 0, 0);
    return canvas as unknown as CanvasImageSource;
  }

  /** A DIB's colour table, resolved through the logical palette for DIB_PAL_COLORS. */
  private dibColors(d: Dib): Uint32Array | undefined {
    if (!d.paletteIndexes) return undefined;
    const pal = this.dc.palette?.colors;
    return Uint32Array.from(d.paletteIndexes, i => (pal && i < pal.length ? pal[i]! : (i & 1 ? 0xffffff : 0)));
  }

  /** Decode, counting the pixels against the work limit; undefined for an undecodable bitmap. */
  decode(d: Dib, opts: DecodeOptions): Pixels | undefined {
    if (d.encoded) return undefined;
    this.spend(d.width * d.height);
    try {
      return decodeDib(d, opts);
    } catch (e) {
      if (e instanceof MetafileError) return undefined;
      throw e;
    }
  }

  /** The image to draw for a blt's bitmap with its raster operation applied to the pixels, cached per variant. */
  private bltImage(rec: BltRecord, plan: ReturnType<typeof ropPlan>, rop: number): CanvasImageSource | undefined {
    const d = rec.dib!;
    if (d.encoded) return this.opts.images?.get(dibImageKey(d.bits));
    const dc = this.dc;
    const fg = this.rgb(dc.textColor), bg = this.rgb(dc.bkColor);
    const brushRgb = this.brushPaint(dc.brush)?.rgb ?? 0xffffff;
    const key = [rop, plan.pixels ?? '', d.colors.length ? '' : `${fg}|${bg}`, plan.pixels ? brushRgb : '', rec.alpha?.perPixel ?? '',
      rec.transparent ?? '', rec.mask ? 'm' : ''].join('|');
    let variants = this.bitmaps.get(d);
    if (!variants) { variants = new Map(); this.bitmaps.set(d, variants); }
    const hit = variants.get(key);
    if (hit !== undefined) return hit ?? undefined;
    const opts: DecodeOptions = { fg, bg, alpha: rec.alpha?.perPixel ? 'premultiplied' : 'ignore' };
    if (d.colors.length) delete opts.fg;
    const colors = this.dibColors(d);
    if (colors) opts.colors = colors;
    const px = this.decode(d, opts);
    if (!px) { variants.set(key, null); return undefined; }
    const data = px.data;
    if (rec.transparent !== undefined) {
      const t = this.rgb(rec.transparent);
      for (let i = 0; i < data.length; i += 4) if (((data[i]! << 16) | (data[i + 1]! << 8) | data[i + 2]!) === t) data[i + 3] = 0;
    }
    if (rec.mask) {
      // MASKBLT and PLGBLT: where the mask is 0 the destination is kept
      const mpx = this.decode(rec.mask.dib, { fg: 0, bg: 0xffffff });
      if (mpx) {
        // the mask's corner is at the source rectangle's corner, and it repeats when smaller
        const wrap = (v: number, n: number) => ((v % n) + n) % n;
        for (let y = 0; y < px.height; y++) {
          const my = wrap(y - rec.sy + rec.mask.y, mpx.height);
          for (let x = 0; x < px.width; x++) {
            const mx = wrap(x - rec.sx + rec.mask.x, mpx.width);
            if (mpx.data[(my * mpx.width + mx) * 4]! <= 127) data[(y * px.width + x) * 4 + 3] = 0;
          }
        }
      }
    }
    if (plan.pixels) {
      const pr = (brushRgb >>> 16) & 255, pg = (brushRgb >>> 8) & 255, pb = brushRgb & 255;
      for (let i = 0; i < data.length; i += 4) {
        if (plan.pixels === 'invert') { data[i] = 255 - data[i]!; data[i + 1] = 255 - data[i + 1]!; data[i + 2] = 255 - data[i + 2]!; continue; }
        if (plan.pixels === 'and-pattern') { data[i] = data[i]! & pr; data[i + 1] = data[i + 1]! & pg; data[i + 2] = data[i + 2]! & pb; continue; }
        // generic: the operation against a black and a white destination, read as a blend of a colour over it
        const r0 = ropBits(rop, 0, pr, data[i]!), g0 = ropBits(rop, 0, pg, data[i + 1]!), b0 = ropBits(rop, 0, pb, data[i + 2]!);
        const r1 = ropBits(rop, 1, pr, data[i]!), g1 = ropBits(rop, 1, pg, data[i + 1]!), b1 = ropBits(rop, 1, pb, data[i + 2]!);
        const alpha = 1 - (Math.max(0, r1 - r0) + Math.max(0, g1 - g0) + Math.max(0, b1 - b0)) / 765;
        data[i] = alpha > 0 ? r0 / alpha : 0;
        data[i + 1] = alpha > 0 ? g0 / alpha : 0;
        data[i + 2] = alpha > 0 ? b0 / alpha : 0;
        data[i + 3] = Math.round(data[i + 3]! * alpha);
      }
    }
    const img = this.image(px);
    variants.set(key, img ?? null);
    return img;
  }

  /** A blt's destination parallelogram: its upper-left, upper-right and lower-left corners in output units. */
  private corners(rec: BltRecord): [[number, number], [number, number], [number, number]] {
    const c = rec.corners;
    if (c) return [this.pt(c[0], c[1]), this.pt(c[2], c[3]), this.pt(c[4], c[5])];
    if (rec.device) {
      // a 1:1 copy: the size is in device pixels from the logical origin
      const [x, y] = this.pt(rec.x, rec.y);
      return [[x, y], [x + rec.w * this.d2o[0], y], [x, y + rec.h * this.d2o[3]]];
    }
    return [this.pt(rec.x, rec.y), this.pt(rec.x + rec.w, rec.y), this.pt(rec.x, rec.y + rec.h)];
  }

  private blt(rec: BltRecord): void {
    const rop = (rec.rop >>> 16) & 0xff;
    if (rop === 0xaa) return;                               // the destination unchanged
    this.syncClip();
    const ctx = this.ctx;
    if (!rec.dib) {
      // the brush and the destination only (PatBlt)
      const [p0, p1, p2] = this.snapCorners(...this.corners(rec));
      const path = new Path();
      path.polygon([...p0, ...p1, p1[0] + p2[0] - p0[0], p1[1] + p2[1] - p0[1], ...p2]);
      const b = this.dc.brush;
      const fill = (style: string | CanvasPattern, op: GlobalCompositeOperation) => {
        ctx.globalCompositeOperation = op;
        ctx.fillStyle = style;
        ctx.beginPath(); path.trace(ctx); ctx.fill();
        ctx.globalCompositeOperation = 'source-over';
      };
      const paint = this.brushPaint(b);
      switch (rop) {
        case 0x00: fill('#000', 'source-over'); return;                         // BLACKNESS
        case 0xff: fill('#fff', 'source-over'); return;                         // WHITENESS
        case 0x55: fill('#fff', 'difference'); return;                          // DSTINVERT
        case 0x5a: if (paint) fill(paint.style, 'difference'); return;          // PATINVERT
        case 0x0f: if (paint) fill(css(~paint.rgb & 0xffffff), 'source-over'); return; // ~P
        case 0xa0: if (paint) fill(paint.style, 'multiply'); return;            // P & D
        case 0xfa: if (paint) fill(paint.style, 'lighten'); return;             // P | D
        default: if (paint) fill(paint.style, 'source-over'); return;           // PATCOPY and the rest
      }
    }
    const d = rec.dib;
    const plan = ropPlan(rop);
    const img = this.bltImage(rec, plan, rop);
    if (!img) return;
    const iw = d.encoded ? (img as { width: number }).width : d.width;
    const ih = d.encoded ? (img as { height: number }).height : d.height;
    // raster operations that combine with the destination stay pixel-exact, so that masks and images line up
    this.place(rec, img, iw, ih, plan.op, rec.alpha ? rec.alpha.constant / 255 : 1, plan.op !== 'source-over' || !!plan.pixels);
  }

  /** Draw a blt's bitmap (iw by ih pixels) into its destination parallelogram. */
  private place(rec: BltRecord, img: CanvasImageSource, iw: number, ih: number, op: GlobalCompositeOperation, alpha: number, exact: boolean): void {
    if (!rec.corners && (rec.w === 0 || rec.h === 0)) return;
    let [p0, p1, p2] = this.corners(rec);
    let { sx, sy, sw, sh } = rec;
    if (rec.dib?.encoded) { sx = 0; sy = 0; sw = iw; sh = ih; }
    // a negative source size mirrors: flip the destination instead
    if (sw < 0) {
      sx += sw; sw = -sw;
      const q = p0; p0 = p1; p1 = q;
      p2 = [p2[0] + p0[0] - p1[0], p2[1] + p0[1] - p1[1]];
    }
    if (sh < 0) {
      sy += sh; sh = -sh;
      const q = p0; p0 = p2; p2 = q;
      p1 = [p1[0] + p0[0] - p2[0], p1[1] + p0[1] - p2[1]];
    }
    if (sw === 0 || sh === 0) { sx = 0; sy = 0; sw = iw; sh = ih; }
    [p0, p1, p2] = this.snapCorners(p0, p1, p2);
    // a source rectangle outside the stored bitmap, when the bitmap is exactly the source size: the bitmap itself
    if ((sx < 0 || sy < 0 || sx + sw > iw || sy + sh > ih) && Math.abs(sw - iw) <= 1 && Math.abs(sh - ih) <= 1) { sx = 0; sy = 0; }
    const ctx = this.ctx;
    this.syncClip();
    ctx.save();
    ctx.globalCompositeOperation = op;
    ctx.globalAlpha = alpha;
    // GDI's stretch modes drop or merge pixels (HALFTONE averages them); the picture is drawn larger than GDI
    // would, so bitmaps are filtered like the rest of it, except where pixels combine with what is below
    ctx.imageSmoothingEnabled = !exact;
    ctx.transform((p1[0] - p0[0]) / sw, (p1[1] - p0[1]) / sw, (p2[0] - p0[0]) / sh, (p2[1] - p0[1]) / sh, p0[0], p0[1]);
    try {
      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);
    } finally {
      ctx.restore();          // a host image that cannot be drawn (closed, broken) must not unbalance the context
    }
  }

  /**
   * The XOR mask idiom: SRCINVERT the image, SRCAND a mask, SRCINVERT the image again leaves the destination
   * where the mask is white and the image where it is black. Drawn as such, since XOR cannot be composited.
   */
  private xorMask(recs: MetaRecord[], i: number): boolean {
    const a = recs[i] as BltRecord, m = recs[i + 1], b = recs[i + 2];
    if (!a.dib || a.dib.encoded || m?.type !== 'blt' || b?.type !== 'blt' || !m.dib || !b.dib) return false;
    const same = (r: BltRecord) => r.x === a.x && r.y === a.y && r.w === a.w && r.h === a.h;
    if (((m.rop >>> 16) & 0xff) !== 0x88 || b.rop !== 0x660046 || !same(m) || !same(b)) return false;
    const fg = this.rgb(this.dc.textColor), bg = this.rgb(this.dc.bkColor);
    const img = this.decode(a.dib, a.dib.colors.length ? {} : { fg, bg });
    const mask = this.decode(m.dib, m.dib.colors.length ? {} : { fg, bg });
    if (!img || !mask || mask.width !== img.width || mask.height !== img.height) return false;
    for (let p = 0; p < img.data.length; p += 4) {
      if (mask.data[p]! + mask.data[p + 1]! + mask.data[p + 2]! > 382) img.data[p + 3] = 0;
    }
    const canvas = this.image(img);
    if (canvas) this.place(a, canvas, img.width, img.height, 'source-over', 1, true);
    return true;
  }
}

/** Paths that WIDENPATH turned into their outline: filling one strokes it. */
const widened = new WeakSet<Path>();

/** A copy of a path with every figure closed, as FILLPATH closes them. */
function closedCopy(p: Path): Path {
  const out = new Path();
  out.append(p);
  out.close();
  return out;
}

/** Characters from scripts that need shaping (Arabic, Hebrew, Indic, Thai...), which per-character drawing breaks. */
const COMPLEX = /[\u0590-\u08ff\u0900-\u0dff\u0e00-\u0eff\u1000-\u109f\u1780-\u17ff]/;

const dibIds = new WeakMap<Dib, number>();
let nextDibId = 0;
const dibKey = (d: Dib) => {
  let id = dibIds.get(d);
  if (id === undefined) { id = nextDibId++; dibIds.set(d, id); }
  return id;
};

function averageColor(px: Pixels): number {
  let r = 0, g = 0, b = 0, n = 0;
  for (let i = 0; i < px.data.length; i += 4) {
    if (px.data[i + 3] === 0) continue;
    r += px.data[i]!; g += px.data[i + 1]!; b += px.data[i + 2]!; n++;
  }
  return n ? (Math.round(r / n) << 16) | (Math.round(g / n) << 8) | Math.round(b / n) : 0;
}

/** Draw a parsed metafile into `dest` (in the context's current user space). */
export function drawMetafile(ctx: MetaCtx, mf: Metafile, dest: Dest, opts: DrawOptions = {}): void {
  new Player(ctx, mf, dest, opts).run();
}

