// packages/metafile/src/types.ts  -- the structural description of a metafile: what the parsers make and the player draws
//
// WMF and EMF records are both turned into one set of GDI operations, so one player draws both. Coordinates stay
// as the file gives them (logical units, or device units where a record says so); the player does all mapping.
import type { View } from './view.ts';

/** A rectangle by its edges. */
export interface Rect { l: number; t: number; r: number; b: number; }

/** An affine transform [a, b, c, d, e, f] in the canvas convention: x' = a x + c y + e, y' = b x + d y + f. */
export type Matrix = [number, number, number, number, number, number];

/** A device-independent bitmap (or a WMF device-dependent one, recast): header fields and still-encoded pixels. */
export interface Dib {
  width: number;
  height: number;
  topDown: boolean;
  bitCount: number;
  compression: number;
  /** Bytes per row when it is not the DIB rule (rows padded to 4 bytes): WMF Bitmap16 rows are padded to 2. */
  stride?: number;
  /** Colour table as 0xRRGGBB. Empty for a 1-bpp device bitmap, which takes the text and background colours. */
  colors: Uint32Array;
  /** DIB_PAL_COLORS: the colour table holds indexes into the selected logical palette instead of colours. */
  paletteIndexes?: Uint16Array;
  /** BI_BITFIELDS (and 16/32-bpp defaults): red, green, blue and alpha masks. */
  masks?: [number, number, number, number];
  bits: Uint8Array;
  /** BI_JPEG and BI_PNG: the bits are a whole JPEG or PNG file, decoded before playback. */
  encoded?: 'image/jpeg' | 'image/png';
}

export interface PenObject {
  kind: 'pen';
  /** PS_* line style (low 4 bits), end cap (0x0F00) and join (0xF000). */
  style: number;
  width: number;
  color: number;
  /** One device pixel wide whatever the transform: PS_COSMETIC, or a width of zero. */
  cosmetic: boolean;
  /** PS_USERSTYLE dash and gap lengths. */
  dashes?: number[];
}

export interface BrushObject {
  kind: 'brush';
  /** BS_* brush style. */
  style: number;
  color: number;
  hatch: number;
  /** Pattern brushes: the tile. */
  dib?: Dib;
  /** CreateMonoBrush and Bitmap16 patterns: 0 bits take the text colour, 1 bits the background colour. */
  mono?: boolean;
}

export interface FontObject {
  kind: 'font';
  height: number;
  width: number;
  escapement: number;
  orientation: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strikeOut: boolean;
  charSet: number;
  quality: number;
  pitchAndFamily: number;
  face: string;
}

export interface PaletteObject { kind: 'palette'; colors: Uint32Array; }
export interface RegionObject { kind: 'region'; rects: Rect[]; }
/** Objects that only hold a slot in the object table (colour spaces and the like). */
export interface OtherObject { kind: 'other'; }

export type GdiObject = PenObject | BrushObject | FontObject | PaletteObject | RegionObject | OtherObject;

export interface TextRecord {
  type: 'text';
  /** Reference point, logical units. */
  x: number;
  y: number;
  /** UTF-16 text (EMF), or `bytes` in the character set of the selected font (WMF and 8-bit EMF text). */
  text?: string;
  bytes?: Uint8Array;
  /** Advance of each character (per string unit), logical units; with ETO_PDY also the vertical advances. */
  dx?: Int32Array;
  dy?: Int32Array;
  /** ETO_* options. */
  options: number;
  /** Opaque and clip rectangle, logical units. */
  box?: Rect;
  /** EMF GM_ADVANCED: the world transform shapes the glyphs too, not only where they go. */
  advanced?: boolean;
}

export interface BltRecord {
  type: 'blt';
  /** Destination rectangle, logical units. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Source rectangle in the bitmap's pixels, rows counted from the top. */
  sx: number;
  sy: number;
  sw: number;
  sh: number;
  /** Ternary raster operation (dwRop); its index is (rop >> 16) & 0xFF. */
  rop: number;
  /** No bitmap: the raster operation uses only the brush and the destination (PatBlt). */
  dib?: Dib;
  /** MASKBLT and PLGBLT: where the mask is 1 the source is drawn, elsewhere the destination is kept. */
  mask?: { dib: Dib; x: number; y: number };
  /** ALPHABLEND: constant alpha 0..255, and whether the source carries per-pixel premultiplied alpha. */
  alpha?: { constant: number; perPixel: boolean };
  /** TRANSPARENTBLT: this source colour is not drawn. */
  transparent?: number;
  /** PLGBLT: the destination parallelogram's upper-left, upper-right and lower-left corners. */
  corners?: [number, number, number, number, number, number];
  /** SETDIBITSTODEVICE: w and h are device pixels (the bitmap is copied, not stretched); x and y stay logical. */
  device?: boolean;
}

/** An EMF+ record: its header and its data, read during playback. */
export interface PlusRecord { type: number; flags: number; data: View; }

export type MetaRecord =
  | { type: 'eof' }
  // mapping
  | { type: 'mapMode'; mode: number }
  | { type: 'windowOrg' | 'windowExt' | 'viewportOrg' | 'viewportExt' | 'offsetWindowOrg' | 'offsetViewportOrg'; x: number; y: number }
  | { type: 'scaleWindowExt' | 'scaleViewportExt'; xNum: number; xDenom: number; yNum: number; yDenom: number }
  | { type: 'worldTransform'; m: Matrix }
  | { type: 'modifyWorldTransform'; m: Matrix; mode: number }
  // state
  | { type: 'saveDC' }
  | { type: 'restoreDC'; n: number }
  | { type: 'bkMode' | 'textAlign' | 'polyFillMode' | 'rop2' | 'stretchMode' | 'arcDirection' | 'miterLimit' | 'charExtra' | 'layout'; value: number }
  | { type: 'bkColor' | 'textColor'; color: number }
  | { type: 'brushOrg'; x: number; y: number }
  // objects; index -1 is the lowest free slot (WMF)
  | { type: 'create'; index: number; object: GdiObject }
  | { type: 'select'; index: number }
  | { type: 'delete'; index: number }
  // drawing, logical units
  | { type: 'moveTo' | 'lineTo'; x: number; y: number }
  | { type: 'poly'; shape: 'polyline' | 'polylineTo' | 'polygon' | 'bezier' | 'bezierTo'; points: Int32Array }
  | { type: 'polyPoly'; shape: 'polyline' | 'polygon'; counts: Uint32Array; points: Int32Array }
  | { type: 'polyDraw'; points: Int32Array; kinds: Uint8Array }
  | { type: 'box'; shape: 'rectangle' | 'ellipse'; box: Rect }
  | { type: 'roundRect'; box: Rect; w: number; h: number }
  | { type: 'arc'; shape: 'arc' | 'arcTo' | 'chord' | 'pie'; box: Rect; start: [number, number]; end: [number, number] }
  | { type: 'angleArc'; x: number; y: number; r: number; start: number; sweep: number }
  | { type: 'pixel'; x: number; y: number; color: number }
  | { type: 'path'; op: 'begin' | 'end' | 'close' | 'abort' | 'fill' | 'stroke' | 'strokeAndFill' | 'flatten' | 'widen' }
  | { type: 'gradient'; mode: number; vertices: Int32Array; colors: Uint32Array; indexes: Uint32Array }
  | { type: 'region'; op: 'fill' | 'paint' | 'frame' | 'invert'; rects?: Rect[]; region?: number; brush?: number; w?: number; h?: number }
  | TextRecord
  | BltRecord
  // clipping
  | { type: 'clipRect'; op: 'intersect' | 'exclude'; box: Rect }
  /** EMF: device units; null rects: the default clip. */
  | { type: 'clipRegion'; mode: number; rects: Rect[] | null }
  /** WMF: an object table index; 0 when it is not a region resets the clip. */
  | { type: 'selectClipRegion'; index: number }
  | { type: 'clipPath'; mode: number }
  | { type: 'offsetClip'; x: number; y: number }
  // EMF+ records carried in an EMF comment
  | { type: 'plus'; records: PlusRecord[] };

export interface Metafile {
  format: 'emf' | 'wmf';
  /**
   * The picture frame in the metafile's device units: it maps onto the destination rectangle. A WMF has no device,
   * so its device units are the logical units of its frame (the placeable header's box, or the first window): the
   * window and viewport start out as the frame, and records that change the window rescale from there.
   */
  frame: Rect;
  /** Natural size in CSS pixels (96 per inch), from the frame. */
  width: number;
  height: number;
  /** Device units per millimetre: the reference device of an EMF, the placeable header's units for a WMF. */
  dpmmX: number;
  dpmmY: number;
  records: MetaRecord[];
  /** EMF+ content: 'dual' files carry the same picture as EMF records, 'only' files need the EMF+ player. */
  plus: 'none' | 'dual' | 'only';
  /** Records skipped because they were inconsistent (counted, not fatal). */
  skipped: number;
}
