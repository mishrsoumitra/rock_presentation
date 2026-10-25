// packages/geometry/src/types.ts  -- shape geometry as DrawingML writes it (ECMA-376 Part 1, section 20.1.9)
//
// Values are kept as the strings the XML holds: a guide name ('hc', 'adj1', 'x2') or a literal (EMU for
// coordinates, 60000ths of a degree for angles, or the path's own units inside a path with w and h).

/** A shape guide: a named formula such as { name: 'a', fmla: 'pin 0 adj 50000' }. Adjust values are guides too. */
export interface Guide {
  readonly name: string;
  readonly fmla: string;
}

export interface AdjPoint {
  readonly x: string;
  readonly y: string;
}

/** a:ahXY. A missing min or max means the handle cannot move along that axis. */
export interface XYAdjustHandle {
  readonly kind: 'xy';
  readonly gdRefX?: string;
  readonly minX?: string;
  readonly maxX?: string;
  readonly gdRefY?: string;
  readonly minY?: string;
  readonly maxY?: string;
  readonly pos: AdjPoint;
}

/** a:ahPolar. The angle limits are angles, the radius limits coordinates. */
export interface PolarAdjustHandle {
  readonly kind: 'polar';
  readonly gdRefR?: string;
  readonly minR?: string;
  readonly maxR?: string;
  readonly gdRefAng?: string;
  readonly minAng?: string;
  readonly maxAng?: string;
  readonly pos: AdjPoint;
}

export type AdjustHandle = XYAdjustHandle | PolarAdjustHandle;

export interface ConnectionSite {
  readonly ang: string;
  readonly pos: AdjPoint;
}

/** a:rect, the text rectangle. */
export interface GeomRect {
  readonly l: string;
  readonly t: string;
  readonly r: string;
  readonly b: string;
}

export type PathFill = 'norm' | 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

export type PathDefCommand =
  | { readonly op: 'moveTo'; readonly pt: AdjPoint }
  | { readonly op: 'lnTo'; readonly pt: AdjPoint }
  | { readonly op: 'arcTo'; readonly wR: string; readonly hR: string; readonly stAng: string; readonly swAng: string }
  | { readonly op: 'quadBezTo'; readonly pts: readonly [AdjPoint, AdjPoint] }
  | { readonly op: 'cubicBezTo'; readonly pts: readonly [AdjPoint, AdjPoint, AdjPoint] }
  | { readonly op: 'close' };

/** a:path. When w or h is not 0 the path has its own coordinate space, scaled to the shape's size. */
export interface PathDef {
  readonly w: number;
  readonly h: number;
  readonly fill: PathFill;
  readonly stroke: boolean;
  readonly extrusionOk: boolean;
  readonly commands: readonly PathDefCommand[];
}

/** a:custGeom, or one preset from Annex D: the same structure. */
export interface ShapeGeometry {
  readonly avLst: readonly Guide[];
  readonly gdLst: readonly Guide[];
  readonly ahLst: readonly AdjustHandle[];
  readonly cxnLst: readonly ConnectionSite[];
  readonly rect: GeomRect | null;
  readonly pathLst: readonly PathDef[];
}

// What evaluateGeometry returns. Every length is in the shape's own space: 0..w and 0..h, in EMU.

export type PathCommand =
  | { op: 'M'; x: number; y: number }
  | { op: 'L'; x: number; y: number }
  | { op: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: 'Q'; x1: number; y1: number; x: number; y: number }
  | { op: 'Z' };

export interface GeometryPath {
  commands: PathCommand[];
  fill: PathFill;
  stroke: boolean;
  extrusionOk: boolean;
}

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Where to draw an adjust handle; ShapeGeometry.ahLst at the same index says what it adjusts. */
export interface HandlePosition {
  x: number;
  y: number;
}

/** A connection site and the direction a connector leaves it in, in 60000ths of a degree (0 = right, clockwise). */
export interface ConnectionSitePosition {
  x: number;
  y: number;
  angle: number;
}

export interface EvaluatedGeometry {
  guides: Map<string, number>;      // every guide by name, built-ins included; a redefined name has its last value
  paths: GeometryPath[];
  textRect: Rect;                   // the whole box when the geometry has no a:rect
  handles: HandlePosition[];
  connectionSites: ConnectionSitePosition[];
}
