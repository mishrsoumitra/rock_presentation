// packages/geometry/src/presets.ts  -- the preset shapes (a:prstGeom) and text warps (a:prstTxWarp) of Annex D
import { decodeGeometry, type CompactGeometry } from './compact.ts';
import { PRESETS } from './presets.generated.ts';
import { TEXT_WARPS } from './textwarps.generated.ts';
import type { ShapeGeometry } from './types.ts';

/** Every ST_ShapeType value, sorted. */
export const PRESET_NAMES: readonly string[] = Object.freeze(Object.keys(PRESETS));
/** Every ST_TextShapeType value except textNoShape, which is plain unwarped text. Sorted. */
export const TEXT_WARP_NAMES: readonly string[] = Object.freeze(Object.keys(TEXT_WARPS));

function deepFreeze<T>(value: T): T {
  if (typeof value === 'object' && value !== null && !Object.isFrozen(value)) {
    for (const v of Object.values(value)) deepFreeze(v);
    Object.freeze(value);
  }
  return value;
}

// Decoded on first use and kept, so each preset is also compiled only once (evaluate.ts caches by object). The
// objects are shared, hence frozen.
function lookup(table: Readonly<Record<string, CompactGeometry>>, cache: Map<string, ShapeGeometry>, name: string): ShapeGeometry | undefined {
  let geom = cache.get(name);
  if (!geom && Object.hasOwn(table, name)) {   // not `name in table`: 'constructor' is no preset
    geom = deepFreeze(decodeGeometry(table[name]!));
    cache.set(name, geom);
  }
  return geom;
}

const presets = new Map<string, ShapeGeometry>();
const warps = new Map<string, ShapeGeometry>();

/** The geometry of a preset shape, by its prst value ('rect', 'roundRect', 'rightArrow'), or undefined. */
export const presetGeometry = (name: string): ShapeGeometry | undefined => lookup(PRESETS, presets, name);

/** The geometry of a preset text warp, by its prst value ('textArchUp', 'textWave1'), or undefined. */
export const textWarpGeometry = (name: string): ShapeGeometry | undefined => lookup(TEXT_WARPS, warps, name);
