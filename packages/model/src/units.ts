// packages/model/src/units.ts  -- the units DrawingML stores numbers in
export const EMU_PER_INCH = 914400;
export const EMU_PER_PT = 12700;
export const EMU_PER_CM = 360000;
export const EMU_PER_PX = 9525;            // at 96 DPI
export const ANGLE_UNIT = 60000;            // ST_Angle: 60000ths of a degree

export const emuToPt = (emu: number) => emu / EMU_PER_PT;
export const ptToEmu = (pt: number) => Math.round(pt * EMU_PER_PT);
export const emuToPx = (emu: number) => emu / EMU_PER_PX;
export const angleToDeg = (st: number) => st / ANGLE_UNIT;
export const degToAngle = (deg: number) => Math.round(deg * ANGLE_UNIT);

/** An integer attribute, or the default when absent or not a number. */
export function int(v: string | undefined, dflt: number): number;
export function int(v: string | undefined): number | undefined;
export function int(v: string | undefined, dflt?: number): number | undefined {
  if (v === undefined) return dflt;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : dflt;
}

/** A decimal attribute, or the default. */
export function num(v: string | undefined, dflt: number): number;
export function num(v: string | undefined): number | undefined;
export function num(v: string | undefined, dflt?: number): number | undefined {
  if (v === undefined) return dflt;
  const n = Number.parseFloat(v);
  return Number.isFinite(n) ? n : dflt;
}

/** xsd:boolean: "1"/"true" and "0"/"false". */
export function bool(v: string | undefined, dflt: boolean): boolean;
export function bool(v: string | undefined): boolean | undefined;
export function bool(v: string | undefined, dflt?: boolean): boolean | undefined {
  if (v === undefined) return dflt;
  if (v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return dflt;
}

/** ST_Percentage and its positive and fixed variants, as a fraction: Transitional writes 50000 for 50%, Strict "50%". */
export function percent(v: string | undefined, dflt: number): number;
export function percent(v: string | undefined): number | undefined;
export function percent(v: string | undefined, dflt?: number): number | undefined {
  if (v === undefined) return dflt;
  const n = v.endsWith('%') ? Number.parseFloat(v) / 100 : Number.parseInt(v, 10) / 100000;
  return Number.isFinite(n) ? n : dflt;
}

/** ST_Coordinate and friends: Transitional writes EMU integers, Strict may write a universal measure ("2.5cm", "12pt"). */
export function coord(v: string | undefined, dflt: number): number;
export function coord(v: string | undefined): number | undefined;
export function coord(v: string | undefined, dflt?: number): number | undefined {
  if (v === undefined) return dflt;
  const m = /^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)?$/.exec(v.trim());
  if (!m) return dflt;
  const n = Number.parseFloat(m[1]!);
  switch (m[2]) {
    case undefined: return n;
    case 'mm': return Math.round(n * EMU_PER_CM / 10);
    case 'cm': return Math.round(n * EMU_PER_CM);
    case 'in': return Math.round(n * EMU_PER_INCH);
    case 'pt': return Math.round(n * EMU_PER_PT);
    default: return Math.round(n * EMU_PER_PT * 12);   // pc and pi: picas
  }
}
