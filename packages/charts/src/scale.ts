// packages/charts/src/scale.ts  -- value axis scaling: automatic minimum, maximum and units the way Office picks them
//
// Office's rules, as observed and documented by chart practitioners: zero is the minimum unless the data's minimum
// is at least 5/6 of its maximum; 5% of the data range is added as headroom before rounding the maximum up to a
// whole major unit; the major unit is the smallest 1, 2 or 5 times a power of ten whose labels fit along the axis;
// the minor unit is a fifth of the major one; 3-D charts get no headroom. Values set in the file (c:min, c:max,
// c:majorUnit) are kept.

export interface Scale {
  min: number;
  max: number;
  major: number;
  minor: number;
  /** Logarithmic axes: the base; min and max are then powers of it. */
  logBase?: number;
}

export interface ScaleInput {
  /** The data's extent; lo <= hi. */
  lo: number;
  hi: number;
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  logBase?: number;
  /** percentStacked: values are fractions and the axis spans 0..100%. */
  percent?: boolean;
  /** 3-D charts get no headroom above their largest value. */
  tight?: boolean;
  /** Whether a candidate scale's labels fit along the axis; candidates come in increasing unit size. */
  fits?: (s: Scale) => boolean;
}

const STEPS = [1, 2, 5];

/** The major units Office chooses among, increasing, from just under `from`. */
function* candidates(from: number): Generator<number> {
  let k = Math.floor(Math.log10(Math.max(from, 1e-300))) - 1;
  for (let n = 0; n < 60; n++, k++) for (const s of STEPS) yield s * 10 ** k;
}

/** Remove float noise from a multiple of a unit: 0.30000000000000004 -> 0.3. */
export const clean = (v: number, unit: number): number => {
  const digits = Math.max(0, Math.min(15, -Math.floor(Math.log10(Math.abs(unit) || 1)) + 2));
  return Number(v.toFixed(digits));
};

const floorTo = (v: number, u: number) => clean(Math.floor(v / u + 1e-9) * u, u);
const ceilTo = (v: number, u: number) => clean(Math.ceil(v / u - 1e-9) * u, u);

function logScale(input: ScaleInput, base: number): Scale {
  const lb = Math.log(base);
  const lo = input.lo > 0 ? input.lo : 1;
  const hi = input.hi > 0 ? Math.max(input.hi, lo) : lo * base;
  let min = input.min !== undefined && input.min > 0 ? input.min : base ** Math.floor(Math.log(lo) / lb + 1e-9);
  let max = input.max !== undefined && input.max > 0 ? input.max : base ** Math.ceil(Math.log(hi) / lb - 1e-9);
  if (!(max > min)) max = min * base;
  if (!Number.isFinite(min) || !Number.isFinite(max)) {
    min = 1;
    max = base;
  }
  const major = input.majorUnit && input.majorUnit > 0 ? input.majorUnit : base;
  return { min, max, major, minor: base, logBase: base };
}

/** Choose an axis scale for data in [lo, hi]. */
export function autoScale(input: ScaleInput): Scale {
  if (input.logBase) return logScale(input, input.logBase);
  let lo = Number.isFinite(input.lo) ? input.lo : 0;
  let hi = Number.isFinite(input.hi) ? input.hi : 1;
  if (lo > hi) [lo, hi] = [hi, lo];
  // without a length to fit, Office's cap of ten major intervals
  const fits = input.fits ?? ((s: Scale) => (s.max - s.min) / s.major <= 10 + 1e-9);

  if (input.percent && input.min === undefined && input.max === undefined) {
    const min = lo < -1e-9 ? -1 : 0;
    const max = hi > 1e-9 || min === 0 ? 1 : 0;
    const steps = input.majorUnit ? [input.majorUnit] : [0.1, 0.2, 0.25, 0.5, 1];
    const major = steps.find(s => fits({ min, max, major: s, minor: s / 5 })) ?? steps.at(-1)!;
    return { min, max, major, minor: input.minorUnit ?? major / 5 };
  }

  // the data's extent, with zero brought in unless the values sit far from it
  if (input.min !== undefined) lo = Math.min(lo, input.min);
  if (input.max !== undefined) hi = Math.max(hi, input.max);
  const free = input.min === undefined && input.max === undefined && !input.majorUnit;
  if (lo === hi) {
    // nothing to plot: Office's empty axis runs 0 to 1 in fifths
    if (lo === 0 && free) return { min: 0, max: 1, major: 0.2, minor: input.minorUnit ?? 0.04 };
    if (lo === 0) hi = 1;
    else if (lo > 0) lo = 0;
    else hi = 0;
  }
  if (lo >= 0 && lo < (hi * 5) / 6) lo = 0;
  if (hi <= 0 && hi > (lo * 5) / 6) hi = 0;
  const range = hi - lo;
  const room = input.tight ? 0 : 0.05 * range;
  const top = hi > 0 ? hi + room : hi;
  const bottom = lo < 0 ? lo - room : lo > 0 ? lo - 0.5 * range : 0;

  const build = (major: number): Scale => {
    // positive data keeps the minimum at or above zero
    const min = input.min ?? (lo === 0 ? 0 : floorTo(lo > 0 ? Math.max(0, bottom) : bottom, major));
    let max = input.max ?? (hi === 0 ? 0 : ceilTo(top, major));
    if (!(max > min)) max = min + major;
    return { min, max, major, minor: input.minorUnit ?? major / 5 };
  };
  const unit = input.majorUnit;
  if (unit && unit > 0 && (input.max ?? hi) - (input.min ?? lo) < unit * 1000) return build(unit);
  let last: Scale | undefined;
  for (const major of candidates(range / 1000)) {
    const s = build(major);
    if ((s.max - s.min) / major > 1000) continue;
    last = s;
    if (fits(s)) return s;
    if ((s.max - s.min) / major <= 1) break;
  }
  return last ?? build(10 ** Math.ceil(Math.log10(range || 1)));
}

/** The major tick values of a scale, from min to max inclusive (at most 1000). */
export function ticks(s: Scale, unit = s.major): number[] {
  const out: number[] = [];
  if (s.logBase) {
    const lb = Math.log(s.logBase);
    const step = s.major > 1 ? Math.log(s.major) / lb : 1;
    const a = Math.log(s.min) / lb;
    const b = Math.log(s.max) / lb;
    for (let k = 0; a + k * step <= b + 1e-9 && k <= 1000; k++) out.push(s.logBase ** (a + k * step));
    return out;
  }
  if (!(unit > 0) || !Number.isFinite(s.min) || !Number.isFinite(s.max)) return [s.min, s.max];
  const n = Math.floor((s.max - s.min) / unit + 1e-9);
  if (n > 1000) return [s.min, s.max];
  for (let k = 0; k <= n; k++) out.push(clean(s.min + k * unit, unit));
  return out;
}

/** Where a value falls along a scale, 0 at min and 1 at max (log axes map logarithmically). */
export function fraction(s: Scale, v: number): number {
  if (s.logBase) {
    if (!(v > 0)) return -1;
    return (Math.log(v) - Math.log(s.min)) / (Math.log(s.max) - Math.log(s.min));
  }
  return (v - s.min) / (s.max - s.min || 1);
}
