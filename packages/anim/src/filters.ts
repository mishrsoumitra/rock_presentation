// packages/anim/src/filters.ts  -- the reveal patterns of animEffect filters, shared with the slide transitions
//
// A pattern says which part of a box is visible at a given progress, as polygons in the unit square (0..1 both
// ways, y down). Polygons are wound clockwise and never overlap, so the nonzero rule fills their union and a
// complement is the unit square plus every polygon reversed.
//
// Directions follow PowerPoint, which names a wipe by the side it starts from: wipe(up) reveals from the top, as
// the presetSubtype PowerPoint writes with it (1, "From Top") says. ECMA-376's illustrations draw it the other way.

export interface Pt { x: number; y: number; }
export type Region = Pt[][];

/** What a filter does at one moment: the visible region (absent: all of it), an opacity, and a content offset
 *  (slide filters move the content within the box), all in units of the box. */
export interface Reveal { region?: Region; opacity?: number; shift?: Pt; }

const rect = (x0: number, y0: number, x1: number, y1: number): Pt[] =>
  [{ x: x0, y: y0 }, { x: x1, y: y0 }, { x: x1, y: y1 }, { x: x0, y: y1 }];
const clamp = (v: number, lo = 0, hi = 1) => (v < lo ? lo : v > hi ? hi : v);
const FULL: Region = [rect(0, 0, 1, 1)];

/** The region outside `r` within the unit square. */
export function complement(r: Region): Region {
  return [rect(0, 0, 1, 1), ...r.map(p => p.slice().reverse())];
}

/** A pseudo-random order of n items, the same for the same seed. */
export function shuffled(n: number, seed: number): number[] {
  let a = seed >>> 0 || 1;
  const rnd = () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = Math.imul(a ^ (a >>> 15), a | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const order = Array.from({ length: n }, (_, i) => i);
  for (let i = n - 1; i > 0; i--) {
    const j = Math.floor(rnd() * (i + 1));
    [order[i], order[j]] = [order[j]!, order[i]!];
  }
  return order;
}

export function hashString(s: string): number {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619);
  return h >>> 0;
}

/** Grid size for cell patterns: `n` cells along the longer side, square cells. aspect = width / height. */
function grid(n: number, aspect: number): { cols: number; rows: number } {
  const a = Number.isFinite(aspect) && aspect > 0 ? aspect : 1;
  return a >= 1 ? { cols: n, rows: Math.max(1, Math.round(n / a)) } : { cols: Math.max(1, Math.round(n * a)), rows: n };
}

/** Bands along one axis: [a, b] in x (vertical bands) or y (horizontal bands). */
const band = (axis: 'x' | 'y', a: number, b: number): Pt[] =>
  (axis === 'x' ? rect(clamp(a), 0, clamp(b), 1) : rect(0, clamp(a), 1, clamp(b)));

/** Cells revealed in a random order, merged into runs along each row. */
function cells(cols: number, rows: number, p: number, seed: number): Region {
  const total = cols * rows;
  const shown = Math.round(clamp(p) * total);
  if (shown >= total) return FULL;
  const on = new Uint8Array(total);
  const order = shuffled(total, seed);
  for (let i = 0; i < shown; i++) on[order[i]!] = 1;
  const out: Region = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (!on[r * cols + c]) continue;
      let e = c;
      while (e + 1 < cols && on[r * cols + e + 1]) e++;
      out.push(rect(c / cols, r / rows, (e + 1) / cols, (r + 1) / rows));
      c = e;
    }
  }
  return out;
}

/** Bars across the box revealed in a random order. */
function randomBars(axis: 'x' | 'y', n: number, p: number, seed: number): Region {
  const shown = Math.round(clamp(p) * n);
  if (shown >= n) return FULL;
  const on = new Uint8Array(n);
  const order = shuffled(n, seed);
  for (let i = 0; i < shown; i++) on[order[i]!] = 1;
  const out: Region = [];
  for (let i = 0; i < n; i++) {
    if (!on[i]) continue;
    let e = i;
    while (e + 1 < n && on[e + 1]) e++;
    out.push(band(axis, i / n, (e + 1) / n));
    i = e;
  }
  return out;
}

/** A polygon for the sector from angle a0 to a1 (radians, clockwise on screen from the x axis) about the centre,
 *  long enough to reach the corners; `aspect` keeps the angles true in the real box. */
function sector(a0: number, a1: number, aspect: number): Pt[] {
  const sx = aspect >= 1 ? 1 : aspect;
  const sy = aspect >= 1 ? 1 / aspect : 1;
  const R = 2;
  const steps = Math.max(2, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 32)));
  const pts: Pt[] = [{ x: 0.5, y: 0.5 }];
  for (let i = 0; i <= steps; i++) {
    const a = a0 + ((a1 - a0) * i) / steps;
    pts.push({ x: 0.5 + (Math.cos(a) * R) / sx, y: 0.5 + (Math.sin(a) * R) / sy });
  }
  return pts;
}

function ellipse(r: number): Pt[] {
  const pts: Pt[] = [];
  for (let i = 0; i < 64; i++) {
    const a = (i / 64) * Math.PI * 2;
    pts.push({ x: 0.5 + Math.cos(a) * r, y: 0.5 + Math.sin(a) * r });
  }
  return pts;
}

const diamondShape = (r: number): Pt[] =>
  [{ x: 0.5, y: 0.5 - r }, { x: 0.5 + r, y: 0.5 }, { x: 0.5, y: 0.5 + r }, { x: 0.5 - r, y: 0.5 }];

function plusShape(p: number): Region {
  const a = 0.5 - p / 2;
  const b = 0.5 + p / 2;
  return [rect(0, a, 1, b), rect(a, 0, b, a), rect(a, b, b, 1)];
}

/** Staircase strips moving from one corner to the opposite one; dx, dy: the direction of travel (+1 or -1). */
function strips(dx: number, dy: number, p: number): Region {
  const n = 10;
  const out: Region = [];
  for (let i = 0; i < n; i++) {
    const reach = clamp(2 * p - (i + 0.5) / n, 0, 1);
    if (reach <= 0) continue;
    const y0 = dy > 0 ? i / n : 1 - (i + 1) / n;
    const x0 = dx > 0 ? 0 : 1 - reach;
    out.push(rect(x0, y0, x0 + reach, y0 + 1 / n));
  }
  return out;
}

/** Bands of a checkerboard: cells of alternate rows (or columns) start half a period apart. */
function checkerboard(axis: 'x' | 'y', p: number, aspect: number): Region {
  const { cols, rows } = grid(8, aspect);
  const along = axis === 'x' ? cols : rows;
  const across = axis === 'x' ? rows : cols;
  const out: Region = [];
  const cell = 1 / along;
  for (let r = 0; r < across; r++) {
    const offset = (r % 2) * cell;
    for (let k = -1; k * 2 * cell < 1; k++) {
      const a = Math.max(0, offset + k * 2 * cell);
      const b = Math.min(1, offset + k * 2 * cell + 2 * cell * clamp(p));
      if (b <= a) continue;
      out.push(axis === 'x' ? rect(a, r / across, b, (r + 1) / across) : rect(r / across, a, (r + 1) / across, b));
    }
  }
  return out;
}

/** The region a filter shows when it has revealed fraction p of a box (p = 1: all of it). */
type Pattern = (p: number, aspect: number, seed: number) => Reveal;

const PATTERN_TABLE: Record<string, Pattern> = {
  'fade': p => ({ opacity: clamp(p) }),
  'dissolve': (p, aspect, seed) => ({ region: cells(grid(24, aspect).cols, grid(24, aspect).rows, p, seed) }),
  'wipe(up)': p => ({ region: [band('y', 0, p)] }),
  'wipe(down)': p => ({ region: [band('y', 1 - p, 1)] }),
  'wipe(left)': p => ({ region: [band('x', 0, p)] }),
  'wipe(right)': p => ({ region: [band('x', 1 - p, 1)] }),
  'slide(fromTop)': p => ({ region: [band('y', 0, p)], shift: { x: 0, y: -(1 - clamp(p)) } }),
  'slide(fromBottom)': p => ({ region: [band('y', 1 - p, 1)], shift: { x: 0, y: 1 - clamp(p) } }),
  'slide(fromLeft)': p => ({ region: [band('x', 0, p)], shift: { x: -(1 - clamp(p)), y: 0 } }),
  'slide(fromRight)': p => ({ region: [band('x', 1 - p, 1)], shift: { x: 1 - clamp(p), y: 0 } }),
  'blinds(horizontal)': p => ({ region: Array.from({ length: 6 }, (_, k) => band('y', k / 6, (k + clamp(p)) / 6)) }),
  'blinds(vertical)': p => ({ region: Array.from({ length: 6 }, (_, k) => band('x', k / 6, (k + clamp(p)) / 6)) }),
  'box(out)': p => ({ region: [rect(0.5 - p / 2, 0.5 - p / 2, 0.5 + p / 2, 0.5 + p / 2)] }),
  'box(in)': p => ({ region: complement([rect(p / 2, p / 2, 1 - p / 2, 1 - p / 2)]) }),
  'circle(out)': p => ({ region: [ellipse(clamp(p) * Math.SQRT1_2)] }),
  'circle(in)': p => ({ region: complement([ellipse((1 - clamp(p)) * Math.SQRT1_2)]) }),
  'diamond(out)': p => ({ region: [diamondShape(clamp(p))] }),
  'diamond(in)': p => ({ region: complement([diamondShape(1 - clamp(p))]) }),
  'plus(out)': p => ({ region: plusShape(clamp(p)) }),
  'plus(in)': p => ({ region: complement(plusShape(1 - clamp(p))) }),
  'barn(inVertical)': p => ({ region: [band('x', 0, p / 2), band('x', 1 - p / 2, 1)] }),
  'barn(outVertical)': p => ({ region: [band('x', 0.5 - p / 2, 0.5 + p / 2)] }),
  'barn(inHorizontal)': p => ({ region: [band('y', 0, p / 2), band('y', 1 - p / 2, 1)] }),
  'barn(outHorizontal)': p => ({ region: [band('y', 0.5 - p / 2, 0.5 + p / 2)] }),
  'randombar(horizontal)': (p, _a, seed) => ({ region: randomBars('y', 60, p, seed) }),
  'randombar(vertical)': (p, _a, seed) => ({ region: randomBars('x', 60, p, seed) }),
  'checkerboard(across)': (p, aspect) => ({ region: checkerboard('x', p, aspect) }),
  'checkerboard(down)': (p, aspect) => ({ region: checkerboard('y', p, aspect) }),
  'strips(downRight)': p => ({ region: strips(1, 1, p) }),
  'strips(downLeft)': p => ({ region: strips(-1, 1, p) }),
  'strips(upRight)': p => ({ region: strips(1, -1, p) }),
  'strips(upLeft)': p => ({ region: strips(-1, -1, p) }),
  // a fan opening both ways from the top
  'wedge': (p, aspect) => ({ region: [sector(-Math.PI / 2 - Math.PI * clamp(p), -Math.PI / 2 + Math.PI * clamp(p), aspect)] }),
};
// looked up through a Map: filter names come from files and must never reach Object.prototype
const PATTERNS = new Map<string, Pattern>(Object.entries(PATTERN_TABLE).map(([k, v]) => [k.toLowerCase(), v]));
const ALIASES: [string, string][] = [
  ['circle', 'circle(out)'], ['diamond', 'diamond(out)'], ['plus', 'plus(out)'],
  ['randombars(horizontal)', 'randombar(horizontal)'], ['randombars(vertical)', 'randombar(vertical)'],
];
for (const [alias, name] of ALIASES) PATTERNS.set(alias, PATTERNS.get(name)!);

function wheel(spokes: number): Pattern {
  const n = Math.max(1, Math.min(64, Math.round(spokes)));
  return (p, aspect) => ({
    region: Array.from({ length: n }, (_, k) => {
      const a0 = -Math.PI / 2 + (k * 2 * Math.PI) / n;
      return sector(a0, a0 + (clamp(p) * 2 * Math.PI) / n, aspect);
    }),
  });
}

function pattern(name: string): Pattern | undefined {
  const known = PATTERNS.get(name.toLowerCase());
  if (known) return known;
  const m = /^wheel\((\d+)\)$/.exec(name);
  return m ? wheel(Number(m[1])) : undefined;
}

/** The first filter of a fallback list ("blinds(horizontal);blinds(vertical)") that is supported. */
export function supportedFilter(filter: string | undefined): string | undefined {
  const names = (filter ?? '').split(';').slice(0, 16).map(s => s.replace(/\s+/g, ''));
  return names.find(f => f && pattern(f));
}

/**
 * What a filter shows at progress p (0..1) of an "in" transition, or of an "out" one: an exit hides what the same
 * filter would reveal (box(in) closes in on the shape), except that slide moves the content back out where it came
 * from. Unknown filters fade.
 */
export function filterReveal(
  filter: string | undefined, p: number, transition: 'in' | 'out' = 'in', aspect = 1, seed?: number,
): Reveal {
  const name = supportedFilter(filter) ?? 'fade';
  const f = pattern(name)!;
  const s = seed ?? hashString(name);
  const t = clamp(p);
  if (transition !== 'out') return f(t, aspect, s);
  if (name.startsWith('slide(')) return f(1 - t, aspect, s);
  const r = f(t, aspect, s);
  return {
    ...(r.region ? { region: complement(r.region) } : {}),
    ...(r.opacity !== undefined ? { opacity: 1 - r.opacity } : {}),
  };
}
