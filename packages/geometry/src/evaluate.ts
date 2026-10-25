// packages/geometry/src/evaluate.ts  -- a shape geometry at one size: guides, paths, text rectangle, handles, sites
import { appendArc, finite } from './arc.ts';
import { GeometryError } from './errors.ts';
import type {
  AdjPoint, ConnectionSitePosition, EvaluatedGeometry, GeometryPath, Guide, PathCommand, PathFill, Rect, ShapeGeometry,
} from './types.ts';

const RAD = Math.PI / 10800000;   // radians per 60000th of a degree

type Fn = (x: number, y: number, z: number) => number;

/** The 17 formulas of section 20.1.9.11, with their argument counts. Angles are in 60000ths of a degree. */
const OPERATORS = new Map<string, readonly [arity: number, fn: Fn]>([
  ['*/', [3, (x, y, z) => (x * y) / z]],
  ['+-', [3, (x, y, z) => x + y - z]],
  ['+/', [3, (x, y, z) => (x + y) / z]],
  ['?:', [3, (x, y, z) => (x > 0 ? y : z)]],
  ['abs', [1, x => Math.abs(x)]],
  ['at2', [2, (x, y) => Math.atan2(y, x) / RAD]],
  ['cat2', [3, (x, y, z) => x * Math.cos(Math.atan2(z, y))]],
  ['cos', [2, (x, y) => x * Math.cos(y * RAD)]],
  ['max', [2, (x, y) => Math.max(x, y)]],
  ['min', [2, (x, y) => Math.min(x, y)]],
  ['mod', [3, (x, y, z) => Math.hypot(x, y, z)]],
  ['pin', [3, (x, y, z) => (y < x ? x : y > z ? z : y)]],
  ['sat2', [3, (x, y, z) => x * Math.sin(Math.atan2(z, y))]],
  ['sin', [2, (x, y) => x * Math.sin(y * RAD)]],
  ['sqrt', [1, x => Math.sqrt(x)]],
  ['tan', [2, (x, y) => x * Math.tan(y * RAD)]],
  ['val', [1, x => x]],
]);

/**
 * The guides every geometry can use (section 20.1.10.56). Annex D also uses hd10 and wd32, which that list lacks,
 * so each of the w, h and ss fractions is defined for the same set of divisors.
 */
const BUILTINS: [string, (w: number, h: number) => number][] = [
  ['l', () => 0], ['t', () => 0], ['r', w => w], ['b', (_, h) => h], ['w', w => w], ['h', (_, h) => h],
  ['hc', w => w / 2], ['vc', (_, h) => h / 2], ['ss', (w, h) => Math.min(w, h)], ['ls', (w, h) => Math.max(w, h)],
  ['cd2', () => 10800000], ['cd4', () => 5400000], ['cd8', () => 2700000], ['3cd4', () => 16200000],
  ['3cd8', () => 8100000], ['5cd8', () => 13500000], ['7cd8', () => 18900000],
];
for (const d of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
  BUILTINS.push([`wd${d}`, w => w / d], [`hd${d}`, (_, h) => h / d], [`ssd${d}`, (w, h) => Math.min(w, h) / d]);
}

// A literal: an integer (ST_Coordinate, ST_Angle) or, in a formula, a decimal like the spec's own "*/ w 1.0 2.0".
const NUMBER = /^[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)$/;
// ST_UniversalMeasure, which a coordinate may also be: "2.5cm".
const MEASURE = /^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$/;
const EMU_PER = new Map([['mm', 36000], ['cm', 360000], ['in', 914400], ['pt', 12700], ['pc', 152400], ['pi', 152400]]);

interface Formula { readonly fn: Fn; readonly args: readonly string[]; }

// Formulas by text. The presets share most of theirs, and a shape's own adjust values ("val 25000") are parsed on
// every evaluation, so this saves parsing them again. Bounded: files can hold any number of distinct values.
const formulas = new Map<string, Formula>();

function parseFormula(fmla: string, where: string): Formula {
  const known = formulas.get(fmla);
  if (known) return known;
  const [op = '', ...args] = fmla.trim().split(/\s+/);
  const operator = OPERATORS.get(op);
  if (!operator) throw new GeometryError(`${where}: unknown formula ${JSON.stringify(op)} in ${JSON.stringify(fmla)}`);
  if (args.length !== operator[0]) {
    throw new GeometryError(`${where}: "${op}" takes ${operator[0]} argument${operator[0] > 1 ? 's' : ''}, not ${args.length}, in ${JSON.stringify(fmla)}`);
  }
  const f = { fn: operator[1], args };
  if (formulas.size >= 10000) formulas.clear();
  formulas.set(fmla, f);
  return f;
}

// ---- compiling a geometry: names resolved to value slots once, so evaluation is a straight run over arrays ----

interface Step {
  readonly name: string;
  readonly fn: Fn;
  readonly a: number;   // argument slots; unused ones point at slot 0
  readonly b: number;
  readonly c: number;
  readonly out: number;
  /** For an adjust value: the names visible where it is defined, for a replacement formula from the shape. */
  readonly scope: ReadonlyMap<string, number> | null;
}

interface PathProgram {
  readonly w: number;
  readonly h: number;
  readonly fill: PathFill;
  readonly stroke: boolean;
  readonly extrusionOk: boolean;
  readonly commands: readonly { readonly op: 'M' | 'L' | 'A' | 'Q' | 'C' | 'Z'; readonly slots: readonly number[] }[];
}

/** One way a handle moves: along x or y (ahXY), or in angle or radius (ahPolar). Limits are slots, -1 if absent. */
export interface Axis {
  readonly ref: string;
  readonly min: number;
  readonly max: number;
  readonly along: 'x' | 'y' | 'polar';
}

export interface Program {
  readonly template: Float64Array;                    // constants in place, built-ins and guides still 0
  readonly steps: readonly Step[];                    // avLst then gdLst, in document order
  readonly names: ReadonlyMap<string, number>;        // every name visible after the last guide
  readonly adjustSlots: ReadonlyMap<string, number>;  // avLst names
  readonly paths: readonly PathProgram[];
  readonly rect: readonly [number, number, number, number] | null;
  readonly sites: readonly (readonly [ang: number, x: number, y: number])[];
  readonly handles: readonly { readonly x: number; readonly y: number; readonly axes: readonly Axis[] }[];
}

function compile(geom: ShapeGeometry): Program {
  const scope = new Map<string, number>();
  const values: number[] = [];
  for (const [name] of BUILTINS) {
    scope.set(name, values.length);
    values.push(0);
  }
  const constants = new Map<number, number>();
  const constant = (value: number) => {
    let slot = constants.get(value);
    if (slot === undefined) {
      slot = values.length;
      values.push(finite(value));
      constants.set(value, slot);
    }
    return slot;
  };
  const ref = (token: string, where: string, coordinate = true): number => {
    if (NUMBER.test(token)) return constant(Number(token));
    const m = coordinate ? MEASURE.exec(token) : null;
    if (m) return constant(Number(m[1]) * EMU_PER.get(m[2]!)!);
    const slot = scope.get(token);
    if (slot === undefined) throw new GeometryError(`${where}: unknown guide ${JSON.stringify(token)}`);
    return slot;
  };

  // Guides see only the built-ins and the guides before them; a name defined again hides the earlier one from then on.
  const steps: Step[] = [];
  const adjustSlots = new Map<string, number>();
  const guide = (g: Guide, adjust: boolean) => {
    const where = `guide ${JSON.stringify(g.name)}`;
    const f = parseFormula(g.fmla, where);
    const [a = 0, b = 0, c = 0] = f.args.map(t => ref(t, where, false));
    const out = values.length;
    values.push(0);
    steps.push({ name: g.name, fn: f.fn, a, b, c, out, scope: adjust ? new Map(scope) : null });
    scope.set(g.name, out);
    if (adjust) adjustSlots.set(g.name, out);
  };
  for (const g of geom.avLst) guide(g, true);
  for (const g of geom.gdLst) guide(g, false);

  const pt = (p: AdjPoint, where: string) => [ref(p.x, where), ref(p.y, where)];
  const paths = geom.pathLst.map((path, i): PathProgram => ({
    w: path.w, h: path.h, fill: path.fill, stroke: path.stroke, extrusionOk: path.extrusionOk,
    commands: path.commands.map(c => {
      const where = `path ${i + 1} ${c.op}`;
      switch (c.op) {
        case 'moveTo': return { op: 'M', slots: pt(c.pt, where) };
        case 'lnTo': return { op: 'L', slots: pt(c.pt, where) };
        case 'arcTo': return { op: 'A', slots: [ref(c.wR, where), ref(c.hR, where), ref(c.stAng, where, false), ref(c.swAng, where, false)] };
        case 'quadBezTo': return { op: 'Q', slots: c.pts.flatMap(p => pt(p, where)) };
        case 'cubicBezTo': return { op: 'C', slots: c.pts.flatMap(p => pt(p, where)) };
        case 'close': return { op: 'Z', slots: [] };
      }
    }),
  }));

  const r = geom.rect;
  const rect = r ? [ref(r.l, 'rect'), ref(r.t, 'rect'), ref(r.r, 'rect'), ref(r.b, 'rect')] as const : null;
  const sites = geom.cxnLst.map(s => [ref(s.ang, 'cxn', false), ...pt(s.pos, 'cxn')] as [number, number, number]);
  const handles = geom.ahLst.map((hd, i) => {
    const where = `handle ${i + 1}`;
    const axis = (gdRef: string | undefined, min: string | undefined, max: string | undefined, along: Axis['along'], angle = false): Axis[] =>
      gdRef === undefined ? [] : [{
        ref: gdRef, along,
        min: min === undefined ? -1 : ref(min, where, !angle),
        max: max === undefined ? -1 : ref(max, where, !angle),
      }];
    const [x, y] = pt(hd.pos, where) as [number, number];
    const axes = hd.kind === 'xy'
      ? [...axis(hd.gdRefX, hd.minX, hd.maxX, 'x'), ...axis(hd.gdRefY, hd.minY, hd.maxY, 'y')]
      : [...axis(hd.gdRefAng, hd.minAng, hd.maxAng, 'polar', true), ...axis(hd.gdRefR, hd.minR, hd.maxR, 'polar')];
    return { x, y, axes };
  });

  return { template: Float64Array.from(values), steps, names: scope, adjustSlots, paths, rect, sites, handles };
}

// Compiled once per geometry object; presetGeometry() hands out the same object every time, so presets compile once.
const programs = new WeakMap<ShapeGeometry, Program | GeometryError>();

export function program(geom: ShapeGeometry): Program {
  let p = programs.get(geom);
  if (!p) {
    try {
      p = compile(geom);
    } catch (e) {
      if (!(e instanceof GeometryError)) throw e;
      p = e;
    }
    programs.set(geom, p);
  }
  if (p instanceof GeometryError) throw new GeometryError(p.message);
  return p;
}

// ---- evaluating ----

/** Replacement adjust values by name: a formula from the shape's a:avLst, or a number (used while solving handles). */
export type Overrides = ReadonlyMap<string, string | number>;

export const overridesOf = (adjust: readonly Guide[]): Map<string, string | number> => new Map(adjust.map(g => [g.name, g.fmla]));

function evalAdjust(name: string, fmla: string, scope: ReadonlyMap<string, number>, v: Float64Array): number {
  const where = `adjust value ${JSON.stringify(name)}`;
  const f = parseFormula(fmla, where);
  const arg = (t: string | undefined): number => {
    if (t === undefined) return 0;
    if (NUMBER.test(t)) return Number(t);
    const slot = scope.get(t);
    if (slot === undefined) throw new GeometryError(`${where}: unknown guide ${JSON.stringify(t)}`);
    return v[slot]!;
  };
  return f.fn(arg(f.args[0]), arg(f.args[1]), arg(f.args[2]));
}

/**
 * Every value slot for one size. A result that is not a finite number becomes 0: degenerate sizes divide by zero
 * (a 0 x 0 shape has ss = 0), and nothing downstream should ever see NaN or Infinity.
 */
export function run(p: Program, w: number, h: number, overrides: Overrides): Float64Array {
  const v = p.template.slice();
  for (let i = 0; i < BUILTINS.length; i++) v[i] = BUILTINS[i]![1](w, h);
  for (const s of p.steps) {
    const o = s.scope ? overrides.get(s.name) : undefined;
    const r = o === undefined ? s.fn(v[s.a]!, v[s.b]!, v[s.c]!) : typeof o === 'number' ? o : evalAdjust(s.name, o, s.scope!, v);
    v[s.out] = finite(r);
  }
  return v;
}

function buildPath(p: PathProgram, v: Float64Array, w: number, h: number): GeometryPath {
  // A path with its own w and h is drawn in that space and scaled to the shape.
  const sx = p.w > 0 ? w / p.w : 1;
  const sy = p.h > 0 ? h / p.h : 1;
  const out: PathCommand[] = [];
  let x = 0, y = 0;             // the pen, in path units; it starts at the origin
  let x0 = 0, y0 = 0;           // where the current subpath started
  let open = false;             // whether a subpath is open, i.e. an M was written since the last Z
  const begin = () => {
    // Every subpath starts with an M, also after a Z, so consumers never need an implicit current point.
    if (open) return;
    out.push({ op: 'M', x: finite(x * sx), y: finite(y * sy) });
    x0 = x; y0 = y; open = true;
  };
  let slots: readonly number[] = [];
  const at = (i: number) => v[slots[i]!]!;
  for (const c of p.commands) {
    slots = c.slots;
    switch (c.op) {
      case 'M':
        x = at(0); y = at(1);
        out.push({ op: 'M', x: finite(x * sx), y: finite(y * sy) });
        x0 = x; y0 = y; open = true;
        break;
      case 'L':
        begin();
        x = at(0); y = at(1);
        out.push({ op: 'L', x: finite(x * sx), y: finite(y * sy) });
        break;
      case 'A':
        begin();
        [x, y] = appendArc(out, x, y, at(0), at(1), at(2), at(3), sx, sy);
        break;
      case 'Q':
        begin();
        out.push({ op: 'Q', x1: finite(at(0) * sx), y1: finite(at(1) * sy), x: finite(at(2) * sx), y: finite(at(3) * sy) });
        x = at(2); y = at(3);
        break;
      case 'C':
        begin();
        out.push({
          op: 'C', x1: finite(at(0) * sx), y1: finite(at(1) * sy), x2: finite(at(2) * sx), y2: finite(at(3) * sy),
          x: finite(at(4) * sx), y: finite(at(5) * sy),
        });
        x = at(4); y = at(5);
        break;
      case 'Z':
        if (!open) break;
        out.push({ op: 'Z' });
        x = x0; y = y0; open = false;   // the pen goes back to the start of the subpath
        break;
    }
  }
  return { commands: out, fill: p.fill, stroke: p.stroke, extrusionOk: p.extrusionOk };
}

/**
 * Evaluate a geometry for a shape of w x h EMU. `adjust` is the shape's own a:avLst: each guide replaces the adjust
 * value of the same name (names the geometry does not have are ignored). Throws a GeometryError for an unknown
 * guide name or formula, including in `adjust`.
 */
export function evaluateGeometry(geom: ShapeGeometry, w: number, h: number, adjust: readonly Guide[] = []): EvaluatedGeometry {
  const p = program(geom);
  w = finite(w);
  h = finite(h);
  const v = run(p, w, h, overridesOf(adjust));
  const at = (slot: number) => v[slot]!;

  const guides = new Map<string, number>();
  for (const [name, slot] of p.names) guides.set(name, at(slot));

  let textRect: Rect = { x: 0, y: 0, w, h };
  if (p.rect) {
    const [l, t, r, b] = p.rect.map(at) as [number, number, number, number];
    textRect = { x: Math.min(l, r), y: Math.min(t, b), w: Math.abs(r - l), h: Math.abs(b - t) };
  }
  return {
    guides,
    paths: p.paths.map(path => buildPath(path, v, w, h)),
    textRect,
    handles: p.handles.map(hd => ({ x: at(hd.x), y: at(hd.y) })),
    connectionSites: p.sites.map(([ang, x, y]): ConnectionSitePosition => ({ x: at(x), y: at(y), angle: at(ang) })),
  };
}
