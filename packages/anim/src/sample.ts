// packages/anim/src/sample.ts  -- every animated shape's state at a playback position, as overrides for the renderer
//
// The schedule says which behaviours have begun and how far each has got; their values are then composed per
// target in the SMIL sandwich: in the order they began, each replaces or adds to what is below it, starting from
// the shape as it is on the slide. Before anything plays, a target whose first visibility change makes it visible
// (an entrance) is hidden, as PowerPoint shows it.
//
// The renderer's override can move, scale, clip, fade and recolour a whole shape, and each paragraph of its text on
// its own: a paragraph's frame is its laid-out lines' rectangle, which formulas (#ppt_x...) and reveals then use as
// a shape's box. A shape's background (bldP animBg), and a paragraph with no laid-out lines, cannot move on its own,
// so while such a part flies, zooms or wipes in, it fades in with the effect's progress instead.
import { attr, type ColorDef } from '@pptx/model';
import type { ParagraphOverride, RenderShape, RenderSlide, ShapeOverride } from '@pptx/render';
import {
  applyMatrix, hslToRgb, invert, multiply, resolveColor, rgbToHsl, type ColorContext, type Matrix, type Paint, type Rgba,
} from '@pptx/resolve';
import { filterReveal, hashString, type Pt, type Region, type Reveal } from './filters.ts';
import { evaluateFormula, type FormulaScope } from './formula.ts';
import { parseMotionPath, pointAt, type MotionPath } from './path.ts';
import { buildSchedule, type Activation, type Schedule, type ScheduledNode } from './schedule.ts';
import {
  isBehavior, isContainer, walk, type AnimClrNode, type AnimNode, type BehaviorNode, type ContainerNode, type PresetClass, type Target,
  type TimeNode, type Timing, type Variant,
} from './timing.ts';

/** Where playback is: clicks done in the main sequence and the time since the last one (or since the slide began);
 *  shapes clicked to start interactive sequences, in the order they were clicked, each with the time since. */
export interface PlaybackPosition {
  clicks: number;
  ms: number;
  triggers?: { spid: number; ms: number }[];
}

export interface SampleOptions {
  /** The slide's colours, for scheme colours in colour animations (resolve's slideContext(doc, part).colors). */
  colors?: ColorContext;
  /** Makes the Path2D objects for clips; defaults to the global Path2D where there is one. */
  path2d?: () => Path2D;
}

/** Overrides by shape id for the slide's own shapes, and by render shape for SmartArt nodes built one by one. */
export interface AnimationFrame {
  shapes: Map<number, ShapeOverride>;
  nodes: Map<RenderShape, ShapeOverride>;
}

// ------------------------------------------------------------------------------------------------- targets

interface Geom { x: number; y: number; w: number; h: number; }

interface Prepared {
  schedule: Schedule;
  /** The render shape for each target key prefix: 's:5', 'n:5:{GUID}'. */
  shapes: Map<string, RenderShape>;
  keys: Map<BehaviorNode, string[]>;
  initiallyHidden: Set<string>;
  slide: RenderSlide;
}

const cache = new WeakMap<Timing, WeakMap<RenderSlide, Prepared>>();

function indexShapes(slide: RenderSlide): { byId: Map<number, RenderShape>; nodes: Map<string, RenderShape> } {
  const byId = new Map<number, RenderShape>();
  const nodes = new Map<string, RenderShape>();
  const visit = (s: RenderShape, frame: RenderShape | undefined) => {
    if (s.layer === 'slide' && s.source.part === slide.part && !byId.has(s.id)) byId.set(s.id, s);
    const modelId = s.source.part !== slide.part ? attr(s.source.view.el, 'modelId') : undefined;
    if (frame && modelId) nodes.set(`n:${frame.id}:${modelId}`, s);
    for (const c of s.children) visit(c, s.source.graphic?.kind === 'diagram' ? s : frame);
  };
  for (const s of slide.shapes) visit(s, undefined);
  return { byId, nodes };
}

const paragraphCount = (s: RenderShape | undefined) => s?.source.text?.paragraphs.length;
const paragraphText = (s: RenderShape | undefined, i: number) => s?.source.text?.paragraphs[i]?.runs.map(r => r.text).join('') ?? '';
const range = (from: number, to: number) => Array.from({ length: Math.max(0, to - from + 1) }, (_, i) => from + i);

/** The paragraphs a character range covers. */
function paragraphsOfChars(s: RenderShape | undefined, st: number, end: number): number[] {
  const out: number[] = [];
  let at = 0;
  for (let i = 0; i < (paragraphCount(s) ?? 0); i++) {
    const len = paragraphText(s, i).length + 1;
    if (at <= end && at + len > st) out.push(i);
    at += len;
  }
  return out;
}

/** The state keys a target addresses: 's:spid' whole shape, 'b:spid' background, 'p:spid:i' paragraph,
 *  'n:spid:modelId' SmartArt node; `degraded` when a part that cannot be drawn alone stands for the whole shape. */
function keysOf(
  t: Target | undefined, shapes: Map<number, RenderShape>, nodes: Map<string, RenderShape>,
): { keys: string[]; degraded: boolean } {
  if (!t || (t.kind !== 'shape' && t.kind !== 'ink')) return { keys: [], degraded: false };
  const whole = { keys: [`s:${t.spid}`], degraded: false };
  if (t.kind === 'ink' || !t.sub) return whole;
  const sub = t.sub;
  const count = paragraphCount(shapes.get(t.spid));
  switch (sub.kind) {
    case 'bg': return { keys: [`b:${t.spid}`], degraded: false };
    case 'text': {
      const r = sub.range;
      const all = range(0, (count ?? 1) - 1);
      const list = !r || (r.unit === 'char' && r.st >= 4294967295) ? all
        : r.unit === 'char' ? paragraphsOfChars(shapes.get(t.spid), r.st, r.end)
        : range(r.st, Math.min(r.end, (count ?? r.end + 1) - 1, r.st + 1000));
      return { keys: list.map(i => `p:${t.spid}:${i}`), degraded: false };
    }
    case 'subShape': return shapes.has(sub.spid) ? { keys: [`s:${sub.spid}`], degraded: false } : { ...whole, degraded: true };
    case 'graphic': {
      const key = sub.diagram ? `n:${t.spid}:${sub.diagram.id}` : undefined;
      return key && nodes.has(key) ? { keys: [key], degraded: false } : { ...whole, degraded: true };
    }
    default: return { ...whole, degraded: true };
  }
}

const visibilityAttr = (n: BehaviorNode) => n.bhvr.attrNames.includes('style.visibility');

/** How a behaviour changes visibility, read in timeline order to find each target's starting state. */
function showsTarget(n: BehaviorNode, cls: PresetClass | undefined): boolean | undefined {
  if (cls === 'entr') return true;
  if (cls === 'exit') return false;
  if (n.kind === 'set' && visibilityAttr(n)) return n.to?.kind === 'string' ? n.to.value === 'visible' : undefined;
  if (n.kind === 'anim' && visibilityAttr(n)) {
    const first = n.tavs[0]?.val;
    return first?.kind === 'string' ? first.value === 'visible' : undefined;
  }
  if (n.kind === 'animEffect' && n.transition !== 'none') return n.transition === 'in';
  return undefined;
}

/** Behaviours in the order they first play: main sequence step by step, then the interactive sequences. */
function staticOrder(schedule: Schedule): BehaviorNode[] {
  const out: { n: BehaviorNode; key: [number, number, number] }[] = [];
  const visit = (s: ScheduledNode, at: number, step: number) => {
    if (!Number.isFinite(s.begin)) return;
    const begin = at + s.begin;
    if (isBehavior(s.node)) out.push({ n: s.node, key: [step, begin, s.order] });
    for (const c of s.children) visit(c, begin, step);
  };
  const acts: Activation[] = [...schedule.main, ...schedule.interactive.flatMap(i => i.steps)];
  acts.forEach((a, step) => a.roots.forEach(r => visit(r, 0, step)));
  out.sort((a, b) => a.key[0] - b.key[0] || a.key[1] - b.key[1] || a.key[2] - b.key[2]);
  return out.map(o => o.n);
}

/** For effects iterated by word or letter: how many words or letters their text has. */
function iterationCount(shapes: Map<number, RenderShape>): (n: TimeNode) => number {
  return n => {
    const type = n.cTn.iterate?.type;
    if (!type || type === 'el') return 1;
    for (const b of walk([n])) {
      if (!isBehavior(b) || b.bhvr.target?.kind !== 'shape') continue;
      const t = b.bhvr.target;
      const shape = shapes.get(t.spid);
      const r = t.sub?.kind === 'text' && t.sub.range?.unit === 'paragraph' ? t.sub.range : undefined;
      const last = (paragraphCount(shape) ?? 0) - 1;
      const text = range(r ? r.st : 0, r ? Math.min(r.end, last) : last).map(i => paragraphText(shape, i)).join(' ');
      const items = type === 'wd' ? text.split(/\s+/).filter(Boolean).length : text.replace(/\s+/g, '').length;
      return Math.min(10_000, Math.max(1, items));
    }
    return 1;
  };
}

function prepare(timing: Timing, slide: RenderSlide): Prepared {
  let bySlide = cache.get(timing);
  if (!bySlide) cache.set(timing, (bySlide = new WeakMap()));
  const hit = bySlide.get(slide);
  if (hit) return hit;
  const { byId, nodes } = indexShapes(slide);
  const schedule = buildSchedule(timing, iterationCount(byId));
  const effectOf = new Map<TimeNode, ContainerNode>();
  const mark = (list: readonly TimeNode[], effect: ContainerNode | undefined) => {
    for (const n of list) {
      const e = isContainer(n) && n.cTn.presetClass ? n : effect;
      if (e) effectOf.set(n, e);
      mark(n.cTn.children, e);
    }
  };
  mark(timing.nodes, undefined);
  const keys = new Map<BehaviorNode, string[]>();
  const degradedFirst = new Map<number, TimeNode>();
  const initiallyHidden = new Set<string>();
  const decided = new Set<string>();
  for (const n of staticOrder(schedule)) {
    const found = keysOf(n.bhvr.target, byId, nodes);
    // behaviours for Internet Explorer's HTML runtime (rctx="IE") are not PowerPoint's to play
    let use = n.bhvr.rctx === 'IE' ? [] : found.keys;
    if (found.degraded && n.bhvr.target?.kind === 'shape') {
      // a chart series or an unknown part stands for the whole shape, but only in the first effect that names one
      const owner = effectOf.get(n) ?? n;
      const first = degradedFirst.get(n.bhvr.target.spid);
      if (!first) degradedFirst.set(n.bhvr.target.spid, owner);
      else if (first !== owner) use = [];
    }
    keys.set(n, use);
    const shows = showsTarget(n, effectOf.get(n)?.cTn.presetClass);
    if (shows === undefined) continue;
    for (const key of use) {
      if (decided.has(key)) continue;
      decided.add(key);
      if (shows) initiallyHidden.add(key);
    }
  }
  const shapes = new Map<string, RenderShape>(nodes);
  for (const [id, s] of byId) shapes.set(`s:${id}`, s);
  const out: Prepared = { schedule, shapes, keys, initiallyHidden, slide };
  bySlide.set(slide, out);
  return out;
}

/** The schedule a slide's timing plays with, for step counts and durations. */
export function scheduleFor(timing: Timing, slide: RenderSlide): Schedule {
  return prepare(timing, slide).schedule;
}

// ------------------------------------------------------------------------------------------------- playback

interface Contribution {
  node: BehaviorNode;
  /** Progress through the simple duration, after easing (0..1). */
  progress: number;
  active: boolean;
  /** The effect (the preset container) the behaviour belongs to, its class and progress. */
  effect?: TimeNode;
  effectProgress: number;
  effectClass?: PresetClass;
  rank: number;
  begin: number;
  order: number;
}

/** SMIL accelerate/decelerate: constant acceleration over the first `a` of the time, deceleration over the last `d`. */
export function ease(t: number, a: number, d: number): number {
  if (a <= 0 && d <= 0) return t;
  const sum = a + d;
  if (sum > 1) {
    a /= sum;
    d /= sum;
  }
  const r = 1 / (1 - a / 2 - d / 2);
  if (t < a) return (r * t * t) / (2 * a);
  if (t <= 1 - d) return r * (t - a / 2);
  const u = 1 - t;
  return 1 - (r * u * u) / (2 * d);
}

/** tmFilter: progress remapped through (time, progress) pairs, linearly between them. */
function filterTime(p: number, f: [number, number][] | undefined): number {
  if (!f) return p;
  if (p <= f[0]![0]) return f[0]![1];
  for (let i = 1; i < f.length; i++) {
    const [t1, v1] = f[i]!;
    const [t0, v0] = f[i - 1]!;
    if (p <= t1) return t1 > t0 ? v0 + ((v1 - v0) * (p - t0)) / (t1 - t0) : v1;
  }
  return f.at(-1)![1];
}

/** The node's local (simple) time and progress, `a` ms of parent time after it began; Infinity seeks to the end. */
function simpleTime(s: ScheduledNode, a: number, ended: boolean): { time: number; progress: number } {
  const c = s.node.cTn;
  const simple = s.simple;
  let local = a * (Math.abs(c.spd) || 1);
  if (!Number.isFinite(simple)) return { time: local, progress: ended ? 1 : 0 };
  if (simple <= 0) return { time: 0, progress: 1 };
  const period = c.autoRev ? simple * 2 : simple;
  if (!Number.isFinite(local)) local = period;                         // seeking an endless repeat: the end of a pass
  let f = local % period;
  if (ended && local > 0 && f < 1e-9) f = period;
  const t = Math.min(simple, Math.max(0, c.autoRev && f > simple ? period - f : f));
  let p = filterTime(ease(t / simple, c.accel, c.decel), c.tmFilter);
  if (c.spd < 0) p = 1 - p;
  return { time: p * simple, progress: p };
}

/** What a node leaves behind when it ends; behaviours are removed unless they say otherwise, containers hold. */
function fillOf(n: TimeNode): 'remove' | 'hold' {
  const f = n.cTn.fill;
  if (f === 'remove') return 'remove';
  if (f === 'hold' || f === 'freeze' || f === 'transition') return 'hold';
  return isBehavior(n) ? 'remove' : 'hold';
}

function collect(act: Activation, t: number, rank: number, out: Contribution[]): void {
  type Effect = { node?: TimeNode; cls?: PresetClass; progress: number };
  const visit = (s: ScheduledNode, tParent: number, at: number, effect: Effect) => {
    if (!Number.isFinite(s.begin) || !(tParent >= s.begin)) return;
    let a = tParent - s.begin;
    let ended = false;
    if (a >= s.active) {
      if (fillOf(s.node) === 'remove') return;
      a = s.active;
      ended = true;
    }
    const { time, progress } = simpleTime(s, a, ended);
    const begin = at + s.begin;
    const n = s.node;
    if (isBehavior(n)) {
      out.push({
        node: n, progress, active: !ended, effectProgress: effect.progress, rank, begin, order: s.order,
        ...(effect.node ? { effect: effect.node } : {}), ...(effect.cls ? { effectClass: effect.cls } : {}),
      });
      return;
    }
    const cls = n.cTn.presetClass;
    const inner = cls ? { node: n, cls, progress: Number.isFinite(s.simple) ? progress : ended ? 1 : 0 } : effect;
    for (const c of s.children) visit(c, time * s.stretch, begin, inner);
  };
  for (const r of act.roots) visit(r, t, 0, { progress: 1 });
}

/** Activations to play, oldest first: completed ones play to their end (Infinity), the latest ones up to now. */
function activations(prep: Prepared, at: PlaybackPosition): { act: Activation; t: number }[] {
  const main = prep.schedule.main;
  const clicks = Math.max(0, Math.floor(at.clicks));
  const out: { act: Activation; t: number }[] = [];
  for (let k = 0; k < Math.min(clicks, main.length); k++) out.push({ act: main[k]!, t: Infinity });
  const timed: { act: Activation; t: number; start: number }[] = [];
  const current = main[clicks];
  if (current) timed.push({ act: current, t: Math.max(0, at.ms), start: -Math.max(0, at.ms) });
  for (const seq of prep.schedule.interactive) {
    if (!seq.steps.length) continue;
    const mine = (at.triggers ?? []).filter(t => t.spid === seq.trigger);
    // each click plays the next group, and starts over after the last; a later click finishes the one before
    mine.forEach((click, j) => {
      const ms = Math.max(0, click.ms);
      timed.push({ act: seq.steps[j % seq.steps.length]!, t: j === mine.length - 1 ? ms : Infinity, start: -ms });
    });
  }
  timed.sort((a, b) => a.start - b.start);
  return [...out, ...timed];
}

// ------------------------------------------------------------------------------------------------- states

/** What a state moves and clips: a box (bw x bh) placed on the slide by X, and its drawn extent in that box. */
interface Frame { X: Matrix; bw: number; bh: number; vb: Bounds; }

interface KeyState {
  shape?: RenderShape;
  /** A paragraph's own frame; whole shapes use their box. */
  frame?: Frame;
  base: Geom;
  visible: boolean;
  opacity: number;
  fade: number;
  x: number;
  y: number;
  w: number;
  h: number;
  dx: number;
  dy: number;
  rot: number;
  sx: number;
  sy: number;
  skewX: number;
  skewY: number;
  reveal?: Reveal;
  fill?: Rgba | 'none';
  line?: Rgba | 'none';
  text?: Rgba;
  /** For paragraphs and backgrounds: how present each moving or wiping effect has made the part, 0..1. */
  presence: Map<unknown, number>;
}

// Maps, not object literals: attribute names come from files and must never reach Object.prototype
const NUMERIC = new Map<string, 'x' | 'y' | 'w' | 'h' | 'rot' | 'opacity' | 'skewX' | 'skewY'>([
  ['ppt_x', 'x'], ['ppt_y', 'y'], ['ppt_w', 'w'], ['ppt_h', 'h'], ['r', 'rot'], ['style.rotation', 'rot'], ['style.opacity', 'opacity'],
  ['xshear', 'skewX'], ['yshear', 'skewY'],
]);
const GEOMETRIC = new Set(['x', 'y', 'w', 'h', 'rot', 'skewX', 'skewY']);
const COLOR_ATTRS = new Map<string, 'fill' | 'line' | 'text'>([
  ['fillcolor', 'fill'], ['fill.color', 'fill'], ['stroke.color', 'line'], ['style.color', 'text'],
]);

/** Where a shape is on the slide, as the formula variables give it: its centre and size, as fractions of the slide. */
function geometryOf(s: RenderShape | undefined, slide: RenderSlide): Geom {
  if (!s) return { x: 0.5, y: 0.5, w: 0, h: 0 };
  const m = s.transform;
  const [cx, cy] = applyMatrix(m, s.box.w / 2, s.box.h / 2);
  const W = slide.size.cx || 1;
  const H = slide.size.cy || 1;
  return { x: cx / W, y: cy / H, w: (Math.hypot(m[0], m[1]) * s.box.w) / W, h: (Math.hypot(m[2], m[3]) * s.box.h) / H };
}

/** Where a frame is on the slide, as the formula variables give it: its centre and size, as fractions of the slide. */
function geometryOfFrame(f: Frame, slide: RenderSlide): Geom {
  const [cx, cy] = applyMatrix(f.X, f.bw / 2, f.bh / 2);
  const W = slide.size.cx || 1;
  const H = slide.size.cy || 1;
  return { x: cx / W, y: cy / H, w: (Math.hypot(f.X[0], f.X[1]) * f.bw) / W, h: (Math.hypot(f.X[2], f.X[3]) * f.bh) / H };
}

/**
 * A paragraph's frame: the rectangle its laid-out lines take in the text box (from the leftmost glyph run to the
 * rightmost, first line's top to last line's bottom), placed by the text's transform. Undefined without lines.
 */
function paragraphFrame(s: RenderShape | undefined, index: number): Frame | undefined {
  const text = s?.prims.find(p => p.kind === 'text');
  if (!text || text.kind !== 'text') return undefined;
  const lines = text.layout.lines.filter(l => l.paragraph === index);
  if (!lines.length) return undefined;
  let x0 = Infinity, x1 = -Infinity;
  for (const l of lines) {
    for (const r of l.runs) {
      x0 = Math.min(x0, r.x);
      x1 = Math.max(x1, r.x + r.advance);
    }
  }
  if (!Number.isFinite(x0)) {
    x0 = lines[0]!.left;
    x1 = x0 + 1;
  }
  const y0 = Math.min(...lines.map(l => l.top)), y1 = Math.max(...lines.map(l => l.top + l.height));
  const bw = Math.max(1, x1 - x0), bh = Math.max(1, y1 - y0);
  const X = multiply(text.transform, translate(x0, y0));
  const k = Math.hypot(X[0], X[1]) || 1;
  const m = 12700 / k;
  return { X, bw, bh, vb: { x0: -m, y0: -m, x1: bw + m, y1: bh + m } };
}

function newState(key: string, prep: Prepared): KeyState {
  const parts = key.split(':');
  const shape = prep.shapes.get(parts[0] === 'n' ? key : `s:${parts[1]}`);
  const frame = parts[0] === 'p' ? paragraphFrame(shape, Number(parts[2])) : undefined;
  const base = frame ? geometryOfFrame(frame, prep.slide) : geometryOf(shape, prep.slide);
  return {
    ...(shape ? { shape } : {}), ...(frame ? { frame } : {}), base, ...base, visible: !prep.initiallyHidden.has(key), opacity: 1, fade: 1,
    dx: 0, dy: 0, rot: 0, sx: 1, sy: 1, skewX: 0, skewY: 0, presence: new Map(),
  };
}

const NO_COLORS: ColorContext = { theme: new Map(), clrMap: {} };
const lerp = (a: number, b: number, t: number) => a + (b - a) * t;
const clamp01 = (v: number) => (v < 0 ? 0 : v > 1 ? 1 : v);

function paintColor(p: Paint | undefined): Rgba | undefined {
  if (!p) return undefined;
  if (p.kind === 'solid') return p.color;
  if (p.kind === 'linear' || p.kind === 'path') return p.stops[0]?.color;
  if (p.kind === 'pattern') return p.fg;
  return undefined;
}

/** The colour an animation without a "from" starts at: what is there now, else the shape's own. */
function baseColor(st: KeyState, which: 'fill' | 'line' | 'text'): Rgba {
  const cur = st[which];
  if (cur && cur !== 'none') return cur;
  const src = st.shape?.source;
  const own = which === 'fill' ? src?.fill : which === 'line' ? src?.line?.paint : src?.text?.paragraphs[0]?.runs[0]?.props.fill;
  return paintColor(own) ?? (which === 'fill' ? { r: 1, g: 1, b: 1, a: 1 } : { r: 0, g: 0, b: 0, a: 1 });
}

function mixColor(from: Rgba, to: Rgba, t: number, space: 'rgb' | 'hsl', dir: 'cw' | 'ccw'): Rgba {
  if (space === 'rgb') return { r: lerp(from.r, to.r, t), g: lerp(from.g, to.g, t), b: lerp(from.b, to.b, t), a: lerp(from.a, to.a, t) };
  const a = rgbToHsl(from.r, from.g, from.b);
  const b = rgbToHsl(to.r, to.g, to.b);
  let dh = b.h - a.h;
  if (dir === 'cw' && dh < 0) dh += 360;
  if (dir === 'ccw' && dh > 0) dh -= 360;
  const [r, g, bl] = hslToRgb(a.h + dh * t, lerp(a.s, b.s, t), lerp(a.l, b.l, t));
  return { r, g, b: bl, a: lerp(from.a, to.a, t) };
}

interface Env { colors: ColorContext; seed: number; }

const scopeOf = (st: KeyState, input?: number): FormulaScope =>
  ({ ppt_x: st.base.x, ppt_y: st.base.y, ppt_w: st.base.w, ppt_h: st.base.h, ...(input !== undefined ? { input } : {}) });

function formulaValue(src: string | undefined, st: KeyState): number | undefined {
  if (src === undefined) return undefined;
  const n = evaluateFormula(src, scopeOf(st));
  return Number.isFinite(n) ? n : undefined;
}

function variantNumber(v: Variant | undefined, st: KeyState): number | undefined {
  if (v?.kind === 'float' || v?.kind === 'int') return v.value;
  if (v?.kind === 'bool') return v.value ? 1 : 0;
  return v?.kind === 'string' ? formulaValue(v.value, st) : undefined;
}

/** Set one attribute: numbers replace, add or multiply; strings and colours replace. True when the shape moved. */
function assign(st: KeyState, name: string, value: number | string | Rgba | undefined, mode: 'replace' | 'add' | 'mult'): boolean {
  if (value === undefined) return false;
  if (name === 'style.visibility') {
    if (typeof value === 'string') st.visible = value !== 'hidden';
    return false;
  }
  const numeric = NUMERIC.get(name);
  if (numeric) {
    if (typeof value !== 'number' || !Number.isFinite(value)) return false;
    st[numeric] = mode === 'add' ? st[numeric] + value : mode === 'mult' ? st[numeric] * value : value;
    return GEOMETRIC.has(numeric);
  }
  const color = COLOR_ATTRS.get(name);
  if (color && typeof value === 'object') st[color] = value;
  if (name === 'fill.on' && typeof value === 'string') st.fill = value === 'false' ? 'none' : st.fill === 'none' ? undefined : st.fill;
  if (name === 'stroke.on' && typeof value === 'string') st.line = value === 'false' ? 'none' : st.line === 'none' ? undefined : st.line;
  return false;
}

const keyTimes = (n: AnimNode) => n.tavs.map((t, i) => t.tm ?? (n.tavs.length > 1 ? i / (n.tavs.length - 1) : 0));

function kindOf(n: AnimNode, name: string): 'num' | 'str' | 'clr' {
  if (n.valueType) return n.valueType;
  if (name === 'style.visibility' || name === 'fill.on' || name === 'stroke.on' || name === 'fill.type') return 'str';
  return COLOR_ATTRS.has(name) ? 'clr' : 'num';
}

type Value = { value: number | string | Rgba | undefined; mode: 'replace' | 'add' | 'mult' };

/**
 * An anim's value at progress p, with how it combines with what is below it. Keyframes interpolate linearly (or
 * step, for discrete and string values); a keyframe's formula takes the value interpolated from it to the next as $.
 * Without keyframes: from, to and by, a missing from starting at the current value.
 */
function animValue(n: AnimNode, name: string, p: number, st: KeyState, env: Env): Value {
  const kind = kindOf(n, name);
  const mode = n.bhvr.additive === 'sum' ? 'add' : n.bhvr.additive === 'mult' ? 'mult' : 'replace';
  if (n.tavs.length) {
    const times = keyTimes(n);
    if (kind !== 'num' || n.calcmode === 'discrete') {
      let pick = 0;
      for (let i = 0; i < times.length; i++) if (times[i]! <= p + 1e-9) pick = i;
      const v = n.tavs[pick]!.val;
      if (kind === 'clr') return { value: v?.kind === 'color' ? resolveColor(v.value, env.colors) : undefined, mode };
      return { value: kind === 'str' ? (v?.kind === 'string' ? v.value : undefined) : variantNumber(v, st), mode };
    }
    let i = 0;
    while (i < times.length - 2 && p > times[i + 1]!) i++;
    const j = Math.min(i + 1, n.tavs.length - 1);
    const a = n.tavs[i]!;
    const va = variantNumber(a.val, st);
    const vb = variantNumber(n.tavs[j]!.val, st) ?? va;
    if (va === undefined || vb === undefined) return { value: undefined, mode };
    const ta = times[i]!;
    const tb = times[j]!;
    let v = lerp(va, vb, tb > ta ? clamp01((p - ta) / (tb - ta)) : p >= tb ? 1 : 0);
    if (a.fmla) {
      const r = evaluateFormula(a.fmla, scopeOf(st, v));
      if (Number.isFinite(r)) v = r;
    }
    return { value: v, mode };
  }
  const from = n.from ?? n.bhvr.from;
  const to = n.to ?? n.bhvr.to;
  if (kind === 'str') return { value: p >= 1 || from === undefined ? to : from, mode: 'replace' };
  if (kind === 'clr') return { value: undefined, mode };
  const fromV = formulaValue(from, st);
  const toV = formulaValue(to, st);
  const byV = formulaValue(n.by ?? n.bhvr.by, st);
  if (toV !== undefined) {
    const k = NUMERIC.get(name);
    const start = fromV ?? (k ? st[k] : undefined);
    return { value: start === undefined ? toV : lerp(start, toV, p), mode };
  }
  if (byV !== undefined) return fromV !== undefined ? { value: fromV + byV * p, mode } : { value: byV * p, mode: 'add' };
  return { value: fromV, mode };
}

const colorOf = (def: ColorDef | undefined, env: Env) => (def ? resolveColor(def, env.colors) : undefined);

function animClrValue(n: AnimClrNode, name: string, p: number, st: KeyState, env: Env): Rgba | undefined {
  const which = COLOR_ATTRS.get(name);
  if (!which) return undefined;
  const from = colorOf(n.from, env) ?? baseColor(st, which);
  let to = colorOf(n.to, env);
  if (!to && n.by?.space === 'rgb') to = { r: clamp01(from.r + n.by.r), g: clamp01(from.g + n.by.g), b: clamp01(from.b + n.by.b), a: from.a };
  else if (!to && n.by?.space === 'hsl') {
    const h = rgbToHsl(from.r, from.g, from.b);
    const [r, g, b] = hslToRgb(h.h + n.by.h, clamp01(h.s + n.by.s), clamp01(h.l + n.by.l));
    to = { r, g, b, a: from.a };
  }
  return mixColor(from, to ?? from, p, n.clrSpc ?? 'rgb', n.dir ?? 'cw');
}

const motionPaths = new Map<string, MotionPath | null>();
function motionPath(src: string): MotionPath | undefined {
  let p = motionPaths.get(src);
  if (p === undefined) {
    p = parseMotionPath(src) ?? null;
    if (motionPaths.size > 1000) motionPaths.clear();
    motionPaths.set(src, p);
  }
  return p ?? undefined;
}

/** The real width / height of the box a filter works in, so circles and wedges keep their angles. */
function aspectOf(st: KeyState): number {
  const s = st.shape;
  if (!s) return 1;
  const vb = visualBounds(s);
  const w = (vb.x1 - vb.x0) * Math.hypot(s.transform[0], s.transform[1]);
  const h = (vb.y1 - vb.y0) * Math.hypot(s.transform[2], s.transform[3]);
  return w > 0 && h > 0 ? w / h : 1;
}

function regionArea(r: Region): number {
  let area = 0;
  for (const poly of r) {
    for (let i = 0; i < poly.length; i++) {
      const p = poly[i]!;
      const q = poly[(i + 1) % poly.length]!;
      area += (p.x * q.y - q.x * p.y) / 2;
    }
  }
  return clamp01(area);
}

function motionPoint(n: Extract<BehaviorNode, { kind: 'animMotion' }>, p: number): Pt {
  const path = n.path ? motionPath(n.path) : undefined;
  if (path) return pointAt(path, p);
  const from = n.from ?? { x: 0, y: 0 };
  const to = n.to ?? (n.by ? { x: from.x + n.by.x, y: from.y + n.by.y } : from);
  return { x: lerp(from.x, to.x, p), y: lerp(from.y, to.y, p) };
}

/** Apply one behaviour to a target's state. `part`: the target is a paragraph or a background. */
function applyContribution(c: Contribution, st: KeyState, part: boolean, env: Env): void {
  const n = c.node;
  const p = c.progress;
  let geometric = false;
  switch (n.kind) {
    case 'set':
      for (const name of n.bhvr.attrNames) {
        const v = n.to;
        const value = v?.kind === 'color' ? resolveColor(v.value, env.colors)
          : NUMERIC.has(name) ? variantNumber(v, st)
          : v?.kind === 'string' ? v.value : v?.kind === 'bool' ? String(v.value) : undefined;
        geometric = assign(st, name, value, 'replace') || geometric;
      }
      break;
    case 'anim': {
      const name = n.bhvr.attrNames[0];
      if (!name) break;
      const { value, mode } = animValue(n, name, p, st, env);
      geometric = assign(st, name, value, mode);
      break;
    }
    case 'animClr':
      for (const name of n.bhvr.attrNames) assign(st, name, animClrValue(n, name, p, st, env), 'replace');
      break;
    case 'animEffect': {
      // "image" is a static filter (Transparency's opacity), not a reveal
      if (n.transition === 'none' || /^\s*image\b/.test(n.filter ?? '')) {
        const m = /opacity\s*:\s*([0-9.]+)/.exec(n.prLst ?? '');
        if (m) st.fade *= clamp01(Number(m[1]));
        break;
      }
      const progress = n.progress ? variantNumber(n.progress, st) ?? p : p;
      const reveal = filterReveal(n.filter, progress, n.transition, aspectOf(st), env.seed ^ hashString(String(c.order)));
      if (reveal.opacity !== undefined) st.fade *= reveal.opacity;
      if (reveal.region && !part) st.reveal = reveal;
      else if (reveal.region && c.active) st.presence.set(n, regionArea(reveal.region));
      break;
    }
    case 'animMotion': {
      const pt = motionPoint(n, p);
      st.dx += pt.x;
      st.dy += pt.y;
      geometric = pt.x !== 0 || pt.y !== 0;
      break;
    }
    case 'animRot': {
      const from = n.from ?? 0;
      const delta = n.to !== undefined ? lerp(from, n.to, p) : from + (n.by ?? 0) * p;
      st.rot += delta;
      geometric = delta !== 0;
      break;
    }
    case 'animScale': {
      const from = n.from ?? { x: 1, y: 1 };
      const to = n.to ?? (n.by ? { x: from.x * n.by.x, y: from.y * n.by.y } : from);
      st.sx *= lerp(from.x, to.x, p);
      st.sy *= lerp(from.y, to.y, p);
      geometric = true;
      break;
    }
    default:
      break;
  }
  // a part that cannot move on its own shows as much of itself as its effect has brought in
  if (part && geometric && c.active) {
    const presence = c.effectClass === 'entr' ? c.effectProgress : c.effectClass === 'exit' ? 1 - c.effectProgress : 1;
    st.presence.set(c.effect ?? n, clamp01(presence));
  }
}

// ------------------------------------------------------------------------------------------------ overrides

const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
const isIdentity = (m: Matrix) => Math.abs(m[0] - 1) < 1e-9 && Math.abs(m[1]) < 1e-9 && Math.abs(m[2]) < 1e-9 &&
  Math.abs(m[3] - 1) < 1e-9 && Math.abs(m[4]) < 0.5 && Math.abs(m[5]) < 0.5;

/**
 * The shape's box to slide, without its own mirroring. Effects work on the shape as drawn, so a wipe from the bottom
 * starts at the bottom of a flipped shape too, while a rotated shape is wiped along its own axes.
 */
function frameOf(s: RenderShape): Matrix {
  const x = s.source.xfrm;
  if (!x.flipH && !x.flipV) return s.transform;
  return multiply(s.transform, [x.flipH ? -1 : 1, 0, 0, x.flipV ? -1 : 1, x.flipH ? s.box.w : 0, x.flipV ? s.box.h : 0]);
}

interface Bounds { x0: number; y0: number; x1: number; y1: number; }
const boundsCache = new WeakMap<RenderShape, Bounds>();

/** The shape's drawn extent in its (unmirrored) box coordinates: the box, its geometry (a callout's tail reaches
 *  outside the box), its text (which may overflow) and its members. */
function visualBounds(s: RenderShape): Bounds {
  const hit = boundsCache.get(s);
  if (hit) return hit;
  const inv = invert(frameOf(s));
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  const point = (sx: number, sy: number) => {
    const [x, y] = applyMatrix(inv, sx, sy);
    x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y);
  };
  const add = (m: Matrix, w: number, h: number) => {
    for (const [px, py] of [[0, 0], [w, 0], [w, h], [0, h]] as const) point(...applyMatrix(m, px, py));
  };
  add(s.transform, s.box.w, s.box.h);
  for (const spec of s.outline) for (const c of spec.commands) if (c.op !== 'Z') point(...applyMatrix(s.transform, c.x, c.y));
  let stroke = 0;
  for (const p of s.prims) {
    if (p.kind === 'text') add(p.transform, p.layout.width, p.layout.height);
    else if (p.kind === 'path' && p.stroke) stroke = Math.max(stroke, p.stroke.width);
  }
  for (const c of s.children) add(c.transform, c.box.w, c.box.h);
  // half the line width, and at least a point, so lines (no height) can be wiped too
  const k = Math.hypot(s.transform[0], s.transform[1]) || 1;
  const m = Math.max(stroke / k, 12700 / k);
  const out = { x0: x0 - m, y0: y0 - m, x1: x1 + m, y1: y1 + m };
  boundsCache.set(s, out);
  return out;
}

const shapeFrame = (s: RenderShape): Frame => ({ X: frameOf(s), bw: s.box.w, bh: s.box.h, vb: visualBounds(s) });

/** The slide-space matrix for a state (a whole shape, or a paragraph), and the one its clip region is mapped with. */
function matrices(st: KeyState, slide: RenderSlide): { content: Matrix; clip: Matrix } | undefined {
  const f = st.frame ?? (st.shape ? shapeFrame(st.shape) : undefined);
  if (!f) return undefined;
  const X = f.X;
  const bw = f.bw;
  const bh = f.bh;
  const [c0x, c0y] = applyMatrix(X, bw / 2, bh / 2);
  // size (ppt_w, ppt_h, animScale) and shear in the shape's own axes, about its centre
  const kx = (st.base.w > 1e-9 ? st.w / st.base.w : 1) * st.sx;
  const ky = (st.base.h > 1e-9 ? st.h / st.base.h : 1) * st.sy;
  const box = multiply(multiply(translate(bw / 2, bh / 2), [kx, st.skewY, st.skewX, ky, 0, 0]), translate(-bw / 2, -bh / 2));
  // then turned about the centre and moved to where ppt_x, ppt_y and motion paths put it
  const rad = (st.rot * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const rotate: Matrix = [cos, sin, -sin, cos, c0x - cos * c0x + sin * c0y, c0y - sin * c0x - cos * c0y];
  const outer = multiply(translate((st.x + st.dx) * slide.size.cx - c0x, (st.y + st.dy) * slide.size.cy - c0y), rotate);
  // a slide filter moves the content inside a clip that stays put
  const vb = f.vb;
  const shift = st.reveal?.shift;
  const inBox = shift ? multiply(box, translate(shift.x * (vb.x1 - vb.x0), shift.y * (vb.y1 - vb.y0))) : box;
  const unit: Matrix = [vb.x1 - vb.x0, 0, 0, vb.y1 - vb.y0, vb.x0, vb.y0];
  return { content: multiply(multiply(outer, X), multiply(inBox, invert(X))), clip: multiply(multiply(outer, X), multiply(box, unit)) };
}

function regionPath(region: Region, m: Matrix, make: () => Path2D): Path2D {
  const path = make();
  for (const poly of region) {
    poly.forEach((p, i) => {
      const [x, y] = applyMatrix(m, p.x, p.y);
      if (i === 0) path.moveTo(x, y);
      else path.lineTo(x, y);
    });
    path.closePath();
  }
  return path;
}

const isFull = (r: Region) => r.length === 1 && regionArea(r) > 0.9999;

function fadePaint(p: Paint, alpha: number): Paint {
  const k = (c: Rgba): Rgba => ({ ...c, a: c.a * alpha });
  switch (p.kind) {
    case 'solid': return { kind: 'solid', color: k(p.color) };
    case 'linear': case 'path': return { ...p, stops: p.stops.map(s => ({ pos: s.pos, color: k(s.color) })) };
    case 'pattern': return { ...p, fg: k(p.fg), bg: k(p.bg) };
    case 'image': return alpha < 0.5 ? { kind: 'none' } : p;
    default: return p;
  }
}

const TRANSPARENT: Rgba = { r: 0, g: 0, b: 0, a: 0 };
const presenceOf = (st: KeyState) => [...st.presence.values()].reduce((a, b) => a * b, 1);

function wholeOverride(st: KeyState, o: ShapeOverride, slide: RenderSlide, make: (() => Path2D) | undefined): void {
  if (!st.visible) {
    o.hidden = true;
    return;
  }
  const opacity = clamp01(st.opacity * st.fade);
  if (opacity < 1) o.opacity = opacity;
  const m = matrices(st, slide);
  if (m && !isIdentity(m.content)) o.matrix = m.content;
  const region = st.reveal?.region;
  if (m && region && !isFull(region) && make) o.clip = regionPath(region, m.clip, make);
  if (st.fill) o.fill = st.fill === 'none' ? { kind: 'none' } : { kind: 'solid', color: st.fill };
  if (st.line) o.lineColor = st.line === 'none' ? TRANSPARENT : st.line;
  if (st.text) o.textColor = st.text;
}

/** A paragraph's own state: hidden, faded, moved, wiped or recoloured, like a whole shape's. */
function paragraphOverride(st: KeyState, slide: RenderSlide, make: (() => Path2D) | undefined): ParagraphOverride {
  if (!st.visible) return { hidden: true };
  const o: ParagraphOverride = {};
  const opacity = clamp01(st.opacity * st.fade);
  if (opacity < 1) o.opacity = opacity;
  const m = matrices(st, slide);
  if (m && !isIdentity(m.content)) o.matrix = m.content;
  const region = st.reveal?.region;
  if (m && region && !isFull(region) && make) o.clip = regionPath(region, m.clip, make);
  // on text, a fill colour is the colour of the letters
  const color = st.text ?? (st.fill && st.fill !== 'none' ? st.fill : undefined);
  if (color) o.textColor = color;
  return o;
}

/** A shape's background animated apart from its text (bldP animBg): its fill and line, faded. */
function backgroundOverride(st: KeyState, o: ShapeOverride): void {
  const alpha = st.visible ? clamp01(st.opacity * st.fade * presenceOf(st)) : 0;
  const src = st.shape?.source;
  const fill: Paint | undefined = st.fill === 'none' ? { kind: 'none' } : st.fill ? { kind: 'solid', color: st.fill } : src?.fill;
  if (fill && fill.kind !== 'none' && (alpha < 1 || st.fill)) o.fill = fadePaint(fill, alpha);
  else if (st.fill === 'none') o.fill = { kind: 'none' };
  const line = st.line === 'none' ? TRANSPARENT : st.line ?? paintColor(src?.line?.paint);
  if (line && (alpha < 1 || st.line)) o.lineColor = { ...line, a: line.a * alpha };
}

const pathFactory = (): (() => Path2D) | undefined => {
  const P = (globalThis as { Path2D?: new () => Path2D }).Path2D;
  return P ? () => new P() : undefined;
};

/** The override colours all of a shape's text or none: it does when the paragraphs recoloured are all those with
 *  text, in one colour. */
function wholeTextColor(shape: RenderShape | undefined, colors: Map<number, Rgba>): Rgba | undefined {
  const withText = (shape?.source.text?.paragraphs ?? []).filter(p => p.runs.some(r => r.text.trim() !== '')).map(p => p.index);
  const first = colors.get(withText[0] ?? -1);
  if (!first) return undefined;
  const same = (c: Rgba | undefined) => !!c && c.r === first.r && c.g === first.g && c.b === first.b && c.a === first.a;
  return withText.every(i => same(colors.get(i))) ? first : undefined;
}

/** The overrides for every animated shape of a slide at a playback position. */
export function sampleFrame(timing: Timing | undefined, slide: RenderSlide, at: PlaybackPosition, options: SampleOptions = {}): AnimationFrame {
  const frame: AnimationFrame = { shapes: new Map(), nodes: new Map() };
  if (!timing) return frame;
  const prep = prepare(timing, slide);
  const contributions: Contribution[] = [];
  activations(prep, at).forEach(({ act, t }, rank) => collect(act, t, rank, contributions));
  contributions.sort((a, b) => a.rank - b.rank || a.begin - b.begin || a.order - b.order);
  const states = new Map<string, KeyState>();
  const stateOf = (key: string) => {
    let st = states.get(key);
    if (!st) states.set(key, (st = newState(key, prep)));
    return st;
  };
  for (const key of prep.initiallyHidden) stateOf(key);
  const env: Env = { colors: options.colors ?? NO_COLORS, seed: 0x9e3779b9 };
  for (const c of contributions) {
    for (const key of prep.keys.get(c.node) ?? []) {
      const st = stateOf(key);
      // a background, or a paragraph that has no lines to move, can only fade
      applyContribution(c, st, key[0] === 'b' || (key[0] === 'p' && !st.frame), env);
    }
  }

  const make = options.path2d ?? pathFactory();
  const textColors = new Map<number, Map<number, Rgba>>();
  const overrideFor = (spid: number) => {
    let o = frame.shapes.get(spid);
    if (!o) frame.shapes.set(spid, (o = {}));
    return o;
  };
  for (const [key, st] of states) {
    const [kind, id, rest] = key.split(':');
    const spid = Number(id);
    if (kind === 's') wholeOverride(st, overrideFor(spid), slide, make);
    else if (kind === 'b') backgroundOverride(st, overrideFor(spid));
    else if (kind === 'n' && st.shape) {
      const o: ShapeOverride = {};
      wholeOverride(st, o, slide, make);
      if (Object.keys(o).length) frame.nodes.set(st.shape, o);
    } else if (kind === 'p' && st.frame) {
      const po = paragraphOverride(st, slide, make);
      if (Object.keys(po).length) {
        const o = overrideFor(spid);
        o.paragraphs = new Map(o.paragraphs ?? []).set(Number(rest), po);
      }
    } else if (kind === 'p') {
      // on text, a fill colour is the colour of the letters
      const color = st.text ?? (st.fill && st.fill !== 'none' ? st.fill : undefined);
      if (color) {
        if (!textColors.has(spid)) textColors.set(spid, new Map());
        textColors.get(spid)!.set(Number(rest), color);
      }
      const opacity = st.visible ? clamp01(st.opacity * st.fade * presenceOf(st)) : 0;
      if (opacity >= 1) continue;
      const o = overrideFor(spid);
      o.paragraphs = new Map(o.paragraphs ?? []).set(Number(rest), opacity <= 0 ? { hidden: true } : { opacity });
    }
  }
  for (const [spid, colors] of textColors) {
    const color = wholeTextColor(prep.shapes.get(`s:${spid}`), colors);
    if (color) overrideFor(spid).textColor = color;
  }
  for (const [spid, o] of frame.shapes) if (!Object.keys(o).length) frame.shapes.delete(spid);
  return frame;
}

/** The overrides by shape id for the slide's shapes at a playback position (sampleFrame's `shapes`). */
export function sample(
  timing: Timing | undefined, slide: RenderSlide, at: PlaybackPosition, options: SampleOptions = {},
): Map<number, ShapeOverride> {
  return sampleFrame(timing, slide, at, options).shapes;
}

/** DrawOptions.override for a frame: slide shapes by id (master and layout shapes keep their own ids out of it),
 *  SmartArt nodes by identity. */
export function overrideLookup(
  frame: AnimationFrame | Map<number, ShapeOverride>, slide: RenderSlide,
): (shape: RenderShape) => ShapeOverride | undefined {
  const f: AnimationFrame = frame instanceof Map ? { shapes: frame, nodes: new Map() } : frame;
  const own = (shape: RenderShape) => shape.layer === 'slide' && shape.source.part === slide.part;
  return shape => f.nodes.get(shape) ?? (own(shape) ? f.shapes.get(shape.id) : undefined);
}
