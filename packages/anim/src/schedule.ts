// packages/anim/src/schedule.ts  -- when each time node plays: the main sequence's click steps and the trigger steps
//
// SMIL as PowerPoint plays it. A node's begin comes from its start conditions, relative to its parent: children of
// a par start with the parent plus their delay; children of a seq start when the previous child ends; a delay of
// "indefinite" waits for a click (the main sequence's next, or the trigger shape of an interactive sequence), and
// onBegin/onEnd conditions follow other nodes. Durations come from dur, or for containers from their children;
// repeatCount, repeatDur, autoRev and spd stretch them. Clicks seek: when the next click comes, everything before it
// has played to its end. Times are milliseconds; a node's begin and end are in its parent's local time.
import {
  interactiveSequences, isBehavior, isContainer, mainSequence, rootOf, type Condition, type SeqNode, type TimeNode, type Timing,
} from './timing.ts';

export interface ScheduledNode {
  node: TimeNode;
  /** In the parent's local time; Infinity when nothing in this activation starts the node. */
  begin: number;
  /** Simple duration in the node's own time; Infinity for indefinite. */
  simple: number;
  /** Active duration in the parent's time: the simple duration repeated, reversed and sped up as the node says. */
  active: number;
  end: number;
  /** Iterated effects (by word, by letter) run their children this much faster than the effect lasts. */
  stretch: number;
  children: ScheduledNode[];
  /** Document order, for ties between nodes that begin together. */
  order: number;
}

export interface Activation {
  roots: ScheduledNode[];
  /** When everything in it has ended (Infinity when something never does). */
  duration: number;
}

export interface Schedule {
  /** steps[0] plays when the slide starts (often nothing); steps[k] after the k-th click. */
  main: Activation[];
  /** One per interactive sequence: steps[k] plays after the (k+1)-th click on the trigger shape. */
  interactive: { trigger: number; seq: SeqNode; steps: Activation[] }[];
}

/** How many items an iterated effect runs over (letters or words of its text); 1 when unknown. */
export type IterationCount = (node: TimeNode) => number;

/** Conditions naming siblings that come later are resolved by going round again, at most this often. */
const MAX_ROUNDS = 8;
const MAX_DEPTH = 64;

interface Ctx { orders: Map<TimeNode, number>; iterations: IterationCount; depth: number; }

const finite = (v: number) => Number.isFinite(v);
const timeOf = (t: number | 'indefinite' | undefined) => (t === undefined ? undefined : t === 'indefinite' ? Infinity : t);

/** A condition's time in the parent's local timeline, if it can be known without a click or another event. */
type Siblings = Map<number, ScheduledNode>;

function conditionTime(c: Condition, base: number, siblings: Siblings, ancestors: ReadonlySet<number>): number | undefined {
  const delay = timeOf(c.delay) ?? 0;
  if (!finite(delay)) return undefined;
  if (c.evt === undefined && !c.target && c.tn === undefined) return base + delay;
  if (c.tn === undefined) return undefined;
  if (c.evt === 'onBegin' || c.evt === 'begin') {
    if (ancestors.has(c.tn)) return delay;
    const s = siblings.get(c.tn);
    return s && finite(s.begin) ? s.begin + delay : undefined;
  }
  if (c.evt === 'onEnd' || c.evt === 'end') {
    const s = siblings.get(c.tn);
    return s && finite(s.end) ? s.end + delay : undefined;
  }
  return undefined;
}

const earliest = (conds: Condition[], base: number, siblings: Siblings, ancestors: ReadonlySet<number>) => {
  let best: number | undefined;
  for (const c of conds) {
    const t = conditionTime(c, base, siblings, ancestors);
    if (t !== undefined && (best === undefined || t < best)) best = t;
  }
  return best;
};

/** Whether a node waits for a click: none of its start conditions can resolve without one. */
export function waitsForClick(n: TimeNode): boolean {
  const start = n.cTn.start;
  const click = (c: Condition) => c.delay === 'indefinite' || c.evt === 'onNext' || c.evt === 'onClick' || c.evt === 'onPrev';
  return start.length > 0 && start.every(click);
}

/** The active duration (in the parent's time) for a simple duration, from repeatCount, repeatDur, autoRev and spd. */
function activeDuration(n: TimeNode, simple: number): number {
  const c = n.cTn;
  const count = c.repeatCount === undefined ? 1 : c.repeatCount === 'indefinite' ? Infinity : c.repeatCount;
  let local = count === 0 ? 0 : (c.autoRev ? simple * 2 : simple) * count;
  if (Number.isNaN(local)) local = 0;                                  // 0 × Infinity: nothing to repeat
  const repeatDur = timeOf(c.repeatDur);
  if (repeatDur !== undefined) local = c.repeatCount === undefined ? repeatDur : Math.min(local, repeatDur);
  return local / (Math.abs(c.spd) || 1);
}

/** A container's duration from its children: when the last one ends (endSync first: the first one). */
function implicitDuration(n: TimeNode, children: ScheduledNode[]): number {
  const started = children.filter(ch => finite(ch.begin));
  if (!started.length) return n.kind === 'seq' && children.length ? Infinity : 0;
  if (n.cTn.endSync?.rtn === 'first') return Math.min(...started.map(ch => ch.end));
  // a seq whose later children wait for clicks lasts until they come
  if (n.kind === 'seq' && started.length < children.length) return Infinity;
  return Math.max(...started.map(ch => ch.end));
}

/** A node's own timeline, whatever starts it: its children, simple and active duration. Begin is resolved later. */
function scheduleSubtree(n: TimeNode, ancestors: ReadonlySet<number>, ctx: Ctx): ScheduledNode {
  const c = n.cTn;
  const inner = new Set(ancestors);
  if (c.id !== undefined) inner.add(c.id);
  ctx.depth++;
  const children = ctx.depth > MAX_DEPTH ? [] : scheduleChildren(n, inner, ctx);
  ctx.depth--;
  let simple = timeOf(c.dur) ?? (isContainer(n) ? implicitDuration(n, children) : 0);
  let stretch = 1;
  if (c.iterate && finite(simple) && simple > 0) {
    const interval = c.iterate.interval;
    const step = interval.kind === 'abs' ? timeOf(interval.ms) ?? 0 : interval.fraction * simple;
    const total = simple + Math.max(0, Math.max(1, ctx.iterations(n)) - 1) * (finite(step) ? step : 0);
    stretch = simple / total;
    simple = total;
  }
  const active = activeDuration(n, simple);
  return { node: n, begin: Infinity, simple, active, end: Infinity, stretch, children, order: ctx.orders.get(n) ?? 0 };
}

/** Place a scheduled node in its parent's timeline: begin from its start conditions, end cut by its end conditions. */
function place(s: ScheduledNode, base: number, siblings: Siblings, ancestors: ReadonlySet<number>, full: number): void {
  const c = s.node.cTn;
  s.begin = c.start.length ? earliest(c.start, base, siblings, ancestors) ?? Infinity : base;
  const cut = earliest(c.end, 0, siblings, ancestors);
  s.active = cut !== undefined && finite(s.begin) ? Math.max(0, Math.min(full, cut - s.begin)) : full;
  s.end = s.begin + s.active;
}

function scheduleChildren(n: TimeNode, ancestors: ReadonlySet<number>, ctx: Ctx): ScheduledNode[] {
  const list = n.cTn.children;
  const out = list.map(child => scheduleSubtree(child, ancestors, ctx));
  const full = out.map(s => s.active);
  const ids = new Set(list.map(c => c.cTn.id).filter((id): id is number => id !== undefined));
  // conditions can name later siblings: only then is another round needed
  const linked = list.some(c => [...c.cTn.start, ...c.cTn.end].some(cond => cond.tn !== undefined && ids.has(cond.tn)));
  const byId = new Map<number, ScheduledNode>();
  for (let round = 0; round < (linked ? MAX_ROUNDS : 1); round++) {
    const before = out.map(s => [s.begin, s.end]);
    let prevEnd = 0;
    out.forEach((s, i) => {
      place(s, n.kind === 'seq' ? prevEnd : 0, byId, ancestors, full[i]!);
      if (s.node.cTn.id !== undefined) byId.set(s.node.cTn.id, s);
      prevEnd = s.end;
    });
    if (round > 0 && out.every((s, i) => s.begin === before[i]![0] && s.end === before[i]![1])) break;
  }
  return out;
}

const activation = (roots: ScheduledNode[]): Activation =>
  ({ roots, duration: roots.reduce((m, r) => Math.max(m, finite(r.begin) ? r.end : 0), 0) });

/**
 * Split a sequence's children into click steps. Step 0 holds what starts by itself; a child that waits for a click
 * opens the next step and begins at that click; the children after it that do not wait follow it, each when the one
 * before ends (plus its delay), as in any seq.
 */
function steps(seq: SeqNode, ctx: Ctx, extra: TimeNode[] = []): Activation[] {
  const ancestors = new Set<number>(seq.cTn.id !== undefined ? [seq.cTn.id] : []);
  const out: ScheduledNode[][] = [[]];
  const byId = new Map<number, ScheduledNode>();
  const add = (n: TimeNode, base: number, click: boolean) => {
    const s = scheduleSubtree(n, ancestors, ctx);
    place(s, base, byId, ancestors, s.active);
    if (click) {
      s.begin = 0;
      s.end = s.active;
    }
    if (n.cTn.id !== undefined) byId.set(n.cTn.id, s);
    out.at(-1)!.push(s);
  };
  for (const n of extra) add(n, 0, false);
  for (const child of seq.cTn.children) {
    const click = waitsForClick(child);
    if (click) out.push([]);
    const prev = out.at(-1)!.at(-1);
    add(child, click || !prev ? 0 : prev.end, click);
  }
  return out.map(activation);
}

function documentOrder(timing: Timing): Map<TimeNode, number> {
  const orders = new Map<TimeNode, number>();
  const visit = (list: readonly TimeNode[]) => {
    for (const n of list) {
      orders.set(n, orders.size);
      visit(n.cTn.children);
    }
  };
  visit(timing.nodes);
  return orders;
}

/** Schedule a slide's timing: its main sequence step by step, and each interactive sequence. */
export function buildSchedule(timing: Timing, iterations: IterationCount = () => 1): Schedule {
  const ctx: Ctx = { orders: documentOrder(timing), iterations, depth: 0 };
  const main = mainSequence(timing);
  const triggered = interactiveSequences(timing);
  const special = new Set<TimeNode>([...(main ? [main] : []), ...triggered.map(t => t.seq)]);
  // containers beside the sequences that start by themselves (rare) play with the slide
  const loose = (rootOf(timing)?.cTn.children ?? []).filter(n => !special.has(n) && !waitsForClick(n) && isContainer(n));
  const mainSteps = main ? steps(main, ctx, loose) : [activation(loose.map(n => {
    const s = scheduleSubtree(n, new Set(), ctx);
    place(s, 0, new Map(), new Set(), s.active);
    return s;
  }))];
  return {
    main: mainSteps,
    // the click on the trigger starts the sequence and plays its first group
    interactive: triggered.map(({ seq, trigger }) => ({ trigger, seq, steps: steps(seq, ctx).filter((s, i) => i > 0 || s.roots.length) })),
  };
}

/** The number of clicks the main sequence needs. */
export const clickCount = (s: Schedule) => Math.max(0, s.main.length - 1);

/**
 * How long an effect plays once it has started, its own start delay left out: `pass` is one pass of its behaviours
 * (a one-millisecond set marks a moment, not a length, as PowerPoint counts it: Appear lasts 0), `total` that
 * repeated, reversed, sped up and iterated as the effect's own timing says.
 */
export function effectLength(node: TimeNode, iterations: IterationCount = () => 1): { pass: number; total: number } {
  const ctx: Ctx = { orders: new Map(), iterations, depth: 0 };
  const children = scheduleChildren(node, new Set<number>(node.cTn.id !== undefined ? [node.cTn.id] : []), ctx);
  let pass = 0;
  const visit = (s: ScheduledNode, at: number) => {
    if (!finite(s.begin)) return;
    const begin = at + s.begin;
    if (isBehavior(s.node) || s.node.kind === 'audio' || s.node.kind === 'video') {
      pass = Math.max(pass, begin + (s.simple <= 1 && finite(s.active) ? 0 : s.active));
      return;
    }
    if (!s.children.length) pass = Math.max(pass, begin + (finite(s.active) ? s.active : 0));
    for (const c of s.children) visit(c, begin);
  };
  for (const c of children) visit(c, 0);
  if (!finite(pass)) return { pass: Infinity, total: Infinity };
  let simple = pass;
  const it = node.cTn.iterate;
  if (it) {
    const step = it.interval.kind === 'abs' ? timeOf(it.interval.ms) ?? 0 : it.interval.fraction * pass;
    simple += Math.max(0, iterations(node) - 1) * (finite(step) ? step : 0);
  }
  return { pass, total: activeDuration(node, simple) };
}
