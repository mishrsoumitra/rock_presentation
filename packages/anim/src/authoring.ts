// packages/anim/src/authoring.ts  -- editing a slide's animations: pure XML edits that the editor wraps in commands
//
// Each operation changes the slide's p:timing in place and touches the slide's owner. Effects are handled as the
// animation pane shows them, a flat list per sequence where each effect starts on click, with the previous one, or
// after it; every edit then rewrites the sequence's groups the way PowerPoint writes them: one par per click, inside
// it one par per "after previous" step whose delay is when the step before ends, and the effects in those. Time
// node ids are renumbered in document order afterwards (references in p:tn follow), the build list keeps an entry
// for each text or graphic build and no other, and p:timing goes away with its last effect.
import type { XmlElement } from '@pptx/xml';
import {
  append, attr, child, children, createElement, ensureChild, insertInOrder, int, is, paragraphsOf, remove, runsOf, setAttr,
  type ShapeView, type SlideView,
} from '@pptx/model';
import { findPreset, presetName, type Preset } from './presets.ts';
import { effectLength, waitsForClick } from './schedule.ts';
import { instantiate, rescale, type EffectTarget } from './template.ts';
import {
  interactiveSequences, isContainer, mainSequence, parseTiming, parseTimeNode, type PresetClass, type SeqNode, type TimeNode, type Timing,
} from './timing.ts';

export type Trigger = 'click' | 'with' | 'after';

export interface EffectInfo {
  /** Position in the animation pane: the main sequence first, then each trigger's sequence. */
  index: number;
  trigger: Trigger;
  /** For effects that play when a shape is clicked: that shape. */
  triggerShape?: number;
  presetClass?: PresetClass;
  presetID?: number;
  presetSubtype?: number;
  name: string;
  spid?: number;
  paragraphs?: { st: number; end: number };
  background?: boolean;
  /** One pass, ms; Infinity for effects that last until the end of the slide (Transparency). */
  duration: number;
  /** After the trigger, ms. */
  delay: number;
  repeat?: number | 'indefinite';
  /** The effect's p:par. */
  el: XmlElement;
}

export interface AddEffectOptions {
  trigger?: Trigger;
  delay?: number;
  /** Length of one pass, ms; the preset's own when absent. */
  duration?: number;
  /** A paragraph range of the shape's text, or 'each' for one effect per paragraph (By Paragraph). */
  paragraphs?: { st: number; end: number } | 'each';
  /** Animate the shape's background (bldP animBg) rather than the shape. */
  background?: boolean;
  /** Position within its sequence; the end when absent. */
  index?: number;
  /** Put the effect in the sequence that clicking this shape starts. */
  triggerShape?: number;
  /** A motion path to follow instead of the preset's. */
  path?: string;
}

export interface EffectTiming {
  trigger?: Trigger;
  delay?: number;
  duration?: number;
  /** Times to play (2.5 is two and a half), 'indefinite', or undefined for once. */
  repeat?: number | 'indefinite' | undefined;
}

type PresetRef = Preset | { presetClass: PresetClass; presetID: number; presetSubtype?: number };

// ------------------------------------------------------------------------------------------------- building

type Spec = [name: string, attrs?: Record<string, string | number>, ...children: Spec[]];

function build(parent: XmlElement, [name, attrs = {}, ...kids]: Spec): XmlElement {
  const el = createElement(parent, name, attrs);
  for (const k of kids) append(el, build(el, k));
  return el;
}

const slideCond = (evt: string): Spec => ['p:cond', { evt, delay: 0 }, ['p:tgtEl', {}, ['p:sldTgt']]];
const shapeCond = (spid: number): Spec => ['p:cond', { evt: 'onClick', delay: 0 }, ['p:tgtEl', {}, ['p:spTgt', { spid }]]];
/** The time root's p:cTn. */
const timeRoot = (timing: XmlElement) => child(child(child(timing, 'p:tnLst'), 'p:par'), 'p:cTn');

/** The time root's child list, read only: undefined when the file self-closes the root (no effects at all). */
const rootListIfAny = (timing: XmlElement) => child(timeRoot(timing), 'p:childTnLst');

/** The time root's child list to add to; created when the file left it out (a self-closed tmRoot). */
function rootList(timing: XmlElement): XmlElement {
  const cTn = timeRoot(timing);
  if (!cTn) throw new Error('p:timing has no time root');
  const list = ensureChild(cTn, 'p:childTnLst');
  cTn.selfClosing = false;
  return list;
}

/** p:timing with its time root, created the way PowerPoint writes it when the slide has none. */
function ensureTiming(slide: SlideView): XmlElement {
  let timing = slide.timing;
  if (!timing) {
    timing = createElement(slide.root, 'p:timing');
    insertInOrder(slide.root, timing);
  }
  let tnLst = child(timing, 'p:tnLst');
  if (!tnLst) {
    tnLst = createElement(timing, 'p:tnLst');
    insertInOrder(timing, tnLst);
  }
  if (!child(tnLst, 'p:par')) {
    append(tnLst, build(tnLst, ['p:par', {}, ['p:cTn', { id: 1, dur: 'indefinite', restart: 'never', nodeType: 'tmRoot' }, ['p:childTnLst']]]));
  }
  return timing;
}

function ensureMainSeq(timing: XmlElement): XmlElement {
  const existing = mainSequence(parseTiming(timing));
  if (existing) return existing.el;
  const list = rootList(timing);
  const seq = build(list, ['p:seq', { concurrent: 1, nextAc: 'seek' },
    ['p:cTn', { id: 2, dur: 'indefinite', nodeType: 'mainSeq' }, ['p:childTnLst']],
    ['p:prevCondLst', {}, slideCond('onPrev')], ['p:nextCondLst', {}, slideCond('onNext')]]);
  // the main sequence comes first
  seq.parent = list;
  list.children.unshift(seq);
  list.selfClosing = false;
  return seq;
}

function ensureInteractiveSeq(timing: XmlElement, spid: number): XmlElement {
  const existing = interactiveSequences(parseTiming(timing)).find(s => s.trigger === spid);
  if (existing) return existing.seq.el;
  const list = rootList(timing);
  const seq = build(list, ['p:seq', { concurrent: 1, nextAc: 'seek' },
    ['p:cTn', { id: 0, restart: 'whenNotActive', fill: 'hold', evtFilter: 'cancelBubble', nodeType: 'interactiveSeq' },
      ['p:stCondLst', {}, shapeCond(spid)], ['p:endSync', { evt: 'end', delay: 0 }, ['p:rtn', { val: 'all' }]], ['p:childTnLst']],
    ['p:nextCondLst', {}, shapeCond(spid)]]);
  // after the other sequences, before media nodes
  const at = list.children.findIndex(c => c.kind === 'element' && !is(c, 'p:seq'));
  seq.parent = list;
  list.children.splice(at < 0 ? list.children.length : at, 0, seq);
  list.selfClosing = false;
  return seq;
}

// --------------------------------------------------------------------------------------------- sequences

interface Entry { el: XmlElement; trigger: Trigger; }
interface Sequence { el: XmlElement; node: SeqNode; trigger?: number; entries: Entry[]; }

const EFFECT_TYPES = new Map<string, Trigger>([['clickEffect', 'click'], ['withEffect', 'with'], ['afterEffect', 'after']]);
const isEffectNode = (n: TimeNode) => isContainer(n) && (n.cTn.presetClass !== undefined || EFFECT_TYPES.has(n.cTn.nodeType ?? ''));

/** A sequence's effects in pane order: click groups, their with/after steps, the effects in each. */
function entriesOf(seq: SeqNode): Entry[] {
  const out: Entry[] = [];
  const push = (n: TimeNode, fallback: Trigger) => out.push({ el: n.el, trigger: EFFECT_TYPES.get(n.cTn.nodeType ?? '') ?? fallback });
  for (const group of seq.cTn.children) {
    if (isEffectNode(group)) {
      push(group, 'click');
      continue;
    }
    group.cTn.children.forEach((step, si) => {
      if (isEffectNode(step)) push(step, si === 0 ? 'click' : 'with');
      else step.cTn.children.filter(isContainer).forEach((e, ei) => push(e, ei > 0 ? 'with' : si > 0 ? 'after' : 'click'));
    });
  }
  return out;
}

function sequencesOf(timing: Timing): Sequence[] {
  const out: Sequence[] = [];
  const main = mainSequence(timing);
  if (main) out.push({ el: main.el, node: main, entries: entriesOf(main) });
  for (const { seq, trigger } of interactiveSequences(timing)) out.push({ el: seq.el, node: seq, trigger, entries: entriesOf(seq) });
  return out;
}

const effectCTn = (par: XmlElement) => child(par, 'p:cTn')!;
const plainCond = (c: XmlElement) => attr(c, 'evt') === undefined && !children(c).length;

function ownDelay(par: XmlElement): number {
  const cond = children(child(effectCTn(par), 'p:stCondLst'), 'p:cond').find(plainCond);
  const d = Number(attr(cond, 'delay') ?? 0);
  return Number.isFinite(d) ? d : 0;
}

function setOwnDelay(par: XmlElement, ms: number): void {
  const cTn = effectCTn(par);
  let list = child(cTn, 'p:stCondLst');
  if (!list) {
    list = createElement(cTn, 'p:stCondLst');
    insertInOrder(cTn, list);
  }
  const delay = Math.max(0, Math.round(ms));
  const cond = children(list, 'p:cond').find(plainCond);
  if (cond) setAttr(cond, 'delay', delay);
  else append(list, createElement(list, 'p:cond', { delay }));
}

const lengthOf = (par: XmlElement) => {
  const n = parseTimeNode(par);
  return n ? effectLength(n) : { pass: 0, total: 0 };
};

const groupSpec = (conds: Spec[]): Spec => ['p:par', {}, ['p:cTn', { id: 0, fill: 'hold' }, ['p:stCondLst', {}, ...conds], ['p:childTnLst']]];

/** Rewrite a sequence's groups from its flat list of effects. */
function rebuild(seq: Sequence, timing: XmlElement): void {
  const cTn = child(seq.el, 'p:cTn')!;
  let list = child(cTn, 'p:childTnLst');
  if (!list) {
    list = createElement(cTn, 'p:childTnLst');
    insertInOrder(cTn, list);
  }
  for (const e of seq.entries) remove(e.el);
  list.children = [];
  if (attr(cTn, 'id') === undefined) setAttr(cTn, 'id', maxId(timing) + 1);
  const seqId = attr(cTn, 'id')!;
  let group: XmlElement | undefined;
  let step: XmlElement | undefined;
  let stepBegin = 0;
  let stepEnd = 0;
  seq.entries.forEach((e, i) => {
    if (!group || e.trigger === 'click') {
      // a group starts on a click; the first one may start with its sequence (the slide, or the trigger's click)
      const onClick: Spec = ['p:cond', { delay: 'indefinite' }];
      const withSequence: Spec = ['p:cond', { evt: 'onBegin', delay: 0 }, ['p:tn', { val: seqId }]];
      const conds: Spec[] = seq.trigger !== undefined ? [['p:cond', { delay: i === 0 ? 0 : 'indefinite' }]]
        : i === 0 && e.trigger !== 'click' ? [onClick, withSequence] : [onClick];
      const g = build(list!, groupSpec(conds));
      append(list!, g);
      group = child(effectCTn(g), 'p:childTnLst')!;
      step = undefined;
      stepBegin = 0;
      stepEnd = 0;
    }
    if (!step || e.trigger === 'after') {
      if (step) stepBegin = stepEnd;
      const s = build(group, groupSpec([['p:cond', { delay: Math.round(stepBegin) }]]));
      append(group, s);
      step = child(effectCTn(s), 'p:childTnLst')!;
    }
    setAttr(effectCTn(e.el), 'nodeType', `${e.trigger}Effect`);
    append(step, e.el);
    const total = lengthOf(e.el).total;
    stepEnd = Math.max(stepEnd, stepBegin + ownDelay(e.el) + (Number.isFinite(total) ? total : 0));
  });
  list.selfClosing = list.children.length === 0;
}

function maxId(timing: XmlElement): number {
  let max = 0;
  const visit = (e: XmlElement) => {
    if (is(e, 'p:cTn')) max = Math.max(max, int(attr(e, 'id'), 0));
    for (const c of children(e)) visit(c);
  };
  visit(timing);
  return max;
}

/** Number every time node 1, 2, 3, ... in document order, as PowerPoint does, keeping p:tn references right. */
function renumber(timing: XmlElement): void {
  const tnLst = child(timing, 'p:tnLst');
  if (!tnLst) return;
  const map = new Map<string, number>();
  let next = 1;
  // PowerPoint leaves the nodes of sub-lists (sounds, dim after) without ids
  const visit = (e: XmlElement, sub: boolean) => {
    const old = is(e, 'p:cTn') ? attr(e, 'id') : undefined;
    if (is(e, 'p:cTn') && (old !== undefined || !sub)) {
      const id = next++;
      if (old !== undefined && !map.has(old)) map.set(old, id);
      if (old === undefined) e.attrs.unshift({ prefix: null, local: 'id', value: String(id), ns: '' });
      else setAttr(e, 'id', id);
    }
    for (const c of children(e)) visit(c, sub || is(c, 'p:subTnLst'));
  };
  visit(tnLst, false);
  const fix = (e: XmlElement) => {
    const v = is(e, 'p:tn') ? attr(e, 'val') : undefined;
    if (v !== undefined && map.has(v)) setAttr(e, 'val', map.get(v)!);
    for (const c of children(e)) fix(c);
  };
  fix(tnLst);
}

// ------------------------------------------------------------------------------------------------- builds

function findShape(slide: SlideView, spid: number): ShapeView | undefined {
  const visit = (list: ShapeView[]): ShapeView | undefined => {
    for (const s of list) {
      if (s.id === spid) return s;
      const inner = visit(s.children);
      if (inner) return inner;
    }
    return undefined;
  };
  return visit(slide.shapes);
}

const hasText = (s: ShapeView | undefined) => !!s?.txBody && paragraphsOf(s.txBody).some(p => runsOf(p).some(r => r.text.trim() !== ''));
const graphicKind = (s: ShapeView | undefined) => /\/(chart|diagram)$/.exec(s?.graphicData?.uri ?? '')?.[1];

/** Whether effects on the shape are builds (bldP for text, bldGraphic for charts and diagrams) with a group id. */
const isBuild = (s: ShapeView | undefined) => (s?.kind === 'sp' ? hasText(s) : s?.kind === 'graphicFrame' && graphicKind(s) !== undefined);

/** Whether a text shape has a fill or line of its own, which PowerPoint animates with the text (animBg). */
function hasBackground(s: ShapeView): boolean {
  const spPr = s.spPr;
  const any = (el: XmlElement | undefined, names: string[]) => children(el).some(c => names.some(n => is(c, n)));
  const fill = any(spPr, ['a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill']);
  const ln = child(spPr, 'a:ln');
  const line = !!ln && !child(ln, 'a:noFill') && any(ln, ['a:solidFill', 'a:gradFill', 'a:pattFill']);
  const ref = (n: string) => int(attr(child(s.style, n), 'idx'), 0) > 0;
  return fill || line || (!child(spPr, 'a:noFill') && ref('a:fillRef')) || (!ln && ref('a:lnRef'));
}

function firstTarget(n: TimeNode): EffectTarget | undefined {
  const find = (el: XmlElement): XmlElement | undefined => {
    for (const c of children(el)) {
      const hit = is(c, 'p:spTgt') ? c : find(c);
      if (hit) return hit;
    }
    return undefined;
  };
  const sp = find(child(n.cTn.el, 'p:childTnLst') ?? n.cTn.el);
  if (!sp) return undefined;
  const pRg = child(child(sp, 'p:txEl'), 'p:pRg');
  return {
    spid: int(attr(sp, 'spid'), 0),
    ...(pRg ? { paragraphs: { st: int(attr(pRg, 'st'), 0), end: int(attr(pRg, 'end'), 0) } } : {}),
    ...(child(sp, 'p:bg') ? { background: true } : {}),
  };
}

const targetsShape = (el: XmlElement, spid: number): boolean =>
  (is(el, 'p:spTgt') && int(attr(el, 'spid'), -1) === spid) || children(el).some(c => targetsShape(c, spid));

interface Use { spid: number; grpId: number; paragraphs: boolean; whole: boolean; background: boolean; }

/** Keep p:bldLst in step with the effects: an entry per (shape, group id) in use, none for anything else. */
function syncBuilds(slide: SlideView, timing: XmlElement): void {
  const uses = new Map<string, Use>();
  for (const seq of sequencesOf(parseTiming(timing))) {
    for (const e of seq.entries) {
      const n = parseTimeNode(e.el);
      const grpId = n?.cTn.grpId;
      const t = n ? firstTarget(n) : undefined;
      if (grpId === undefined || !t) continue;
      const k = `${t.spid}|${grpId}`;
      const u = uses.get(k) ?? { spid: t.spid, grpId, paragraphs: false, whole: false, background: false };
      if (t.paragraphs) u.paragraphs = true;
      else if (t.background) u.background = true;
      else u.whole = true;
      uses.set(k, u);
    }
  }
  let bldLst = child(timing, 'p:bldLst');
  for (const b of children(bldLst)) {
    const k = `${attr(b, 'spid')}|${attr(b, 'grpId')}`;
    if (!uses.has(k)) remove(b);
    else uses.delete(k);
  }
  for (const u of uses.values()) {
    const shape = findShape(slide, u.spid);
    if (!isBuild(shape)) continue;
    if (!bldLst) {
      bldLst = createElement(timing, 'p:bldLst');
      insertInOrder(timing, bldLst);
    }
    if (graphicKind(shape)) {
      append(bldLst, build(bldLst, ['p:bldGraphic', { spid: u.spid, grpId: u.grpId }, ['p:bldAsOne']]));
      continue;
    }
    const attrs: Record<string, string | number> = { spid: u.spid, grpId: u.grpId };
    if (u.paragraphs) attrs.build = 'p';
    if (u.background || (u.whole && shape && hasBackground(shape))) attrs.animBg = 1;
    append(bldLst, build(bldLst, ['p:bldP', attrs]));
  }
  if (bldLst && !children(bldLst).length) remove(bldLst);
}

function nextGroupId(timing: XmlElement, spid: number): number {
  let max = -1;
  for (const b of children(child(timing, 'p:bldLst'))) if (int(attr(b, 'spid'), -1) === spid) max = Math.max(max, int(attr(b, 'grpId'), 0));
  for (const seq of sequencesOf(parseTiming(timing))) {
    for (const e of seq.entries) {
      const n = parseTimeNode(e.el);
      if (n?.cTn.grpId !== undefined && firstTarget(n)?.spid === spid) max = Math.max(max, n.cTn.grpId);
    }
  }
  return max + 1;
}

// ------------------------------------------------------------------------------------------------- the API

function info(seq: Sequence, e: Entry, index: number): EffectInfo {
  const n = parseTimeNode(e.el);
  const c = n?.cTn;
  const t = n ? firstTarget(n) : undefined;
  const out: EffectInfo = {
    index, trigger: e.trigger, name: presetName(c?.presetClass, c?.presetID),
    duration: n ? effectLength(n).pass : 0, delay: ownDelay(e.el), el: e.el,
  };
  if (seq.trigger !== undefined) out.triggerShape = seq.trigger;
  if (c?.presetClass) out.presetClass = c.presetClass;
  if (c?.presetID !== undefined) out.presetID = c.presetID;
  if (c?.presetSubtype !== undefined) out.presetSubtype = c.presetSubtype;
  if (t) out.spid = t.spid;
  if (t?.paragraphs) out.paragraphs = t.paragraphs;
  if (t?.background) out.background = true;
  if (c?.repeatCount !== undefined) out.repeat = c.repeatCount;
  return out;
}

/** The slide's effects in animation-pane order. */
export function listEffects(slide: SlideView): EffectInfo[] {
  const timing = slide.timing;
  if (!timing) return [];
  const out: EffectInfo[] = [];
  for (const seq of sequencesOf(parseTiming(timing))) for (const e of seq.entries) out.push(info(seq, e, out.length));
  return out;
}

/** Rebuild the given sequences, drop empty ones (and p:timing when nothing is left), sync builds, renumber, touch. */
function commit(slide: SlideView, timing: XmlElement, seqs: Sequence[]): void {
  for (const seq of seqs) {
    if (seq.entries.length) rebuild(seq, timing);
    else remove(seq.el);
  }
  if (!children(rootListIfAny(timing)).length) remove(timing);
  else {
    syncBuilds(slide, timing);
    renumber(timing);
  }
  slide.owner.touch();
}

/** The sequence holding the effect at a pane index, and its position in that sequence. */
function locate(slide: SlideView, index: number): { timing: XmlElement; seqs: Sequence[]; seq: Sequence; at: number } {
  const timing = slide.timing;
  if (!timing) throw new RangeError(`no effect ${index}: the slide has no animations`);
  const seqs = sequencesOf(parseTiming(timing));
  let i = index;
  for (const seq of seqs) {
    if (i >= 0 && i < seq.entries.length) return { timing, seqs, seq, at: i };
    i -= seq.entries.length;
  }
  throw new RangeError(`no effect ${index}: the slide has ${seqs.reduce((n, s) => n + s.entries.length, 0)}`);
}

/**
 * Add an effect for shape `spid` from a preset, and return what was added (several with paragraphs: 'each'). The
 * effect goes at `index` in its sequence, the end by default. With paragraphs 'each' the first paragraph starts on
 * the trigger and each next one on a click (after the previous one when the trigger is not a click).
 */
export function addEffect(slide: SlideView, spid: number, preset: PresetRef, options: AddEffectOptions = {}): EffectInfo[] {
  const p = 'behaviors' in preset ? preset : findPreset(preset.presetClass, preset.presetID, preset.presetSubtype);
  if (!p) {
    const subtype = preset.presetSubtype !== undefined ? `/${preset.presetSubtype}` : '';
    throw new Error(`no preset ${preset.presetClass} ${preset.presetID}${subtype}`);
  }
  const timing = ensureTiming(slide);
  const seqEl = options.triggerShape !== undefined ? ensureInteractiveSeq(timing, options.triggerShape) : ensureMainSeq(timing);
  const seq = sequencesOf(parseTiming(timing)).find(s => s.el === seqEl)!;
  const shape = findShape(slide, spid);
  const grpId = isBuild(shape) ? nextGroupId(timing, spid) : undefined;
  let targets: EffectTarget[];
  if (options.paragraphs === 'each') {
    const paras = shape?.txBody ? paragraphsOf(shape.txBody) : [];
    targets = paras.flatMap((para, i) => (runsOf(para).some(r => r.text.trim() !== '') ? [{ spid, paragraphs: { st: i, end: i } }] : []));
    if (!targets.length) targets = [{ spid }];
  } else {
    targets = [{ spid, ...(options.paragraphs ? { paragraphs: options.paragraphs } : {}), ...(options.background ? { background: true } : {}) }];
  }
  const trigger = options.trigger ?? 'click';
  const list = child(child(seqEl, 'p:cTn'), 'p:childTnLst') ?? seqEl;
  const settings = {
    ...(options.duration !== undefined ? { duration: options.duration } : {}),
    ...(options.path !== undefined ? { path: options.path } : {}),
  };
  const added: Entry[] = targets.map((t, i) => {
    const el = instantiate(list, p, p, t, settings);
    if (grpId !== undefined) setAttr(effectCTn(el), 'grpId', grpId);
    if (options.delay) setOwnDelay(el, options.delay);
    return { el, trigger: i === 0 ? trigger : trigger === 'click' ? 'click' : 'after' };
  });
  seq.entries.splice(Math.max(0, Math.min(options.index ?? seq.entries.length, seq.entries.length)), 0, ...added);
  commit(slide, timing, [seq]);
  return listEffects(slide).filter(e => added.some(a => a.el === e.el));
}

/** Remove the effect at `index` (animation-pane order). */
export function removeEffect(slide: SlideView, index: number): void {
  const { timing, seq, at } = locate(slide, index);
  const [gone] = seq.entries.splice(at, 1);
  if (gone) remove(gone.el);
  commit(slide, timing, [seq]);
}

/** Move the effect at `from` to position `to` (both in animation-pane order); it joins the sequence it lands in. */
export function moveEffect(slide: SlideView, from: number, to: number): void {
  const src = locate(slide, from);
  const total = src.seqs.reduce((n, s) => n + s.entries.length, 0);
  const [moved] = src.seq.entries.splice(src.at, 1);
  if (!moved) return;
  let i = Math.max(0, Math.min(to, total - 1));
  let dest = src.seqs.at(-1)!;
  let pos = dest.entries.length;
  for (const seq of src.seqs) {
    // an index just past a sequence's end belongs to it only when the effect came from there
    if (i < seq.entries.length || (i === seq.entries.length && seq === src.seq)) {
      dest = seq;
      pos = i;
      break;
    }
    i -= seq.entries.length;
  }
  dest.entries.splice(pos, 0, moved);
  commit(slide, src.timing, dest === src.seq ? [src.seq] : [src.seq, dest]);
}

/** Change an effect's trigger, delay, duration (the whole effect's timeline is scaled) or repeat count. */
export function setEffectTiming(slide: SlideView, index: number, timing: EffectTiming): void {
  const { timing: el, seq, at } = locate(slide, index);
  const e = seq.entries[at]!;
  if (timing.trigger) e.trigger = timing.trigger;
  if (timing.delay !== undefined) setOwnDelay(e.el, timing.delay);
  if (timing.duration !== undefined) rescale(e.el, lengthOf(e.el).pass, Math.max(0, timing.duration));
  if ('repeat' in timing) {
    const r = timing.repeat;
    setAttr(effectCTn(e.el), 'repeatCount', r === undefined ? undefined : r === 'indefinite' ? r : Math.round(r * 1000));
  }
  commit(slide, el, [seq]);
}

/** Remove every effect on a shape and the sequences it triggers, as when the shape is deleted. */
export function removeEffectsOf(slide: SlideView, spid: number): void {
  const timing = slide.timing;
  if (!timing) return;
  const seqs = sequencesOf(parseTiming(timing));
  for (const seq of seqs) {
    if (seq.trigger === spid) {
      seq.entries = [];
      continue;
    }
    seq.entries = seq.entries.filter(e => {
      if (!targetsShape(e.el, spid)) return true;
      remove(e.el);
      return false;
    });
  }
  // media nodes beside the sequences that play the shape
  for (const c of children(rootListIfAny(timing))) if ((is(c, 'p:audio') || is(c, 'p:video')) && targetsShape(c, spid)) remove(c);
  commit(slide, timing, seqs);
}

/** How many main-sequence clicks a slide's effects need. */
export function clicksOf(slide: SlideView): number {
  const timing = slide.timing;
  const main = timing ? mainSequence(parseTiming(timing)) : undefined;
  return main ? main.cTn.children.filter(waitsForClick).length : 0;
}
