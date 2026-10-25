// packages/anim/src/timing.ts  -- p:timing as plain data: time nodes, conditions, behaviours, targets and builds
//
// ECMA-376 Part 1, section 19.5. Reading never changes the tree: every item keeps the element it came from, so the
// XML stays the source of truth and the authoring functions edit it directly. Values outside the schema's
// enumerations read as undefined (the element is still there to look at).
import type { XmlElement } from '@pptx/xml';
import {
  attr, attrNs, bool, child, children, colorIn, effectiveChildren, int, is, num, percent, textOf, type ColorDef, type NsKey,
} from '@pptx/model';

/** Milliseconds, or 'indefinite' (ST_TLTime). */
export type Time = number | 'indefinite';

const PRESET_CLASSES = ['entr', 'exit', 'emph', 'path', 'verb', 'mediacall'] as const;
const NODE_TYPES = [
  'clickEffect', 'withEffect', 'afterEffect', 'mainSeq', 'interactiveSeq', 'clickPar', 'withGroup', 'afterGroup', 'tmRoot',
] as const;
const EVENTS = [
  'onBegin', 'onEnd', 'begin', 'end', 'onClick', 'onDblClick', 'onMouseOver', 'onMouseOut', 'onNext', 'onPrev', 'onStopAudio',
] as const;

export type PresetClass = (typeof PRESET_CLASSES)[number];
export type NodeType = (typeof NODE_TYPES)[number];
export type TriggerEvent = (typeof EVENTS)[number];
export type FillMode = 'remove' | 'freeze' | 'hold' | 'transition';
export type Additive = 'base' | 'sum' | 'repl' | 'mult' | 'none';
export type CalcMode = 'discrete' | 'lin' | 'fmla';
export type ValueType = 'str' | 'num' | 'clr';

/** Branches of mc:AlternateContent read inside p:timing and p:transition. */
export const ANIM_UNDERSTOOD: ReadonlySet<NsKey> = new Set<NsKey>(['p14', 'p15', 'p159']);

export interface Condition {
  el: XmlElement;
  evt?: TriggerEvent;
  /** 0 when absent. */
  delay: Time;
  target?: Target;
  /** p:tn: the id of the time node whose begin or end triggers this one. */
  tn?: number;
  /** p:rtn: a runtime node (endSync uses it: all, first or last child). */
  rtn?: 'first' | 'last' | 'all';
}

export interface Iterate {
  el: XmlElement;
  /** Element, word or letter. */
  type: 'el' | 'wd' | 'lt';
  backwards: boolean;
  /** Time between the starts of successive items: absolute, or a fraction of the effect's duration. */
  interval: { kind: 'abs'; ms: Time } | { kind: 'pct'; fraction: number };
}

export type ShapeSub =
  | { kind: 'bg' }
  | { kind: 'subShape'; spid: number }
  | { kind: 'oleChart'; type: string; lvl: number }
  /** Text: a paragraph or character range (inclusive), or the whole text when range is absent. */
  | { kind: 'text'; range?: { unit: 'paragraph' | 'char'; st: number; end: number } }
  | {
      kind: 'graphic'; el: XmlElement;
      diagram?: { id: string; bldStep?: string };
      chart?: { seriesIdx: number; categoryIdx: number; bldStep?: string };
    };

export type Target =
  | { kind: 'slide' }
  | { kind: 'sound'; el: XmlElement; embed?: string; name?: string }
  | { kind: 'shape'; spid: number; sub?: ShapeSub }
  | { kind: 'ink'; spid: number };

export type Variant =
  | { kind: 'bool'; value: boolean }
  | { kind: 'int'; value: number }
  | { kind: 'float'; value: number }
  | { kind: 'string'; value: string }
  | { kind: 'color'; value: ColorDef };

/** p:cTn: what every time node has. Percentages are fractions (1 = 100%), times milliseconds. */
export interface CommonTimeNode {
  el: XmlElement;
  id?: number;
  presetClass?: PresetClass;
  presetID?: number;
  presetSubtype?: number;
  nodeType?: NodeType;
  dur?: Time;
  fill?: FillMode;
  restart?: 'always' | 'whenNotActive' | 'never';
  /** Iterations (the file stores thousandths: 3000 is three times), or 'indefinite'. */
  repeatCount?: Time;
  repeatDur?: Time;
  autoRev: boolean;
  accel: number;
  decel: number;
  /** Playback speed: 1 is normal, 2 twice as fast, negative plays backwards. */
  spd: number;
  grpId?: number;
  afterEffect?: boolean;
  masterRel?: 'sameClick' | 'lastClick' | 'nextClick';
  bldLvl?: number;
  display?: boolean;
  syncBehavior?: 'canSlip' | 'locked';
  /** Time filter: (time, progress) pairs that remap the simple time, e.g. "0,0; 0.5,0.2; 1,1". */
  tmFilter?: [number, number][];
  evtFilter?: string;
  nodePh?: boolean;
  start: Condition[];
  end: Condition[];
  endSync?: Condition;
  iterate?: Iterate;
  children: TimeNode[];
  /** p:subTnLst: nodes started by their master relation (dim after, sounds), not by their parent. */
  sub: TimeNode[];
}

/** p:cBhvr: what every behaviour has. */
export interface Behavior {
  el: XmlElement;
  target?: Target;
  attrNames: string[];
  additive?: Additive;
  accumulate?: 'none' | 'always';
  xfrmType?: 'pt' | 'img';
  from?: string;
  to?: string;
  by?: string;
  /** Runtime context: "PPT" or "IE"; behaviours for IE's HTML runtime are not PowerPoint's to play. */
  rctx?: string;
  override?: 'normal' | 'childStyle';
}

/** A point of p:animMotion or p:animScale (CT_TLPoint), as fractions: of the slide for motion, 1 = 100% for scale. */
export interface Point { x: number; y: number; }

export interface Tav {
  el: XmlElement;
  /** 0..1 along the simple duration; undefined for "indefinite" (spread evenly). */
  tm?: number;
  fmla?: string;
  val?: Variant;
}

interface NodeBase { el: XmlElement; cTn: CommonTimeNode; }
interface BehaviorBase extends NodeBase { bhvr: Behavior; }

export interface ParNode extends NodeBase { kind: 'par'; }
export interface ExclNode extends NodeBase { kind: 'excl'; }
export interface SeqNode extends NodeBase {
  kind: 'seq';
  concurrent?: boolean;
  prevAc?: 'none' | 'skipTimed';
  nextAc?: 'none' | 'seek';
  prev: Condition[];
  next: Condition[];
}
export interface SetNode extends BehaviorBase { kind: 'set'; to?: Variant; }
export interface AnimNode extends BehaviorBase {
  kind: 'anim';
  by?: string;
  from?: string;
  to?: string;
  calcmode?: CalcMode;
  valueType?: ValueType;
  tavs: Tav[];
}
export interface AnimClrNode extends BehaviorBase {
  kind: 'animClr';
  clrSpc?: 'rgb' | 'hsl';
  dir?: 'cw' | 'ccw';
  /** rgb: channel offsets as fractions; hsl: hue offset in degrees, saturation and luminance offsets as fractions. */
  by?: { space: 'rgb'; r: number; g: number; b: number } | { space: 'hsl'; h: number; s: number; l: number };
  from?: ColorDef;
  to?: ColorDef;
}
export interface AnimEffectNode extends BehaviorBase {
  kind: 'animEffect';
  transition: 'in' | 'out' | 'none';
  filter?: string;
  prLst?: string;
  progress?: Variant;
}
export interface AnimMotionNode extends BehaviorBase {
  kind: 'animMotion';
  origin?: 'parent' | 'layout';
  path?: string;
  pathEditMode?: 'relative' | 'fixed';
  /** Degrees. */
  rAng?: number;
  ptsTypes?: string;
  by?: Point;
  from?: Point;
  to?: Point;
  rCtr?: Point;
}
/** Degrees; positive turns clockwise. */
export interface AnimRotNode extends BehaviorBase { kind: 'animRot'; by?: number; from?: number; to?: number; }
export interface AnimScaleNode extends BehaviorBase { kind: 'animScale'; zoomContents?: boolean; by?: Point; from?: Point; to?: Point; }
export interface CmdNode extends BehaviorBase { kind: 'cmd'; type?: 'evt' | 'call' | 'verb'; cmd?: string; }
export interface MediaNode extends NodeBase {
  kind: 'audio' | 'video';
  target?: Target;
  vol: number;
  mute: boolean;
  numSld: number;
  showWhenStopped: boolean;
  isNarration?: boolean;
  fullScrn?: boolean;
}

export type BehaviorNode = SetNode | AnimNode | AnimClrNode | AnimEffectNode | AnimMotionNode | AnimRotNode | AnimScaleNode | CmdNode;
export type ContainerNode = ParNode | SeqNode | ExclNode;
export type TimeNode = ContainerNode | BehaviorNode | MediaNode;

interface BuildBase { el: XmlElement; spid: number; grpId: number; uiExpand: boolean; }
export type Build =
  | BuildBase & {
      kind: 'paragraph';
      build: 'allAtOnce' | 'p' | 'cust' | 'whole';
      bldLvl: number;
      animBg: boolean;
      autoUpdateAnimBg: boolean;
      rev: boolean;
      advAuto: Time;
      templates: { lvl: number; nodes: TimeNode[] }[];
    }
  | BuildBase & { kind: 'diagram'; bld: string }
  | BuildBase & { kind: 'oleChart'; bld: string; animBg: boolean }
  | BuildBase & {
      kind: 'graphic'; asOne: boolean;
      sub?: { el: XmlElement; kind: 'diagram' | 'chart'; bld?: string; rev?: boolean; animBg?: boolean };
    };

export interface Timing {
  el: XmlElement;
  /** The children of p:tnLst; PowerPoint writes one, the tmRoot par. */
  nodes: TimeNode[];
  builds: Build[];
}

// ------------------------------------------------------------------------------------------------- simple types

const oneOf = <T extends string>(v: string | undefined, values: readonly T[]): T | undefined =>
  v !== undefined && (values as readonly string[]).includes(v) ? (v as T) : undefined;

/** ST_TLTime: an unsigned number of milliseconds, or "indefinite". */
export function parseTime(v: string | undefined): Time | undefined {
  if (v === undefined) return undefined;
  if (v.trim() === 'indefinite') return 'indefinite';
  const n = Number(v.trim());
  return Number.isFinite(n) && n >= 0 ? n : undefined;
}

/** "0,0; 0.25,0.07; 1,1" to pairs, sorted by time; undefined when nothing usable is there. */
export function parseTmFilter(v: string | undefined): [number, number][] | undefined {
  if (!v) return undefined;
  const pairs: [number, number][] = [];
  for (const item of v.split(';').slice(0, 1000)) {
    const [t, p] = item.split(',').map(s => Number(s.trim()));
    if (t !== undefined && p !== undefined && Number.isFinite(t) && Number.isFinite(p)) pairs.push([Math.min(1, Math.max(0, t)), p]);
  }
  pairs.sort((a, b) => a[0] - b[0]);
  return pairs.length >= 2 ? pairs : undefined;
}

const angle = (v: string | undefined) => (v === undefined ? undefined : int(v, 0) / 60000);
const point = (el: XmlElement | undefined): Point | undefined =>
  (el ? { x: percent(attr(el, 'x'), 0), y: percent(attr(el, 'y'), 0) } : undefined);

export function parseVariant(el: XmlElement | undefined): Variant | undefined {
  for (const c of children(el)) {
    const v = attr(c, 'val');
    if (is(c, 'p:boolVal')) return { kind: 'bool', value: bool(v, false) };
    if (is(c, 'p:intVal')) return { kind: 'int', value: int(v, 0) };
    if (is(c, 'p:fltVal')) return { kind: 'float', value: num(v, 0) };
    if (is(c, 'p:strVal')) return { kind: 'string', value: v ?? '' };
    if (is(c, 'p:clrVal')) {
      const color = colorIn(c);
      if (color) return { kind: 'color', value: color };
    }
  }
  return undefined;
}

// ---------------------------------------------------------------------------------------------- targets

function graphicSub(el: XmlElement): ShapeSub {
  const dgm = child(el, 'a:dgm');
  const chart = child(el, 'a:chart');
  const step = (e: XmlElement) => (attr(e, 'bldStep') !== undefined ? { bldStep: attr(e, 'bldStep')! } : {});
  return {
    kind: 'graphic', el,
    ...(dgm ? { diagram: { id: attr(dgm, 'id') ?? '', ...step(dgm) } } : {}),
    ...(chart
      ? { chart: { seriesIdx: int(attr(chart, 'seriesIdx'), -1), categoryIdx: int(attr(chart, 'categoryIdx'), -1), ...step(chart) } }
      : {}),
  };
}

function shapeSub(spTgt: XmlElement): ShapeSub | undefined {
  for (const c of children(spTgt)) {
    if (is(c, 'p:bg')) return { kind: 'bg' };
    if (is(c, 'p:subSp')) return { kind: 'subShape', spid: int(attr(c, 'spid'), 0) };
    if (is(c, 'p:oleChartEl')) return { kind: 'oleChart', type: attr(c, 'type') ?? 'series', lvl: int(attr(c, 'lvl'), 0) };
    if (is(c, 'p:graphicEl')) return graphicSub(c);
    if (is(c, 'p:txEl')) {
      const rg = child(c, 'p:pRg') ?? child(c, 'p:charRg');
      if (!rg) return { kind: 'text' };
      return { kind: 'text', range: { unit: is(rg, 'p:pRg') ? 'paragraph' : 'char', st: num(attr(rg, 'st'), 0), end: num(attr(rg, 'end'), 0) } };
    }
  }
  return undefined;
}

/** p:tgtEl (CT_TLTimeTargetElement). */
export function parseTarget(tgtEl: XmlElement | undefined): Target | undefined {
  for (const c of children(tgtEl)) {
    if (is(c, 'p:sldTgt')) return { kind: 'slide' };
    if (is(c, 'p:spTgt')) {
      const sub = shapeSub(c);
      return { kind: 'shape', spid: int(attr(c, 'spid'), 0), ...(sub ? { sub } : {}) };
    }
    if (is(c, 'p:inkTgt')) return { kind: 'ink', spid: int(attr(c, 'spid'), 0) };
    if (is(c, 'p:sndTgt')) {
      const embed = attrNs(c, 'r', 'embed');
      const name = attr(c, 'name');
      return { kind: 'sound', el: c, ...(embed ? { embed } : {}), ...(name ? { name } : {}) };
    }
  }
  return undefined;
}

// ------------------------------------------------------------------------------------------- conditions

export function parseCondition(el: XmlElement): Condition {
  const evt = oneOf(attr(el, 'evt'), EVENTS);
  const target = parseTarget(child(el, 'p:tgtEl'));
  const tn = child(el, 'p:tn');
  const rtn = oneOf(attr(child(el, 'p:rtn'), 'val'), ['first', 'last', 'all'] as const);
  return {
    el, delay: parseTime(attr(el, 'delay')) ?? 0,
    ...(evt ? { evt } : {}), ...(target ? { target } : {}), ...(tn ? { tn: int(attr(tn, 'val'), 0) } : {}), ...(rtn ? { rtn } : {}),
  };
}

const conditions = (list: XmlElement | undefined) => children(list, 'p:cond').map(parseCondition);

// ------------------------------------------------------------------------------------------- time nodes

/** Deeper trees than this are cut off: PowerPoint nests five or six levels; this only stops hostile input. */
const MAX_DEPTH = 64;
const MAX_NODES = 50_000;

interface ParseState { nodes: number; }

/** Set `out[key]` when the value is defined (the types keep optional properties absent rather than undefined). */
function put<T extends object, K extends keyof T>(out: T, key: K, value: T[K] | undefined): void {
  if (value !== undefined) out[key] = value;
}

function parseIterate(el: XmlElement): Iterate {
  const abs = child(el, 'p:tmAbs');
  return {
    el, type: oneOf(attr(el, 'type'), ['el', 'wd', 'lt'] as const) ?? 'el', backwards: bool(attr(el, 'backwards'), false),
    interval: abs
      ? { kind: 'abs', ms: parseTime(attr(abs, 'val')) ?? 0 }
      : { kind: 'pct', fraction: percent(attr(child(el, 'p:tmPct'), 'val'), 0) },
  };
}

function parseCTn(el: XmlElement, depth: number, state: ParseState): CommonTimeNode {
  const a = (n: string) => attr(el, n);
  const out: CommonTimeNode = {
    el, autoRev: bool(a('autoRev'), false), accel: percent(a('accel'), 0), decel: percent(a('decel'), 0), spd: percent(a('spd'), 1),
    start: conditions(child(el, 'p:stCondLst')), end: conditions(child(el, 'p:endCondLst')),
    children: nodeList(child(el, 'p:childTnLst'), depth + 1, state), sub: nodeList(child(el, 'p:subTnLst'), depth + 1, state),
  };
  const repeat = parseTime(a('repeatCount'));
  put(out, 'id', int(a('id')));
  put(out, 'presetClass', oneOf(a('presetClass'), PRESET_CLASSES));
  put(out, 'presetID', int(a('presetID')));
  put(out, 'presetSubtype', int(a('presetSubtype')));
  put(out, 'nodeType', oneOf(a('nodeType'), NODE_TYPES));
  put(out, 'dur', parseTime(a('dur')));
  put(out, 'fill', oneOf(a('fill'), ['remove', 'freeze', 'hold', 'transition'] as const));
  put(out, 'restart', oneOf(a('restart'), ['always', 'whenNotActive', 'never'] as const));
  put(out, 'repeatCount', repeat === undefined || repeat === 'indefinite' ? repeat : repeat / 1000);
  put(out, 'repeatDur', parseTime(a('repeatDur')));
  put(out, 'grpId', int(a('grpId')));
  put(out, 'afterEffect', bool(a('afterEffect')));
  put(out, 'masterRel', oneOf(a('masterRel'), ['sameClick', 'lastClick', 'nextClick'] as const));
  put(out, 'bldLvl', int(a('bldLvl')));
  put(out, 'display', bool(a('display')));
  put(out, 'syncBehavior', oneOf(a('syncBehavior'), ['canSlip', 'locked'] as const));
  put(out, 'tmFilter', parseTmFilter(a('tmFilter')));
  put(out, 'evtFilter', a('evtFilter'));
  put(out, 'nodePh', bool(a('nodePh')));
  const endSync = child(el, 'p:endSync');
  if (endSync) out.endSync = parseCondition(endSync);
  const iterate = child(el, 'p:iterate');
  if (iterate) out.iterate = parseIterate(iterate);
  return out;
}

function parseBehavior(cBhvr: XmlElement | undefined, depth: number, state: ParseState): { cTn: CommonTimeNode; bhvr: Behavior } | undefined {
  const cTnEl = child(cBhvr, 'p:cTn');
  if (!cBhvr || !cTnEl) return undefined;
  const a = (n: string) => attr(cBhvr, n);
  const attrNames = children(child(cBhvr, 'p:attrNameLst'), 'p:attrName').map(n => textOf(n).trim()).filter(Boolean);
  const bhvr: Behavior = { el: cBhvr, attrNames };
  put(bhvr, 'target', parseTarget(child(cBhvr, 'p:tgtEl')));
  put(bhvr, 'additive', oneOf(a('additive'), ['base', 'sum', 'repl', 'mult', 'none'] as const));
  put(bhvr, 'accumulate', oneOf(a('accumulate'), ['none', 'always'] as const));
  put(bhvr, 'xfrmType', oneOf(a('xfrmType'), ['pt', 'img'] as const));
  put(bhvr, 'override', oneOf(a('override'), ['normal', 'childStyle'] as const));
  for (const k of ['from', 'to', 'by', 'rctx'] as const) put(bhvr, k, a(k));
  return { cTn: parseCTn(cTnEl, depth, state), bhvr };
}

function parseTav(t: XmlElement): Tav {
  const out: Tav = { el: t };
  const tm = attr(t, 'tm');
  if (tm !== undefined && tm !== 'indefinite') out.tm = Math.min(1, Math.max(0, percent(tm, 0)));
  put(out, 'fmla', attr(t, 'fmla') || undefined);
  put(out, 'val', parseVariant(child(t, 'p:val')));
  return out;
}

function parseMedia(el: XmlElement, depth: number, state: ParseState): MediaNode | undefined {
  const media = child(el, 'p:cMediaNode');
  const cTnEl = child(media, 'p:cTn');
  if (!media || !cTnEl) return undefined;
  const node: MediaNode = {
    kind: is(el, 'p:audio') ? 'audio' : 'video', el, cTn: parseCTn(cTnEl, depth, state),
    vol: percent(attr(media, 'vol'), 0.5), mute: bool(attr(media, 'mute'), false), numSld: int(attr(media, 'numSld'), 1),
    showWhenStopped: bool(attr(media, 'showWhenStopped'), true),
  };
  put(node, 'target', parseTarget(child(media, 'p:tgtEl')));
  if (node.kind === 'audio') node.isNarration = bool(attr(el, 'isNarration'), false);
  else node.fullScrn = bool(attr(el, 'fullScrn'), false);
  return node;
}

function parseContainer(el: XmlElement, depth: number, state: ParseState): ContainerNode | undefined {
  const cTnEl = child(el, 'p:cTn');
  if (!cTnEl) return undefined;
  const cTn = parseCTn(cTnEl, depth, state);
  if (is(el, 'p:par')) return { kind: 'par', el, cTn };
  if (is(el, 'p:excl')) return { kind: 'excl', el, cTn };
  const seq: SeqNode = { kind: 'seq', el, cTn, prev: conditions(child(el, 'p:prevCondLst')), next: conditions(child(el, 'p:nextCondLst')) };
  put(seq, 'concurrent', bool(attr(el, 'concurrent')));
  put(seq, 'prevAc', oneOf(attr(el, 'prevAc'), ['none', 'skipTimed'] as const));
  put(seq, 'nextAc', oneOf(attr(el, 'nextAc'), ['none', 'seek'] as const));
  return seq;
}

/** A behaviour element, given its parsed cBhvr. */
function parseBehaviorNode(el: XmlElement, base: { el: XmlElement; cTn: CommonTimeNode; bhvr: Behavior }): BehaviorNode | undefined {
  const a = (n: string) => attr(el, n);
  if (is(el, 'p:set')) {
    const node: SetNode = { kind: 'set', ...base };
    put(node, 'to', parseVariant(child(el, 'p:to')));
    return node;
  }
  if (is(el, 'p:anim')) {
    const node: AnimNode = { kind: 'anim', ...base, tavs: children(child(el, 'p:tavLst'), 'p:tav').slice(0, 1000).map(parseTav) };
    for (const k of ['by', 'from', 'to'] as const) put(node, k, a(k));
    put(node, 'calcmode', oneOf(a('calcmode'), ['discrete', 'lin', 'fmla'] as const));
    put(node, 'valueType', oneOf(a('valueType'), ['str', 'num', 'clr'] as const));
    return node;
  }
  if (is(el, 'p:animClr')) {
    const node: AnimClrNode = { kind: 'animClr', ...base };
    put(node, 'clrSpc', oneOf(a('clrSpc'), ['rgb', 'hsl'] as const));
    put(node, 'dir', oneOf(a('dir'), ['cw', 'ccw'] as const));
    const rgb = child(child(el, 'p:by'), 'p:rgb');
    const hsl = child(child(el, 'p:by'), 'p:hsl');
    const pc = (e: XmlElement | undefined, k: string) => percent(attr(e, k), 0);
    if (rgb) node.by = { space: 'rgb', r: pc(rgb, 'r'), g: pc(rgb, 'g'), b: pc(rgb, 'b') };
    else if (hsl) node.by = { space: 'hsl', h: angle(attr(hsl, 'h')) ?? 0, s: pc(hsl, 's'), l: pc(hsl, 'l') };
    put(node, 'from', colorIn(child(el, 'p:from')));
    put(node, 'to', colorIn(child(el, 'p:to')));
    return node;
  }
  if (is(el, 'p:animEffect')) {
    const node: AnimEffectNode = { kind: 'animEffect', ...base, transition: oneOf(a('transition'), ['in', 'out', 'none'] as const) ?? 'in' };
    put(node, 'filter', a('filter'));
    put(node, 'prLst', a('prLst'));
    put(node, 'progress', parseVariant(child(el, 'p:progress')));
    return node;
  }
  if (is(el, 'p:animMotion')) {
    const node: AnimMotionNode = { kind: 'animMotion', ...base };
    put(node, 'origin', oneOf(a('origin'), ['parent', 'layout'] as const));
    put(node, 'path', a('path'));
    put(node, 'pathEditMode', oneOf(a('pathEditMode'), ['relative', 'fixed'] as const));
    put(node, 'rAng', angle(a('rAng')));
    put(node, 'ptsTypes', a('ptsTypes'));
    for (const k of ['by', 'from', 'to', 'rCtr'] as const) put(node, k, point(child(el, `p:${k}`)));
    return node;
  }
  if (is(el, 'p:animRot')) {
    const node: AnimRotNode = { kind: 'animRot', ...base };
    for (const k of ['by', 'from', 'to'] as const) put(node, k, angle(a(k)));
    return node;
  }
  if (is(el, 'p:animScale')) {
    const node: AnimScaleNode = { kind: 'animScale', ...base };
    put(node, 'zoomContents', bool(a('zoomContents')));
    for (const k of ['by', 'from', 'to'] as const) put(node, k, point(child(el, `p:${k}`)));
    return node;
  }
  if (is(el, 'p:cmd')) {
    const node: CmdNode = { kind: 'cmd', ...base };
    put(node, 'type', oneOf(a('type'), ['evt', 'call', 'verb'] as const));
    put(node, 'cmd', a('cmd'));
    return node;
  }
  return undefined;
}

function parseNode(el: XmlElement, depth: number, state: ParseState): TimeNode | undefined {
  if (++state.nodes > MAX_NODES) return undefined;
  if (is(el, 'p:par') || is(el, 'p:excl') || is(el, 'p:seq')) return parseContainer(el, depth, state);
  if (is(el, 'p:audio') || is(el, 'p:video')) return parseMedia(el, depth, state);
  const b = parseBehavior(child(el, 'p:cBhvr'), depth, state);
  return b ? parseBehaviorNode(el, { el, ...b }) : undefined;
}

function nodeList(list: XmlElement | undefined, depth: number, state: ParseState): TimeNode[] {
  if (!list || depth > MAX_DEPTH) return [];
  const out: TimeNode[] = [];
  for (const c of effectiveChildren(list, ANIM_UNDERSTOOD)) {
    const n = parseNode(c, depth, state);
    if (n) out.push(n);
  }
  return out;
}

// ------------------------------------------------------------------------------------------------- builds

function parseBuild(el: XmlElement, state: ParseState): Build | undefined {
  const base: BuildBase = { el, spid: int(attr(el, 'spid'), 0), grpId: int(attr(el, 'grpId'), 0), uiExpand: bool(attr(el, 'uiExpand'), false) };
  if (is(el, 'p:bldP')) {
    return {
      ...base, kind: 'paragraph', build: oneOf(attr(el, 'build'), ['allAtOnce', 'p', 'cust', 'whole'] as const) ?? 'whole',
      bldLvl: int(attr(el, 'bldLvl'), 1), animBg: bool(attr(el, 'animBg'), false), autoUpdateAnimBg: bool(attr(el, 'autoUpdateAnimBg'), true),
      rev: bool(attr(el, 'rev'), false), advAuto: parseTime(attr(el, 'advAuto')) ?? 'indefinite',
      templates: children(child(el, 'p:tmplLst'), 'p:tmpl')
        .map(t => ({ lvl: int(attr(t, 'lvl'), 0), nodes: nodeList(child(t, 'p:tnLst'), 1, state) })),
    };
  }
  if (is(el, 'p:bldDgm')) return { ...base, kind: 'diagram', bld: attr(el, 'bld') ?? 'whole' };
  if (is(el, 'p:bldOleChart')) return { ...base, kind: 'oleChart', bld: attr(el, 'bld') ?? 'allAtOnce', animBg: bool(attr(el, 'animBg'), true) };
  if (!is(el, 'p:bldGraphic')) return undefined;
  const sub = child(el, 'p:bldSub');
  if (!sub) return { ...base, kind: 'graphic', asOne: true };
  const props = child(sub, 'a:bldDgm') ?? child(sub, 'a:bldChart');
  const out: { el: XmlElement; kind: 'diagram' | 'chart'; bld?: string; rev?: boolean; animBg?: boolean } =
    { el: sub, kind: is(props, 'a:bldChart') ? 'chart' : 'diagram' };
  put(out, 'bld', attr(props, 'bld'));
  put(out, 'rev', bool(attr(props, 'rev')));
  put(out, 'animBg', bool(attr(props, 'animBg')));
  return { ...base, kind: 'graphic', asOne: false, sub: out };
}

/** One time node element (p:par, p:seq, p:anim, ...) read on its own; undefined for anything else. */
export function parseTimeNode(el: XmlElement): TimeNode | undefined {
  return parseNode(el, 0, { nodes: 0 });
}

/** Read p:timing into plain data. Never throws for well-formed XML; unknown elements are skipped. */
export function parseTiming(timing: XmlElement): Timing {
  const state: ParseState = { nodes: 0 };
  const bldLst = child(timing, 'p:bldLst');
  const builds = (bldLst ? effectiveChildren(bldLst, ANIM_UNDERSTOOD) : []).map(b => parseBuild(b, state)).filter((b): b is Build => !!b);
  return { el: timing, nodes: nodeList(child(timing, 'p:tnLst'), 0, state), builds };
}

// -------------------------------------------------------------------------------------------------- queries

/** Every node of a list, depth first in document order, children (and sub-lists) after their parent. */
export function* walk(nodes: readonly TimeNode[]): Generator<TimeNode> {
  for (const n of nodes) {
    yield n;
    yield* walk(n.cTn.children);
    yield* walk(n.cTn.sub);
  }
}

export const isBehavior = (n: TimeNode): n is BehaviorNode => 'bhvr' in n;
export const isContainer = (n: TimeNode): n is ContainerNode => n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl';

/** The time root (nodeType tmRoot, else the first par of tnLst). */
export function rootOf(timing: Timing): ContainerNode | undefined {
  const pars = timing.nodes.filter(isContainer);
  return pars.find(n => n.cTn.nodeType === 'tmRoot') ?? pars[0];
}

/** The main sequence: the seq under the root with nodeType mainSeq, else its first seq that is not interactive. */
export function mainSequence(timing: Timing): SeqNode | undefined {
  const seqs = (rootOf(timing)?.cTn.children ?? []).filter((n): n is SeqNode => n.kind === 'seq');
  return seqs.find(s => s.cTn.nodeType === 'mainSeq') ?? seqs.find(s => s.cTn.nodeType !== 'interactiveSeq' && triggerOf(s) === undefined);
}

/** The shape that starts an interactive sequence (its onClick condition), if it is one. */
export function triggerOf(seq: SeqNode): number | undefined {
  for (const c of seq.cTn.start) if (c.evt === 'onClick' && c.target?.kind === 'shape') return c.target.spid;
  return undefined;
}

/** The interactive sequences under the root, with the shape that triggers each. */
export function interactiveSequences(timing: Timing): { seq: SeqNode; trigger: number }[] {
  const out: { seq: SeqNode; trigger: number }[] = [];
  for (const n of rootOf(timing)?.cTn.children ?? []) {
    if (n.kind !== 'seq' || (n.cTn.nodeType !== 'interactiveSeq' && n.cTn.nodeType !== undefined)) continue;
    const trigger = triggerOf(n);
    if (trigger !== undefined) out.push({ seq: n, trigger });
  }
  return out;
}
