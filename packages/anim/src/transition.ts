// packages/anim/src/transition.ts  -- p:transition as typed data, PowerPoint 2010 and 2013 kinds included, and writing it
//
// ECMA-376 Part 1, 19.3.1.50 and 19.5; [MS-PPTX] for the p14 kinds (vortex, ripple, prism, ...), p15:prstTrans
// (curtains, airplane, ...) and p159:morph. Those sit in mc:AlternateContent: the Choice holds them with the
// duration in p14:dur, the Fallback a fade for older readers. Written transitions follow the same pattern.
import { XMLNS_NS, type XmlElement } from '@pptx/xml';
import {
  append, attr, attrNs, bool, child, children, clone, createElement, insertInOrder, int, is, lookupPrefix, nsKey, remove, setAttr, setAttrNs,
  uriFor, type SlideView,
} from '@pptx/model';
import { ANIM_UNDERSTOOD } from './timing.ts';

export type TransitionKind =
  | 'none' | 'blinds' | 'checker' | 'circle' | 'dissolve' | 'comb' | 'cover' | 'cut' | 'diamond' | 'fade' | 'newsflash' | 'plus'
  | 'pull' | 'push' | 'random' | 'randomBar' | 'split' | 'strips' | 'wedge' | 'wheel' | 'wipe' | 'zoom'
  // PowerPoint 2010 (p14)
  | 'vortex' | 'switch' | 'flip' | 'ripple' | 'honeycomb' | 'prism' | 'doors' | 'window' | 'ferris' | 'gallery' | 'conveyor'
  | 'pan' | 'glitter' | 'warp' | 'flythrough' | 'flash' | 'shred' | 'reveal' | 'wheelReverse'
  // PowerPoint 2013 (p15:prstTrans) and 2016 (p159:morph)
  | 'curtains' | 'drape' | 'wind' | 'prestige' | 'fracture' | 'crush' | 'peelOff' | 'pageCurlDouble' | 'pageCurlSingle'
  | 'airplane' | 'origami' | 'fallOver' | 'morph'
  | 'unknown';

export type TransitionSpeed = 'fast' | 'med' | 'slow';

export interface TransitionOptions {
  /** l, r, u, d (side), lu, ru, ld, rd (corner), in, out, horz, vert or center, depending on the kind. */
  dir?: string;
  /** split: the orientation of the split line. */
  orient?: 'horz' | 'vert';
  /** fade, cut, reveal: through black. */
  thruBlk?: boolean;
  /** wheel, wheelReverse. */
  spokes?: number;
  /** glitter: diamond or hexagon; shred: strip or rectangle. */
  pattern?: string;
  /** prism: only the content turns (PowerPoint's Rotate and Orbit); inverted: seen from inside (Box, Orbit). */
  isContent?: boolean;
  isInverted?: boolean;
  /** flythrough. */
  hasBounce?: boolean;
  /** p15:prstTrans mirroring. */
  invX?: boolean;
  invY?: boolean;
  /** morph: byObject, byWord or byChar. */
  option?: string;
}

export interface Transition extends TransitionOptions {
  kind: TransitionKind;
  /** The p:transition in effect (the branch of mc:AlternateContent this reader takes). */
  el: XmlElement;
  /** The kind's own element (p:wipe, p14:vortex, ...), absent for 'none'. */
  effectEl?: XmlElement;
  speed: TransitionSpeed;
  /** Milliseconds: p14:dur when given, else the speed's duration. */
  duration: number;
  /** Whether a click advances to the next slide (advClick, default true). */
  advanceOnClick: boolean;
  /** Advance automatically after this many milliseconds (advTm). */
  advanceAfter?: number;
  sound?: { kind: 'start'; embed?: string; name?: string; loop: boolean } | { kind: 'stop' };
}

/** What setTransition writes. Duration in milliseconds; omitted: the speed's (fast when no speed either). */
export interface TransitionSpec extends TransitionOptions {
  kind: TransitionKind;
  speed?: TransitionSpeed;
  duration?: number;
  advanceOnClick?: boolean;
  advanceAfter?: number;
}

/** PowerPoint's durations for spd. */
export const SPEED_MS: Readonly<Record<TransitionSpeed, number>> = { fast: 500, med: 750, slow: 1000 };

const ECMA_KINDS = new Set<string>([
  'blinds', 'checker', 'circle', 'dissolve', 'comb', 'cover', 'cut', 'diamond', 'fade', 'newsflash', 'plus', 'pull', 'push', 'random',
  'randomBar', 'split', 'strips', 'wedge', 'wheel', 'wipe', 'zoom',
]);
const P14_KINDS = new Set<string>([
  'vortex', 'switch', 'flip', 'ripple', 'honeycomb', 'prism', 'doors', 'window', 'ferris', 'gallery', 'conveyor', 'pan', 'glitter',
  'warp', 'flythrough', 'flash', 'shred', 'reveal', 'wheelReverse',
]);
const PRESET_KINDS = new Set<string>([
  'curtains', 'drape', 'wind', 'prestige', 'fracture', 'crush', 'peelOff', 'pageCurlDouble', 'pageCurlSingle', 'airplane', 'origami', 'fallOver',
]);

/** Default attribute values by kind, from the schemas. */
const DEFAULTS = new Map<TransitionKind, TransitionOptions>([
  ['blinds', { dir: 'horz' }], ['checker', { dir: 'horz' }], ['comb', { dir: 'horz' }],
  ['randomBar', { dir: 'horz' }], ['cover', { dir: 'l' }], ['pull', { dir: 'l' }], ['push', { dir: 'l' }],
  ['wipe', { dir: 'l' }], ['strips', { dir: 'lu' }], ['zoom', { dir: 'out' }],
  ['split', { orient: 'horz', dir: 'out' }], ['wheel', { spokes: 4 }], ['wheelReverse', { spokes: 4 }],
  ['cut', { thruBlk: false }], ['fade', { thruBlk: false }], ['vortex', { dir: 'l' }], ['switch', { dir: 'r' }],
  ['flip', { dir: 'r' }], ['ripple', { dir: 'center' }],
  ['prism', { dir: 'l', isContent: false, isInverted: false }], ['doors', { dir: 'horz' }],
  ['window', { dir: 'horz' }], ['ferris', { dir: 'l' }], ['gallery', { dir: 'l' }], ['conveyor', { dir: 'l' }],
  ['pan', { dir: 'l' }], ['glitter', { dir: 'l', pattern: 'diamond' }], ['warp', { dir: 'out' }],
  ['flythrough', { dir: 'in', hasBounce: false }], ['shred', { pattern: 'strip', dir: 'in' }],
  ['reveal', { dir: 'l', thruBlk: false }],
]);

function kindOf(el: XmlElement): TransitionKind {
  const k = nsKey(el.ns);
  if (k === 'p' && ECMA_KINDS.has(el.local)) return el.local as TransitionKind;
  if (k === 'p14' && P14_KINDS.has(el.local)) return el.local as TransitionKind;
  if (k === 'p15' && el.local === 'prstTrans') {
    const prst = attr(el, 'prst') ?? '';
    return PRESET_KINDS.has(prst) ? (prst as TransitionKind) : 'unknown';
  }
  return k === 'p159' && el.local === 'morph' ? 'morph' : 'unknown';
}

function readSound(snd: XmlElement | undefined): Transition['sound'] | undefined {
  const st = child(snd, 'p:stSnd');
  if (!st) return child(snd, 'p:endSnd') ? { kind: 'stop' } : undefined;
  const s = child(st, 'p:snd');
  const embed = attrNs(s, 'r', 'embed');
  const name = attr(s, 'name');
  return { kind: 'start', loop: bool(attr(st, 'loop'), false), ...(embed ? { embed } : {}), ...(name ? { name } : {}) };
}

/** Read a p:transition element. */
export function readTransition(el: XmlElement): Transition {
  const effectEl = children(el).find(c => !is(c, 'p:sndAc') && !is(c, 'p:extLst'));
  const kind = effectEl ? kindOf(effectEl) : 'none';
  const speed = (['fast', 'med', 'slow'] as const).find(s => s === attr(el, 'spd')) ?? 'fast';
  const dur = int(attrNs(el, 'p14', 'dur'));
  const out: Transition = {
    kind, el, speed, duration: dur !== undefined && dur >= 0 ? dur : SPEED_MS[speed], advanceOnClick: bool(attr(el, 'advClick'), true),
    ...DEFAULTS.get(kind),
  };
  const advTm = int(attr(el, 'advTm'));
  if (advTm !== undefined && advTm >= 0) out.advanceAfter = advTm;
  if (effectEl) {
    out.effectEl = effectEl;
    const a = (n: string) => attr(effectEl, n);
    if (a('dir') !== undefined) out.dir = a('dir')!;
    if (a('orient') === 'horz' || a('orient') === 'vert') out.orient = a('orient') as 'horz' | 'vert';
    for (const k of ['thruBlk', 'isContent', 'isInverted', 'hasBounce', 'invX', 'invY'] as const) {
      const v = bool(a(k));
      if (v !== undefined) out[k] = v;
    }
    const spokes = int(a('spokes'));
    if (spokes !== undefined) out.spokes = Math.max(1, Math.min(64, spokes));
    if (a('pattern') !== undefined) out.pattern = a('pattern')!;
    if (a('option') !== undefined) out.option = a('option')!;
  }
  const sound = readSound(child(el, 'p:sndAc'));
  if (sound) out.sound = sound;
  return out;
}

/** The slide's transition as this reader plays it (p14, p15 and p159 understood), or undefined when it has none. */
export function slideTransition(slide: SlideView): Transition | undefined {
  const el = slide.transition(ANIM_UNDERSTOOD);
  return el ? readTransition(el) : undefined;
}

const NAMES = new Map<TransitionKind, string>([
  ['none', 'None'], ['cut', 'Cut'], ['fade', 'Fade'], ['push', 'Push'], ['wipe', 'Wipe'], ['split', 'Split'],
  ['reveal', 'Reveal'], ['randomBar', 'Random Bars'], ['circle', 'Shape'], ['diamond', 'Shape'], ['plus', 'Shape'],
  ['zoom', 'Shape'], ['cover', 'Cover'], ['pull', 'Uncover'], ['flash', 'Flash'], ['dissolve', 'Dissolve'],
  ['checker', 'Checkerboard'], ['blinds', 'Blinds'], ['wheel', 'Clock'], ['wheelReverse', 'Clock'],
  ['wedge', 'Clock'], ['ripple', 'Ripple'], ['honeycomb', 'Honeycomb'], ['glitter', 'Glitter'], ['vortex', 'Vortex'],
  ['shred', 'Shred'], ['switch', 'Switch'], ['flip', 'Flip'], ['gallery', 'Gallery'], ['doors', 'Doors'],
  ['warp', 'Zoom'], ['ferris', 'Ferris Wheel'], ['conveyor', 'Conveyor'], ['window', 'Window'],
  ['flythrough', 'Fly Through'], ['pan', 'Pan'], ['comb', 'Comb'], ['strips', 'Strips'], ['newsflash', 'Newsflash'],
  ['random', 'Random'], ['morph', 'Morph'], ['curtains', 'Curtains'], ['drape', 'Drape'], ['wind', 'Wind'],
  ['prestige', 'Prestige'], ['fracture', 'Fracture'], ['crush', 'Crush'], ['peelOff', 'Peel Off'],
  ['pageCurlDouble', 'Page Curl'], ['pageCurlSingle', 'Page Curl'], ['airplane', 'Airplane'], ['origami', 'Origami'],
  ['fallOver', 'Fall Over'], ['unknown', 'Unknown'],
]);

/** The name PowerPoint's Transitions gallery gives the kind (Cube, Box, Rotate and Orbit are prisms). */
export function transitionName(t: { kind: TransitionKind; isContent?: boolean; isInverted?: boolean }): string {
  if (t.kind === 'prism') return t.isContent ? (t.isInverted ? 'Orbit' : 'Rotate') : t.isInverted ? 'Box' : 'Cube';
  return NAMES.get(t.kind) ?? t.kind;
}

// ------------------------------------------------------------------------------------------------- writing

/** The spd PowerPoint writes beside a p14:dur. */
export function speedFor(ms: number): TransitionSpeed {
  return ms < 625 ? 'fast' : ms < 875 ? 'med' : 'slow';
}

/** The p:transition, or the mc:AlternateContent holding one in any branch, directly under the slide's root. */
function transitionHolders(root: XmlElement): XmlElement[] {
  const holds = (c: XmlElement) => is(c, 'mc:AlternateContent') && children(c).some(b => !!child(b, 'p:transition'));
  return children(root).filter(c => is(c, 'p:transition') || holds(c));
}

/** Remove the slide transition, whichever form it is written in. */
export function clearTransition(slide: SlideView): void {
  const holders = transitionHolders(slide.root);
  for (const h of holders) remove(h);
  if (holders.length) slide.owner.touch();
}

/** A prefix bound to the namespace at `el`, declaring it there when nothing in scope does. */
function declare(el: XmlElement, key: 'p14' | 'p15' | 'p159'): string {
  const uri = uriFor(key);
  const bound = lookupPrefix(el, uri);
  if (bound) return bound;
  el.attrs.push({ prefix: 'xmlns', local: key, value: uri, ns: XMLNS_NS });
  return key;
}

/** Fill a p:transition: speed, duration, advance, the kind's element, and a sound action kept from before. */
function fill(t: XmlElement, spec: TransitionSpec, kindName: string | undefined, sound: XmlElement | undefined, withDuration: boolean): void {
  const ms = spec.duration ?? SPEED_MS[spec.speed ?? 'fast'];
  const speed = spec.speed ?? speedFor(ms);
  setAttr(t, 'spd', speed === 'fast' ? undefined : speed);
  if (withDuration) setAttrNs(t, 'p14', 'dur', String(Math.max(0, Math.round(ms))));
  if (spec.advanceOnClick === false) setAttr(t, 'advClick', '0');
  if (spec.advanceAfter !== undefined) setAttr(t, 'advTm', Math.max(0, Math.round(spec.advanceAfter)));
  if (kindName) {
    const e = createElement(t, kindName);
    const d = DEFAULTS.get(spec.kind) ?? {};
    const keep = (k: keyof TransitionOptions) => spec[k] !== undefined && spec[k] !== d[k];
    if (kindName === 'p15:prstTrans') setAttr(e, 'prst', spec.kind);
    if (keep('dir')) setAttr(e, 'dir', spec.dir);
    if (keep('orient')) setAttr(e, 'orient', spec.orient);
    for (const k of ['thruBlk', 'isContent', 'isInverted', 'hasBounce', 'invX', 'invY'] as const) {
      if (keep(k)) setAttr(e, k, spec[k] ? '1' : '0');
    }
    if (keep('spokes')) setAttr(e, 'spokes', spec.spokes);
    if (keep('pattern')) setAttr(e, 'pattern', spec.pattern);
    if (spec.kind === 'morph') setAttr(e, 'option', spec.option ?? 'byObject');
    append(t, e);
  }
  if (sound) append(t, sound);
}

/**
 * Set the slide transition. ECMA kinds at their speed's duration are written as a plain p:transition; other durations
 * and the PowerPoint 2010+ kinds as mc:AlternateContent, a Choice with p14:dur and a Fallback PowerPoint 2007 can
 * play (the same kind, or a fade). A sound action already on the slide is kept.
 */
export function setTransition(slide: SlideView, transition: TransitionSpec): void {
  const spec: TransitionSpec = transition.kind === 'unknown' ? { ...transition, kind: 'none' } : transition;
  const root = slide.root;
  const sound = child(slide.transition(ANIM_UNDERSTOOD), 'p:sndAc');
  // a fallback may name the same sound through a relationship of its own: keep that one there
  const fallbackSound = child(slide.transition(new Set()), 'p:sndAc');
  for (const h of transitionHolders(root)) remove(h);
  const ecma = ECMA_KINDS.has(spec.kind);
  const t = createElement(root, 'p:transition');
  insertInOrder(root, t);
  const defaultDuration = spec.duration === undefined || spec.duration === SPEED_MS[spec.speed ?? speedFor(spec.duration)];
  if ((ecma || spec.kind === 'none') && defaultDuration) {
    fill(t, spec, ecma ? `p:${spec.kind}` : undefined, sound, false);
    slide.owner.touch();
    return;
  }
  // the AlternateContent takes the transition's place in schema order
  const ac = createElement(root, 'mc:AlternateContent');
  root.children[root.children.indexOf(t)] = ac;
  t.parent = null;
  const choice = createElement(ac, 'mc:Choice');
  const fallback = createElement(ac, 'mc:Fallback');
  append(ac, choice);
  append(ac, fallback);
  const key = P14_KINDS.has(spec.kind) ? 'p14' : PRESET_KINDS.has(spec.kind) ? 'p15' : spec.kind === 'morph' ? 'p159' : 'p14';
  // where PowerPoint declares them: p15 on the Choice, p14 and p159 on the AlternateContent (p14:dur declares its own)
  setAttr(choice, 'Requires', declare(key === 'p15' ? choice : ac, key));
  const inChoice = createElement(choice, 'p:transition');
  append(choice, inChoice);
  const choiceKind = ecma ? `p:${spec.kind}` : key === 'p15' ? 'p15:prstTrans' : key === 'p159' ? 'p159:morph' : `p14:${spec.kind}`;
  fill(inChoice, spec, spec.kind === 'none' ? undefined : choiceKind, sound, true);
  const inFallback = createElement(fallback, 'p:transition');
  append(fallback, inFallback);
  // older readers get the same ECMA kind, or a plain fade for the newer kinds
  const fallbackSpec: TransitionSpec = ecma || spec.kind === 'none' ? spec : {
    kind: 'fade', ...(spec.speed ? { speed: spec.speed } : {}), ...(spec.duration !== undefined ? { duration: spec.duration } : {}),
    ...(spec.advanceOnClick !== undefined ? { advanceOnClick: spec.advanceOnClick } : {}),
    ...(spec.advanceAfter !== undefined ? { advanceAfter: spec.advanceAfter } : {}),
  };
  const otherSound = fallbackSound && fallbackSound !== sound ? fallbackSound : sound ? clone(sound) : undefined;
  fill(inFallback, fallbackSpec, spec.kind === 'none' ? undefined : `p:${fallbackSpec.kind}`, otherSound, false);
  slide.owner.touch();
}
