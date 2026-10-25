// packages/anim/src/template.ts  -- an effect as a reusable template, and a template as a new effect in a slide
//
// A template is an effect's p:par with the slide-specific parts taken out: no time node ids, every target written
// as <p:spTgt spid="0"/>, no iteration or sub-node list (those are options), and the outer timing (trigger,
// delay, group id) left to the sequence it goes into. Elements are written with their logical prefixes (p:, a:)
// whatever the source file used, so templates compare and instantiate the same everywhere.
import { parseXml, type XmlElement } from '@pptx/xml';
import { attr, child, children, createElement, is, lname, setAttr, uriFor } from '@pptx/model';
import { effectLength } from './schedule.ts';
import { parseTiming, type ContainerNode, type Timing } from './timing.ts';

export interface EffectTemplate {
  /** Attributes of the effect's p:cTn besides id, presetID, presetClass, presetSubtype, grpId and nodeType. */
  attrs: Record<string, string>;
  /** The effect's behaviours (its p:childTnLst content). */
  behaviors: string;
  /** How long one pass of the effect lasts, ms. */
  duration: number;
}

/** A template extracted from a deck (presets.generated.ts), with where it came from. */
export interface PresetSource extends EffectTemplate {
  presetClass: 'entr' | 'exit' | 'emph' | 'path' | 'verb' | 'mediacall';
  presetID: number;
  presetSubtype: number;
  /** corpus/external path of the deck and the slide part. */
  source: string;
  /** How many effects in the corpus have this preset, and how many of them have exactly this template. */
  occurrences: number;
  agreeing: number;
}

/** Where an instantiated effect points: a shape, its background, or a paragraph range of its text. */
export interface EffectTarget {
  spid: number;
  paragraphs?: { st: number; end: number };
  background?: boolean;
}

/** Effect attributes that belong to the slide (ids, builds, the trigger) or are user options (repeat), not to the preset. */
const EFFECT_ATTRS_DROPPED = new Set([
  'id', 'presetID', 'presetClass', 'presetSubtype', 'grpId', 'nodeType', 'repeatCount', 'repeatDur', 'restart', 'afterEffect', 'bldLvl',
]);
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Logical names and unprefixed attributes only; text kept. Returns undefined for elements of unknown namespaces. */
function canonical(el: XmlElement, out: string[], strip: (e: XmlElement) => boolean): boolean {
  const name = lname(el);
  if (name.startsWith('{')) return false;
  out.push('<', name);
  for (const a of el.attrs) {
    if (a.ns !== '' || (is(el, 'p:cTn') && a.local === 'id')) continue;
    out.push(' ', a.local, '="', esc(is(el, 'p:spTgt') && a.local === 'spid' ? '0' : a.value), '"');
  }
  const kids = el.children.filter(c => c.kind !== 'comment' && c.kind !== 'pi' && !(c.kind === 'element' && strip(c)));
  const content = is(el, 'p:spTgt') ? [] : kids;
  if (!content.some(c => c.kind === 'element' || ((c.kind === 'text' || c.kind === 'cdata') && c.value !== ''))) {
    out.push('/>');
    return true;
  }
  out.push('>');
  for (const c of content) {
    if (c.kind === 'element') {
      if (!canonical(c, out, strip)) return false;
    } else if (c.kind === 'text' || c.kind === 'cdata') out.push(esc(c.value));
  }
  out.push('</', name, '>');
  return true;
}

/** An element's subtree in template form; undefined when it holds something a template cannot (relationships,
 *  extension namespaces). */
export function templateXml(el: XmlElement, strip: (e: XmlElement) => boolean = () => false): string | undefined {
  const out: string[] = [];
  const hasRel = (e: XmlElement): boolean =>
    e.attrs.some(a => a.ns !== '' && a.prefix !== 'xmlns') || children(e).some(hasRel);
  if (hasRel(el)) return undefined;
  return canonical(el, out, strip) ? out.join('') : undefined;
}

const NS_DECL = `xmlns:p="${uriFor('p')}" xmlns:a="${uriFor('a')}"`;

function parseFragment(xml: string): XmlElement {
  return parseXml(`<p:timing ${NS_DECL}>${xml}</p:timing>`).root;
}

/** How long one pass of a template's behaviours lasts, scheduled as one effect. */
export function templateDuration(behaviors: string): number {
  const xml = `<p:tnLst><p:par><p:cTn><p:childTnLst>${behaviors}</p:childTnLst></p:cTn></p:par></p:tnLst>`;
  const timing: Timing = parseTiming(parseFragment(xml));
  const root = timing.nodes[0];
  return root ? effectLength(root).pass : 0;
}

/** The template of an effect (a container with a presetClass); undefined when it cannot be made one. */
export function templateOf(effect: ContainerNode): EffectTemplate | undefined {
  const strip = (e: XmlElement) => is(e, 'p:subTnLst') || is(e, 'p:iterate');
  const list = child(effect.cTn.el, 'p:childTnLst');
  if (!list) return undefined;
  const parts: string[] = [];
  for (const c of children(list)) {
    const x = templateXml(c, strip);
    if (x === undefined) return undefined;
    parts.push(x);
  }
  const attrs: Record<string, string> = {};
  for (const a of effect.cTn.el.attrs) if (a.ns === '' && !EFFECT_ATTRS_DROPPED.has(a.local)) attrs[a.local] = a.value;
  const behaviors = parts.join('');
  return { attrs, behaviors, duration: templateDuration(behaviors) };
}

/** Copy a parsed template element into a document under `parent`, with the document's own prefixes. */
function importElement(parent: XmlElement, src: XmlElement): XmlElement {
  const el = createElement(parent, lname(src));
  for (const a of src.attrs) if (a.ns === '') setAttr(el, a.local, a.value);
  for (const c of src.children) {
    if (c.kind === 'element') {
      el.children.push(importElement(el, c));
      el.selfClosing = false;
    } else if (c.kind === 'text' || c.kind === 'cdata') {
      el.children.push({ kind: 'text', value: c.value });
      el.selfClosing = false;
    }
  }
  return el;
}

function retarget(el: XmlElement, target: EffectTarget): void {
  for (const c of children(el)) retarget(c, target);
  if (!is(el, 'p:spTgt')) return;
  setAttr(el, 'spid', target.spid);
  el.children = [];
  el.selfClosing = true;
  if (target.background) {
    el.children.push(createElement(el, 'p:bg'));
    el.selfClosing = false;
  } else if (target.paragraphs) {
    const txEl = createElement(el, 'p:txEl');
    const pRg = createElement(txEl, 'p:pRg', { st: target.paragraphs.st, end: target.paragraphs.end });
    txEl.children.push(pRg);
    txEl.selfClosing = false;
    el.children.push(txEl);
    el.selfClosing = false;
  }
  for (const c of el.children) if (c.kind === 'element') c.parent = el;
}

/** Scale the timeline inside an effect (its p:childTnLst, or the effect's p:par) to a new length: durations and
 *  delays; a set one millisecond before the end stays one millisecond before the end. */
export function rescale(el: XmlElement, from: number, to: number): void {
  if (is(el, 'p:par') || is(el, 'p:cTn')) {
    const list = child(is(el, 'p:par') ? child(el, 'p:cTn') : el, 'p:childTnLst');
    if (list) rescale(list, from, to);
    return;
  }
  if (from <= 0 || to < 0 || from === to) return;
  const k = to / from;
  const visit = (e: XmlElement) => {
    if (is(e, 'p:cTn')) {
      const dur = attr(e, 'dur');
      if (dur !== undefined && dur !== 'indefinite' && Number(dur) > 1) {
        setAttr(e, 'dur', Math.max(1, Math.round(Number(dur) * k)));
      }
    }
    if (is(e, 'p:cond')) {
      const delay = attr(e, 'delay');
      if (delay !== undefined && delay !== 'indefinite' && Number(delay) > 0) {
        const d = Number(delay);
        setAttr(e, 'delay', d === from - 1 ? Math.max(0, Math.round(to) - 1) : Math.round(d * k));
      }
    }
    for (const c of children(e)) visit(c);
  };
  for (const c of children(el)) visit(c);
}

export interface InstantiateOptions {
  /** New length of one pass, ms; the template's timeline is scaled to it. */
  duration?: number;
  /** For motion paths: the path to follow instead of the template's. */
  path?: string;
}

/**
 * A new effect p:par for a template, created for `parent` (its prefixes are bound there) but not inserted. Its
 * p:cTn has the preset attributes and the template's, a start delay of 0 and no id yet.
 */
export function instantiate(
  parent: XmlElement, preset: { presetClass: string; presetID: number; presetSubtype: number }, template: EffectTemplate,
  target: EffectTarget, options: InstantiateOptions = {},
): XmlElement {
  const par = createElement(parent, 'p:par');
  const { presetID, presetClass, presetSubtype } = preset;
  const cTn = createElement(par, 'p:cTn', { id: 0, presetID, presetClass, presetSubtype });
  const { fill, ...rest } = template.attrs;
  for (const [k, v] of Object.entries(rest)) setAttr(cTn, k, v);
  setAttr(cTn, 'fill', fill ?? 'hold');
  par.children.push(cTn);
  par.selfClosing = false;
  const st = createElement(cTn, 'p:stCondLst');
  st.children.push(createElement(st, 'p:cond', { delay: 0 }));
  st.selfClosing = false;
  const list = createElement(cTn, 'p:childTnLst');
  cTn.children.push(st, list);
  cTn.selfClosing = false;
  const source = parseFragment(template.behaviors);
  for (const c of children(source)) list.children.push(importElement(list, c));
  list.selfClosing = list.children.length === 0;
  retarget(list, target);
  if (options.duration !== undefined) rescale(list, template.duration, options.duration);
  if (options.path !== undefined) for (const m of descendantsNamed(list, 'p:animMotion')) setAttr(m, 'path', options.path);
  return par;
}

function descendantsNamed(el: XmlElement, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  const visit = (e: XmlElement) => {
    for (const c of children(e)) {
      if (is(c, name)) out.push(c);
      visit(c);
    }
  };
  visit(el);
  return out;
}
