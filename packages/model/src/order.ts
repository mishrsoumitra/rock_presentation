// packages/model/src/order.ts  -- schema child order, so the editor never writes an element where PowerPoint rejects it
import type { XmlElement, XmlNode } from '@pptx/xml';
import { createElement, is, lname } from './dom.ts';
import { GLOBALS, TYPES } from './schema.generated.ts';

interface Slot { type: string | null; rank: number; repeat: boolean; }

const parsed = new Map<string, Map<string, Slot>>();

function slotsOf(type: string): Map<string, Slot> | undefined {
  const cached = parsed.get(type);
  if (cached) return cached;
  const src = TYPES[type];
  if (src === undefined) return undefined;
  const slots = new Map<string, Slot>();
  for (const entry of src.split(';')) {
    const eq = entry.indexOf('=');
    const at = entry.lastIndexOf('@');
    const repeat = entry.endsWith('*');
    const name = entry.slice(0, eq);
    if (!slots.has(name)) slots.set(name, { type: entry.slice(eq + 1, at) || null, rank: Number(entry.slice(at + 1, repeat ? -1 : undefined)), repeat });
  }
  parsed.set(type, slots);
  return slots;
}

/** mc:AlternateContent, mc:Choice and mc:Fallback are transparent: their children belong to the element around them. */
function schemaParent(el: XmlElement): XmlElement | null {
  let p = el.parent;
  while (p && (is(p, 'mc:AlternateContent') || is(p, 'mc:Choice') || is(p, 'mc:Fallback'))) p = p.parent;
  return p;
}

const typeCache = new WeakMap<XmlElement, string | null>();

/**
 * The schema complex type of an element ('a:CT_ShapeProperties'), found by walking down from its part's root
 * element. Null for what the ECMA schemas do not cover: extension elements, and roots of other vocabularies.
 */
export function schemaType(el: XmlElement): string | null {
  const cached = typeCache.get(el);
  if (cached !== undefined) return cached;
  const parent = schemaParent(el);
  let type: string | null;
  if (!parent) type = GLOBALS[lname(el)] ?? null;
  else {
    const pt = schemaType(parent);
    type = pt ? slotsOf(pt)?.get(lname(el))?.type ?? null : null;
  }
  typeCache.set(el, type);
  return type;
}

/** True when the schema allows `name` as a child of `parent`. */
export function allowsChild(parent: XmlElement, name: string): boolean {
  const t = schemaType(parent);
  return t !== null && slotsOf(t)?.has(name) === true;
}

/**
 * Insert `child` into `parent` where the schema puts it. A slot that holds at most one element is replaced if
 * occupied, including by another alternative of the same choice (a:solidFill replaces a:gradFill). Elements the
 * schema does not know, such as Office extension elements, keep their place. Throws when the schema does not allow
 * the child here, so a bad write is caught before PowerPoint would repair the file.
 */
export function insertInOrder(parent: XmlElement, child: XmlElement): void {
  const type = schemaType(parent);
  const slots = type ? slotsOf(type) : undefined;
  if (!slots) throw new Error(`no schema order known for ${lname(parent)}`);
  const mine = slots.get(lname(child));
  if (!mine) throw new Error(`${lname(child)} is not allowed in ${lname(parent)} (${type})`);
  const rankOf = (n: XmlNode) => (n.kind === 'element' ? slots.get(lname(n))?.rank : undefined);
  child.parent = parent;
  parent.selfClosing = false;
  if (!mine.repeat) {
    const existing = parent.children.findIndex(n => rankOf(n) === mine.rank);
    if (existing >= 0) {
      const old = parent.children[existing]!;
      if (old.kind === 'element') old.parent = null;
      parent.children[existing] = child;
      return;
    }
  }
  let at = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const r = rankOf(parent.children[i]!);
    if (r !== undefined && r > mine.rank) {
      at = i;
      break;
    }
  }
  parent.children.splice(at, 0, child);
}

/** The child `name` of `parent`, created and inserted in schema order when missing. */
export function ensureChild(parent: XmlElement, name: string, attrs: Record<string, string | number | boolean> = {}): XmlElement {
  for (const c of parent.children) if (c.kind === 'element' && is(c, name)) return c;
  const el = createElement(parent, name, attrs);
  insertInOrder(parent, el);
  return el;
}

/** Remove every child of `parent` with one of these names. Returns how many were removed. */
export function removeChildren(parent: XmlElement, ...names: string[]): number {
  const before = parent.children.length;
  parent.children = parent.children.filter(c => !(c.kind === 'element' && names.some(n => is(c, n))));
  return before - parent.children.length;
}
