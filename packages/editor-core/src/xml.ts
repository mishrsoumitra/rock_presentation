// packages/editor-core/src/xml.ts  -- building and placing elements: fragments, ids, the units of a shape tree
import { attr, effectiveChildren, is, isStrictUri, lookupNamespace, nsKey, shapeKind, UNDERSTOOD, uriFor, type NsKey } from '@pptx/model';
import { parseXml, XMLNS_NS, type XmlElement, type XmlNode } from '@pptx/xml';

const FRAGMENT_KEYS: NsKey[] = ['a', 'p', 'r', 'mc', 'p14', 'a14', 'a16'];

/**
 * Elements from XML text written with the usual prefixes (a:, p:, r:, ...), in the document's flavour
 * (Transitional or Strict). Declarations the place they will go already makes are dropped, so the result looks like
 * PowerPoint's own XML; anything else is declared on each top element. The elements are not inserted anywhere.
 */
export function fragment(context: XmlElement, xml: string, strict: boolean): XmlElement[] {
  const decls = FRAGMENT_KEYS.map(k => `xmlns:${k}="${uriFor(k, strict)}"`).join(' ');
  const root = parseXml(`<f ${decls}>${xml}</f>`).root;
  const out: XmlElement[] = [];
  for (const node of root.children) {
    if (node.kind !== 'element') continue;
    const used = new Map<string, string>();
    const walk = (e: XmlElement) => {
      if (e.prefix) used.set(e.prefix, e.ns);
      for (const a of e.attrs) if (a.prefix && a.ns !== XMLNS_NS) used.set(a.prefix, a.ns);
      for (const c of e.children) if (c.kind === 'element') walk(c);
    };
    walk(node);
    // declarations the context already makes are dropped; anything else the subtree uses is declared on it
    node.attrs = node.attrs.filter(a => !(a.ns === XMLNS_NS && a.prefix === 'xmlns' && lookupNamespace(context, a.local) === a.value));
    for (const [prefix, uri] of used) {
      if (lookupNamespace(context, prefix) === uri || node.attrs.some(a => a.ns === XMLNS_NS && a.local === prefix)) continue;
      node.attrs.unshift({ prefix: 'xmlns', local: prefix, value: uri, ns: XMLNS_NS });
    }
    node.parent = null;
    out.push(node);
  }
  return out;
}

/**
 * Bring XML from a deck of the other flavour into this one: every Transitional namespace URI becomes its Strict
 * twin or the reverse, in names and declarations, so a Transitional deck never gets Strict elements pasted in.
 */
export function convertFlavour(root: XmlElement, strict: boolean): void {
  const conv = (uri: string) => {
    const key = nsKey(uri);
    return key && isStrictUri(uri) !== strict ? uriFor(key, strict) : uri;
  };
  const walk = (e: XmlElement) => {
    e.ns = conv(e.ns);
    for (const a of e.attrs) {
      if (a.ns === XMLNS_NS) a.value = conv(a.value);
      else if (a.ns) a.ns = conv(a.ns);
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(root);
}

/** One element from XML text; see `fragment`. */
export function element(context: XmlElement, xml: string, strict: boolean): XmlElement {
  const els = fragment(context, xml, strict);
  if (els.length !== 1) throw new Error(`expected one element, got ${els.length}`);
  return els[0]!;
}

export const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Every cNvPr id in a part, wherever it is (groups, alternate content branches, graphic frames). */
export function shapeIds(root: XmlElement): Set<number> {
  const ids = new Set<number>();
  const walk = (e: XmlElement) => {
    if (e.local === 'cNvPr') {
      const id = Number(attr(e, 'id'));
      if (Number.isInteger(id)) ids.add(id);
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(root);
  return ids;
}

/** A shape id not used in the part: one more than the largest, as PowerPoint numbers them. */
export function nextShapeId(root: XmlElement, taken: ReadonlySet<number> = new Set()): number {
  let max = 0;
  for (const id of shapeIds(root)) max = Math.max(max, id);
  for (const id of taken) max = Math.max(max, id);
  return max + 1;
}

/**
 * The element of a shape tree that a shape is drawn as: the shape itself, or the mc:AlternateContent around it.
 * Z-order, grouping and deleting move these units, so every branch of an alternate content block goes together.
 */
export function unitOf(el: XmlElement): XmlElement {
  let u = el;
  while (u.parent && (is(u.parent, 'mc:Choice') || is(u.parent, 'mc:Fallback') || is(u.parent, 'mc:AlternateContent'))) u = u.parent;
  return u;
}

/** The shapes a unit holds: the shape itself, or the shape of each branch of an alternate content block. */
export function shapesOfUnit(unit: XmlElement): XmlElement[] {
  if (shapeKind(unit)) return [unit];
  const out: XmlElement[] = [];
  const walk = (e: XmlElement) => {
    for (const c of e.children) {
      if (c.kind !== 'element') continue;
      if (shapeKind(c)) out.push(c);
      else if (is(c, 'mc:Choice') || is(c, 'mc:Fallback')) walk(c);
    }
  };
  walk(unit);
  return out;
}

/** The units of a shape tree or group, in drawing order (non-shape children such as nvGrpSpPr left out). */
export function unitsIn(container: XmlElement): XmlElement[] {
  return container.children.filter((c): c is XmlElement => c.kind === 'element' && (shapeKind(c) !== undefined || (is(c, 'mc:AlternateContent') && shapesOfUnit(c).length > 0)));
}

/** The same shape in the other branches of the alternate content block it sits in (matched by id). */
export function twins(el: XmlElement): XmlElement[] {
  const unit = unitOf(el);
  if (unit === el) return [el];
  const id = idOf(el);
  return shapesOfUnit(unit).filter(s => idOf(s) === id);
}

export function idOf(shape: XmlElement): number {
  for (const c of shape.children) {
    if (c.kind !== 'element' || !c.local.startsWith('nv')) continue;
    for (const d of c.children) if (d.kind === 'element' && d.local === 'cNvPr') return Number(attr(d, 'id') ?? 0);
  }
  return 0;
}

/** Insert `node` into `parent` at child position `index` (into the raw child list). */
export function insertAt(parent: XmlElement, index: number, node: XmlNode): void {
  parent.children.splice(Math.max(0, Math.min(index, parent.children.length)), 0, node);
  if (node.kind === 'element') node.parent = parent;
  parent.selfClosing = false;
}

/** The visible shape elements under a container, through alternate content as the renderer reads it. */
export function effectiveShapes(container: XmlElement): XmlElement[] {
  return effectiveChildren(container, UNDERSTOOD).filter(c => shapeKind(c) !== undefined);
}
