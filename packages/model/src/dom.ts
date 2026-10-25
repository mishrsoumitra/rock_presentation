// packages/model/src/dom.ts  -- helpers over the lossless tree that match by namespace, never by prefix
import { XMLNS_NS, type XmlAttr, type XmlElement, type XmlNode } from '@pptx/xml';
import { isStrictUri, NAMESPACES, nsKey, uriFor, type NsKey } from './ns.ts';

/** 'a:spPr' for an element in DrawingML whatever prefix the file used, '{uri}local' for unknown namespaces. */
export function lname(el: { ns: string; local: string }): string {
  const k = nsKey(el.ns);
  return k ? `${k}:${el.local}` : `{${el.ns}}${el.local}`;
}

function split(name: string): [NsKey, string] {
  const i = name.indexOf(':');
  return [name.slice(0, i) as NsKey, name.slice(i + 1)];
}

/** True when `node` is the element `name` ('a:spPr'), in either flavour of its namespace. Deliberately not a type
 *  guard: on a value already typed XmlElement, a guard would narrow the else branch of an if/else chain to never. */
export function is(node: XmlNode | null | undefined, name: string): boolean {
  if (!node || node.kind !== 'element') return false;
  const [key, local] = split(name);
  return node.local === local && nsKey(node.ns) === key;
}

export function child(el: XmlElement | null | undefined, name: string): XmlElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) if (c.kind === 'element' && is(c, name)) return c;
  return undefined;
}

export function children(el: XmlElement | null | undefined, name?: string): XmlElement[] {
  if (!el) return [];
  const out: XmlElement[] = [];
  for (const c of el.children) if (c.kind === 'element' && (name === undefined || is(c, name))) out.push(c);
  return out;
}

/** First element along a path of names: path(spPr, 'a:ln', 'a:solidFill'). */
export function path(el: XmlElement | null | undefined, ...names: string[]): XmlElement | undefined {
  let cur = el ?? undefined;
  for (const n of names) cur = child(cur, n);
  return cur;
}

/** Every descendant element named `name`, in document order. */
export function descendants(el: XmlElement, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (e: XmlElement) => {
    for (const c of e.children) {
      if (c.kind !== 'element') continue;
      if (is(c, name)) out.push(c);
      walk(c);
    }
  };
  walk(el);
  return out;
}

/** An unprefixed attribute. */
export function attr(el: XmlElement | null | undefined, local: string): string | undefined {
  if (!el) return undefined;
  for (const a of el.attrs) if (a.local === local && a.ns === '') return a.value;
  return undefined;
}

/** A namespaced attribute such as r:id or r:embed. */
export function attrNs(el: XmlElement | null | undefined, key: NsKey, local: string): string | undefined {
  if (!el) return undefined;
  for (const a of el.attrs) if (a.local === local && a.ns !== '' && nsKey(a.ns) === key) return a.value;
  return undefined;
}

/** Set (or with `undefined`, remove) an unprefixed attribute. An existing attribute keeps its position. */
export function setAttr(el: XmlElement, local: string, value: string | number | boolean | undefined): void {
  const at = el.attrs.findIndex(a => a.local === local && a.ns === '');
  if (value === undefined) {
    if (at >= 0) el.attrs.splice(at, 1);
    return;
  }
  const v = typeof value === 'boolean' ? (value ? '1' : '0') : String(value);
  if (at >= 0) el.attrs[at]!.value = v;
  else el.attrs.push({ prefix: null, local, value: v, ns: '' });
}

/** Set (or remove) a namespaced attribute, declaring the namespace on `el` if nothing in scope binds it. */
export function setAttrNs(el: XmlElement, key: NsKey, local: string, value: string | undefined): void {
  const at = el.attrs.findIndex(a => a.local === local && a.ns !== '' && nsKey(a.ns) === key);
  if (value === undefined) {
    if (at >= 0) el.attrs.splice(at, 1);
    return;
  }
  if (at >= 0) {
    el.attrs[at]!.value = value;
    return;
  }
  const { prefix, uri } = bindPrefix(el, key, el);
  if (prefix === null) throw new Error(`cannot bind ${key} as an attribute prefix`);
  el.attrs.push({ prefix, local, value, ns: uri });
}

/** The prefix bound to `uri` at `el` ('' for the default namespace), or undefined when nothing binds it. */
export function lookupPrefix(el: XmlElement | null, uri: string): string | undefined {
  const shadowed = new Set<string>();
  for (let e = el; e; e = e.parent) {
    for (const a of e.attrs) {
      if (a.ns !== XMLNS_NS) continue;
      const p = a.prefix === null ? '' : a.local;
      if (shadowed.has(p)) continue;
      shadowed.add(p);
      if (a.value === uri) return p;
    }
  }
  return undefined;
}

function boundUri(el: XmlElement | null, prefix: string): string | undefined {
  for (let e = el; e; e = e.parent) {
    for (const a of e.attrs) {
      if (a.ns !== XMLNS_NS) continue;
      if ((a.prefix === null ? '' : a.local) === prefix) return a.value;
    }
  }
  return undefined;
}

/** The flavour of the document `el` belongs to, judged from the nearest known namespace on the way up. */
function strictAt(el: XmlElement | null): boolean {
  for (let e = el; e; e = e.parent) if (nsKey(e.ns)) return isStrictUri(e.ns);
  return false;
}

/**
 * A prefix and URI to use for `key` at `context`: a prefix already bound there to either flavour of the namespace,
 * else a new declaration added to `declareOn`. Returns prefix null for the default namespace.
 */
function bindPrefix(context: XmlElement, key: NsKey, declareOn: XmlElement): { prefix: string | null; uri: string } {
  for (const uri of NAMESPACES[key] as readonly string[]) {
    const p = lookupPrefix(context, uri);
    if (p !== undefined) return { prefix: p === '' ? null : p, uri };
  }
  const uri = uriFor(key, strictAt(context));
  let prefix: string = key;
  for (let n = 2; boundUri(context, prefix) !== undefined && boundUri(context, prefix) !== uri; n++) prefix = `${key}${n}`;
  if (!declareOn.attrs.some(a => a.ns === XMLNS_NS && a.prefix === 'xmlns' && a.local === prefix)) {
    declareOn.attrs.unshift({ prefix: 'xmlns', local: prefix, value: uri, ns: XMLNS_NS });
  }
  return { prefix, uri };
}

/**
 * A new element named `name` ('a:solidFill') to be inserted under `parent`. Its prefix is one already bound at
 * `parent` when possible, so the written XML looks like PowerPoint's; otherwise it declares its own.
 * The element's parent pointer is set to `parent`, but it is not inserted: use insertInOrder or append.
 */
export function createElement(parent: XmlElement, name: string, attrs: Record<string, string | number | boolean> = {}): XmlElement {
  const [key, local] = split(name);
  const el: XmlElement = { kind: 'element', prefix: null, local, ns: '', attrs: [], children: [], selfClosing: true, parent };
  const { prefix, uri } = bindPrefix(parent, key, el);
  el.prefix = prefix;
  el.ns = uri;
  for (const [k, v] of Object.entries(attrs)) setAttr(el, k, v);
  return el;
}

/** Append `node` as the last child of `parent`. */
export function append(parent: XmlElement, node: XmlNode): void {
  parent.children.push(node);
  if (node.kind === 'element') node.parent = parent;
  parent.selfClosing = false;
}

export function remove(node: XmlElement): void {
  const p = node.parent;
  if (!p) return;
  const i = p.children.indexOf(node);
  if (i >= 0) p.children.splice(i, 1);
  node.parent = null;
}

/** Replace `old` by `node` in place. */
export function replace(old: XmlElement, node: XmlElement): void {
  const p = old.parent;
  if (!p) throw new Error('replace: element has no parent');
  const i = p.children.indexOf(old);
  p.children[i] = node;
  node.parent = p;
  old.parent = null;
}

/** Concatenated text of an element's direct text and CDATA children. */
export function textOf(el: XmlElement | null | undefined): string {
  if (!el) return '';
  let s = '';
  for (const c of el.children) if (c.kind === 'text' || c.kind === 'cdata') s += c.value;
  return s;
}

export function setText(el: XmlElement, text: string): void {
  el.children = text === '' ? [] : [{ kind: 'text', value: text }];
}

/** A deep copy. Namespace declarations in scope at the original are copied onto the clone's root when it needs them. */
export function clone(el: XmlElement, parent: XmlElement | null = null): XmlElement {
  const copy: XmlElement = {
    kind: 'element', prefix: el.prefix, local: el.local, ns: el.ns, selfClosing: el.selfClosing, parent,
    attrs: el.attrs.map((a): XmlAttr => ({ ...a })), children: [],
  };
  copy.children = el.children.map(c => (c.kind === 'element' ? clone(c, copy) : { ...c }));
  return copy;
}

/** Declare on `el` every prefix its subtree uses but that was declared outside it, so the subtree can be moved into
 *  another part. Elements and attributes already carry their resolved URIs, which are the bindings to declare. */
export function declareUsedNamespaces(el: XmlElement): void {
  const needed = new Map<string, string>();
  const walk = (e: XmlElement, declared: ReadonlySet<string>) => {
    let scope = declared;
    for (const a of e.attrs) {
      if (a.ns !== XMLNS_NS) continue;
      if (scope === declared) scope = new Set(declared);
      (scope as Set<string>).add(a.prefix === null ? '' : a.local);
    }
    const use = (prefix: string, uri: string) => {
      if (!scope.has(prefix) && !needed.has(prefix)) needed.set(prefix, uri);
    };
    if (e.prefix !== null) use(e.prefix, e.ns);
    else if (e.ns !== '') use('', e.ns);
    for (const a of e.attrs) if (a.prefix !== null && a.prefix !== 'xmlns' && a.prefix !== 'xml') use(a.prefix, a.ns);
    for (const c of e.children) if (c.kind === 'element') walk(c, scope);
  };
  walk(el, new Set());
  for (const [prefix, uri] of needed) {
    el.attrs.unshift(prefix === ''
      ? { prefix: null, local: 'xmlns', value: uri, ns: XMLNS_NS }
      : { prefix: 'xmlns', local: prefix, value: uri, ns: XMLNS_NS });
  }
}
