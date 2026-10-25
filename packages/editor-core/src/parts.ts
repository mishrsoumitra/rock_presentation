// packages/editor-core/src/parts.ts  -- relationships and parts behind the XML: references, garbage, copies, media
import { extensionOf, normalizeName, partKey, relativeTarget, relsNameFor, resolveTarget, type OpcPackage } from '@pptx/opc';
import { relKind } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';

const REL_NS = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
]);

/** Attributes in the relationships namespace (r:id, r:embed, r:link, r:dm, ...) under an element. */
export function relRefs(root: XmlElement): { el: XmlElement; local: string; value: string }[] {
  const out: { el: XmlElement; local: string; value: string }[] = [];
  const walk = (e: XmlElement) => {
    for (const a of e.attrs) if (REL_NS.has(a.ns)) out.push({ el: e, local: a.local, value: a.value });
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(root);
  return out;
}

export const referencedIds = (root: XmlElement) => new Set(relRefs(root).map(r => r.value));

/** Point every relationship reference under `root` named `from` at `to` instead. */
export function renameRelRefs(root: XmlElement, map: ReadonlyMap<string, string>): void {
  const walk = (e: XmlElement) => {
    for (const a of e.attrs) if (REL_NS.has(a.ns) && map.has(a.value)) a.value = map.get(a.value)!;
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(root);
}

/**
 * Remove a part's relationships among `candidates` that nothing in the part refers to any more, then every part
 * that was their target and that no relationship in the package reaches now (and so on down its own relationships).
 * Parts reached from elsewhere (a layout, a picture another slide shows) stay. Returns the parts removed.
 */
export function dropUnreferenced(pkg: OpcPackage, part: string, candidates: Iterable<string>): string[] {
  const still = pkg.getPart(part) ? referencedIds(pkg.xml(part).root) : new Set<string>();
  const rels = pkg.relationships(part);
  const targets: string[] = [];
  for (const id of new Set(candidates)) {
    if (still.has(id)) continue;
    const rel = rels.find(r => r.id === id);
    if (!rel) continue;
    pkg.removeRelationship(part, id);
    if (!rel.external) targets.push(resolveTarget(normalizeName(part), rel.target));
  }
  return collectOrphans(pkg, targets);
}

/** Remove the parts among `names` that no relationship reaches (except their own), following their relationships. */
export function collectOrphans(pkg: OpcPackage, names: string[]): string[] {
  const removed: string[] = [];
  const queue = [...names];
  while (queue.length) {
    const name = queue.shift()!;
    const p = pkg.getPart(name);
    if (!p) continue;
    if (pkg.incomingRelationships(p.name).some(r => partKey(r.source) !== partKey(p.name))) continue;
    const next = pkg.relationships(p.name).filter(r => !r.external).map(r => resolveTarget(p.name, r.target));
    pkg.removePart(p.name, { force: true });
    removed.push(p.name);
    queue.push(...next);
  }
  return removed;
}

/** A free part name from a pattern with one `#`: /ppt/media/image#.png gives the next number after the largest. */
export function freeName(pkg: OpcPackage, pattern: string): string {
  const [head, tail] = pattern.split('#') as [string, string];
  let max = 0;
  for (const n of pkg.partNames()) {
    if (!n.toLowerCase().startsWith(head.toLowerCase()) || !n.toLowerCase().endsWith(tail.toLowerCase())) continue;
    const num = Number(n.slice(head.length, n.length - tail.length));
    if (Number.isInteger(num)) max = Math.max(max, num);
  }
  for (let i = max + 1; ; i++) if (!pkg.getPart(`${head}${i}${tail}`)) return `${head}${i}${tail}`;
}

/** The pattern a copy of a part is named by: its folder and stem with the number taken out. */
export function patternOf(name: string): string {
  const m = /^(.*\/)([^/]*?)(\d*)(\.[^./]*)?$/.exec(name);
  if (!m) return `${name}#`;
  return `${m[1]}${m[2]}#${m[4] ?? ''}`;
}

/** Relationship kinds whose targets are shared, not copied, when content is copied within one package. */
const SHARED = new Set(['slideLayout', 'slideMaster', 'notesMaster', 'handoutMaster', 'theme', 'image', 'media', 'video', 'audio',
  'hyperlink', 'slide', 'font', 'presProps', 'viewProps', 'tableStyles', 'commentAuthors']);

/**
 * Copy a part and what it owns into `pkg` under a new name, returning the name. Relationships whose kind is shared
 * (layouts, pictures, links) keep pointing where they did when the copy stays in the same package; `mapShared`
 * decides for shared targets when copying from another package: the name to point at, undefined to copy the
 * target too, or null to leave the relationship out. `copied` remembers parts already copied in this operation,
 * so shared sub-parts are copied once.
 */
export function copyPart(
  pkg: OpcPackage, from: OpcPackage, name: string, copied: Map<string, string>,
  mapShared?: (kind: string, target: string) => string | null | undefined,
): string {
  const key = partKey(name);
  const done = copied.get(key);
  if (done) return done;
  const src = from.getPart(name);
  if (!src) throw new Error(`missing part ${name}`);
  const newName = freeName(pkg, patternOf(src.name));
  copied.set(key, newName);
  const bytes = src.dirty && src.xml ? undefined : src.bytes;
  const part = pkg.addPart(newName, src.contentType, bytes ?? new Uint8Array(0));
  if (!bytes) {
    // a part edited in this session: copy its tree, not its stale bytes
    part.xml = structuredCloneDoc(from.xml(src.name));
    part.encoding = src.encoding ?? { charset: 'utf-8', bom: false };
  }
  for (const r of from.relationships(src.name)) {
    let target: string;
    if (r.external) target = r.target;
    else {
      const t = resolveTarget(src.name, r.target);
      const kind = relKind(r.type);
      const shared = !SHARED.has(kind) ? undefined : from === pkg ? t : mapShared?.(kind, t);
      if (shared === null || !from.getPart(t)) continue;
      target = shared ?? copyPart(pkg, from, t, copied, mapShared);
    }
    addRelationshipWithId(pkg, newName, r.id, r.type, target, r.external);
  }
  return newName;
}

/** A deep copy of a parsed document, parent pointers rebuilt. */
function structuredCloneDoc(doc: ReturnType<OpcPackage['xml']>): ReturnType<OpcPackage['xml']> {
  const cloneEl = (e: XmlElement, parent: XmlElement | null): XmlElement => {
    const c: XmlElement = { ...e, attrs: e.attrs.map(a => ({ ...a })), children: [], parent };
    c.children = e.children.map(n => (n.kind === 'element' ? cloneEl(n, c) : { ...n }));
    return c;
  };
  return { declaration: doc.declaration, prolog: doc.prolog.map(n => ({ ...n })), root: cloneEl(doc.root, null), epilog: doc.epilog.map(n => ({ ...n })) };
}

/**
 * Add a relationship that must keep a given id (a copied part's XML refers to it). The package numbers new
 * relationships itself, so this adds one and renames it when the numbers differ.
 */
export function addRelationshipWithId(pkg: OpcPackage, source: string, id: string, type: string, target: string, external: boolean): void {
  const got = pkg.addRelationship(source, type, target, external ? { external: true } : {});
  if (got === id) return;
  const root = pkg.xml(relsNameFor(normalizeName(source))).root;
  for (const c of root.children) {
    if (c.kind !== 'element') continue;
    const a = c.attrs.find(x => x.local === 'Id' && x.ns === '');
    if (a?.value === got) a.value = id;
  }
}

/** Relative target text for a relationship from `source` to part `target`. */
export const targetText = (source: string, target: string) => relativeTarget(normalizeName(source), normalizeName(target));

const hashes = new WeakMap<Uint8Array, string>();

async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes as Uint8Array<ArrayBuffer>);
  return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, '0')).join('');
}

/** Hash some bytes (cached per buffer), and every media part of the package, so `findMedia` can answer at once. */
export async function hashMedia(ed: Editor, bytes: Uint8Array): Promise<string> {
  for (const n of ed.pkg.partNames()) {
    if (!n.startsWith('/ppt/media/')) continue;
    const p = ed.pkg.getPart(n)!;
    if (!hashes.has(p.bytes)) hashes.set(p.bytes, await sha256(p.bytes));
  }
  let h = hashes.get(bytes);
  if (!h) {
    h = await sha256(bytes);
    hashes.set(bytes, h);
  }
  return h;
}

/** A media part with these contents, if the package has one (PowerPoint stores a picture inserted twice once). */
export function findMedia(pkg: OpcPackage, hash: string): string | undefined {
  for (const n of pkg.partNames()) {
    if (!n.startsWith('/ppt/media/')) continue;
    if (hashes.get(pkg.getPart(n)!.bytes) === hash) return n;
  }
  return undefined;
}

/** Store media bytes as a new part (with a Default content type for its extension), or reuse an identical one. */
export function mediaPart(pkg: OpcPackage, bytes: Uint8Array, ext: string, contentType: string, hash: string, stem = 'image'): string {
  const existing = findMedia(pkg, hash);
  if (existing) return existing;
  const name = freeName(pkg, `/ppt/media/${stem}#.${ext}`);
  pkg.ensureDefaultContentType(ext, contentType);
  pkg.addPart(name, contentType, bytes);
  hashes.set(bytes, hash);
  return name;
}

export { extensionOf };
