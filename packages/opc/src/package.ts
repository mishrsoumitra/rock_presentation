// packages/opc/src/package.ts
import {
  childElements, decodeXmlBytes, encodeXmlText, getAttr, parseXml, serializeXml,
  type XmlDocument, type XmlEncoding,
} from '@pptx/xml';
import { ContentTypes, RELS_CONTENT_TYPE, type ContentTypesCheckpoint } from './content-types.ts';
import { OpcError } from './errors.ts';
import { extensionOf, normalizeName, partKey, relativeTarget, relsNameFor, resolveTarget, sourceOfRels, unsafeEntryName } from './names.ts';
import { compoundFileError, DEFAULT_LIMITS, readZip, sniff, writeZip, type ZipLimits } from './zip.ts';

export const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface Relationship { id: string; type: string; target: string; external: boolean; }

export interface Part {
  readonly name: string;      // '/ppt/slides/slide1.xml'
  contentType: string;        // '' when [Content_Types].xml has nothing for it (reported in diagnostics)
  bytes: Uint8Array;          // as read, or as last given to setPartBytes
  dirty: boolean;             // write from `xml` when parsed, else from `bytes`
  xml?: XmlDocument;          // parsed lazily by OpcPackage.xml(); this cache is the editable tree
  encoding?: XmlEncoding;     // what the XML was read in, so it is written back the same way
}

/** ZIP limits plus the largest XML part that will be parsed. The largest part in the test corpus is 1.7 MB. */
export interface PackageLimits extends ZipLimits { maxXmlPartBytes: number; }
export const DEFAULT_PACKAGE_LIMITS: PackageLimits = { ...DEFAULT_LIMITS, maxXmlPartBytes: 64 * (1 << 20) };

export interface OpenOptions { limits?: Partial<PackageLimits>; }

/**
 * The package's structure at one moment: which parts exist, in what order, with what content types, bytes,
 * parsed trees and dirty flags. Part objects and trees are shared, not copied, so restoring brings back the very
 * objects that were there (editors hold references into them). Opaque outside this module.
 */
export interface PackageCheckpoint {
  readonly parts: readonly [string, Part, Omit<Part, 'name'>][];
  readonly zipNames: readonly [string, string][];
  readonly order: readonly string[];
  readonly contentTypes: ContentTypesCheckpoint;
}

// Formats that are already compressed: storing them saves time on large decks and loses nothing.
const PRECOMPRESSED = new Set([
  'png', 'jpg', 'jpeg', 'jpe', 'jfif', 'gif', 'webp', 'avif', 'heic',
  'mp4', 'm4v', 'mov', 'm4a', 'mp3', 'aac', 'wma', 'wmv', 'webm', 'ogg', 'oga', 'ogv',
  'zip', 'xlsx', 'xlsm', 'docx', 'docm', 'pptx', 'pptm',
]);

/**
 * An Open Packaging Conventions package (ECMA-376 Part 2). Every part keeps the bytes it was read with; a part is
 * re-serialized only when it is dirty, so whatever the editor does not understand or did not touch is written
 * back exactly as it came in.
 */
export class OpcPackage {
  readonly diagnostics: string[] = [];
  private readonly parts = new Map<string, Part>();       // partKey -> part
  private readonly zipNames = new Map<string, string>();  // partKey -> ZIP item name as read, e.g. 'ppt/slides/slide1.xml'
  private order: string[] = [];                           // ZIP item names in original order, the content types stream included
  private readonly contentTypes: ContentTypes;
  private readonly ctZipName: string;
  private readonly limits: PackageLimits;

  private constructor(contentTypes: ContentTypes, ctZipName: string, limits: PackageLimits) {
    this.contentTypes = contentTypes;
    this.ctZipName = ctZipName;
    this.limits = limits;
  }

  /** An empty package: just [Content_Types].xml with the rels and xml defaults. */
  static create(): OpcPackage {
    const pkg = new OpcPackage(ContentTypes.create(), '[Content_Types].xml', DEFAULT_PACKAGE_LIMITS);
    pkg.order.push(pkg.ctZipName);
    return pkg;
  }

  static open(bytes: Uint8Array, options: OpenOptions = {}): OpcPackage {
    const kind = sniff(bytes);
    if (kind === 'cfb') throw compoundFileError(bytes);
    if (kind === 'unknown') throw new OpcError('not-a-package', 'Not a PowerPoint file: the data is neither a ZIP package nor a compound file.');
    const limits = { ...DEFAULT_PACKAGE_LIMITS, ...options.limits };
    const entries = readZip(bytes, limits);

    const ct = entries.find(e => e.name.toLowerCase() === '[content_types].xml');
    if (!ct) throw new OpcError('invalid-package', 'Not an Open XML package: [Content_Types].xml is missing.');
    const pkg = new OpcPackage(ContentTypes.parse(ct.data), ct.name, limits);
    for (const e of entries) {
      if (e === ct) {
        pkg.order.push(e.name);
        continue;
      }
      const why = unsafeEntryName(e.name);
      if (why) throw new OpcError('unsafe-name', `ZIP entry "${e.name}" cannot be a part name: ${why}`);
      const name = normalizeName(e.name);
      const key = partKey(name);
      if (pkg.parts.has(key)) throw new OpcError('invalid-package', `two ZIP entries name the same part: ${pkg.parts.get(key)!.name} and ${name}`);
      const contentType = pkg.contentTypes.lookup(name);
      if (contentType === undefined) pkg.diagnostics.push(`no content type for ${name}`);
      // fflate does not expose entry CRCs; a short entry is the corruption that can still be seen
      if (e.data.length !== e.declaredSize) pkg.diagnostics.push(`ZIP entry ${e.name} holds ${e.data.length} bytes but declares ${e.declaredSize}`);
      pkg.parts.set(key, { name, contentType: contentType ?? '', bytes: e.data, dirty: false });
      pkg.zipNames.set(key, e.name);
      pkg.order.push(e.name);
    }
    pkg.diagnostics.push(...pkg.contentTypes.diagnostics);
    return pkg;
  }

  /** Part names in package order. */
  partNames(): string[] {
    return this.order.filter(n => n !== this.ctZipName).map(n => this.parts.get(partKey(n))!.name);
  }

  getPart(name: string): Part | undefined {
    return this.parts.get(partKey(name));
  }

  private requirePart(name: string): Part {
    const p = this.getPart(name);
    if (!p) throw new OpcError('missing-part', `missing part ${normalizeName(name)}`);
    return p;
  }

  /** Parsed lazily and cached; the cache is the editable tree. Edits must be followed by markDirty(). */
  xml(name: string): XmlDocument {
    const p = this.requirePart(name);
    if (!p.xml) {
      if (p.bytes.length > this.limits.maxXmlPartBytes) {
        throw new OpcError('limit', `${p.name} is ${p.bytes.length} bytes, over the ${this.limits.maxXmlPartBytes}-byte limit for XML parts`);
      }
      try {
        const { text, encoding } = decodeXmlBytes(p.bytes);
        p.xml = parseXml(text);
        p.encoding = encoding;
      } catch (e) {
        throw new OpcError('invalid-part', `${p.name}: ${e instanceof Error ? e.message : String(e)}`, { cause: e });
      }
    }
    return p.xml;
  }

  /** Snapshot the package structure (not the inside of parsed trees: the editor snapshots the elements it edits). */
  checkpoint(): PackageCheckpoint {
    return {
      parts: [...this.parts].map(([key, p]) => [key, p, { contentType: p.contentType, bytes: p.bytes, dirty: p.dirty, xml: p.xml, encoding: p.encoding }]),
      zipNames: [...this.zipNames],
      order: this.order.slice(),
      contentTypes: this.contentTypes.checkpoint(),
    };
  }

  /** Put the package structure back as `checkpoint` found it. */
  restore(cp: PackageCheckpoint): void {
    this.parts.clear();
    for (const [key, part, state] of cp.parts) {
      part.contentType = state.contentType;
      part.bytes = state.bytes;
      part.dirty = state.dirty;
      if (state.xml) part.xml = state.xml;
      else delete part.xml;
      if (state.encoding) part.encoding = state.encoding;
      else delete part.encoding;
      this.parts.set(key, part);
    }
    this.zipNames.clear();
    for (const [k, v] of cp.zipNames) this.zipNames.set(k, v);
    this.order = cp.order.slice();
    this.contentTypes.restore(cp.contentTypes);
  }

  /** Change a part's content type, e.g. a template's main part when it is opened as a presentation. */
  setContentType(name: string, contentType: string): void {
    const p = this.requirePart(name);
    this.contentTypes.assign(p.name, contentType);
    p.contentType = contentType;
  }

  /** Add a Default content type for an extension that has none (PowerPoint gives each media extension one). */
  ensureDefaultContentType(extension: string, contentType: string): void {
    this.contentTypes.ensureDefault(extension, contentType);
  }

  /** Whether a part is dirty: it will be written from its tree rather than from the bytes it was read with. */
  isDirty(name: string): boolean {
    return this.requirePart(name).dirty;
  }

  markDirty(name: string): void {
    this.requirePart(name).dirty = true;
  }

  /** Replace a part's content outright, e.g. a picture's bytes. Drops any parsed tree. */
  setPartBytes(name: string, bytes: Uint8Array): void {
    const p = this.requirePart(name);
    p.bytes = bytes;
    p.xml = undefined;
    p.encoding = undefined;
    p.dirty = true;
  }

  relationships(source: string): Relationship[] {
    const rels = this.getPart(relsNameFor(normalizeName(source)));
    if (!rels) return [];
    return childElements(this.xml(rels.name).root)
      .filter(c => c.local === 'Relationship')
      .map(c => ({
        id: getAttr(c, 'Id') ?? '',
        type: getAttr(c, 'Type') ?? '',
        target: getAttr(c, 'Target') ?? '',
        external: getAttr(c, 'TargetMode') === 'External',
      }));
  }

  /** Follow one relationship from a part (or '/' for the package) to the part it points at. Internal targets only. */
  related(source: string, relId: string): Part {
    const src = normalizeName(source);
    const rel = this.relationships(src).find(r => r.id === relId);
    if (!rel) throw new OpcError('missing-relationship', `${src} has no relationship ${relId}`);
    if (rel.external) throw new OpcError('missing-relationship', `relationship ${relId} of ${src} is external (${rel.target})`);
    const target = this.getPart(resolveTarget(src, rel.target));
    if (!target) throw new OpcError('dangling-relationship', `relationship ${relId} of ${src} points at ${rel.target}, which is not in the package`);
    return target;
  }

  /** Every internal relationship, from any part, whose target is `target`. */
  incomingRelationships(target: string): { source: string; id: string; type: string }[] {
    const key = partKey(target);
    const out: { source: string; id: string; type: string }[] = [];
    for (const p of [...this.parts.values()]) {
      const source = sourceOfRels(p.name);
      if (source === null) continue;
      for (const r of this.relationships(source)) {
        if (!r.external && partKey(resolveTarget(source, r.target)) === key) out.push({ source, id: r.id, type: r.type });
      }
    }
    return out;
  }

  /**
   * Add a relationship and return its new id (rId1, rId2, ... as PowerPoint numbers them). For an internal
   * relationship `target` is the target's part name and must exist; the relative reference is computed here.
   * For an external one it is the URI as written.
   */
  addRelationship(source: string, type: string, target: string, options: { external?: boolean } = {}): string {
    const src = normalizeName(source);
    if (src !== '/') this.requirePart(src);
    if (!options.external) this.requirePart(target);
    const relsName = relsNameFor(src);
    let rels = this.getPart(relsName);
    if (!rels) {
      rels = this.addPart(relsName, RELS_CONTENT_TYPE, new Uint8Array(0));
      rels.xml = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${RELATIONSHIPS_NS}"></Relationships>`);
      rels.encoding = { charset: 'utf-8', bom: false };
    }
    const root = this.xml(rels.name).root;
    const taken = new Set(childElements(root).map(c => getAttr(c, 'Id') ?? ''));
    let next = 1;
    for (const id of taken) {
      const m = /^rId(\d+)$/.exec(id);
      if (m) next = Math.max(next, Number(m[1]) + 1);
    }
    while (taken.has(`rId${next}`)) next++;
    const id = `rId${next}`;
    const attrs: [string, string][] = [['Id', id], ['Type', type], ['Target', options.external ? target : relativeTarget(src, this.requirePart(target).name)]];
    if (options.external) attrs.push(['TargetMode', 'External']);
    root.children.push({
      kind: 'element', prefix: root.prefix, local: 'Relationship', ns: root.ns, children: [], selfClosing: true, parent: root,
      attrs: attrs.map(([local, value]) => ({ prefix: null, local, value, ns: '' })),
    });
    root.selfClosing = false;
    rels.dirty = true;
    return id;
  }

  removeRelationship(source: string, relId: string): void {
    const src = normalizeName(source);
    const rels = this.getPart(relsNameFor(src));
    const root = rels && this.xml(rels.name).root;
    const at = root ? root.children.findIndex(c => c.kind === 'element' && c.local === 'Relationship' && getAttr(c, 'Id') === relId) : -1;
    if (!rels || !root || at < 0) throw new OpcError('missing-relationship', `${src} has no relationship ${relId}`);
    root.children.splice(at, 1);
    rels.dirty = true;
  }

  addPart(name: string, contentType: string, bytes: Uint8Array): Part {
    const partName = normalizeName(name);
    const why = unsafeEntryName(partName.slice(1));
    if (why) throw new OpcError('unsafe-name', `${partName} cannot be a part name: ${why}`);
    const key = partKey(partName);
    if (this.parts.has(key)) throw new OpcError('part-exists', `part exists: ${partName}`);
    this.contentTypes.assign(partName, contentType);
    const part: Part = { name: partName, contentType, bytes, dirty: true };
    this.parts.set(key, part);
    this.zipNames.set(key, partName.slice(1));
    this.order.push(partName.slice(1));
    return part;
  }

  /**
   * Remove a part, its content type override and its own relationships part. Refused while another part still
   * has a relationship to it, unless `force` is set: removing a referenced part is what makes PowerPoint repair.
   */
  removePart(name: string, options: { force?: boolean } = {}): void {
    const part = this.requirePart(name);
    if (!options.force) {
      const refs = this.incomingRelationships(part.name).filter(r => partKey(r.source) !== partKey(part.name));
      if (refs.length) throw new OpcError('referenced-part', `${part.name} is still the target of ${refs.map(r => `${r.source} ${r.id}`).join(', ')}`);
    }
    const rels = this.getPart(relsNameFor(part.name));
    this.dropPart(part.name);
    if (rels) this.dropPart(rels.name);
  }

  private dropPart(name: string): void {
    const key = partKey(name);
    this.parts.delete(key);
    this.zipNames.delete(key);
    this.order = this.order.filter(n => n === this.ctZipName || partKey(n) !== key);
    this.contentTypes.removeOverride(name);
  }

  /** The package as ZIP bytes. Untouched parts, and an untouched [Content_Types].xml, are written from their original bytes. */
  save(): Uint8Array {
    return writeZip(this.entries());
  }

  /**
   * What `save` zips, in order: dirty XML is serialized now, everything else is the bytes as read. An editor hands
   * these to a worker that runs writeZip, so the main thread only pays for serializing what was edited.
   */
  entries(): { name: string; data: Uint8Array; store: boolean }[] {
    return this.order.map(zipName => {
      if (zipName === this.ctZipName) return { name: zipName, data: this.contentTypes.bytes(), store: false };
      const p = this.parts.get(partKey(zipName))!;
      const data = p.dirty && p.xml ? encodeXmlText(serializeXml(p.xml), p.encoding) : p.bytes;
      return { name: zipName, data, store: PRECOMPRESSED.has(extensionOf(p.name)) };
    });
  }
}
