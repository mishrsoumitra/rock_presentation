// packages/opc/src/content-types.ts  -- [Content_Types].xml (ECMA-376 Part 2, section 10.1.2)
import {
  childElements, decodeXmlBytes, encodeXmlText, getAttr, parseXml, serializeXml,
  type XmlDocument, type XmlElement, type XmlEncoding,
} from '@pptx/xml';
import { OpcError } from './errors.ts';
import { asciiLower, extensionOf, partKey } from './names.ts';

export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';

export interface ContentTypesCheckpoint {
  readonly children: XmlElement['children'];
  readonly selfClosing: boolean;
  readonly dirty: boolean;
  readonly defaults: ReadonlyMap<string, string>;
  readonly overrides: ReadonlyMap<string, string>;
}

/**
 * The content types stream, kept as the document it was read from. Lookups go through two maps; edits change the
 * document in place (append or remove an Override, add a Default) so everything else in it survives, and an
 * unedited stream is written back from its original bytes.
 */
export class ContentTypes {
  private readonly doc: XmlDocument;
  private readonly encoding: XmlEncoding;
  private readonly original: Uint8Array | null;
  private dirty: boolean;
  private readonly defaults = new Map<string, string>();    // extension, ASCII lower case -> content type
  private readonly overrides = new Map<string, string>();   // partKey -> content type
  readonly diagnostics: string[] = [];

  private constructor(doc: XmlDocument, encoding: XmlEncoding, original: Uint8Array | null) {
    this.doc = doc;
    this.encoding = encoding;
    this.original = original;
    this.dirty = original === null;
    if (doc.root.local !== 'Types') throw new OpcError('invalid-package', `[Content_Types].xml has root <${doc.root.local}>, expected <Types>`);
    if (doc.root.ns !== CONTENT_TYPES_NS) this.diagnostics.push(`[Content_Types].xml uses namespace "${doc.root.ns}"`);
    for (const c of childElements(doc.root)) {
      if (c.local === 'Default') {
        const ext = asciiLower(getAttr(c, 'Extension') ?? '');
        if (this.defaults.has(ext)) this.diagnostics.push(`[Content_Types].xml repeats the Default for "${ext}"; the first one is used`);
        else this.defaults.set(ext, getAttr(c, 'ContentType') ?? '');
      } else if (c.local === 'Override') {
        const key = partKey(getAttr(c, 'PartName') ?? '');
        if (this.overrides.has(key)) this.diagnostics.push(`[Content_Types].xml repeats the Override for ${getAttr(c, 'PartName')}; the first one is used`);
        else this.overrides.set(key, getAttr(c, 'ContentType') ?? '');
      }
    }
  }

  static parse(bytes: Uint8Array): ContentTypes {
    let decoded: { text: string; encoding: XmlEncoding };
    let doc: XmlDocument;
    try {
      decoded = decodeXmlBytes(bytes);
      doc = parseXml(decoded.text);
    } catch (e) {
      throw new OpcError('invalid-package', `[Content_Types].xml is not well-formed: ${e instanceof Error ? e.message : String(e)}`, { cause: e });
    }
    return new ContentTypes(doc, decoded.encoding, bytes);
  }

  /** The stream of a new package, with the two defaults every package needs. */
  static create(): ContentTypes {
    const doc = parseXml(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      `<Types xmlns="${CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="${RELS_CONTENT_TYPE}"/>` +
      '<Default Extension="xml" ContentType="application/xml"/></Types>');
    return new ContentTypes(doc, { charset: 'utf-8', bom: false }, null);
  }

  /** Override by part name first, then Default by extension; undefined when neither applies. */
  lookup(partName: string): string | undefined {
    return this.overrides.get(partKey(partName)) ?? this.defaults.get(extensionOf(partName));
  }

  /** Make `partName` resolve to `contentType`, adding an Override only when the extension's Default says otherwise. */
  assign(partName: string, contentType: string): void {
    const key = partKey(partName);
    const fromDefault = this.defaults.get(extensionOf(partName));
    if (fromDefault === contentType) {
      if (this.overrides.has(key)) this.removeOverride(partName);
      return;
    }
    if (this.overrides.get(key) === contentType) return;
    this.removeOverride(partName);
    this.overrides.set(key, contentType);
    this.append(this.element('Override', { PartName: partName, ContentType: contentType }), false);
  }

  /** Add a Default for an extension that has none. An existing Default is never changed: other parts rely on it. */
  ensureDefault(extension: string, contentType: string): void {
    const ext = asciiLower(extension);
    if (this.defaults.has(ext)) return;
    this.defaults.set(ext, contentType);
    this.append(this.element('Default', { Extension: ext, ContentType: contentType }), true);
  }

  removeOverride(partName: string): void {
    const key = partKey(partName);
    if (!this.overrides.delete(key)) return;
    const root = this.doc.root;
    root.children = root.children.filter(c => !(c.kind === 'element' && c.local === 'Override' && partKey(getAttr(c, 'PartName') ?? '') === key));
    this.dirty = true;
  }

  /** The stream's whole state, for undo: what `restore` puts back. Elements are shared, never copied. */
  checkpoint(): ContentTypesCheckpoint {
    const root = this.doc.root;
    return {
      children: root.children.slice(), selfClosing: root.selfClosing, dirty: this.dirty,
      defaults: new Map(this.defaults), overrides: new Map(this.overrides),
    };
  }

  restore(cp: ContentTypesCheckpoint): void {
    const root = this.doc.root;
    root.children = cp.children.slice();
    root.selfClosing = cp.selfClosing;
    this.dirty = cp.dirty;
    this.defaults.clear();
    for (const [k, v] of cp.defaults) this.defaults.set(k, v);
    this.overrides.clear();
    for (const [k, v] of cp.overrides) this.overrides.set(k, v);
  }

  /** The stream to write: the original bytes when nothing changed. */
  bytes(): Uint8Array {
    if (!this.dirty && this.original) return this.original;
    return encodeXmlText(serializeXml(this.doc), this.encoding);
  }

  private element(local: string, attrs: Record<string, string>): XmlElement {
    const root = this.doc.root;
    return {
      kind: 'element', prefix: root.prefix, local, ns: root.ns, children: [], selfClosing: true, parent: root,
      attrs: Object.entries(attrs).map(([k, v]) => ({ prefix: null, local: k, value: v, ns: '' })),
    };
  }

  /** PowerPoint lists every Default before the first Override; keep that grouping. */
  private append(el: XmlElement, isDefault: boolean): void {
    const root = this.doc.root;
    root.selfClosing = false;
    let at = root.children.length;
    if (isDefault) {
      const firstOverride = root.children.findIndex(c => c.kind === 'element' && c.local === 'Override');
      if (firstOverride >= 0) at = firstOverride;
    }
    root.children.splice(at, 0, el);
    this.dirty = true;
  }
}
