// packages/xml/src/parse.ts
import { describeChar, FORBIDDEN_CHAR, isXmlChar } from './chars.ts';
import { XML_NS, XMLNS_NS, qname, type XmlAttr, type XmlDocument, type XmlElement, type XmlMisc } from './tree.ts';

export class XmlParseError extends Error {
  readonly offset: number;
  constructor(message: string, offset: number) {
    super(`${message} (offset ${offset})`);
    this.name = 'XmlParseError';
    this.offset = offset;
  }
}

export interface ParseOptions {
  /** Deepest element nesting accepted. PowerPoint's own parts stay within a few dozen levels. */
  maxDepth?: number;
}

// A Map, not an object literal: `&constructor;` must be an undefined entity, not Object.prototype.constructor.
const ENTITY = new Map([['amp', '&'], ['lt', '<'], ['gt', '>'], ['quot', '"'], ['apos', "'"]]);
const REF = /&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|([A-Za-z][A-Za-z0-9]*));|&/g;

// XML 1.0 fifth edition NameStartChar and NameChar without ':', so names are matched as namespace QNames.
const START = 'A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D\\u037F-\\u1FFF\\u200C\\u200D' +
  '\\u2070-\\u218F\\u2C00-\\u2FEF\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD\\u{10000}-\\u{EFFFF}';
const CHAR = START + '\\-.0-9\\u00B7\\u0300-\\u036F\\u203F\\u2040';
const QNAME = new RegExp(`[${START}][${CHAR}]*(?::[${START}][${CHAR}]*)?`, 'uy');
const WS_ONLY = /^[ \t\r\n]*$/;
// XML 1.0 production [23] XMLDecl: version, then optional encoding and standalone, in that order.
const XML_DECL = /^<\?xml[ \t\r\n]+version[ \t\r\n]*=[ \t\r\n]*(["'])1\.[0-9]+\1([ \t\r\n]+encoding[ \t\r\n]*=[ \t\r\n]*(["'])[A-Za-z][A-Za-z0-9._-]*\3)?([ \t\r\n]+standalone[ \t\r\n]*=[ \t\r\n]*(["'])(?:yes|no)\5)?[ \t\r\n]*\?>$/;

type Scope = ReadonlyMap<string, string>;   // prefix ('' for the default namespace) -> namespace URI
const BASE_SCOPE: Scope = new Map([['xml', XML_NS]]);

const isWs = (c: number) => c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d;
const skipWs = (s: string, i: number) => {
  while (i < s.length && isWs(s.charCodeAt(i))) i++;
  return i;
};
/** XML 1.0 section 2.11: a reader sees every CR LF pair and every lone CR as LF. */
const normalizeEol = (s: string) => (s.indexOf('\r') < 0 ? s : s.replace(/\r\n?/g, '\n'));
/** Section 3.3.3 without a DTD: each tab, LF, CR or CR LF pair in an attribute value becomes one space. */
const normalizeAttr = (s: string) => (/[\t\n\r]/.test(s) ? s.replace(/\r\n|[\t\n\r]/g, ' ') : s);

/**
 * Parse one OOXML part. OOXML parts are a narrow subset of XML 1.0: no DOCTYPE, so no external or custom
 * entities. A DOCTYPE is rejected outright, which removes entity expansion and XXE as a class. Processing
 * instructions are inert and are kept (Office writes <?mso-contentType?> into SharePoint customXml parts).
 * Everything else is checked as strictly as PowerPoint's own parser: well-formedness and namespace
 * well-formedness.
 */
export function parseXml(src: string, options: ParseOptions = {}): XmlDocument {
  const maxDepth = options.maxDepth ?? 256;
  const n = src.length;
  let i = 0;

  function fail(message: string, at: number = i): never {
    throw new XmlParseError(message, at);
  }

  function scanQName(): [string | null, string] {
    QNAME.lastIndex = i;
    const m = QNAME.exec(src);
    if (!m) fail(i < n ? `unexpected ${JSON.stringify(src[i])} where a name was expected` : 'unexpected end of input');
    const name = m[0];
    i += name.length;
    const colon = name.indexOf(':');
    return colon < 0 ? [null, name] : [name.slice(0, colon), name.slice(colon + 1)];
  }

  function decodeRefs(raw: string, at: number): string {
    if (raw.indexOf('&') < 0) return raw;
    return raw.replace(REF, (m: string, hex: string | undefined, dec: string | undefined, name: string | undefined, offset: number) => {
      if (m === '&') fail('bare & (write &amp;)', at + offset);
      if (name !== undefined) {
        const value = ENTITY.get(name);
        if (value === undefined) fail(`undefined entity ${m}`, at + offset);   // no DTD, so no custom entities
        return value;
      }
      const cp = hex !== undefined ? parseInt(hex, 16) : parseInt(dec ?? '', 10);
      if (!isXmlChar(cp)) fail(`character reference ${m} is not an XML character`, at + offset);
      return String.fromCodePoint(cp);
    });
  }

  const forbidden = src.search(FORBIDDEN_CHAR);
  if (forbidden >= 0) fail(`character ${describeChar(src, forbidden)} is not allowed in XML`, forbidden);
  if (src.charCodeAt(0) === 0xfeff) i = 1;   // a byte order mark the caller left in the string

  let declaration: string | null = null;
  if (src.startsWith('<?xml', i) && isWs(src.charCodeAt(i + 5))) {
    const end = src.indexOf('?>', i + 5);
    if (end < 0) fail('unterminated XML declaration');
    declaration = src.slice(i, end + 2);
    if (!XML_DECL.test(declaration)) fail('malformed XML declaration');
    i = end + 2;
  }

  const prolog: XmlMisc[] = [];
  const epilog: XmlMisc[] = [];
  let root: XmlElement | null = null;
  const stack: XmlElement[] = [];
  const scopes: Scope[] = [];               // scopes[k] is the namespace scope inside stack[k]

  while (i < n) {
    const lt = src.indexOf('<', i);
    const end = lt < 0 ? n : lt;
    if (end > i) {
      const raw = src.slice(i, end);
      const top = stack[stack.length - 1];
      if (top) {
        const bad = raw.indexOf(']]>');
        if (bad >= 0) fail("']]>' is not allowed in text (write ]]&gt;)", i + bad);
        top.children.push({ kind: 'text', value: decodeRefs(normalizeEol(raw), i) });
      } else if (!WS_ONLY.test(raw)) {
        fail(root ? 'text after the root element' : 'text before the root element');
      } else {
        (root ? epilog : prolog).push({ kind: 'text', value: raw });
      }
      i = end;
    }
    if (lt < 0) break;

    const next = src.charCodeAt(i + 1);
    if (next === 0x21 /* ! */) {
      if (src.startsWith('<!--', i)) {
        const close = src.indexOf('-->', i + 4);
        if (close < 0) fail('unterminated comment');
        const value = src.slice(i + 4, close);
        if (value.indexOf('--') >= 0 || value.endsWith('-')) fail("'--' inside a comment");
        const top = stack[stack.length - 1];
        if (top) top.children.push({ kind: 'comment', value });
        else (root ? epilog : prolog).push({ kind: 'comment', value });
        i = close + 3;
        continue;
      }
      if (src.startsWith('<![CDATA[', i)) {
        const top = stack[stack.length - 1];
        if (!top) fail('CDATA outside the root element');
        const close = src.indexOf(']]>', i + 9);
        if (close < 0) fail('unterminated CDATA section');
        top.children.push({ kind: 'cdata', value: normalizeEol(src.slice(i + 9, close)) });
        i = close + 3;
        continue;
      }
      fail('DOCTYPE and markup declarations are not allowed in OOXML parts');
    }
    if (next === 0x3f /* ? */) {
      // <?target data?>. Inert, so harmless to accept, and Office writes them: SharePoint customXml parts start
      // with <?mso-contentType?>. The target "xml", in any case, is reserved for the declaration at the very start.
      const at = i;
      i += 2;
      const [prefix, target] = scanQName();
      if (prefix !== null) fail('a processing instruction target cannot contain a colon', at);
      if (target.toLowerCase() === 'xml') fail('the XML declaration is only allowed at the very start of a part', at);
      const close = src.indexOf('?>', i);
      if (close < 0) fail('unterminated processing instruction', at);
      if (close > i && !isWs(src.charCodeAt(i))) fail('expected whitespace after a processing instruction target');
      const pi = { kind: 'pi' as const, target, data: normalizeEol(src.slice(skipWs(src, i), close)) };
      const top = stack[stack.length - 1];
      if (top) top.children.push(pi);
      else (root ? epilog : prolog).push(pi);
      i = close + 2;
      continue;
    }

    if (next === 0x2f /* / */) {
      const at = i;
      i += 2;
      const [prefix, local] = scanQName();
      i = skipWs(src, i);
      if (src.charCodeAt(i) !== 0x3e) fail('expected > at the end of a close tag');
      i++;
      const el = stack.pop();
      scopes.pop();
      if (!el) fail('close tag without an open tag', at);
      if (el.local !== local || el.prefix !== prefix) fail(`mismatched close tag </${qname({ prefix, local })}> for <${qname(el)}>`, at);
      continue;
    }

    // open tag
    const tagStart = i;
    i++;
    const [prefix, local] = scanQName();
    const attrs: XmlAttr[] = [];
    let declared: Map<string, string> | null = null;
    let selfClosing = false;
    for (;;) {
      const gap = i;
      i = skipWs(src, i);
      if (i >= n) fail('unterminated open tag', tagStart);
      const c = src.charCodeAt(i);
      if (c === 0x3e /* > */) { i++; break; }
      if (c === 0x2f /* / */) {
        if (src.charCodeAt(i + 1) !== 0x3e) fail('expected />');
        selfClosing = true;
        i += 2;
        break;
      }
      if (i === gap) fail('expected whitespace before an attribute');
      const attrStart = i;
      const [ap, al] = scanQName();
      i = skipWs(src, i);
      if (src.charCodeAt(i) !== 0x3d /* = */) fail('expected = after an attribute name');
      i = skipWs(src, i + 1);
      const q = src[i];
      if (q !== '"' && q !== "'") fail('attribute value must be quoted');
      const vEnd = src.indexOf(q, i + 1);
      if (vEnd < 0) fail('unterminated attribute value');
      const raw = src.slice(i + 1, vEnd);
      const lt2 = raw.indexOf('<');
      if (lt2 >= 0) fail('< in an attribute value', i + 1 + lt2);
      for (const a of attrs) if (a.local === al && a.prefix === ap) fail(`duplicate attribute ${qname({ prefix: ap, local: al })}`, attrStart);
      const value = decodeRefs(normalizeAttr(raw), i + 1);
      if (ap === 'xmlns' || (ap === null && al === 'xmlns')) {
        const p = ap === null ? '' : al;
        if (p === 'xmlns' || (p === 'xml') !== (value === XML_NS) || value === XMLNS_NS) fail(`xmlns:${p} cannot be bound to ${value}`, attrStart);
        if (p !== '' && value === '') fail(`xmlns:${p} cannot be empty`, attrStart);
        (declared ??= new Map()).set(p, value);
      }
      attrs.push({ prefix: ap, local: al, value, ns: '' });
      i = vEnd + 1;
    }

    const parentScope = scopes[scopes.length - 1] ?? BASE_SCOPE;
    let scope = parentScope;
    if (declared) {
      const m = new Map(parentScope);
      for (const [p, uri] of declared) {
        if (uri === '') m.delete(p);       // xmlns="" undeclares the default namespace
        else m.set(p, uri);
      }
      scope = m;
    }
    if (prefix === 'xmlns') fail('xmlns cannot be an element prefix', tagStart);
    const ns = prefix === null ? scope.get('') ?? '' : scope.get(prefix);
    if (ns === undefined) fail(`undeclared namespace prefix ${prefix}`, tagStart);
    for (const a of attrs) {
      if (a.prefix === null) a.ns = a.local === 'xmlns' ? XMLNS_NS : '';
      else if (a.prefix === 'xmlns') a.ns = XMLNS_NS;
      else {
        const uri = scope.get(a.prefix);
        if (uri === undefined) fail(`undeclared namespace prefix ${a.prefix}`, tagStart);
        a.ns = uri;
      }
    }
    for (let x = 0; x < attrs.length; x++) {
      const a = attrs[x]!;
      if (a.prefix === null || a.ns === XMLNS_NS) continue;
      for (let y = x + 1; y < attrs.length; y++) {
        const b = attrs[y]!;
        if (b.ns === a.ns && b.local === a.local) fail(`attributes ${qname(a)} and ${qname(b)} have the same expanded name`, tagStart);
      }
    }

    const parent = stack[stack.length - 1] ?? null;
    if (stack.length >= maxDepth) fail(`elements nested deeper than ${maxDepth} levels`, tagStart);
    const el: XmlElement = { kind: 'element', prefix, local, ns, attrs, children: [], selfClosing, parent };
    if (parent) parent.children.push(el);
    else if (root) fail('more than one root element', tagStart);
    else root = el;
    if (!selfClosing) {
      stack.push(el);
      scopes.push(scope);
    }
  }

  const open = stack[stack.length - 1];
  if (open) fail(`unclosed element <${qname(open)}>`, n);
  if (!root) fail('no root element', n);
  return { declaration, prolog, root, epilog };
}
