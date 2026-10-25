// packages/xml/src/serialize.ts
import { describeChar, FORBIDDEN_CHAR } from './chars.ts';
import { qname, type XmlDocument, type XmlElement, type XmlMisc, type XmlNode, type XmlPI } from './tree.ts';

export class XmlSerializeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'XmlSerializeError';
  }
}

// Escaping follows what PowerPoint writes, so an unedited part serializes byte for byte: `"` and `'` stay raw in
// text, `'` stays raw in attributes, and whitespace in attributes uses hex references (`descr="a&#xA;b"`).
// CR must be escaped in text too: written raw, a reader would turn it into LF (XML 1.0 section 2.11).
const TEXT_ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '\r': '&#xD;' };
const ATTR_ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\t': '&#x9;', '\n': '&#xA;', '\r': '&#xD;' };
const TEXT_ANY = /[&<>\r]/;
const ATTR_ANY = /[&<>"\t\n\r]/;

function checkChars(s: string, where: string): void {
  const bad = s.search(FORBIDDEN_CHAR);
  if (bad >= 0) throw new XmlSerializeError(`character ${describeChar(s, bad)} cannot be written in XML (${where})`);
}
const escText = (s: string) => (TEXT_ANY.test(s) ? s.replace(/[&<>\r]/g, c => TEXT_ESC[c]!) : s);
const escAttr = (s: string) => (ATTR_ANY.test(s) ? s.replace(/[&<>"\t\n\r]/g, c => ATTR_ESC[c]!) : s);

function writeMisc(m: XmlMisc, out: string[]): void {
  if (m.kind === 'text') {
    if (!/^[ \t\r\n]*$/.test(m.value)) throw new XmlSerializeError('only whitespace may appear outside the root element');
    out.push(m.value);
  } else if (m.kind === 'pi') writePI(m, out);
  else writeComment(m.value, out);
}

function writePI(pi: XmlPI, out: string[]): void {
  if (!/^[^\s?:]+$/.test(pi.target) || pi.target.toLowerCase() === 'xml') throw new XmlSerializeError(`invalid processing instruction target "${pi.target}"`);
  if (pi.data.includes('?>')) throw new XmlSerializeError("processing instruction data cannot contain '?>'");
  checkChars(pi.data, 'processing instruction');
  out.push('<?', pi.target, pi.data ? ' ' + pi.data : '', '?>');
}

function writeComment(value: string, out: string[]): void {
  if (value.includes('--') || value.endsWith('-')) throw new XmlSerializeError("a comment cannot contain '--' or end with '-'");
  checkChars(value, 'comment');
  out.push('<!--', value, '-->');
}

function writeNode(c: XmlNode, out: string[]): void {
  if (c.kind === 'element') writeElement(c, out);
  else if (c.kind === 'text') {
    checkChars(c.value, 'text');
    out.push(escText(c.value));
  } else if (c.kind === 'cdata') {
    checkChars(c.value, 'CDATA');
    // ']]>' cannot appear inside one CDATA section; split it across two.
    out.push('<![CDATA[', c.value.replaceAll(']]>', ']]]]><![CDATA[>'), ']]>');
  } else if (c.kind === 'pi') writePI(c, out);
  else writeComment(c.value, out);
}

function writeElement(el: XmlElement, out: string[]): void {
  out.push('<', qname(el));
  for (const a of el.attrs) {
    checkChars(a.value, `attribute ${qname(a)}`);
    out.push(' ', qname(a), '="', escAttr(a.value), '"');
  }
  if (el.children.length === 0 && el.selfClosing) {
    out.push('/>');
    return;
  }
  out.push('>');
  for (const c of el.children) writeNode(c, out);
  out.push('</', qname(el), '>');
}

export function serializeXml(doc: XmlDocument): string {
  const out: string[] = [];
  if (doc.declaration !== null) out.push(doc.declaration);
  for (const m of doc.prolog) writeMisc(m, out);
  writeElement(doc.root, out);
  for (const m of doc.epilog) writeMisc(m, out);
  return out.join('');
}

/** One element and its subtree, without a declaration. Namespace declarations are written only where the tree has them. */
export function serializeElement(el: XmlElement): string {
  const out: string[] = [];
  writeElement(el, out);
  return out.join('');
}
