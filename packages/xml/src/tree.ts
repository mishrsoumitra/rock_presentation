// packages/xml/src/tree.ts
export const XML_NS = 'http://www.w3.org/XML/1998/namespace';
export const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

export interface XmlDocument {
  declaration: string | null;   // kept verbatim, e.g. <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  prolog: XmlMisc[];            // whitespace and comments between the declaration and the root; PowerPoint writes "\r\n"
  root: XmlElement;
  epilog: XmlMisc[];            // whitespace and comments after the root element
}
export interface XmlAttr {
  prefix: string | null;
  local: string;
  value: string;
  ns: string;                   // resolved namespace URI; '' for an unprefixed attribute
}
export interface XmlElement {
  kind: 'element';
  prefix: string | null;        // 'a', 'p', 'r', 'mc', or null
  local: string;                // 'sp', 'txBody', 'AlternateContent'
  ns: string;                   // resolved namespace URI, so views can match Strict and odd prefixes alike
  attrs: XmlAttr[];             // document order, xmlns declarations included
  children: XmlNode[];
  selfClosing: boolean;         // <a:t/> versus <a:t></a:t>; both are legal, we keep what we read
  parent: XmlElement | null;
}
export interface XmlText    { kind: 'text';    value: string; }
export interface XmlCData   { kind: 'cdata';   value: string; }
export interface XmlComment { kind: 'comment'; value: string; }
/** A processing instruction, e.g. the <?mso-contentType?> Office writes at the top of SharePoint customXml parts. */
export interface XmlPI      { kind: 'pi';      target: string; data: string; }
export type XmlNode = XmlElement | XmlText | XmlCData | XmlComment | XmlPI;
export type XmlMisc = XmlText | XmlComment | XmlPI;   // what may appear outside the root element

export const qname = (n: { prefix: string | null; local: string }) => (n.prefix ? n.prefix + ':' : '') + n.local;

/** Element children only, skipping text, CDATA and comments. */
export const childElements = (el: XmlElement): XmlElement[] =>
  el.children.filter((c): c is XmlElement => c.kind === 'element');

/** Value of the attribute with this local name and namespace ('' means unprefixed). */
export function getAttr(el: XmlElement, local: string, ns = ''): string | undefined {
  for (const a of el.attrs) if (a.local === local && a.ns === ns) return a.value;
  return undefined;
}
