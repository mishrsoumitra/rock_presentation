export * from './tree.ts';
export { parseXml, XmlParseError, type ParseOptions } from './parse.ts';
export { serializeXml, serializeElement, XmlSerializeError } from './serialize.ts';
export { canonical, canonicalDiff } from './canonical.ts';
export { decodeXmlBytes, encodeXmlText, XmlEncodingError, type XmlCharset, type XmlEncoding } from './encoding.ts';
export { FORBIDDEN_CHAR, isXmlChar } from './chars.ts';
