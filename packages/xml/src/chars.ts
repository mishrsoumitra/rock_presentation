// packages/xml/src/chars.ts

/** Characters XML 1.0 forbids anywhere in a document, raw or escaped: C0 controls other than tab, LF and CR,
 *  U+FFFE, U+FFFF, and unpaired surrogates (with the u flag a well-formed pair is one code point outside the range). */
export const FORBIDDEN_CHAR = /[\0-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF￾￿]/u;

export const isXmlChar = (cp: number) =>
  cp === 0x9 || cp === 0xa || cp === 0xd || (cp >= 0x20 && cp <= 0xd7ff) ||
  (cp >= 0xe000 && cp <= 0xfffd) || (cp >= 0x10000 && cp <= 0x10ffff);

export const describeChar = (s: string, at: number) =>
  'U+' + (s.codePointAt(at) ?? 0).toString(16).toUpperCase().padStart(4, '0');
