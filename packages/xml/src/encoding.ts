// packages/xml/src/encoding.ts
export type XmlCharset = 'utf-8' | 'utf-16le' | 'utf-16be';
export interface XmlEncoding { charset: XmlCharset; bom: boolean; }

export class XmlEncodingError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'XmlEncodingError';
  }
}

/**
 * Detect and decode the encodings OOXML allows for XML parts (ECMA-376 Part 2: UTF-8 or UTF-16), the way
 * XML 1.0 Appendix F does: by byte order mark, else by how '<?' is encoded, else UTF-8. PowerPoint writes
 * UTF-8 without a BOM; other producers occasionally write a BOM or UTF-16. Invalid sequences throw.
 */
export function decodeXmlBytes(bytes: Uint8Array): { text: string; encoding: XmlEncoding } {
  const [b0, b1, b2, b3] = bytes;
  let charset: XmlCharset = 'utf-8';
  let bom = false;
  let skip = 0;
  if (b0 === 0xef && b1 === 0xbb && b2 === 0xbf) { bom = true; skip = 3; }
  else if (b0 === 0xff && b1 === 0xfe) { charset = 'utf-16le'; bom = true; skip = 2; }
  else if (b0 === 0xfe && b1 === 0xff) { charset = 'utf-16be'; bom = true; skip = 2; }
  else if (b0 === 0x3c && b1 === 0x00 && b2 === 0x3f && b3 === 0x00) charset = 'utf-16le';
  else if (b0 === 0x00 && b1 === 0x3c && b2 === 0x00 && b3 === 0x3f) charset = 'utf-16be';
  let text: string;
  try {
    text = new TextDecoder(charset, { fatal: true, ignoreBOM: true }).decode(bytes.subarray(skip));
  } catch {
    throw new XmlEncodingError(`not valid ${charset.toUpperCase()}`);
  }
  return { text, encoding: { charset, bom } };
}

/** Encode serialized XML back into the encoding it was read in. */
export function encodeXmlText(text: string, encoding: XmlEncoding = { charset: 'utf-8', bom: false }): Uint8Array {
  if (encoding.charset === 'utf-8') {
    const body = new TextEncoder().encode(text);
    if (!encoding.bom) return body;
    const out = new Uint8Array(body.length + 3);
    out.set([0xef, 0xbb, 0xbf]);
    out.set(body, 3);
    return out;
  }
  const le = encoding.charset === 'utf-16le';
  const at = encoding.bom ? 2 : 0;
  const out = new Uint8Array(at + text.length * 2);
  if (encoding.bom) out.set(le ? [0xff, 0xfe] : [0xfe, 0xff]);
  for (let k = 0; k < text.length; k++) {
    const u = text.charCodeAt(k);
    out[at + 2 * k] = le ? u & 0xff : u >>> 8;
    out[at + 2 * k + 1] = le ? u >>> 8 : u & 0xff;
  }
  return out;
}
