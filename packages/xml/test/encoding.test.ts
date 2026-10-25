import { describe, expect, test } from 'vitest';
import { decodeXmlBytes, encodeXmlText, XmlEncodingError, type XmlEncoding } from '../src/index.ts';

const TEXT = '<?xml version="1.0"?><r>Grüße 😀</r>';
const utf16 = (s: string, le: boolean) => {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const u = s.charCodeAt(i);
    out[2 * i] = le ? u & 0xff : u >>> 8;
    out[2 * i + 1] = le ? u >>> 8 : u & 0xff;
  }
  return out;
};
const cat = (...parts: Uint8Array[]) => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
};

describe('decodeXmlBytes and encodeXmlText', () => {
  const cases: [string, Uint8Array, XmlEncoding][] = [
    ['UTF-8', new TextEncoder().encode(TEXT), { charset: 'utf-8', bom: false }],
    ['UTF-8 with BOM', cat(new Uint8Array([0xef, 0xbb, 0xbf]), new TextEncoder().encode(TEXT)), { charset: 'utf-8', bom: true }],
    ['UTF-16LE with BOM', cat(new Uint8Array([0xff, 0xfe]), utf16(TEXT, true)), { charset: 'utf-16le', bom: true }],
    ['UTF-16BE with BOM', cat(new Uint8Array([0xfe, 0xff]), utf16(TEXT, false)), { charset: 'utf-16be', bom: true }],
    ['UTF-16LE without BOM', utf16(TEXT, true), { charset: 'utf-16le', bom: false }],
    ['UTF-16BE without BOM', utf16(TEXT, false), { charset: 'utf-16be', bom: false }],
  ];

  test.each(cases)('%s is detected, decoded and re-encoded to the same bytes', (_name, bytes, encoding) => {
    const decoded = decodeXmlBytes(bytes);
    expect(decoded.text).toBe(TEXT);
    expect(decoded.encoding).toEqual(encoding);
    expect(encodeXmlText(decoded.text, decoded.encoding)).toEqual(bytes);
  });

  test('invalid UTF-8 is an error, not replacement characters', () => {
    expect(() => decodeXmlBytes(new Uint8Array([0x3c, 0x72, 0x3e, 0xff, 0x3c]))).toThrow(XmlEncodingError);
  });

  test('a lone surrogate in UTF-16 is an error', () => {
    expect(() => decodeXmlBytes(cat(new Uint8Array([0xff, 0xfe]), utf16('<r>\uD800</r>', true)))).toThrow(XmlEncodingError);
  });
});
