// packages/metafile/src/text.ts  -- fonts and character sets: 8-bit text to Unicode, LOGFONT to a CSS font
import type { FontObject } from './types.ts';

export const SYMBOL_CHARSET = 2;

// DEFAULT_CHARSET text is in the code page of the system that wrote it, which the face name usually gives away.
const CJK_FACES: [RegExp, string][] = [
  [/yahei|simsun|simhei|nsimsun|dengxian|kaiti|fangsong|\u5b8b\u4f53|\u9ed1\u4f53|\u5fae\u8f6f\u96c5\u9ed1/i, 'gbk'],
  [/mingliu|jhenghei|\u7d30\u660e\u9ad4|\u65b0\u7d30\u660e\u9ad4/i, 'big5'],
  [/ms gothic|ms mincho|ms pgothic|ms pmincho|meiryo|yu gothic|yu mincho|ms ui gothic|\uff2d\uff33/i, 'shift_jis'],
  [/malgun|gulim|batang|dotum|gungsuh|\uad74\ub9bc|\ubc14\ud0d5|\ub3cb\uc6c0/i, 'euc-kr'],
];

/** The WHATWG encoding of a Windows character set ([MS-WMF] 2.1.1.5), for 8-bit text. */
export function encodingOf(charSet: number, face = ''): string {
  if (charSet === 1) for (const [re, enc] of CJK_FACES) if (re.test(face)) return enc;
  switch (charSet) {
    case 77: return 'macintosh';
    case 128: return 'shift_jis';
    case 129: case 130: return 'euc-kr';
    case 134: return 'gbk';
    case 136: return 'big5';
    case 161: return 'windows-1253';
    case 162: return 'windows-1254';
    case 163: return 'windows-1258';
    case 177: return 'windows-1255';
    case 178: return 'windows-1256';
    case 186: return 'windows-1257';
    case 204: return 'windows-1251';
    case 222: return 'windows-874';
    case 238: return 'windows-1250';
    default: return 'windows-1252';       // ANSI, DEFAULT, OEM (code page 437 is not a web encoding) and unknown
  }
}

const decoders = new Map<string, TextDecoder>();
function decoder(label: string): TextDecoder {
  let d = decoders.get(label);
  if (!d) {
    try {
      d = new TextDecoder(label);
    } catch {
      d = new TextDecoder('windows-1252');
    }
    decoders.set(label, d);
  }
  return d;
}

const DOUBLE_BYTE = new Set(['shift_jis', 'euc-kr', 'gbk', 'big5']);

/**
 * 8-bit text in a font's character set as Unicode, with the index of the first byte of each character so that
 * per-byte advances can be summed per character. Symbol fonts keep their codes (see `symbolText`).
 */
export function decodeAnsiUnits(bytes: Uint8Array, charSet: number, face = ''): { text: string; starts: number[] } {
  if (charSet === SYMBOL_CHARSET) {
    let text = '';
    for (const b of bytes) text += String.fromCharCode(b);
    return { text, starts: Array.from(bytes, (_, i) => i) };
  }
  const label = encodingOf(charSet, face);
  const d = decoder(label);
  if (!DOUBLE_BYTE.has(label)) {
    const text = d.decode(bytes);
    if (text.length === bytes.length) return { text, starts: Array.from(bytes, (_, i) => i) };
  }
  // one character at a time: a lead byte that does not decode alone takes the next byte with it
  let text = '';
  const starts: number[] = [];
  for (let i = 0; i < bytes.length;) {
    const one = d.decode(bytes.subarray(i, i + 1));
    if (one !== '\ufffd' || i + 1 >= bytes.length) {
      text += one; starts.push(i); i++;
      continue;
    }
    const two = d.decode(bytes.subarray(i, i + 2));
    text += two.length === 1 ? two : one;
    starts.push(i);
    i += two.length === 1 ? 2 : 1;
  }
  return { text, starts };
}

export const decodeAnsi = (bytes: Uint8Array, charSet: number, face = '') => decodeAnsiUnits(bytes, charSet, face).text;

// Adobe Symbol encoding, codes 0x20..0xFF (the Windows Symbol font puts the euro at 0xA0).
const SYMBOL_HIGH =
  ' !\u2200#\u2203%&\u220b()\u2217+,\u2212./0123456789:;<=>?' +
  '\u2245\u0391\u0392\u03a7\u0394\u0395\u03a6\u0393\u0397\u0399\u03d1\u039a\u039b\u039c\u039d\u039f' +
  '\u03a0\u0398\u03a1\u03a3\u03a4\u03a5\u03c2\u03a9\u039e\u03a8\u0396[\u2234]\u22a5_' +
  '\u203e\u03b1\u03b2\u03c7\u03b4\u03b5\u03c6\u03b3\u03b7\u03b9\u03d5\u03ba\u03bb\u03bc\u03bd\u03bf' +
  '\u03c0\u03b8\u03c1\u03c3\u03c4\u03c5\u03d6\u03c9\u03be\u03c8\u03b6{|}\u223c ' +
  '                                ' +
  '\u20ac\u03d2\u2032\u2264\u2044\u221e\u0192\u2663\u2666\u2665\u2660\u2194\u2190\u2191\u2192\u2193' +
  '\u00b0\u00b1\u2033\u2265\u00d7\u221d\u2202\u2022\u00f7\u2260\u2261\u2248\u2026\u23d0\u23af\u21b5' +
  '\u2135\u2111\u211c\u2118\u2297\u2295\u2205\u2229\u222a\u2283\u2287\u2284\u2282\u2286\u2208\u2209' +
  '\u2220\u2207\u00ae\u00a9\u2122\u220f\u221a\u22c5\u00ac\u2227\u2228\u21d4\u21d0\u21d1\u21d2\u21d3' +
  '\u25ca\u2329\u00ae\u00a9\u2122\u2211\u239b\u239c\u239d\u23a1\u23a2\u23a3\u23a7\u23a8\u23a9\u23aa' +
  ' \u232a\u222b\u2320\u23ae\u2321\u239e\u239f\u23a0\u23a4\u23a5\u23a6\u23ab\u23ac\u23ad ';

// The few Webdings and Marlett codes that controls and tables draw (arrows, check marks, bullets), as Unicode.
const ARROWS: Record<number, string> = { 0x33: '\u25c0', 0x34: '\u25b6', 0x35: '\u25b2', 0x36: '\u25bc', 0x61: '\u2714', 0x72: '\u2716' };
const WEBDINGS: Record<number, string> = { ...ARROWS, 0x6e: '\u25cf' };
const MARLETT: Record<number, string> = { ...ARROWS, 0x74: '\u25b2', 0x75: '\u25bc' };

/** Which code table a font's text is in: the Symbol font's, a pictograph font's, or Unicode. */
export function symbolTable(f: FontObject): 'symbol' | 'webdings' | 'marlett' | undefined {
  const face = f.face.toLowerCase();
  if (face === 'symbol' || face === 'symbol mt') return 'symbol';
  if (face === 'webdings') return 'webdings';
  if (face === 'marlett') return 'marlett';
  return undefined;
}

/** Text in a symbol font as Unicode: codes 0x20..0xFF, or the same codes in the private-use range U+F020..U+F0FF. */
export function symbolText(s: string, table: 'symbol' | 'webdings' | 'marlett'): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    let c = s.charCodeAt(i);
    if (c >= 0xf020 && c <= 0xf0ff) c -= 0xf000;
    if (table === 'symbol') out += c >= 0x20 && c <= 0xff ? SYMBOL_HIGH[c - 0x20] ?? String.fromCharCode(c) : String.fromCharCode(c);
    else out += (table === 'webdings' ? WEBDINGS : MARLETT)[c] ?? String.fromCharCode(c);
  }
  return out;
}

// Faces that are rarely installed, with what to try in their place.
const SUBSTITUTES: Record<string, string[]> = {
  'ms sans serif': ['Microsoft Sans Serif', 'Tahoma', 'Arial'],
  'ms shell dlg': ['Microsoft Sans Serif', 'Tahoma', 'Arial'],
  'ms shell dlg 2': ['Tahoma', 'Arial'],
  'helv': ['Arial', 'Helvetica'],
  'helvetica': ['Arial'],
  'ms serif': ['Times New Roman', 'Times'],
  'tms rmn': ['Times New Roman', 'Times'],
  'times': ['Times New Roman'],
  'courier': ['Courier New'],
  'system': ['Microsoft Sans Serif', 'Tahoma', 'Arial'],
  'fixedsys': ['Courier New'],
  'calibri': ['Carlito'],
  'cambria': ['Caladea'],
  'arial': ['Liberation Sans', 'Helvetica'],
  'times new roman': ['Liberation Serif', 'Times'],
  'courier new': ['Liberation Mono', 'Courier'],
  'symbol': ['Times New Roman', 'Times'],
};

/** The CSS generic family for LOGFONT's pitch and family byte. */
function generic(f: FontObject): string {
  switch (f.pitchAndFamily & 0xf0) {
    case 0x10: return 'serif';              // FF_ROMAN
    case 0x30: return 'monospace';          // FF_MODERN
    case 0x40: return 'cursive';            // FF_SCRIPT
    default: return (f.pitchAndFamily & 3) === 1 ? 'monospace' : 'sans-serif';
  }
}

const quote = (name: string) => `"${name.replace(/["\\]/g, '')}"`;

/** A CSS font shorthand for a LOGFONT at `px` pixels: style, weight, size and a family list with fallbacks. */
export function cssFont(f: FontObject, px: number): string {
  const face = f.face.replace(/^@/, '');                                  // vertical-writing names start with @
  const table = symbolTable(f);
  // symbol text is drawn as Unicode, so the symbol font itself (with its own code layout) is not asked for
  const families = table === 'symbol' ? SUBSTITUTES.symbol! : table ? [] : face ? [face, ...(SUBSTITUTES[face.toLowerCase()] ?? [])] : [];
  const family = [...families.map(quote), table === 'symbol' ? 'serif' : generic(f)].join(', ');
  const weight = f.weight <= 0 ? 400 : Math.min(900, Math.max(100, Math.round(f.weight / 100) * 100));
  const size = Math.max(0.01, Math.min(10_000, px));
  return `${f.italic ? 'italic ' : ''}${weight} ${+size.toFixed(3)}px ${family}`;
}
