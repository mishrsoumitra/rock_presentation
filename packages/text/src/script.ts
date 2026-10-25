// packages/text/src/script.ts  -- which of a run's three fonts (latin, ea, cs) draws a character
//
// DrawingML gives every run a Latin, an East Asian and a complex-script typeface. Which applies depends on the
// character: CJK and the like take the ea font, right-to-left and Indic scripts the cs font, everything else latin.
// Spaces, digits and punctuation take the class of the text around them.

export type FontClass = 'latin' | 'ea' | 'cs';

const EA: [number, number][] = [
  [0x1100, 0x11ff], [0x2e80, 0x2fdf], [0x2ff0, 0x303f], [0x3040, 0x309f], [0x30a0, 0x30ff], [0x3100, 0x312f], [0x3130, 0x318f],
  [0x3190, 0x31ff], [0x3200, 0x32ff], [0x3300, 0x33ff], [0x3400, 0x4dbf], [0x4e00, 0x9fff], [0xa000, 0xa4cf], [0xa960, 0xa97f],
  [0xac00, 0xd7af], [0xd7b0, 0xd7ff], [0xf900, 0xfaff], [0xfe30, 0xfe4f], [0xff00, 0xffef], [0x1b000, 0x1b16f], [0x20000, 0x3134f],
];
const CS: [number, number][] = [
  [0x0590, 0x05ff], [0x0600, 0x06ff], [0x0700, 0x074f], [0x0750, 0x077f], [0x0780, 0x07bf], [0x07c0, 0x07ff], [0x0800, 0x08ff],
  [0x0900, 0x0dff], [0x0e00, 0x0eff], [0x0f00, 0x0fff], [0x1000, 0x109f], [0x1780, 0x17ff], [0x1800, 0x18af], [0xfb1d, 0xfdff],
  [0xfe70, 0xfeff],
];

const inRanges = (cp: number, ranges: [number, number][]) => ranges.some(([a, b]) => cp >= a && cp <= b);

/** Characters that belong to no script of their own: they follow their neighbours. */
export function isNeutral(cp: number): boolean {
  return cp <= 0x40 || (cp >= 0x5b && cp <= 0x60) || (cp >= 0x7b && cp <= 0xbf) || (cp >= 0x2000 && cp <= 0x206f) || cp === 0xa0 || cp === 0x200b;
}

export function fontClassOf(cp: number): FontClass {
  if (inRanges(cp, EA)) return 'ea';
  if (inRanges(cp, CS)) return 'cs';
  return 'latin';
}

/** One font class per UTF-16 index of `text`; neutrals take the class before them (or after, at the start). */
export function fontClasses(text: string): FontClass[] {
  const out: FontClass[] = new Array(text.length);
  let last: FontClass | undefined;
  const pending: number[] = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    const n = cp > 0xffff ? 2 : 1;
    if (isNeutral(cp)) {
      for (let k = 0; k < n; k++) pending.push(i + k);
    } else {
      const c = fontClassOf(cp);
      for (const p of pending) out[p] = last ?? c;
      pending.length = 0;
      for (let k = 0; k < n; k++) out[i + k] = c;
      last = c;
    }
    i += n;
  }
  for (const p of pending) out[p] = last ?? 'latin';
  return out;
}

/** The ISO 15924 code of a language tag's usual script, as theme font tables key them ("ja-JP" -> "Jpan"). */
export function scriptOfLang(lang: string | undefined): string | undefined {
  if (!lang) return undefined;
  const l = lang.toLowerCase();
  if (l.startsWith('ja')) return 'Jpan';
  if (l === 'zh-tw' || l === 'zh-hk' || l === 'zh-mo' || l.startsWith('zh-hant')) return 'Hant';
  if (l.startsWith('zh')) return 'Hans';
  if (l.startsWith('ko')) return 'Hang';
  const two = l.slice(0, 2);
  const map: Record<string, string> = {
    ar: 'Arab', fa: 'Arab', ur: 'Arab', he: 'Hebr', yi: 'Hebr', th: 'Thai', lo: 'Laoo', km: 'Khmr', my: 'Mymr', hi: 'Deva', mr: 'Deva',
    ne: 'Deva', sa: 'Deva', bn: 'Beng', as: 'Beng', pa: 'Guru', gu: 'Gujr', or: 'Orya', ta: 'Taml', te: 'Telu', kn: 'Knda', ml: 'Mlym',
    si: 'Sinh', bo: 'Tibt', dv: 'Thaa', am: 'Ethi', ka: 'Geor', hy: 'Armn', el: 'Grek', ru: 'Cyrl', uk: 'Cyrl', bg: 'Cyrl',
  };
  return map[two];
}

/** The script of a character as HarfBuzz wants it (ISO 15924), for buffers that hold one script. */
export function hbScriptOf(cp: number): string | undefined {
  if (cp >= 0x0590 && cp <= 0x05ff) return 'Hebr';
  if ((cp >= 0x0600 && cp <= 0x06ff) || (cp >= 0x0750 && cp <= 0x077f) || (cp >= 0xfb50 && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff)) return 'Arab';
  if (cp >= 0x0900 && cp <= 0x097f) return 'Deva';
  if (cp >= 0x0980 && cp <= 0x09ff) return 'Beng';
  if (cp >= 0x0b80 && cp <= 0x0bff) return 'Taml';
  if (cp >= 0x0e00 && cp <= 0x0e7f) return 'Thai';
  if ((cp >= 0x3040 && cp <= 0x309f)) return 'Hira';
  if ((cp >= 0x30a0 && cp <= 0x30ff)) return 'Kana';
  if ((cp >= 0xac00 && cp <= 0xd7af) || (cp >= 0x1100 && cp <= 0x11ff)) return 'Hang';
  if ((cp >= 0x4e00 && cp <= 0x9fff) || (cp >= 0x3400 && cp <= 0x4dbf)) return 'Hani';
  if (cp >= 0x0400 && cp <= 0x04ff) return 'Cyrl';
  if (cp >= 0x0370 && cp <= 0x03ff) return 'Grek';
  return undefined;
}
