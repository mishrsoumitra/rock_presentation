// packages/charts/src/numfmt.ts  -- number format codes for tick and data labels (ECMA-376 Part 1, 18.8.30)
//
// Covers what charts use: General, digit placeholders (0 # ?), grouping and scaling commas, percent, scientific,
// simple fractions, literals and currency symbols ([$€-407]), up to four sections with optional conditions, and
// dates and times. Colours ([Red]), fill characters (*x) and locale-specific names are ignored: en-US output.

type Token =
  | { t: 'lit'; s: string }
  | { t: 'digit'; c: '0' | '#' | '?' }
  | { t: 'point' }
  | { t: 'comma' }
  | { t: 'percent' }
  | { t: 'exp'; sign: boolean }
  | { t: 'slash' }
  | { t: 'general' }
  | { t: 'text' }
  | { t: 'date'; s: string }
  | { t: 'ampm'; s: string }
  | { t: 'elapsed'; s: string };

interface Section { tokens: Token[]; cond?: { op: string; v: number }; isDate: boolean; }

const DATE_CHARS = /[ymdhs]/i;

function tokenize(code: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const lit = (s: string) => {
    const last = out.at(-1);
    if (last?.t === 'lit') last.s += s;
    else out.push({ t: 'lit', s });
  };
  while (i < code.length) {
    const ch = code[i]!;
    const rest = code.slice(i);
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      lit(code.slice(i + 1, end < 0 ? code.length : end));
      i = end < 0 ? code.length : end + 1;
    } else if (ch === '\\') {
      lit(code[i + 1] ?? '');
      i += 2;
    } else if (ch === '_') {
      lit(' ');                                        // a space as wide as the next character
      i += 2;
    } else if (ch === '*') {
      i += 2;                                          // fill character: nothing to fill in a label
    } else if (ch === '[') {
      const end = code.indexOf(']', i);
      const inner = code.slice(i + 1, end < 0 ? code.length : end);
      i = end < 0 ? code.length : end + 1;
      if (inner.startsWith('$')) {
        // [$€-407]: a currency symbol and a locale
        const sym = inner.slice(1).split('-')[0] ?? '';
        if (sym) lit(sym);
      } else if (/^(h+|m+|s+)$/i.test(inner)) {
        out.push({ t: 'elapsed', s: inner.toLowerCase() });
      }
      // colours, conditions (read separately) and [DBNum] are dropped
    } else if (/^general/i.test(rest)) {
      out.push({ t: 'general' });
      i += 7;
    } else if (/^(am\/pm|a\/p)/i.test(rest)) {
      const s = /^am\/pm/i.test(rest) ? rest.slice(0, 5) : rest.slice(0, 3);
      out.push({ t: 'ampm', s });
      i += s.length;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      out.push({ t: 'digit', c: ch });
      i++;
    } else if (ch === '.') {
      out.push({ t: 'point' });
      i++;
    } else if (ch === ',') {
      out.push({ t: 'comma' });
      i++;
    } else if (ch === '%') {
      out.push({ t: 'percent' });
      i++;
    } else if ((ch === 'E' || ch === 'e') && (code[i + 1] === '+' || code[i + 1] === '-')) {
      out.push({ t: 'exp', sign: code[i + 1] === '+' });
      i += 2;
    } else if (ch === '/') {
      out.push({ t: 'slash' });
      i++;
    } else if (ch === '@') {
      out.push({ t: 'text' });
      i++;
    } else if (DATE_CHARS.test(ch)) {
      let j = i;
      while (j < code.length && code[j]!.toLowerCase() === ch.toLowerCase()) j++;
      out.push({ t: 'date', s: code.slice(i, j).toLowerCase() });
      i = j;
    } else {
      lit(ch);
      i++;
    }
  }
  return out;
}

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (ch === '\\' && !quoted) {
      cur += ch + (code[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"') quoted = !quoted;
    else if (ch === '[' && !quoted) bracket = true;
    else if (ch === ']' && !quoted) bracket = false;
    if (ch === ';' && !quoted && !bracket) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

const cache = new Map<string, Section[]>();

function sectionsOf(code: string): Section[] {
  let s = cache.get(code);
  if (s) return s;
  s = splitSections(code).slice(0, 4).map(text => {
    const m = /\[(<=|>=|<>|<|>|=)(-?[0-9.]+)\]/.exec(text);
    const tokens = tokenize(text);
    const isDate = tokens.some(t => t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm');
    return { tokens, isDate, ...(m ? { cond: { op: m[1]!, v: Number(m[2]) } } : {}) };
  });
  if (cache.size > 500) cache.clear();
  cache.set(code, s);
  return s;
}

const COMPARE: Record<string, (a: number, b: number) => boolean> = {
  '<': (a, b) => a < b, '<=': (a, b) => a <= b, '>': (a, b) => a > b, '>=': (a, b) => a >= b, '=': (a, b) => a === b,
  '<>': (a, b) => a !== b,
};
/** Whether a value meets a section's condition ([<0], [>=100]...). */
const test = (op: string, a: number, b: number) => COMPARE[op]?.(a, b) ?? false;

/** Round half away from zero at `d` decimals, the way spreadsheets do, avoiding binary artefacts (1.005 -> 1.01). */
export function roundDecimal(v: number, d: number): number {
  if (!Number.isFinite(v)) return v;
  const p = Number(Math.abs(v).toPrecision(15));
  const r = Number(`${Math.round(Number(`${p}e${d}`))}e-${d}`);
  return v < 0 ? -r : r;
}

/** The General format: up to 11 characters, switching to scientific notation for very large or small numbers. */
export function formatGeneral(v: number): string {
  if (v === 0 || !Number.isFinite(v)) return Number.isFinite(v) ? '0' : String(v);
  const a = Math.abs(v);
  const sign = v < 0 ? '-' : '';
  if (a >= 1e11 || a < 1e-9) {
    for (let p = 6; p >= 1; p--) {
      const [m = '0', e = '0'] = a.toExponential(p - 1).split('e');
      const mant = m.includes('.') ? m.replace(/\.?0+$/, '') : m;
      const exp = Number(e);
      const s = `${mant}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
      if (s.length <= 11 - sign.length || p === 1) return sign + s;
    }
  }
  for (let p = 10; p >= 1; p--) {
    let s = Number(a.toPrecision(p)).toString();
    if (s.includes('e')) s = a.toFixed(Math.min(20, Math.max(0, p - Math.floor(Math.log10(a)) - 1)));
    if (s.includes('.')) s = s.replace(/\.?0+$/, '');
    if (s.length <= 11 - sign.length) return sign + s;
  }
  return sign + String(Math.round(a));
}

function groupThousands(int: string): string {
  let out = '';
  for (let i = 0; i < int.length; i++) {
    if (i > 0 && (int.length - i) % 3 === 0) out += ',';
    out += int[i];
  }
  return out;
}

/** Digits for one run of placeholders: '0' pads with zeros, '?' with spaces, '#' with nothing. */
function padInt(digits: string, ph: string[]): string {
  const need = ph.filter(c => c !== '#').length;
  let s = digits === '0' && !ph.includes('0') ? '' : digits;
  if (s.length < need) {
    const pad = ph.slice(0, ph.length - s.length).filter(c => c !== '#').map(c => (c === '0' ? '0' : ' ')).join('');
    s = pad.slice(Math.max(0, pad.length - (need - s.length))) + s;
  }
  return s;
}

/** Decimals for the placeholders after the point: trailing zeros go for '#' and become spaces for '?'. */
function fracDigits(digits: string, ph: string[]): string {
  const out = digits.split('');
  for (let i = ph.length - 1; i >= 0; i--) {
    if (out[i] !== '0' || ph[i] === '0') break;
    out[i] = ph[i] === '?' ? ' ' : '';
  }
  return out.join('');
}

function formatNumberSection(v: number, tokens: Token[]): string {
  const percent = tokens.filter(t => t.t === 'percent').length;
  let x = v * 100 ** percent;
  const expAt = tokens.findIndex(t => t.t === 'exp');
  const slashAt = tokens.findIndex(t => t.t === 'slash');
  const pointAt = tokens.findIndex(t => t.t === 'point');
  const numEnd = expAt >= 0 ? expAt : tokens.length;
  // scaling commas: commas right after the last integer placeholder divide by 1000 each
  let lastDigit = -1;
  for (let i = 0; i < (pointAt >= 0 ? pointAt : numEnd); i++) if (tokens[i]!.t === 'digit') lastDigit = i;
  let scale = 0;
  for (let i = lastDigit + 1; i < tokens.length && tokens[i]!.t === 'comma'; i++) scale++;
  x /= 1000 ** scale;
  const intPh: string[] = [];
  const fracPh: string[] = [];
  let grouping = false;
  tokens.forEach((t, i) => {
    if (t.t === 'digit' && i < numEnd) (pointAt >= 0 && i > pointAt ? fracPh : intPh).push(t.c);
    if (t.t === 'comma' && i < lastDigit) grouping = true;
  });

  if (slashAt >= 0) return formatFraction(x, tokens, slashAt);

  let mant = Math.abs(x);
  let exp = 0;
  if (expAt >= 0 && mant !== 0) {
    // engineering formats (##0.0E+0) keep the exponent a multiple of their integer placeholder count
    const step = intPh.length > 1 && intPh.includes('#') ? intPh.length : 1;
    exp = Math.floor(Math.log10(mant) / step) * step;
    mant /= 10 ** exp;
    if (roundDecimal(mant, fracPh.length) >= 10 ** step) {
      mant /= 10 ** step;
      exp += step;
    }
  }
  if (mant >= 1e21) return formatGeneral(x);
  if (fracPh.length > 30) fracPh.length = 30;
  const rounded = roundDecimal(mant, fracPh.length);
  const fixed = rounded.toFixed(fracPh.length);
  const [ip = '0', fp = ''] = fixed.split('.');
  let intStr = padInt(ip, intPh.length ? intPh : ['#']);
  if (grouping) intStr = groupThousands(intStr.trim()).padStart(intStr.length, ' ');
  const fracStr = fracDigits(fp, fracPh);

  let out = '';
  let intDone = false;
  let fracDone = false;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (i >= numEnd && expAt >= 0) {
      if (i === expAt) {
        const e = tokens.slice(expAt + 1).filter(k => k.t === 'digit').length;
        out += `E${exp < 0 ? '-' : (t.t === 'exp' && t.sign ? '+' : '')}${String(Math.abs(exp)).padStart(e, '0')}`;
      } else if (t.t === 'lit') out += t.s;
      continue;
    }
    switch (t.t) {
      case 'lit': out += t.s; break;
      case 'digit':
        if (pointAt >= 0 && i > pointAt) {
          if (!fracDone) { out += fracStr; fracDone = true; }
        } else if (!intDone) { out += intStr; intDone = true; }
        break;
      case 'point':
        if (!intDone) { out += intStr; intDone = true; }
        out += '.';
        break;
      case 'percent': out += '%'; break;
      case 'general': out += formatGeneral(Math.abs(x)); break;
      default: break;
    }
  }
  return out;
}

/** Fractions: "# ?/?", "?/8"; the whole part is printed when the format has placeholders before the fraction. */
function formatFraction(x: number, tokens: Token[], slashAt: number): string {
  const before = tokens.slice(0, slashAt);
  const after = tokens.slice(slashAt + 1);
  const denDigits = after.filter(t => t.t === 'digit').length;
  const fixedDen = Number(after.filter(t => t.t === 'lit').map(t => (t.t === 'lit' ? t.s : '')).join('').trim());
  // a space literal splits the whole number from the numerator: "# ?/?"
  const spaceAt = before.findIndex(t => t.t === 'lit' && t.s.includes(' '));
  const hasWhole = spaceAt >= 0 && before.slice(0, spaceAt).some(t => t.t === 'digit');
  const a = Math.abs(x);
  const whole = hasWhole ? Math.floor(a) : 0;
  const f = a - whole;
  let num = 0;
  let den = 1;
  if (Number.isFinite(fixedDen) && fixedDen > 0) {
    den = fixedDen;
    num = Math.round(f * den);
  } else {
    const maxDen = 10 ** Math.max(1, Math.min(3, denDigits)) - 1;
    let best = Infinity;
    for (let d = 1; d <= maxDen; d++) {
      const n = Math.round(f * d);
      const err = Math.abs(f - n / d);
      if (err < best - 1e-12) { best = err; num = n; den = d; }
      if (err === 0) break;
    }
  }
  if (num === den) return `${whole + 1}`;
  const w = hasWhole && whole ? `${whole}` : '';
  if (num === 0) return w || '0';
  return `${w}${w ? ' ' : ''}${num}/${den}`;
}

// ---------------------------------------------------------------------------------------------------------- dates

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** A date and time: year, month 1-12, day, weekday 0-6 from Sunday, hours, minutes, seconds and their fraction. */
export interface DateParts {
  y: number; m: number; d: number; wd: number;
  h: number; mi: number; s: number; ms: number;
}

/** A spreadsheet date serial as calendar parts (1900 system with its phantom 29 Feb 1900, or the 1904 system). */
export function serialToDate(serial: number, date1904 = false): DateParts {
  let days = Math.floor(serial);
  const secs = Math.round((serial - days) * 86400 * 1000) / 1000;
  let t: number;
  if (date1904) t = Date.UTC(1904, 0, 1) + days * 86400000;
  else if (days === 60) return { y: 1900, m: 2, d: 29, wd: 3, ...timeParts(secs) };
  else {
    if (days > 60) days -= 1;
    t = Date.UTC(1899, 11, 31) + days * 86400000;
  }
  const dt = new Date(t);
  const date = { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay() };
  return { ...date, ...timeParts(secs) };
}

/** The serial of a calendar date (the inverse of serialToDate for whole days; month may run past 12). */
export function dateToSerial(y: number, m: number, d: number, date1904 = false): number {
  const t = Date.UTC(y, m - 1, d);
  if (date1904) return Math.round((t - Date.UTC(1904, 0, 1)) / 86400000);
  const days = Math.round((t - Date.UTC(1899, 11, 31)) / 86400000);
  return days >= 60 ? days + 1 : days;
}

function timeParts(secs: number): { h: number; mi: number; s: number; ms: number } {
  const whole = Math.floor(secs);
  return { h: Math.floor(whole / 3600) % 24, mi: Math.floor(whole / 60) % 60, s: whole % 60, ms: secs - whole };
}

const timeTok = (k: Token): k is Extract<Token, { t: 'date' | 'elapsed' }> => k.t === 'date' || k.t === 'elapsed';
const pad2 = (n: number, wide: boolean) => (wide ? String(n).padStart(2, '0') : String(n));

function formatDateSection(v: number, tokens: Token[], date1904: boolean): string {
  // round to the precision shown, whole seconds unless the format has fractional seconds: a point right after
  // seconds, followed by zeros (ss.00)
  const secondsAt = (i: number) => {
    const prev = tokens[i - 1];
    return tokens[i]?.t === 'point' && prev?.t === 'date' && prev.s[0] === 's' && tokens[i + 1]?.t === 'digit';
  };
  const fracSec = tokens.some((_, i) => secondsAt(i));
  const serial = fracSec ? v : Math.round(v * 86400) / 86400;
  const p = serialToDate(serial, date1904);
  const ampm = tokens.some(t => t.t === 'ampm');
  let out = '';
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.t === 'lit') { out += t.s; continue; }
    if (t.t === 'ampm') {
      const pm = p.h >= 12;
      const s = t.s.toLowerCase() === 'a/p' ? (pm ? 'P' : 'A') : pm ? 'PM' : 'AM';
      out += t.s[0] === t.s[0]!.toLowerCase() ? s.toLowerCase() : s;
      continue;
    }
    if (t.t === 'elapsed') {
      const hours = v * 24;
      const n = Math.floor(t.s[0] === 'h' ? hours : t.s[0] === 'm' ? hours * 60 : hours * 3600);
      out += String(n).padStart(t.s.length, '0');
      continue;
    }
    if (secondsAt(i)) {
      let n = 0;
      while (tokens[i + 1 + n]?.t === 'digit') n++;
      out += `.${String(Math.floor(p.ms * 10 ** n)).padStart(n, '0')}`;
      i += n;
      continue;
    }
    if (t.t === 'digit') { out += t.c === '0' ? '0' : ''; continue; }
    // in a date, separators are just characters: m/d/yy, d.m.yyyy, mmm d, yyyy
    if (t.t === 'slash') { out += '/'; continue; }
    if (t.t === 'point') { out += '.'; continue; }
    if (t.t === 'comma') { out += ','; continue; }
    if (t.t === 'percent') { out += '%'; continue; }
    if (t.t !== 'date') continue;
    const s = t.s;
    const month = MONTHS[p.m - 1]!;
    const day = DAYS[p.wd]!;
    switch (s[0]) {
      case 'y': out += s.length <= 2 ? pad2(p.y % 100, true) : String(p.y); break;
      case 'd': out += s.length <= 2 ? pad2(p.d, s.length === 2) : s.length === 3 ? day.slice(0, 3) : day; break;
      case 'h': out += pad2(ampm ? ((p.h + 11) % 12) + 1 : p.h, s.length >= 2); break;
      case 's': out += pad2(p.s, s.length >= 2); break;
      case 'm': {
        // minutes when right after an hour or right before a second, months otherwise
        const prev = tokens.slice(0, i).reverse().find(timeTok);
        const next = tokens.slice(i + 1).find(timeTok);
        if (s.length <= 2 && (prev?.s[0] === 'h' || next?.s[0] === 's')) out += pad2(p.mi, s.length === 2);
        else if (s.length <= 2) out += pad2(p.m, s.length === 2);
        else out += s.length === 3 ? month.slice(0, 3) : s.length === 5 ? month[0]! : month;
        break;
      }
      default: break;
    }
  }
  return out;
}

/** True when the format code shows dates or times (axes then treat their numbers as dates). */
export function isDateFormat(code: string): boolean {
  return sectionsOf(code)[0]?.isDate ?? false;
}

/** Format a number with a spreadsheet number format code. */
export function formatValue(v: number, code: string, date1904 = false): string {
  if (!Number.isFinite(v)) return '';
  const trimmed = code.trim();
  if (!trimmed || /^general$/i.test(trimmed)) return formatGeneral(v);
  // the fourth section, and any section that only places the text ('@'), is for text
  const forText = (s: Section) =>
    s.tokens.some(t => t.t === 'text') && s.tokens.every(t => t.t === 'text' || t.t === 'lit');
  const sections = sectionsOf(trimmed).slice(0, 3).filter(s => !forText(s));
  let section: Section | undefined;
  const value = Math.abs(v);
  let minus = v < 0;
  if (sections.some(s => s.cond)) {
    // up to two conditions, then a section for everything else; the sign shows wherever the value lands
    section = sections.find(s => !s.cond || test(s.cond.op, v, s.cond.v)) ?? sections.at(-1);
  } else if (v < 0 && sections.length >= 2) {
    section = sections[1];
    minus = false;                                     // the negative section writes its own sign, if any
  } else if (v === 0 && sections.length >= 3) {
    section = sections[2];
  } else {
    section = sections[0];
  }
  if (!section) return formatGeneral(v);
  if (section.isDate) return v < 0 ? '#####' : formatDateSection(v, section.tokens, date1904).trim();
  if (!section.tokens.some(t => t.t === 'digit' || t.t === 'general')) {
    // a section of literals only, such as the "-" of an accounting format's zero section
    const s = section.tokens.map(t => (t.t === 'lit' ? t.s : t.t === 'percent' ? '%' : '')).join('').trim();
    return (minus ? '-' : '') + s;
  }
  const s = formatNumberSection(value, section.tokens).trim();
  // a negative number that rounds to zero shows no sign
  return minus && /[1-9]/.test(s) ? `-${s}` : s;
}
