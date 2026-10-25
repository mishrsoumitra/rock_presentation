// packages/text/src/layout.ts  -- paragraph and text body layout (plan section 3.4)
/// <reference path="./vendor.d.ts" />
//
// The browser's own text engine differs between platforms, so text is shaped here with HarfBuzz and drawn from
// glyph outlines: the same layout on every machine, used for rendering, caret placement, hit-testing and the
// line-break comparison against PowerPoint.
//
// Units: EMU throughout, y down; font sizes come in points.
import * as hb from 'harfbuzzjs';
import LineBreaker from 'linebreak';
import { getEmbeddingLevels } from 'bidi-js/src/index.js';
import type { FontMatch, FontRegistry } from './registry.ts';
import { fontClasses, hbScriptOf, scriptOfLang, type FontClass } from './script.ts';

export const EMU_PER_PT = 12700;

/** What layout needs from a run's resolved properties; anything else (paint, links) passes through untouched. */
export interface RunStyle {
  size: number;
  bold: boolean;
  italic: boolean;
  caps: 'none' | 'small' | 'all';
  spacing: number;                 // points
  kern: number;                    // points; kerning applies at and above this size, 0 = never
  baseline: number;                // fraction: positive superscript, negative subscript
  latin: string;
  ea: string;
  cs: string;
  scriptFonts: Record<string, string>;
  lang: string;
  altLang?: string;
}
export interface RunInput<S extends RunStyle> { kind: 'text' | 'break' | 'field'; text: string; props: S; }
export type Spacing = { kind: 'pct'; val: number } | { kind: 'pts'; val: number };
export interface BulletStyle {
  kind: 'char' | 'autonum' | 'picture';
  char?: string;
  scheme?: string;
  startAt?: number;
  font: string | null;
  size: { kind: 'pct'; val: number } | { kind: 'pts'; val: number } | null;
}
export interface ParaStyle {
  level: number;
  align: string;
  marL: number;
  marR: number;
  indent: number;
  lineSpacing: Spacing;
  spaceBefore: Spacing;
  spaceAfter: Spacing;
  defTab: number;
  tabs: { pos: number; algn: string }[];
  rtl: boolean;
  bullet: BulletStyle | null;
}
export interface ParagraphInput<S extends RunStyle> { props: ParaStyle; runs: RunInput<S>[]; endProps: S; }
export type Autofit = { kind: 'none' } | { kind: 'norm'; fontScale: number; lnSpcReduction: number } | { kind: 'shape' };
export interface BodyStyle {
  insets: { l: number; t: number; r: number; b: number };
  anchor: string;
  anchorCtr: boolean;
  wrap: boolean;
  vert: string;
  numCol: number;
  spcCol: number;
  autofit: Autofit;
  spcFirstLastPara: boolean;
}
export interface TextInput<S extends RunStyle> { body: BodyStyle; paragraphs: ParagraphInput<S>[]; }

export interface GlyphRun<S extends RunStyle> {
  match: FontMatch;
  /** The em size in EMU (font size after autofit, and after super/subscript or small-caps reduction). */
  size: number;
  glyphs: number[];
  /** Glyph origins relative to the run origin (x right, y down), in EMU. */
  xs: number[];
  ys: number[];
  advance: number;
  /** Run origin relative to the text box: x of its left edge, y of its baseline. */
  x: number;
  baseline: number;
  props: S;
  paragraph: number;
  /** Paragraph character range [start, end) the run draws, and the paragraph offset of each glyph. */
  start: number;
  end: number;
  clusters: number[];
  rtl: boolean;
  bullet?: boolean;
}
export interface Line<S extends RunStyle> {
  paragraph: number;
  start: number;
  end: number;
  top: number;
  height: number;
  baseline: number;
  /** The tallest ascent and descent on the line, before line spacing: what a caret spans. */
  ascent: number;
  descent: number;
  /** Where text starts on the line before alignment, and the width available to it. */
  left: number;
  available: number;
  /** Width of the line's text (trailing spaces excluded). */
  width: number;
  runs: GlyphRun<S>[];
  /** x of the caret before each paragraph offset from start to end inclusive. */
  carets: { offset: number; x: number }[];
}
export interface TextLayout<S extends RunStyle> {
  lines: Line<S>[];
  /** The box the text was laid out in (vertical text: width and height swapped). */
  width: number;
  height: number;
  /** Top of the first line and the height of all lines, inside the box. */
  contentTop: number;
  contentHeight: number;
  paragraphText: string[];
  fontScale: number;
  lineSpacingReduction: number;
  /** Rotation the renderer applies for vertical text: 0, 90 or 270 degrees. */
  rotation: 0 | 90 | 270;
}

interface Item {
  start: number;
  end: number;
  run: number;
  match: FontMatch;
  level: number;
  smallCap: boolean;
  tab: boolean;
}

interface ShapedItem extends Item {
  size: number;                       // em size, EMU
  rise: number;                       // baseline shift up, EMU
  glyphs: number[];
  advances: number[];                 // EMU, letter spacing included
  offsetsX: number[];
  offsetsY: number[];
  clusters: number[];
  ascent: number;
  descent: number;
}

const isSpace = (c: string | undefined) => c === ' ' || c === ' ' || c === '　';

/** The face for one character of a run, by its font class, with fallback for characters the face lacks. */
function faceFor<S extends RunStyle>(props: S, cls: FontClass, cp: number, fonts: FontRegistry): FontMatch | undefined {
  const scriptFont = (lang: string | undefined) => props.scriptFonts[scriptOfLang(lang) ?? ''] ?? '';
  const family = cls === 'ea'
    ? props.ea || scriptFont(props.altLang ?? props.lang) || props.latin
    : cls === 'cs' ? props.cs || scriptFont(props.lang) || props.latin : props.latin;
  const m = fonts.resolve(family, props.bold, props.italic);
  if (!m || cp <= 0x20 || m.face.hasGlyph(cp)) return m;
  return fonts.fallback(cp, props.bold, props.italic, [props.latin, props.ea, props.cs].filter(Boolean)) ?? m;
}

function itemize<S extends RunStyle>(text: string, runOf: Int32Array, runs: RunInput<S>[], levels: Uint8Array, fonts: FontRegistry): Item[] {
  const classes = fontClasses(text);
  const items: Item[] = [];
  let cur: Item | null = null;
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!;
    const n = cp > 0xffff ? 2 : 1;
    const run = runOf[i]!;
    const props = runs[run]!.props;
    const match = faceFor(props, classes[i]!, cp, fonts);
    const tab = cp === 9;
    const ch = text[i]!;
    const smallCap = props.caps === 'small' && ch !== ch.toUpperCase();
    if (!match) { i += n; continue; }
    if (cur && !tab && !cur.tab && cur.run === run && cur.match.face === match.face && cur.match.syntheticBold === match.syntheticBold
      && cur.level === levels[i] && cur.smallCap === smallCap && text[i - 1] !== '\n' && ch !== '\n') {
      cur.end = i + n;
    } else {
      cur = { start: i, end: i + n, run, match, level: levels[i] ?? 0, smallCap, tab };
      items.push(cur);
    }
    i += n;
  }
  return items;
}

function shapeItem<S extends RunStyle>(item: Item, text: string, props: S, scale: number): ShapedItem {
  const reduce = (props.baseline !== 0 ? 2 / 3 : 1) * (item.smallCap ? 0.8 : 1);
  const size = props.size * EMU_PER_PT * scale * reduce;
  const rise = props.baseline * props.size * EMU_PER_PT * scale;
  const face = item.match.face;
  const m = face.metrics;
  const unit = size / m.upem;
  const shaped: ShapedItem = {
    ...item, size, rise, glyphs: [], advances: [], offsetsX: [], offsetsY: [], clusters: [],
    ascent: m.ascent * unit, descent: m.descent * unit,
  };
  if (item.tab || text[item.start] === '\n') return shaped;
  let content = text.slice(item.start, item.end);
  if (props.caps === 'all' || item.smallCap) {
    const upper = content.toUpperCase();
    if (upper.length === content.length) content = upper;
  }
  const buffer = new hb.Buffer();
  buffer.addText(content);
  buffer.guessSegmentProperties();
  const script = hbScriptOf(content.codePointAt(0) ?? 0);
  if (script) buffer.setScript(script);
  buffer.setDirection(item.level % 2 ? hb.Direction.RTL : hb.Direction.LTR);
  if (props.lang) buffer.setLanguage(props.lang);
  const kern = props.kern > 0 && props.size * scale >= props.kern;
  const features = [`${kern ? '+' : '-'}kern`];
  // PowerPoint draws Latin text without standard ligatures; complex scripts need theirs
  if (!script || script === 'Latn' || script === 'Cyrl' || script === 'Grek') features.push('-liga', '-clig');
  hb.shape(face.hbFont, buffer, features.map(f => hb.Feature.fromString(f)!).filter(Boolean));
  const infos = buffer.getGlyphInfos();
  const positions = buffer.getGlyphPositions();
  const extra = props.spacing * EMU_PER_PT * scale;
  for (let g = 0; g < infos.length; g++) {
    const info = infos[g]!;
    const pos = positions[g]!;
    const cluster = item.start + info.cluster;
    shaped.glyphs.push(info.codepoint);
    shaped.clusters.push(cluster);
    shaped.offsetsX.push(pos.xOffset * unit);
    shaped.offsetsY.push(-pos.yOffset * unit);
    // letter spacing goes after the last glyph of each cluster
    const lastOfCluster = g === infos.length - 1 || infos[g + 1]!.cluster !== info.cluster;
    shaped.advances.push(pos.xAdvance * unit + (lastOfCluster ? extra : 0));
  }
  return shaped;
}

function lineHeightOf(asc: number, desc: number, spacing: Spacing, reduction: number): { height: number; above: number } {
  if (spacing.kind === 'pts') {
    const height = (spacing.val / 100) * EMU_PER_PT;
    return { height, above: asc + desc > 0 ? height * (asc / (asc + desc)) : height };
  }
  const pct = Math.max(0, spacing.val - reduction);
  return { height: (asc + desc) * pct, above: asc * pct };
}

function spacingOf(s: Spacing, naturalLine: number, scale: number): number {
  return s.kind === 'pts' ? (s.val / 100) * EMU_PER_PT * scale : s.val * naturalLine;
}

// ------------------------------------------------------------------------------------------------- auto numbering

const ROMAN: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
function roman(n: number): string {
  let out = '';
  let v = Math.max(1, Math.min(3999, n));
  for (const [k, s] of ROMAN) while (v >= k) { out += s; v -= k; }
  return out;
}
function alpha(n: number): string {
  let out = '';
  let v = Math.max(1, n);
  while (v > 0) { v--; out = String.fromCharCode(97 + (v % 26)) + out; v = Math.floor(v / 26); }
  return out;
}

/** The label of an ST_TextAutonumberScheme for number n: arabicPeriod -> "3.", romanUcParenR -> "III)". */
export function autonumberLabel(scheme: string, n: number): string {
  const m = /^(arabic|alphaLc|alphaUc|romanLc|romanUc|circleNumDbPlain|circleNumWdBlackPlain|circleNumWdWhitePlain)(.*)$/.exec(scheme);
  const kind = m?.[1] ?? 'arabic';
  const tail = m?.[2] ?? 'Period';
  let num: string;
  switch (kind) {
    case 'alphaLc': num = alpha(n); break;
    case 'alphaUc': num = alpha(n).toUpperCase(); break;
    case 'romanLc': num = roman(n); break;
    case 'romanUc': num = roman(n).toUpperCase(); break;
    case 'circleNumDbPlain': case 'circleNumWdWhitePlain': return n >= 1 && n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : String(n);
    case 'circleNumWdBlackPlain': return n >= 1 && n <= 10 ? String.fromCodePoint(0x2776 + n - 1) : String(n);
    default: num = String(n);
  }
  if (tail.startsWith('ParenBoth')) return `(${num})`;
  if (tail.startsWith('ParenR')) return `${num})`;
  if (tail.startsWith('Plain')) return num;
  if (tail.startsWith('Minus')) return `- ${num} -`;
  return `${num}.`;
}

// ------------------------------------------------------------------------------------------------------ the layout

/** Visual order of runs on one line: reverse every maximal sequence at or above each odd level (UAX #9, L2). */
function visualOrder<T extends { level: number }>(runs: T[]): T[] {
  if (!runs.some(r => r.level % 2)) return runs;
  const out = [...runs];
  const max = Math.max(...out.map(r => r.level));
  const minOdd = Math.min(...out.map(r => r.level).filter(l => l % 2)) ;
  for (let level = max; level >= minOdd; level--) {
    for (let i = 0; i < out.length;) {
      if (out[i]!.level >= level) {
        let j = i;
        while (j < out.length && out[j]!.level >= level) j++;
        out.splice(i, j - i, ...out.slice(i, j).reverse());
        i = j;
      } else i++;
    }
  }
  return out;
}

interface ParagraphLines<S extends RunStyle> { lines: Line<S>[]; spaceBefore: number; spaceAfter: number; }

function layoutParagraph<S extends RunStyle>(
  p: ParagraphInput<S>, index: number, width: number, wrap: boolean, scale: number, reduction: number, fonts: FontRegistry,
  bulletLabel: string | undefined,
): ParagraphLines<S> & { text: string } {
  const text = p.runs.map(r => (r.kind === 'break' ? '\n' : r.text)).join('');
  const runOf = new Int32Array(text.length);
  let at = 0;
  p.runs.forEach((r, k) => {
    const len = r.kind === 'break' ? 1 : r.text.length;
    runOf.fill(k, at, at + len);
    at += len;
  });
  const levels = text.length ? getEmbeddingLevels(text, p.props.rtl ? 'rtl' : 'ltr').levels : new Uint8Array(0);
  const items = itemize(text, runOf, p.runs, levels, fonts).map(it => shapeItem(it, text, p.runs[it.run]!.props, scale));

  // advance of each character position (a ligature's whole advance sits on its first character)
  const adv = new Float64Array(text.length + 1);
  for (const it of items) {
    for (let g = 0; g < it.glyphs.length; g++) {
      const c = it.clusters[g]!;
      adv[c] = (adv[c] ?? 0) + it.advances[g]!;
    }
  }

  const props = p.props;
  const firstRunProps = p.runs.find(r => r.kind !== 'break')?.props ?? p.endProps;
  // the bullet: shaped as its own item, placed at marL + indent on the first line
  let bullet: ShapedItem | undefined;
  if (props.bullet && bulletLabel && text.trim().length) {
    const b = props.bullet;
    const size = b.size === null ? firstRunProps.size : b.size.kind === 'pct' ? firstRunProps.size * b.size.val : b.size.val / 100;
    const bprops = { ...firstRunProps, size, latin: b.font ?? firstRunProps.latin, ea: b.font ?? firstRunProps.ea, cs: b.font ?? firstRunProps.cs, caps: 'none' as const, spacing: 0, baseline: 0 };
    const cp = bulletLabel.codePointAt(0) ?? 0x2022;
    const match = faceFor(bprops, 'latin', cp, fonts);
    if (match) bullet = shapeItem({ start: 0, end: bulletLabel.length, run: -1, match, level: 0, smallCap: false, tab: false }, bulletLabel, bprops, scale);
  }
  const bulletAdvance = bullet ? bullet.advances.reduce((a, b) => a + b, 0) : 0;
  const bulletX = props.marL + props.indent;
  const firstLeft = bullet ? Math.max(props.marL, bulletX + bulletAdvance) : props.marL + props.indent;
  const leftOf = (lineNo: number) => (lineNo === 0 ? firstLeft : props.marL);
  const availableOf = (lineNo: number) => Math.max(0, width - leftOf(lineNo) - props.marR);

  // tab stops: custom stops first, then every defTab, measured from the left of the text box
  const nextTab = (x: number) => {
    for (const t of [...props.tabs].sort((a, b) => a.pos - b.pos)) if (t.pos > x + 1) return t.pos;
    const step = props.defTab > 0 ? props.defTab : 914400;
    return (Math.floor(x / step) + 1) * step;
  };

  // break opportunities (UAX #14); mandatory ones at our '\n'
  const breaks: { pos: number; required: boolean }[] = [];
  if (text.length) {
    const breaker = new LineBreaker(text);
    for (let b = breaker.nextBreak(); b; b = breaker.nextBreak()) breaks.push({ pos: b.position, required: b.required });
  }
  if (!breaks.length || breaks[breaks.length - 1]!.pos !== text.length) breaks.push({ pos: text.length, required: true });

  // greedy line filling over break opportunities, with character breaks for words wider than a line
  const ranges: { start: number; end: number }[] = [];
  let lineStart = 0;
  let lineNo = 0;
  let x = 0;                                          // pen position from the line's left edge
  const widthOf = (s: number, e: number, startX: number) => {
    let w = startX;
    for (let i = s; i < e; i++) {
      if (text[i] === '\t') w = nextTab(leftOf(lineNo) + w) - leftOf(lineNo);
      else if (text[i] !== '\n') w += adv[i]!;
    }
    return w - startX;
  };
  const trimmedEnd = (s: number, e: number) => {
    let k = e;
    while (k > s && (isSpace(text[k - 1]) || text[k - 1] === '\n')) k--;
    return k;
  };
  let segStart = 0;
  for (const b of breaks) {
    const segEnd = b.pos;
    const visibleEnd = trimmedEnd(segStart, segEnd);
    const segWidth = widthOf(segStart, visibleEnd, x);
    if (wrap && lineStart < segStart && x + segWidth > availableOf(lineNo) + 1) {
      ranges.push({ start: lineStart, end: segStart });
      lineStart = segStart;
      lineNo++;
      x = 0;
    }
    // a single segment wider than the line is broken between characters
    if (wrap && x === 0 && widthOf(segStart, visibleEnd, 0) > availableOf(lineNo) + 1) {
      let w = 0;
      for (let i = segStart; i < visibleEnd; i++) {
        const cw = text[i] === '\t' ? 0 : adv[i]!;
        if (w + cw > availableOf(lineNo) + 1 && i > lineStart) {
          ranges.push({ start: lineStart, end: i });
          lineStart = i;
          lineNo++;
          w = 0;
        }
        w += cw;
      }
      x = w + widthOf(visibleEnd, segEnd, w);
    } else {
      x += widthOf(segStart, segEnd, x);
    }
    segStart = segEnd;
    if (b.required && segEnd < text.length) {
      ranges.push({ start: lineStart, end: segEnd });
      lineStart = segEnd;
      lineNo++;
      x = 0;
    }
  }
  ranges.push({ start: lineStart, end: text.length });

  // place each line
  const lines: Line<S>[] = [];
  ranges.forEach((range, li) => {
    const contentEnd = text[range.end - 1] === '\n' ? range.end - 1 : range.end;
    const pieces: (ShapedItem & { gStart: number; gEnd: number })[] = [];
    for (const it of items) {
      if (it.end <= range.start || it.start >= contentEnd) continue;
      if (it.tab || text[it.start] === '\n') {
        pieces.push({ ...it, gStart: 0, gEnd: 0, start: Math.max(it.start, range.start), end: Math.min(it.end, contentEnd) });
        continue;
      }
      // glyphs whose cluster falls on this line
      const idx: number[] = [];
      for (let g = 0; g < it.glyphs.length; g++) {
        const c = it.clusters[g]!;
        if (c >= range.start && c < contentEnd) idx.push(g);
      }
      if (!idx.length) continue;
      pieces.push({ ...it, gStart: Math.min(...idx), gEnd: Math.max(...idx) + 1, start: Math.max(it.start, range.start), end: Math.min(it.end, contentEnd) });
    }
    const ordered = visualOrder(pieces);
    const left = leftOf(li);
    let pen = 0;
    let asc = 0;
    let desc = 0;
    const runs: GlyphRun<S>[] = [];
    const carets: { offset: number; x: number }[] = [];
    for (const piece of ordered) {
      if (piece.tab) {
        const to = nextTab(left + pen) - left;
        carets.push({ offset: piece.start, x: left + pen });
        pen = Math.max(pen, to);
        continue;
      }
      if (piece.gEnd === piece.gStart) continue;
      const xs: number[] = [];
      const ys: number[] = [];
      let runPen = 0;
      const glyphs = piece.glyphs.slice(piece.gStart, piece.gEnd);
      const advances = piece.advances.slice(piece.gStart, piece.gEnd);
      const clusters = piece.clusters.slice(piece.gStart, piece.gEnd);
      for (let g = 0; g < glyphs.length; g++) {
        xs.push(runPen + piece.offsetsX[piece.gStart + g]!);
        ys.push(piece.offsetsY[piece.gStart + g]! - piece.rise);
        runPen += advances[g]!;
      }
      const rtl = piece.level % 2 === 1;
      for (let g = 0; g < glyphs.length; g++) {
        const gx = left + pen + xs[g]! - piece.offsetsX[piece.gStart + g]!;
        carets.push({ offset: clusters[g]!, x: rtl ? gx + advances[g]! : gx });
      }
      runs.push({
        match: piece.match, size: piece.size, glyphs, xs, ys, advance: runPen, x: left + pen, baseline: 0,
        props: p.runs[piece.run]!.props, paragraph: index, start: piece.start, end: piece.end, clusters, rtl,
      });
      asc = Math.max(asc, piece.ascent + Math.max(0, piece.rise));
      desc = Math.max(desc, piece.descent + Math.max(0, -piece.rise));
      pen += runPen;
    }
    if (!runs.length) {
      // an empty line takes its height from the paragraph mark
      const match = faceFor(p.endProps, 'latin', 0x41, fonts);
      if (match) {
        const unit = (p.endProps.size * EMU_PER_PT * scale) / match.face.metrics.upem;
        asc = match.face.metrics.ascent * unit;
        desc = match.face.metrics.descent * unit;
      }
    }
    if (li === 0 && bullet) {
      asc = Math.max(asc, bullet.ascent);
      desc = Math.max(desc, bullet.descent);
    }
    // trailing spaces hang past the edge: they do not count for alignment
    let width = pen;
    for (let i = contentEnd - 1; i >= range.start && isSpace(text[i]); i--) width -= adv[i]!;
    carets.push({ offset: contentEnd, x: left + pen });
    carets.sort((a, b) => a.offset - b.offset || a.x - b.x);
    const { height, above } = lineHeightOf(asc, desc, props.lineSpacing, reduction);
    const available = availableOf(li);
    const line: Line<S> = { paragraph: index, start: range.start, end: contentEnd, top: 0, height, baseline: above, ascent: asc, descent: desc, left, available, width, runs, carets };

    // alignment
    const lastLine = li === ranges.length - 1 || text[contentEnd] === '\n';
    const extra = available - width;
    let shift = 0;
    if (props.align === 'ctr') shift = extra / 2;
    else if (props.align === 'r') shift = extra;
    else if ((props.align === 'just' && !lastLine) || props.align === 'dist') {
      // Justification spreads the slack after each gap: every inner space for "just", every character for "dist".
      // A glyph moves right by the slack of all the gaps before its character.
      let lastVisible = contentEnd;
      while (lastVisible > range.start && isSpace(text[lastVisible - 1])) lastVisible--;
      const gaps: number[] = [];
      for (let i = range.start; i < lastVisible - 1; i++) if (props.align === 'dist' || isSpace(text[i])) gaps.push(i);
      const per = gaps.length && extra > 0 ? extra / gaps.length : 0;
      if (per) {
        const before = (offset: number) => {
          let n = 0;
          for (const gp of gaps) if (gp < offset) n++;
          return n;
        };
        for (const r of line.runs) {
          const base = before(r.start);
          for (let g = 0; g < r.glyphs.length; g++) r.xs[g]! += (before(r.clusters[g]!) - base) * per;
          r.x += base * per;
          r.advance += (before(r.end) - base) * per;
        }
        for (const c of line.carets) c.x += before(c.offset) * per;
        line.width = available;
      }
    }
    if (shift && extra > 0) {
      for (const r of line.runs) r.x += shift;
      for (const c of line.carets) c.x += shift;
    }
    if (li === 0 && bullet) {
      line.runs.unshift({
        match: bullet.match, size: bullet.size, glyphs: bullet.glyphs, xs: bullet.advances.map((_, g) => bullet!.advances.slice(0, g).reduce((a, b) => a + b, 0)),
        ys: bullet.glyphs.map(() => 0), advance: bulletAdvance, x: bulletX + (props.align === 'ctr' || props.align === 'r' ? shift : 0), baseline: 0,
        props: p.runs.find(r => r.kind !== 'break')?.props ?? p.endProps, paragraph: index, start: 0, end: 0, clusters: [], rtl: false, bullet: true,
      });
    }
    lines.push(line);
  });

  // percentages in space before and after count lines of the paragraph's natural height
  const natural = (l: Line<S> | undefined) => (l ? l.ascent + l.descent : 0);
  return {
    text, lines,
    spaceBefore: spacingOf(props.spaceBefore, natural(lines[0]), scale),
    spaceAfter: spacingOf(props.spaceAfter, natural(lines.at(-1)), scale),
  };
}

/**
 * Lay out a text body in a box of the given size (EMU). Line positions are relative to the box's top-left corner.
 * The stored autofit scale applies as PowerPoint saved it; recomputing it after an edit is the editor's job.
 */
export function layoutText<S extends RunStyle>(input: TextInput<S>, box: { w: number; h: number }, fonts: FontRegistry): TextLayout<S> {
  const body = input.body;
  const rotation: 0 | 90 | 270 = body.vert === 'vert' || body.vert === 'eaVert' || body.vert === 'wordArtVertRtl' ? 90 : body.vert === 'vert270' ? 270 : 0;
  const w = rotation ? box.h : box.w;
  const h = rotation ? box.w : box.h;
  const scale = body.autofit.kind === 'norm' ? body.autofit.fontScale : 1;
  const reduction = body.autofit.kind === 'norm' ? body.autofit.lnSpcReduction : 0;
  const ins = body.insets;
  const innerW = Math.max(0, w - ins.l - ins.r);
  const innerH = Math.max(0, h - ins.t - ins.b);
  const numCol = Math.max(1, body.numCol);
  const colW = numCol > 1 ? Math.max(0, (innerW - (numCol - 1) * body.spcCol) / numCol) : innerW;

  // auto numbering: consecutive paragraphs at a level with the same scheme count up
  const counters: { scheme: string; n: number }[] = [];
  const labels = input.paragraphs.map(p => {
    const b = p.props.bullet;
    const level = p.props.level;
    counters.length = Math.min(counters.length, level + 1);
    const hasText = p.runs.some(r => r.kind !== 'break' && r.text.trim().length);
    if (!b) return undefined;
    if (b.kind === 'char') return b.char ?? '•';
    if (b.kind === 'picture') return '•';
    if (!hasText) return undefined;
    const prev = counters[level];
    const n = prev && prev.scheme === b.scheme ? prev.n + 1 : b.startAt ?? 1;
    counters[level] = { scheme: b.scheme ?? 'arabicPeriod', n };
    return autonumberLabel(b.scheme ?? 'arabicPeriod', n);
  });

  const laid = input.paragraphs.map((p, i) => layoutParagraph(p, i, colW, body.wrap, scale, reduction, fonts, labels[i]));
  const all: Line<S>[] = [];
  let y = 0;
  laid.forEach((pl, i) => {
    const first = i === 0;
    const last = i === laid.length - 1;
    if (!first || body.spcFirstLastPara) y += pl.spaceBefore;
    for (const line of pl.lines) {
      line.top = y;
      line.baseline = y + line.baseline;
      for (const r of line.runs) r.baseline = line.baseline;
      y += line.height;
      all.push(line);
    }
    if (!last || body.spcFirstLastPara) y += pl.spaceAfter;
  });
  let contentHeight = y;

  // columns: fill each to the box height, then move on
  if (numCol > 1 && all.length) {
    let col = 0;
    let colTop = 0;
    for (const line of all) {
      if (col < numCol - 1 && line.top - colTop + line.height > innerH && line.top > colTop) {
        col++;
        colTop = line.top;
      }
      const dx = col * (colW + body.spcCol);
      const dy = -colTop;
      line.top += dy;
      line.baseline += dy;
      for (const r of line.runs) { r.x += dx; r.baseline += dy; }
      for (const c of line.carets) c.x += dx;
      line.left += dx;
    }
    contentHeight = Math.max(...all.map(l => l.top + l.height));
  }

  const anchor = body.anchor;
  const top = ins.t + (anchor === 'ctr' ? (innerH - contentHeight) / 2 : anchor === 'b' ? innerH - contentHeight : 0);
  // anchorCtr centres the block of text horizontally as a whole
  let dxAll = ins.l;
  if (body.anchorCtr && all.length) {
    const blockW = Math.max(...all.map(l => l.left + l.width)) - Math.min(...all.map(l => l.left));
    dxAll += Math.max(0, (innerW - blockW) / 2) - Math.min(...all.map(l => l.left));
  }
  for (const line of all) {
    line.top += top;
    line.baseline += top;
    line.left += dxAll;
    for (const r of line.runs) { r.x += dxAll; r.baseline += top; }
    for (const c of line.carets) c.x += dxAll;
  }
  return {
    lines: all, width: w, height: h, contentTop: top, contentHeight, paragraphText: laid.map(l => l.text),
    fontScale: scale, lineSpacingReduction: reduction, rotation,
  };
}

/** The paragraph offset nearest to a point in the text box (EMU), for clicks and drags in the editor. */
export function hitTest<S extends RunStyle>(layout: TextLayout<S>, x: number, y: number): { paragraph: number; offset: number } {
  const lines = layout.lines;
  if (!lines.length) return { paragraph: 0, offset: 0 };
  let line = lines[0]!;
  for (const l of lines) if (y >= l.top) line = l;
  let best = line.carets[0] ?? { offset: line.start, x: line.left };
  for (const c of line.carets) if (Math.abs(c.x - x) < Math.abs(best.x - x)) best = c;
  return { paragraph: line.paragraph, offset: best.offset };
}

/** Where to draw the caret for a paragraph offset: x, top and height of the line, in the text box (EMU). */
export function caretAt<S extends RunStyle>(layout: TextLayout<S>, paragraph: number, offset: number): { x: number; top: number; height: number } | undefined {
  const lines = layout.lines.filter(l => l.paragraph === paragraph);
  if (!lines.length) return undefined;
  const line = lines.find(l => offset >= l.start && offset <= l.end && (offset < l.end || l === lines[lines.length - 1])) ?? lines[lines.length - 1]!;
  let best = line.carets[0] ?? { offset: line.start, x: line.left };
  for (const c of line.carets) if (Math.abs(c.offset - offset) < Math.abs(best.offset - offset) || (c.offset === offset && best.offset !== offset)) best = c;
  return { x: best.x, top: line.top, height: line.height };
}
