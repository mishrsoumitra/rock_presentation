import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { beforeAll, describe, expect, test } from 'vitest';
import {
  autonumberLabel, caretAt, fontClasses, FontRegistry, hitTest, layoutText, scriptOfLang,
  type BodyStyle, type ParaStyle, type ParagraphInput, type RunStyle, type TextLayout,
} from '../src/index.ts';

const FONTS = join(import.meta.dirname, '../../../assets/fonts');
const fonts = new FontRegistry();

beforeAll(() => {
  for (const f of ['carlito/Carlito-Regular.ttf', 'carlito/Carlito-Bold.ttf', 'carlito/Carlito-Italic.ttf', 'liberation/LiberationSans-Regular.ttf', 'liberation/LiberationSerif-Regular.ttf']) {
    expect(fonts.addBytes(new Uint8Array(readFileSync(join(FONTS, f))))).toBeDefined();
  }
});

const style = (over: Partial<RunStyle> = {}): RunStyle => ({
  size: 18, bold: false, italic: false, caps: 'none', spacing: 0, kern: 0, baseline: 0,
  latin: 'Calibri', ea: '', cs: '', scriptFonts: {}, lang: 'en-US', ...over,
});
const para = (over: Partial<ParaStyle> = {}): ParaStyle => ({
  level: 0, align: 'l', marL: 0, marR: 0, indent: 0, lineSpacing: { kind: 'pct', val: 1 }, spaceBefore: { kind: 'pts', val: 0 },
  spaceAfter: { kind: 'pts', val: 0 }, defTab: 914400, tabs: [], rtl: false, bullet: null, ...over,
});
const body = (over: Partial<BodyStyle> = {}): BodyStyle => ({
  insets: { l: 0, t: 0, r: 0, b: 0 }, anchor: 't', anchorCtr: false, wrap: true, vert: 'horz', numCol: 1, spcCol: 0,
  autofit: { kind: 'none' }, spcFirstLastPara: false, ...over,
});
const p = (text: string, ps: Partial<ParaStyle> = {}, rs: Partial<RunStyle> = {}): ParagraphInput<RunStyle> =>
  ({ props: para(ps), runs: text ? [{ kind: 'text', text, props: style(rs) }] : [], endProps: style(rs) });
const lay = (paragraphs: ParagraphInput<RunStyle>[], w = 3_000_000, h = 3_000_000, b: Partial<BodyStyle> = {}) =>
  layoutText({ body: body(b), paragraphs }, { w, h }, fonts);
const lineTexts = (l: TextLayout<RunStyle>) => l.lines.map(x => l.paragraphText[x.paragraph]!.slice(x.start, x.end));

describe('fonts', () => {
  test('Calibri resolves to its metric-compatible substitute, with synthetic italic only when no italic face exists', () => {
    const m = fonts.resolve('Calibri', false, false)!;
    expect([m.face.family, m.substituted]).toEqual(['Carlito', true]);
    expect(fonts.resolve('Calibri', true, false)!.face.metrics.bold).toBe(true);
    expect(fonts.resolve('Arial', false, true)).toMatchObject({ syntheticItalic: true, substituted: true });
    expect(fonts.resolve('No Such Font', false, false)!.face.family).toBe('Carlito');
  });

  test('line metrics come from the Windows ascent and descent', () => {
    const m = fonts.resolve('Carlito', false, false)!.face.metrics;
    expect([m.upem, m.ascent, m.descent]).toEqual([2048, 1950, 550]);
  });
});

describe('line breaking', () => {
  test('a line that fits stays one line, at the natural line height', () => {
    const l = lay([p('Hello world')]);
    expect(lineTexts(l)).toEqual(['Hello world']);
    const line = l.lines[0]!;
    // Carlito: (1950 + 550) / 2048 of 18 pt
    expect(line.height).toBeCloseTo((2500 / 2048) * 18 * 12700, 0);
    expect(line.baseline).toBeCloseTo((1950 / 2048) * 18 * 12700, 0);
  });

  test('text wraps at spaces, and trailing spaces do not push a word onto the next line', () => {
    const l = lay([p('The quick brown fox jumps over the lazy dog')], 1_200_000);
    expect(l.lines.length).toBeGreaterThan(1);
    expect(lineTexts(l).join('')).toBe('The quick brown fox jumps over the lazy dog');
    for (const line of l.lines) expect(line.width).toBeLessThanOrEqual(line.available + 1);
    for (const t of lineTexts(l).slice(0, -1)) expect(t.endsWith(' ')).toBe(true);
  });

  test('a word wider than the box is broken between characters', () => {
    const l = lay([p('Supercalifragilisticexpialidocious')], 600_000);
    expect(l.lines.length).toBeGreaterThan(1);
    expect(lineTexts(l).join('')).toBe('Supercalifragilisticexpialidocious');
  });

  test('wrap="none" keeps everything on one line; a line break still breaks', () => {
    const l = lay([p('The quick brown fox jumps over the lazy dog')], 300_000, 3_000_000, { wrap: false });
    expect(l.lines.length).toBe(1);
    const withBreak = lay([{ props: para(), runs: [{ kind: 'text', text: 'one', props: style() }, { kind: 'break', text: '\n', props: style() }, { kind: 'text', text: 'two', props: style() }], endProps: style() }]);
    expect(lineTexts(withBreak)).toEqual(['one', 'two']);
  });

  test('an empty paragraph still takes a line, sized by its paragraph mark', () => {
    const l = lay([p('a'), p('', {}, { size: 36 }), p('b')]);
    expect(l.lines.length).toBe(3);
    expect(l.lines[1]!.height).toBeCloseTo(2 * l.lines[0]!.height, 0);
  });
});

describe('spacing, alignment, anchoring', () => {
  test('percentage and exact line spacing', () => {
    const single = lay([p('x')]).lines[0]!.height;
    expect(lay([p('x', { lineSpacing: { kind: 'pct', val: 0.9 } })]).lines[0]!.height).toBeCloseTo(single * 0.9, 0);
    expect(lay([p('x', { lineSpacing: { kind: 'pts', val: 3000 } })]).lines[0]!.height).toBeCloseTo(30 * 12700, 0);
  });

  test('space before is skipped on the first paragraph unless spcFirstLastPara', () => {
    const before = { kind: 'pts' as const, val: 1200 };
    const a = lay([p('a', { spaceBefore: before }), p('b', { spaceBefore: before })]);
    expect(a.lines[0]!.top).toBe(0);
    expect(a.lines[1]!.top - a.lines[0]!.height).toBeCloseTo(12 * 12700, 0);
    expect(lay([p('a', { spaceBefore: before })], 3e6, 3e6, { spcFirstLastPara: true }).lines[0]!.top).toBeCloseTo(12 * 12700, 0);
  });

  test('centre and right alignment, and the insets', () => {
    const w = 3_000_000;
    const ins = { l: 91440, t: 45720, r: 91440, b: 45720 };
    const l = lay([p('Hi'), p('Hi', { align: 'ctr' }), p('Hi', { align: 'r' })], w, 3e6, { insets: ins });
    const [left, ctr, right] = l.lines.map(x => x.runs[0]!);
    expect(left!.x).toBeCloseTo(ins.l, 0);
    expect(ctr!.x + ctr!.advance / 2).toBeCloseTo(ins.l + (w - ins.l - ins.r) / 2, 0);
    expect(right!.x + right!.advance).toBeCloseTo(w - ins.r, 0);
    expect(l.lines[0]!.top).toBeCloseTo(ins.t, 0);
  });

  test('justified lines fill the width, except the last', () => {
    const l = lay([p('The quick brown fox jumps over the lazy dog and keeps running far away', { align: 'just' })], 1_500_000);
    for (const line of l.lines.slice(0, -1)) expect(line.width).toBeCloseTo(line.available, 0);
    const last = l.lines.at(-1)!;
    expect(last.width).toBeLessThan(last.available);
  });

  test('vertical anchoring', () => {
    const top = lay([p('x')], 3e6, 3e6, { anchor: 't' }).lines[0]!.top;
    const mid = lay([p('x')], 3e6, 3e6, { anchor: 'ctr' }).lines[0]!;
    const bot = lay([p('x')], 3e6, 3e6, { anchor: 'b' }).lines[0]!;
    expect(top).toBe(0);
    expect(mid.top + mid.height / 2).toBeCloseTo(1.5e6, 0);
    expect(bot.top + bot.height).toBeCloseTo(3e6, 0);
  });

  test('the stored autofit scale shrinks fonts and line spacing', () => {
    const full = lay([p('x')]).lines[0]!;
    const fit = lay([p('x')], 3e6, 3e6, { autofit: { kind: 'norm', fontScale: 0.5, lnSpcReduction: 0.2 } });
    expect(fit.lines[0]!.runs[0]!.size).toBeCloseTo(full.runs[0]!.size / 2, 0);
    expect(fit.lines[0]!.height).toBeCloseTo((full.height / 2) * 0.8, 0);
  });
});

describe('bullets, numbering, tabs, kerning', () => {
  test('a hanging bullet sits at marL + indent and the text at marL', () => {
    const l = lay([p('Item', { marL: 342900, indent: -342900, bullet: { kind: 'char', char: '•', font: 'Arial', size: null } })]);
    const [bullet, text] = l.lines[0]!.runs;
    expect(bullet!.bullet).toBe(true);
    expect(bullet!.x).toBeCloseTo(0, 0);
    expect(text!.x).toBeCloseTo(342900, 0);
    expect(bullet!.match.face.family).toBe('Liberation Sans');
  });

  test('auto numbering counts consecutive paragraphs at a level and skips empty ones', () => {
    const num = { kind: 'autonum' as const, scheme: 'arabicPeriod', startAt: 1, font: null, size: null };
    const l = lay([p('a', { bullet: num }), p('', { bullet: num }), p('b', { bullet: num }), p('c', { bullet: { ...num, scheme: 'alphaLcParenR' } })]);
    const labels = l.lines.filter(x => x.runs[0]?.bullet).map(x => x.runs[0]!.glyphs.length);
    expect(labels).toEqual([2, 2, 2]);
    expect([autonumberLabel('arabicPeriod', 3), autonumberLabel('romanUcPeriod', 4), autonumberLabel('alphaLcParenR', 28), autonumberLabel('arabicParenBoth', 2)])
      .toEqual(['3.', 'IV.', 'ab)', '(2)']);
  });

  test('a tab moves to the next default stop, a custom stop wins', () => {
    const l = lay([p('a\tb')]);
    const b = l.lines[0]!.runs.at(-1)!;
    expect(b.x).toBeCloseTo(914400, 0);
    const c = lay([p('a\tb', { tabs: [{ pos: 400000, algn: 'l' }] })]).lines[0]!.runs.at(-1)!;
    expect(c.x).toBeCloseTo(400000, 0);
  });

  test('kerning applies only at or above the run\'s kern size', () => {
    const plain = lay([p('AVAVAV', {}, { kern: 0 })]).lines[0]!.width;
    const kerned = lay([p('AVAVAV', {}, { kern: 12 })]).lines[0]!.width;
    expect(kerned).toBeLessThan(plain);
    const above = lay([p('AVAVAV', {}, { kern: 24 })]).lines[0]!.width;
    expect(above).toBeCloseTo(plain, 0);
  });

  test('Latin text is drawn without standard ligatures', () => {
    const run = lay([p('office')]).lines[0]!.runs[0]!;
    expect(run.glyphs.length).toBe(6);
  });

  test('letter spacing widens every character', () => {
    const a = lay([p('abc')]).lines[0]!.width;
    const b = lay([p('abc', {}, { spacing: 2 })]).lines[0]!.width;
    expect(b - a).toBeCloseTo(3 * 2 * 12700, 0);
  });
});

describe('carets and hit testing', () => {
  test('caret positions run left to right and hit testing finds the nearest', () => {
    const l = lay([p('Hello world'), p('Second')]);
    const line = l.lines[0]!;
    const xs = line.carets.map(c => c.x);
    expect([...xs].sort((a, b) => a - b)).toEqual(xs);
    expect(line.carets[0]).toEqual({ offset: 0, x: 0 });
    expect(line.carets.at(-1)!.offset).toBe(11);
    const mid = caretAt(l, 0, 5)!;
    expect(hitTest(l, mid.x + 1, line.top + 1)).toEqual({ paragraph: 0, offset: 5 });
    expect(hitTest(l, 0, l.lines[1]!.top + 1)).toEqual({ paragraph: 1, offset: 0 });
  });
});

describe('scripts', () => {
  test('font classes: CJK is ea, Arabic is cs, neutrals follow their neighbours', () => {
    expect(fontClasses('ab 日本 x')).toEqual(['latin', 'latin', 'latin', 'ea', 'ea', 'ea', 'latin']);
    expect(fontClasses('١ مرحبا')[3]).toBe('cs');
    expect([scriptOfLang('ja-JP'), scriptOfLang('zh-TW'), scriptOfLang('zh-CN'), scriptOfLang('ar-SA'), scriptOfLang('hi-IN')]).toEqual(['Jpan', 'Hant', 'Hans', 'Arab', 'Deva']);
  });

  test('right-to-left text shapes and reports itself as such', () => {
    const l = lay([p('abc مرحبا def', { rtl: false }, { cs: 'Liberation Sans' })]);
    expect(l.lines[0]!.runs.some(r => r.rtl)).toBe(true);
  });
});
