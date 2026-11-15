// The text formatting the ribbon offers beyond bold and size: highlight, super- and subscript, character spacing,
// line and paragraph spacing, the text box's anchor, margins and columns, and Clear Formatting. Each is checked in
// the XML it writes, in what the resolver then reads, and by exact undo and redo.
import { child, children, runsOf } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement, type XmlElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { bodyPlainText, clearRunFormatting, formatBody, formatParagraphs, formatRuns, insertTextBox, retypeText, typeText, type Editor, type TextRange } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const WORLD: TextRange = { start: { p: 0, offset: 6 }, end: { p: 0, offset: 11 } };
const ALL: TextRange = { start: { p: 0, offset: 0 }, end: { p: 1, offset: 6 } };

/** A deck with a text box holding "Hello world" and "Second"; returns the editor and the box's text body. */
function withBox(): { ed: Editor; id: number; body: () => XmlElement } {
  const ed = openEditor(deckBytes());
  const r = insertTextBox(ed, SLIDE, { x: 1000000, y: 1000000, w: 4000000, h: 800000 }, 'Hello world\nSecond');
  ed.run(r.command);
  ed.log.clear();
  return { ed, id: r.id, body: () => ed.shape(SLIDE, r.id)!.txBody! };
}

/** What the resolver reads for the box: its body properties and paragraphs. */
const resolved = (ed: Editor, id: number) => resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.text!;
const para = (body: XmlElement, i = 0) => children(body, 'a:p')[i]!;
const runXml = (body: XmlElement, i: number) => serializeElement(child(runsOf(para(body))[i]!.el, 'a:rPr')!);

describe('character formatting', () => {
  test('highlight: a colour on the selected words only, and removed again', () => {
    const { ed, id, body } = withBox();
    expect(cycle(ed, () => formatRuns(ed, SLIDE, body(), WORLD, { highlight: { kind: 'srgb', hex: 'FFFF00' } }))).toEqual([]);
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { highlight: { kind: 'srgb', hex: 'FFFF00' } }));
    expect(runsOf(para(body())).map(r => r.text)).toEqual(['Hello ', 'world']);
    expect(runXml(body(), 1)).toContain('<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>');
    const runs = resolved(ed, id).paragraphs[0]!.runs;
    expect(runs[0]!.props.highlight).toBeNull();
    expect(runs[1]!.props.highlight).toMatchObject({ r: 1, g: 1, b: 0 });
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { highlight: null }));
    expect(runXml(body(), 1)).not.toContain('highlight');
    expect(resolved(ed, id).paragraphs[0]!.runs[1]!.props.highlight).toBeNull();
  });

  test('superscript and subscript raise and lower the baseline; 0 puts the text back on the line', () => {
    const { ed, id, body } = withBox();
    expect(cycle(ed, () => formatRuns(ed, SLIDE, body(), WORLD, { baseline: 30000 }))).toEqual([]);
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { baseline: 30000 }));
    expect(runXml(body(), 1)).toContain('baseline="30000"');
    expect(resolved(ed, id).paragraphs[0]!.runs[1]!.props.baseline).toBeCloseTo(0.3);
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { baseline: -25000 }));
    expect(resolved(ed, id).paragraphs[0]!.runs[1]!.props.baseline).toBeCloseTo(-0.25);
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { baseline: 0 }));
    expect(resolved(ed, id).paragraphs[0]!.runs[1]!.props.baseline).toBe(0);
    expect(resolved(ed, id).paragraphs[0]!.runs[0]!.props.baseline).toBe(0);
  });

  test('character spacing is stored in hundredths of a point and read in points', () => {
    const { ed, id, body } = withBox();
    expect(cycle(ed, () => formatRuns(ed, SLIDE, body(), WORLD, { spc: 300 }))).toEqual([]);
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { spc: 300 }));
    expect(runXml(body(), 1)).toContain('spc="300"');
    expect(resolved(ed, id).paragraphs[0]!.runs[1]!.props.spacing).toBe(3);
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { spc: -150 }));
    expect(resolved(ed, id).paragraphs[0]!.runs[1]!.props.spacing).toBe(-1.5);
    ed.run(formatRuns(ed, SLIDE, body(), WORLD, { spc: null }));
    expect(runXml(body(), 1)).not.toContain('spc=');
  });

  test('Clear Formatting drops what the runs set themselves, keeps the language and the text, and leaves paragraphs alone', () => {
    const { ed, id, body } = withBox();
    ed.run(formatRuns(ed, SLIDE, body(), ALL, { b: true, i: true, u: 'sng', sz: 4000, spc: 300, baseline: 30000, latin: 'Georgia', fill: { kind: 'srgb', hex: 'FF0000' }, highlight: { kind: 'srgb', hex: 'FFFF00' } }));
    ed.run(formatParagraphs(ed, SLIDE, body(), ALL, { algn: 'ctr' }));
    ed.log.seal();
    const plain = { bold: false, italic: false, underline: 'none', size: 18, spacing: 0, baseline: 0, highlight: null };
    expect(resolved(ed, id).paragraphs[0]!.runs[0]!.props).not.toMatchObject(plain);
    expect(cycle(ed, () => clearRunFormatting(ed, SLIDE, body(), ALL))).toEqual([]);
    ed.run(clearRunFormatting(ed, SLIDE, body(), ALL));
    for (const p of resolved(ed, id).paragraphs) {
      expect(p.runs[0]!.props).toMatchObject(plain);
      expect(p.runs[0]!.props.latin).not.toBe('Georgia');
      expect(p.endProps).toMatchObject(plain);
      expect(p.props.align).toBe('ctr');                               // paragraph formatting is not character formatting
    }
    expect(runXml(body(), 0)).toBe('<a:rPr lang="en-US"/>');
    expect(runsOf(para(body())).map(r => r.text).join('')).toBe('Hello world');
    expect(ed.log.undoLabel).toBe('Clear Formatting');
  });

  test('Clear Formatting on part of a paragraph clears only that part', () => {
    const { ed, id, body } = withBox();
    ed.run(formatRuns(ed, SLIDE, body(), { start: { p: 0, offset: 0 }, end: { p: 0, offset: 11 } }, { b: true }));
    ed.run(clearRunFormatting(ed, SLIDE, body(), WORLD));
    expect(resolved(ed, id).paragraphs[0]!.runs.map(r => [r.text, r.props.bold])).toEqual([['Hello ', true], ['world', false]]);
  });
});

describe('paragraph spacing', () => {
  test('line spacing as a multiple, space before and after in points, each undone exactly', () => {
    const { ed, id, body } = withBox();
    const first: TextRange = { start: { p: 0, offset: 0 }, end: { p: 0, offset: 0 } };
    expect(cycle(ed, () => formatParagraphs(ed, SLIDE, body(), first, { lnSpc: { pct: 150000 }, spcBef: { pts: 600 }, spcAft: { pts: 1200 } }))).toEqual([]);
    ed.run(formatParagraphs(ed, SLIDE, body(), first, { lnSpc: { pct: 150000 }, spcBef: { pts: 600 }, spcAft: { pts: 1200 } }));
    expect(serializeElement(child(para(body()), 'a:pPr')!)).toBe('<a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="1200"/></a:spcAft></a:pPr>');
    const ps = resolved(ed, id).paragraphs;
    expect(ps[0]!.props.lineSpacing).toEqual({ kind: 'pct', val: 1.5 });
    expect(ps[0]!.props.spaceBefore).toEqual({ kind: 'pts', val: 600 });
    expect(ps[0]!.props.spaceAfter).toEqual({ kind: 'pts', val: 1200 });
    expect(ps[1]!.props.lineSpacing).toEqual({ kind: 'pct', val: 1 });     // the caret's paragraph only
    ed.run(formatParagraphs(ed, SLIDE, body(), first, { lnSpc: null, spcBef: null, spcAft: null }));
    expect(resolved(ed, id).paragraphs[0]!.props).toMatchObject({ lineSpacing: { kind: 'pct', val: 1 }, spaceBefore: { kind: 'pts', val: 0 } });
  });
});

describe('text box options', () => {
  test('vertical anchor, margins and columns', () => {
    const { ed, id, body } = withBox();
    const patch = { anchor: 'ctr', lIns: 182880, tIns: 91440, rIns: 182880, bIns: 91440, numCol: 2 };
    expect(cycle(ed, () => formatBody(ed, SLIDE, body(), patch))).toEqual([]);
    ed.run(formatBody(ed, SLIDE, body(), patch));
    expect(resolved(ed, id).body).toMatchObject({ anchor: 'ctr', numCol: 2, insets: { l: 182880, t: 91440, r: 182880, b: 91440 } });
    ed.run(formatBody(ed, SLIDE, body(), { anchor: 'b', numCol: null, lIns: 0, tIns: 0, rIns: 0, bIns: 0 }));
    expect(resolved(ed, id).body).toMatchObject({ anchor: 'b', numCol: 1, insets: { l: 0, t: 0, r: 0, b: 0 } });
    ed.undo();
    expect(resolved(ed, id).body).toMatchObject({ anchor: 'ctr', numCol: 2 });
  });
});

describe('typing over a selection', () => {
  test('text typed over everything looks like the first character it replaces, not the last', () => {
    const { ed, id, body } = withBox();
    ed.run(formatRuns(ed, SLIDE, body(), { start: { p: 1, offset: 0 }, end: { p: 1, offset: 6 } }, { u: 'sng', b: true }));
    ed.log.seal();
    expect(cycle(ed, () => typeText(ed, SLIDE, body(), ALL, 'New').command)).toEqual([]);
    ed.run(typeText(ed, SLIDE, body(), ALL, 'New').command);
    const p = resolved(ed, id).paragraphs;
    expect(p.length).toBe(1);
    expect(p[0]!.runs.map(r => [r.text, r.props.bold, r.props.underline])).toEqual([['New', false, 'none']]);
  });

  test('and over a selection that starts in formatted text, like that text', () => {
    const { ed, id, body } = withBox();
    ed.run(formatRuns(ed, SLIDE, body(), { start: { p: 0, offset: 0 }, end: { p: 0, offset: 11 } }, { i: true }));
    ed.run(typeText(ed, SLIDE, body(), ALL, 'New').command);
    expect(resolved(ed, id).paragraphs[0]!.runs.map(r => [r.text, r.props.italic])).toEqual([['New', true]]);
  });

  test('a selection that leaves text in the first paragraph still joins the paragraphs as before', () => {
    const { ed, id, body } = withBox();
    ed.run(typeText(ed, SLIDE, body(), { start: { p: 0, offset: 5 }, end: { p: 1, offset: 3 } }, '-').command);
    expect(resolved(ed, id).paragraphs.map(p => p.runs.map(r => r.text).join(''))).toEqual(['Hello-ond']);
  });
});

describe('correcting text through a plain copy of it', () => {
  /** "Hello world" with "world" bold, then "Second". Runs as [text, bold] per paragraph. */
  function formatted() {
    const t = withBox();
    t.ed.run(formatRuns(t.ed, SLIDE, t.body(), WORLD, { b: true }));
    t.ed.log.clear();
    return t;
  }
  const runs = (ed: Editor, id: number) => resolved(ed, id).paragraphs.map(p => p.runs.map(r => [r.text, r.props.bold]));

  test('a fix inside a word changes only that word\'s run; every other run keeps its text and look', () => {
    const { ed, id, body } = formatted();
    expect(cycle(ed, () => retypeText(ed, SLIDE, body(), 'Hello wurld\nSecond')!)).toEqual([]);
    ed.run(retypeText(ed, SLIDE, body(), 'Helo world\nSecond')!);
    expect(runs(ed, id)).toEqual([[['Helo ', false], ['world', true]], [['Second', false]]]);
    ed.run(retypeText(ed, SLIDE, body(), 'Helo wide world\nSecond')!);
    expect(runs(ed, id)).toEqual([[['Helo ', false], ['wide world', true]], [['Second', false]]]);
    expect(ed.log.undoLabel).toBe('Correct Text');
  });

  test('text that already reads so needs no command', () => {
    const { ed, body } = formatted();
    expect(retypeText(ed, SLIDE, body(), 'Hello world\nSecond')).toBeUndefined();
    expect(retypeText(ed, SLIDE, body(), 'Hello world\r\nSecond')).toBeUndefined();
  });

  test('a change across paragraphs, a paragraph added, and paragraphs joined', () => {
    const { ed, id, body } = formatted();
    ed.run(retypeText(ed, SLIDE, body(), 'Hello world\nA new line\nSecond')!);
    expect(bodyPlainText(body())).toBe('Hello world\nA new line\nSecond');
    expect(runs(ed, id)[0]).toEqual([['Hello ', false], ['world', true]]);
    ed.run(retypeText(ed, SLIDE, body(), 'Hello world Second')!);
    expect(bodyPlainText(body())).toBe('Hello world Second');
    expect(resolved(ed, id).paragraphs.length).toBe(1);
    ed.run(retypeText(ed, SLIDE, body(), '')!);
    expect(bodyPlainText(body())).toBe('');
  });

  test('the whole text replaced with nothing in common, and text added at the very start and end', () => {
    const { ed, body } = formatted();
    ed.run(retypeText(ed, SLIDE, body(), 'XYZ')!);
    expect(bodyPlainText(body())).toBe('XYZ');
    ed.run(retypeText(ed, SLIDE, body(), '>XYZ<')!);
    expect(bodyPlainText(body())).toBe('>XYZ<');
  });

  test('characters made of two code units are never cut in half', () => {
    const t = withBox();
    t.ed.run(retypeText(t.ed, SLIDE, t.body(), 'a😀b')!);
    t.ed.run(retypeText(t.ed, SLIDE, t.body(), 'a😁b')!);              // the two faces share their first code unit
    expect(bodyPlainText(t.body())).toBe('a😁b');
    const again = openEditor(t.ed.save());
    expect(bodyPlainText(again.shape(SLIDE, t.id)!.txBody!)).toBe('a😁b');
  });
});

