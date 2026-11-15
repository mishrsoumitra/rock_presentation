// WordArt: text painted with its own fill, outline and shadow; inserted as a text box, or put on text that is there.
import { checkOrder } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { formatRuns, insertTextBox, insertWordArt, WORDART_STYLES, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const BOX = { x: 1000000, y: 2000000, w: 6000000, h: 900000 };
const style = (id: string) => WORDART_STYLES.find(s => s.id === id)!;
const runProps = (ed: Editor, id: number) => resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.text!.paragraphs[0]!.runs[0]!.props;
const xml = (ed: Editor, id: number) => serializeElement(ed.shape(SLIDE, id)!.el);

describe('the gallery', () => {
  test('six looks, each with an id and a name of its own', () => {
    expect(WORDART_STYLES.length).toBe(6);
    expect(new Set(WORDART_STYLES.map(s => s.id)).size).toBe(6);
    expect(new Set(WORDART_STYLES.map(s => s.name)).size).toBe(6);
  });
});

describe('inserting WordArt', () => {
  test('a text box of large bold text that does not wrap, in the look chosen; undone exactly', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => insertWordArt(ed, SLIDE, style('fill-accent1-shadow'), BOX).command)).toEqual([]);
    const w = insertWordArt(ed, SLIDE, style('fill-accent1-shadow'), BOX);
    ed.run(w.command);
    const x = xml(ed, w.id);
    expect(x).toContain('<p:cNvSpPr txBox="1"/>');
    expect(x).toContain('<a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr>');
    expect(x).toContain('<a:t>Your text here</a:t>');
    expect(x).toMatch(/<a:rPr lang="en-US" sz="5400" b="1"><a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:effectLst><a:outerShdw /);
    // the paragraph mark looks the same, so text typed after it does too
    expect(x).toMatch(/<a:endParaRPr lang="en-US" sz="5400" b="1"><a:solidFill><a:schemeClr val="accent1"\/>/);
    const props = runProps(ed, w.id);
    expect(props.size).toBe(54);
    expect(props.bold).toBe(true);
    expect(props.fill).toMatchObject({ kind: 'solid', color: { r: 0x44 / 255, g: 0x72 / 255, b: 0xc4 / 255 } });
    expect(props.effects).toBeDefined();
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Insert WordArt');
  });

  test('each look resolves to what it says: an outline with no fill, a gradient, a fill with an outline', () => {
    const ed = openEditor(deckBytes());
    const insert = (id: string, text?: string) => { const w = insertWordArt(ed, SLIDE, style(id), BOX, text); ed.run(w.command); return w.id; };
    const hollow = runProps(ed, insert('outline-accent2', 'Hollow & <bold>'));
    expect(hollow.fill.kind).toBe('none');
    expect(hollow.outline).toMatchObject({ width: 22225, paint: { kind: 'solid' } });
    expect(hollow.effects).toBeUndefined();
    const grad = runProps(ed, insert('gradient-accent4'));
    expect(grad.fill.kind).toBe('linear');
    expect(grad.outline).toBeNull();
    const both = runProps(ed, insert('fill-light-outline-accent1'));
    expect(both.fill).toMatchObject({ kind: 'solid', color: { r: 1, g: 1, b: 1 } });
    expect(both.outline).not.toBeNull();
    expect(both.effects).toBeDefined();
    for (const s of WORDART_STYLES) expect(checkOrder(ed.doc.root(SLIDE)), s.id).toEqual([]);
    expect(xml(ed, ed.doc.slide(SLIDE).shapes.at(-3)!.id)).toContain('<a:t>Hollow &amp; &lt;bold&gt;</a:t>');
  });
});

describe('a WordArt look on text that is there', () => {
  function withText(): { ed: Editor; id: number } {
    const ed = openEditor(deckBytes());
    const t = insertTextBox(ed, SLIDE, BOX, 'plain words');
    ed.run(t.command);
    ed.log.clear();
    return { ed, id: t.id };
  }
  const body = (ed: Editor, id: number) => ed.shape(SLIDE, id)!.txBody!;
  const WORD = { start: { p: 0, offset: 0 }, end: { p: 0, offset: 5 } };

  test('the look goes on the range only, in schema order (outline, fill, effects); undone exactly', () => {
    const { ed, id } = withText();
    const effect = style('fill-accent5-outline-light').effect;
    expect(cycle(ed, () => formatRuns(ed, SLIDE, body(ed, id), WORD, { effect }, { label: 'WordArt Style' }))).toEqual([]);
    ed.run(formatRuns(ed, SLIDE, body(ed, id), WORD, { effect }, { label: 'WordArt Style' }));
    const x = xml(ed, id);
    expect(x).toMatch(/<a:rPr[^>]*><a:ln w="9525"><a:solidFill><a:schemeClr val="bg1"\/><\/a:solidFill><a:prstDash val="solid"\/><\/a:ln><a:solidFill><a:schemeClr val="accent5"\/><\/a:solidFill><a:effectLst>.*?<\/a:effectLst><\/a:rPr><a:t>plain<\/a:t>/);
    expect(x).toMatch(/<a:t> words<\/a:t>/);
    expect(x.match(/<a:ln /g)?.length).toBe(1);
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('WordArt Style');
  });

  test('a new look replaces the old one whole; clearing takes fill, outline and shadow off and leaves size and bold', () => {
    const { ed, id } = withText();
    const all = { start: { p: 0, offset: 0 }, end: { p: 0, offset: 11 } };
    ed.run(formatRuns(ed, SLIDE, body(ed, id), all, { b: true, sz: 4000, effect: style('fill-light-outline-accent1').effect }));
    ed.run(formatRuns(ed, SLIDE, body(ed, id), all, { effect: style('gradient-accent4').effect }));
    let x = xml(ed, id);
    expect(x).toContain('<a:gradFill>');
    expect(x).not.toContain('<a:ln ');
    expect(x).not.toContain('<a:effectLst>');
    expect(x).not.toContain('<a:solidFill><a:schemeClr val="bg1"/>');
    expect(cycle(ed, () => formatRuns(ed, SLIDE, body(ed, id), all, { effect: null }))).toEqual([]);
    ed.run(formatRuns(ed, SLIDE, body(ed, id), all, { effect: null }));
    x = xml(ed, id);
    expect(x).not.toContain('gradFill');
    expect(x).toContain('sz="4000"');
    expect(x).toContain('b="1"');
    expect(runProps(ed, id).fill.kind).toBe('solid');                     // the deck's own text colour again
  });

  test('a saved deck opens with the look', () => {
    const ed = openEditor(deckBytes());
    const w = insertWordArt(ed, SLIDE, style('gradient-accent4'), BOX, 'Kept');
    ed.run(w.command);
    const back = openEditor(ed.save());
    expect(runProps(back, w.id).fill.kind).toBe('linear');
    expect(serializeElement(back.shape(SLIDE, w.id)!.el)).toContain('<a:lin ang="5400000" scaled="0"/>');
  });
});
