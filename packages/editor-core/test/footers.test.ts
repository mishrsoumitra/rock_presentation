// Header & Footer: the date, the footer text and the slide number as placeholders on a slide, and on every slide.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, PresentationDocument, setAttr } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { addSlide, headerFooter, setHeaderFooter, setHeaderFooterAll, type Editor, type HeaderFooter } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const NONE: HeaderFooter = { date: null, footer: null, slideNumber: false };
const ph = (ed: Editor, part: string, type: string) => ed.doc.slide(part).shapes.filter(s => s.placeholder?.type === type);
const xml = (ed: Editor, part: string, type: string) => serializeElement(ph(ed, part, type)[0]!.el);

describe('what a slide shows at its foot', () => {
  test('the test deck\'s slide shows its number and nothing else', () => {
    const ed = openEditor(deckBytes());
    expect(headerFooter(ed, SLIDE)).toEqual({ date: null, footer: null, slideNumber: true });
  });

  test('a fixed date and a footer are added as placeholders along the bottom; undone exactly', () => {
    const ed = openEditor(deckBytes());
    const spec: HeaderFooter = { date: { auto: false, text: '16 November 2026' }, footer: 'Quarterly review <draft>', slideNumber: true };
    expect(cycle(ed, () => setHeaderFooter(ed, SLIDE, spec))).toEqual([]);
    ed.run(setHeaderFooter(ed, SLIDE, spec));
    expect(headerFooter(ed, SLIDE)).toEqual(spec);
    expect(ph(ed, SLIDE, 'dt').length).toBe(1);
    expect(ph(ed, SLIDE, 'ftr').length).toBe(1);
    expect(ph(ed, SLIDE, 'ftr')[0]!.name).toMatch(/^Footer Placeholder \d+$/);
    expect(xml(ed, SLIDE, 'ftr')).toContain('<a:t>Quarterly review &lt;draft&gt;</a:t>');
    // the deck's layout has no place for them, so each has its own: the date at the left, the footer in the middle
    const drawn = resolveSlide(ed.doc, SLIDE).shapes.filter(s => s.layer === 'slide');
    const date = drawn.find(s => s.placeholder?.type === 'dt')!, footer = drawn.find(s => s.placeholder?.type === 'ftr')!;
    expect(date.xfrm.x).toBeLessThan(footer.xfrm.x);
    expect(date.xfrm.y).toBeGreaterThan(6858000 * 0.9);
    expect(footer.xfrm.x + footer.xfrm.cx / 2).toBeCloseTo(12192000 / 2, -2);
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Header & Footer');
  });

  test('a date kept up to date is a field with the text it showed last; made fixed, it is plain text again', () => {
    const ed = openEditor(deckBytes());
    ed.run(setHeaderFooter(ed, SLIDE, { ...NONE, date: { auto: true, text: '11/16/2026' } }));
    expect(xml(ed, SLIDE, 'dt')).toMatch(/<a:fld id="\{[0-9A-F-]{36}\}" type="datetime1"><a:rPr[^>]*\/><a:t>11\/16\/2026<\/a:t><\/a:fld>/);
    expect(headerFooter(ed, SLIDE)).toEqual({ ...NONE, date: { auto: true, text: '11/16/2026' } });
    expect(cycle(ed, () => setHeaderFooter(ed, SLIDE, { ...NONE, date: { auto: false, text: 'Autumn term' } }))).toEqual([]);
    ed.run(setHeaderFooter(ed, SLIDE, { ...NONE, date: { auto: false, text: 'Autumn term' } }));
    expect(xml(ed, SLIDE, 'dt')).not.toContain('a:fld');
    expect(headerFooter(ed, SLIDE).date).toEqual({ auto: false, text: 'Autumn term' });
    expect(ph(ed, SLIDE, 'dt').length).toBe(1);
  });

  test('the number is removed, and comes back as a field that takes its place from the layout', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => setHeaderFooter(ed, SLIDE, NONE))).toEqual([]);
    ed.run(setHeaderFooter(ed, SLIDE, NONE));
    expect(ph(ed, SLIDE, 'sldNum')).toEqual([]);
    ed.run(setHeaderFooter(ed, SLIDE, { ...NONE, slideNumber: true }));
    const num = xml(ed, SLIDE, 'sldNum');
    expect(num).toContain('<p:ph type="sldNum" sz="quarter" idx="12"/>');   // as the layout's
    expect(num).toContain('<p:spPr/>');
    expect(num).toMatch(/<a:fld id="\{[0-9A-F-]{36}\}" type="slidenum">/);
    const layoutBox = resolveSlide(ed.doc, SLIDE).shapes.find(s => s.layer === 'slide' && s.placeholder?.type === 'sldNum')!.xfrm;
    expect(layoutBox.cx).toBeGreaterThan(0);
  });

  test('changed footer text replaces the old, keeping the look of its run; the same again changes nothing', () => {
    const ed = openEditor(deckBytes());
    ed.run(setHeaderFooter(ed, SLIDE, { ...NONE, footer: 'First' }));
    const run = ph(ed, SLIDE, 'ftr')[0]!.el;
    const rPr = serializeElement(run).match(/<a:rPr[^>]*\/>/)![0];
    ed.run(setHeaderFooter(ed, SLIDE, { ...NONE, footer: 'Second' }));
    expect(ph(ed, SLIDE, 'ftr').length).toBe(1);
    expect(xml(ed, SLIDE, 'ftr')).toContain(`${rPr}<a:t>Second</a:t>`);
    expect(xml(ed, SLIDE, 'ftr')).not.toContain('First');
    const before = serializeElement(ed.doc.root(SLIDE));
    ed.run(setHeaderFooter(ed, SLIDE, { ...NONE, footer: 'Second' }));
    expect(serializeElement(ed.doc.root(SLIDE))).toBe(before);
    // an empty footer is still a footer: the placeholder stays, with no text
    ed.run(setHeaderFooter(ed, SLIDE, { ...NONE, footer: '' }));
    expect(headerFooter(ed, SLIDE).footer).toBe('');
  });
});

describe('Apply to All', () => {
  function threeSlides(): { ed: Editor; parts: string[] } {
    const ed = openEditor(deckBytes());
    ed.run(addSlide(ed, LAYOUT).command);
    ed.run(addSlide(ed, LAYOUT).command);
    ed.log.clear();
    return { ed, parts: ed.doc.slideParts() };
  }

  test('every slide shows the same, as one undo step', () => {
    const { ed, parts } = threeSlides();
    const spec: HeaderFooter = { date: null, footer: 'Team offsite', slideNumber: true };
    expect(cycle(ed, () => setHeaderFooterAll(ed, spec))).toEqual([]);
    ed.run(setHeaderFooterAll(ed, spec));
    expect(parts.map(p => headerFooter(ed, p))).toEqual([spec, spec, spec]);
    ed.undo();
    expect(parts.map(p => headerFooter(ed, p).footer)).toEqual([null, null, null]);
    expect(ed.log.redoLabel).toBe('Header & Footer');
  });

  test('"Don\'t show on title slide" leaves slides on a title layout with nothing', () => {
    const { ed, parts } = threeSlides();
    setAttr(ed.doc.root(LAYOUT), 'type', 'title');                       // every slide of this deck is now a title slide
    ed.run(setHeaderFooterAll(ed, { date: null, footer: 'x', slideNumber: true }, true));
    expect(parts.map(p => headerFooter(ed, p))).toEqual([NONE, NONE, NONE]);
    ed.run(setHeaderFooterAll(ed, { date: null, footer: 'x', slideNumber: true }, false));
    expect(parts.map(p => headerFooter(ed, p).footer)).toEqual(['x', 'x', 'x']);
  });

  test('a saved deck opens showing what was set', () => {
    const { ed, parts } = threeSlides();
    const spec: HeaderFooter = { date: { auto: true, text: '16/11/2026' }, footer: 'Kept', slideNumber: true };
    ed.run(setHeaderFooterAll(ed, spec));
    const back = openEditor(ed.save());
    expect(parts.map(p => headerFooter(back, p))).toEqual([spec, spec, spec]);
    expect(new PresentationDocument(OpcPackage.open(ed.save())).slide(parts[2]!).shapes.filter(s => s.placeholder?.type === 'ftr').length).toBe(1);
  });
});
