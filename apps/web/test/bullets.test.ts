// The bullet and number styles of the ribbon: what each writes, and the indent that goes on and off with them.
import { child, children } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { formatParagraphs, insertTextBox, type Editor } from '../../../packages/editor-core/src/index.ts';
import { openEditor } from '../../../packages/editor-core/test/helpers.ts';
import { deckBytes } from '../../../packages/resolve/test/deck.ts';
import { BULLET_STYLES, bulletPatch, HANGING, NUMBER_STYLES, numberSample } from '../src/ui/bullets.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const ALL = { start: { p: 0, offset: 0 }, end: { p: 2, offset: 0 } };

function withBox(): { ed: Editor; id: number } {
  const ed = openEditor(deckBytes());
  const r = insertTextBox(ed, SLIDE, { x: 1000000, y: 1000000, w: 4000000, h: 1200000 }, 'one\ntwo\nthree');
  ed.run(r.command);
  return { ed, id: r.id };
}
const body = (ed: Editor, id: number) => ed.shape(SLIDE, id)!.txBody!;
const pPr = (ed: Editor, id: number, i = 0) => serializeElement(child(children(body(ed, id), 'a:p')[i]!, 'a:pPr')!);
const resolved = (ed: Editor, id: number) => resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.text!;

describe('the indent that goes with bullets', () => {
  test('on, in text with no indent: a hanging indent, wider for numbers; off: the indent goes again', () => {
    expect(bulletPatch({ kind: 'char', char: '•' }, { marL: 0, indent: 0 })).toEqual({ bullet: { kind: 'char', char: '•' }, marL: HANGING.char, indent: -HANGING.char });
    expect(bulletPatch({ kind: 'autonum', scheme: 'arabicPeriod' }, undefined)).toMatchObject({ marL: HANGING.autonum, indent: -HANGING.autonum });
    expect(bulletPatch({ kind: 'none' }, { marL: HANGING.char, indent: -HANGING.char })).toEqual({ bullet: { kind: 'none' }, marL: 0, indent: 0 });
    expect(bulletPatch({ kind: 'none' }, { marL: HANGING.autonum, indent: -HANGING.autonum })).toMatchObject({ marL: 0, indent: 0 });
  });

  test('from bullets to numbers and back, the indent follows; another bullet character changes nothing else', () => {
    expect(bulletPatch({ kind: 'autonum', scheme: 'romanUcPeriod' }, { marL: HANGING.char, indent: -HANGING.char })).toMatchObject({ marL: HANGING.autonum, indent: -HANGING.autonum });
    expect(bulletPatch({ kind: 'char', char: '■' }, { marL: HANGING.autonum, indent: -HANGING.autonum })).toMatchObject({ marL: HANGING.char, indent: -HANGING.char });
    expect(bulletPatch({ kind: 'char', char: '■' }, { marL: HANGING.char, indent: -HANGING.char })).toEqual({ bullet: { kind: 'char', char: '■' } });
  });

  test("a placeholder's or a hand-set indent is never touched", () => {
    const own = { marL: 457200, indent: -457200 };
    expect(bulletPatch({ kind: 'char', char: '•' }, own)).toEqual({ bullet: { kind: 'char', char: '•' } });
    expect(bulletPatch({ kind: 'autonum', scheme: 'arabicPeriod' }, own)).toEqual({ bullet: { kind: 'autonum', scheme: 'arabicPeriod' } });
    expect(bulletPatch({ kind: 'none' }, own)).toEqual({ bullet: { kind: 'none' } });
    expect(bulletPatch({ kind: 'none' }, { marL: 285750, indent: 0 })).toEqual({ bullet: { kind: 'none' } });
  });
});

describe('bullet styles', () => {
  test('every style is one character, all different, and each is written and read back as chosen', () => {
    expect(new Set(BULLET_STYLES.map(b => b.char)).size).toBe(BULLET_STYLES.length);
    for (const style of BULLET_STYLES) {
      expect([...style.char].length).toBe(1);
      const { ed, id } = withBox();
      ed.run(formatParagraphs(ed, SLIDE, body(ed, id), ALL, bulletPatch({ kind: 'char', char: style.char, font: 'Arial' }, { marL: 0, indent: 0 })));
      expect(pPr(ed, id)).toBe(`<a:pPr marL="285750" indent="-285750"><a:buFont typeface="Arial"/><a:buChar char="${style.char}"/></a:pPr>`);
      const ps = resolved(ed, id).paragraphs;
      expect(ps.map(p => p.props.bullet?.char)).toEqual([style.char, style.char, style.char]);
      expect(ps[0]!.props).toMatchObject({ marL: 285750, indent: -285750 });
    }
  });

  test('None writes a:buNone and puts the text back at the edge', () => {
    const { ed, id } = withBox();
    ed.run(formatParagraphs(ed, SLIDE, body(ed, id), ALL, bulletPatch({ kind: 'char', char: '•', font: 'Arial' }, { marL: 0, indent: 0 })));
    ed.run(formatParagraphs(ed, SLIDE, body(ed, id), ALL, bulletPatch({ kind: 'none' }, resolved(ed, id).paragraphs[0]!.props)));
    expect(pPr(ed, id)).toBe('<a:pPr marL="0" indent="0"><a:buFont typeface="Arial"/><a:buNone/></a:pPr>');
    expect(resolved(ed, id).paragraphs[0]!.props.bullet).toBeNull();
  });
});

describe('number styles', () => {
  test('the samples show how each scheme counts', () => {
    expect(NUMBER_STYLES.map(numberSample)).toEqual(['1. 2. 3.', '1) 2) 3)', 'I. II. III.', 'i. ii. iii.', 'A. B. C.', 'a. b. c.', 'a) b) c)']);
  });

  test('each scheme is written as chosen, on every paragraph, with the wider indent numbers get', () => {
    for (const scheme of NUMBER_STYLES) {
      const { ed, id } = withBox();
      ed.run(formatParagraphs(ed, SLIDE, body(ed, id), ALL, bulletPatch({ kind: 'autonum', scheme }, { marL: 0, indent: 0 })));
      expect(pPr(ed, id, 2)).toBe(`<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="${scheme}"/></a:pPr>`);
      expect(resolved(ed, id).paragraphs.map(p => p.props.bullet?.scheme)).toEqual([scheme, scheme, scheme]);
    }
  });
});
