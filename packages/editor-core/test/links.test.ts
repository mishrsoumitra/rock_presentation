// Hyperlinks: on a stretch of text and on a whole shape, to an address and to a slide; set, changed, read, removed.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, children, PresentationDocument, runsOf } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import type { XmlElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  addSlide, deleteSlide, insertShape, insertTextBox, setShapeLink, setTextLink, shapeLink, textLinkAt, type Editor, type TextRange,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const WORLD: TextRange = { start: { p: 0, offset: 6 }, end: { p: 0, offset: 11 } };

/** A deck with a text box "Hello world" / "Second" on its slide; `rels` counts the slide's relationships. */
function withBox(): { ed: Editor; id: number; body: () => XmlElement; rels: () => number } {
  const ed = openEditor(deckBytes());
  const r = insertTextBox(ed, SLIDE, { x: 1000000, y: 1000000, w: 4000000, h: 800000 }, 'Hello world\nSecond');
  ed.run(r.command);
  ed.log.clear();
  return { ed, id: r.id, body: () => ed.shape(SLIDE, r.id)!.txBody!, rels: () => ed.pkg.relationships(SLIDE).length };
}
/** What the resolver reads for each run of the box's first paragraph: its text and where its link goes. */
function linksDrawn(ed: Editor, id: number): [string, string | undefined][] {
  const text = resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.text!;
  return text.paragraphs[0]!.runs.map(r => [r.text, r.props.hyperlink?.url ?? r.props.hyperlink?.slide]);
}

describe('links on text', () => {
  test('a link to an address: the words are linked, the address is a relationship outside the file, and undo is exact', () => {
    const { ed, id, body, rels } = withBox();
    const before = rels();
    expect(cycle(ed, () => setTextLink(ed, SLIDE, body(), WORLD, { url: 'https://example.com/a?b=1&c=2' }))).toEqual([]);
    ed.run(setTextLink(ed, SLIDE, body(), WORLD, { url: 'https://example.com/a?b=1&c=2' }));
    expect(linksDrawn(ed, id)).toEqual([['Hello ', undefined], ['world', 'https://example.com/a?b=1&c=2']]);
    expect(rels()).toBe(before + 1);
    const rel = ed.pkg.relationships(SLIDE).find(r => r.external && r.target.startsWith('https://example.com'))!;
    expect(rel.external).toBe(true);
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Hyperlink');
    // the saved file opens with the link
    const back = new PresentationDocument(OpcPackage.open(ed.save()));
    expect(back.pkg.relationships(SLIDE).some(r => r.external && r.target === 'https://example.com/a?b=1&c=2')).toBe(true);
  });

  test('a link across formatted runs uses one relationship; a tooltip is kept', () => {
    const { ed, id, body, rels } = withBox();
    const before = rels();
    ed.run(setTextLink(ed, SLIDE, body(), { start: { p: 0, offset: 0 }, end: { p: 0, offset: 5 } }, { url: 'https://one.example' }));
    ed.run(setTextLink(ed, SLIDE, body(), { start: { p: 0, offset: 2 }, end: { p: 0, offset: 9 } }, { url: 'https://two.example', tooltip: 'Go "there"' }));
    expect(linksDrawn(ed, id)).toEqual([['He', 'https://one.example'], ['llo', 'https://two.example'], [' wor', 'https://two.example'], ['ld', undefined]]);
    expect(rels()).toBe(before + 2);
    expect(textLinkAt(ed, SLIDE, body(), { p: 0, offset: 4 })).toEqual({ range: { start: { p: 0, offset: 2 }, end: { p: 0, offset: 9 } }, target: { url: 'https://two.example', tooltip: 'Go "there"' } });
  });

  test('changing a link replaces it and drops the old relationship; removing it leaves the text and no relationship', () => {
    const { ed, id, body, rels } = withBox();
    const before = rels();
    ed.run(setTextLink(ed, SLIDE, body(), WORLD, { url: 'https://old.example' }));
    ed.run(setTextLink(ed, SLIDE, body(), WORLD, { url: 'https://new.example' }));
    expect(rels()).toBe(before + 1);
    expect(ed.pkg.relationships(SLIDE).some(r => r.target === 'https://old.example')).toBe(false);
    expect(cycle(ed, () => setTextLink(ed, SLIDE, body(), WORLD, null))).toEqual([]);
    ed.run(setTextLink(ed, SLIDE, body(), WORLD, null));
    expect(linksDrawn(ed, id)).toEqual([['Hello ', undefined], ['world', undefined]]);
    expect(rels()).toBe(before);
    expect(runsOf(children(body(), 'a:p')[0]!).map(r => r.text).join('')).toBe('Hello world');
    expect(ed.log.undoLabel).toBe('Remove Hyperlink');
  });

  test('removing part of a link leaves the rest linked, and its relationship in place', () => {
    const { ed, id, body, rels } = withBox();
    const before = rels();
    ed.run(setTextLink(ed, SLIDE, body(), { start: { p: 0, offset: 0 }, end: { p: 0, offset: 11 } }, { url: 'https://example.com' }));
    ed.run(setTextLink(ed, SLIDE, body(), WORLD, null));
    expect(linksDrawn(ed, id)).toEqual([['Hello ', 'https://example.com'], ['world', undefined]]);
    expect(rels()).toBe(before + 1);
  });

  test('a link to a slide of the deck; deleting that slide takes the link away', () => {
    const { ed, id, body } = withBox();
    const second = addSlide(ed, LAYOUT);
    ed.run(second.command);
    expect(cycle(ed, () => setTextLink(ed, SLIDE, body(), WORLD, { slide: second.part }))).toEqual([]);
    ed.run(setTextLink(ed, SLIDE, body(), WORLD, { slide: second.part }));
    expect(linksDrawn(ed, id)).toEqual([['Hello ', undefined], ['world', second.part]]);
    expect(textLinkAt(ed, SLIDE, body(), { p: 0, offset: 8 })?.target).toEqual({ slide: second.part });
    expect(() => setTextLink(ed, SLIDE, body(), WORLD, { slide: '/ppt/slides/slide99.xml' }).do()).toThrow(/not in this presentation/);
    ed.run(deleteSlide(ed, second.part));
    expect(linksDrawn(ed, id).map(l => l[1])).toEqual([undefined, undefined]);
  });

  test('the link at a caret: inside it, at its edges, and not next to it', () => {
    const { ed, body } = withBox();
    ed.run(setTextLink(ed, SLIDE, body(), WORLD, { url: 'https://example.com' }));
    const range = { start: { p: 0, offset: 6 }, end: { p: 0, offset: 11 } };
    for (const offset of [6, 7, 11]) expect(textLinkAt(ed, SLIDE, body(), { p: 0, offset })?.range, `offset ${offset}`).toEqual(range);
    expect(textLinkAt(ed, SLIDE, body(), { p: 0, offset: 3 })).toBeUndefined();
    expect(textLinkAt(ed, SLIDE, body(), { p: 1, offset: 2 })).toBeUndefined();
    expect(textLinkAt(ed, SLIDE, body(), { p: 7, offset: 0 })).toBeUndefined();
  });
});

describe('links on shapes', () => {
  test('a shape is linked whole; the link is read back, changed and removed', () => {
    const ed = openEditor(deckBytes());
    const r = insertShape(ed, SLIDE, 'rect', { x: 1000000, y: 1000000, w: 1000000, h: 500000 });
    ed.run(r.command);
    ed.log.clear();
    const shape = () => ed.shape(SLIDE, r.id)!;
    const before = ed.pkg.relationships(SLIDE).length;
    expect(shapeLink(ed, SLIDE, shape())).toBeUndefined();
    expect(cycle(ed, () => setShapeLink(ed, SLIDE, [shape()], { url: 'mailto:someone@example.com' }))).toEqual([]);
    ed.run(setShapeLink(ed, SLIDE, [shape()], { url: 'mailto:someone@example.com' }));
    expect(shapeLink(ed, SLIDE, shape())).toEqual({ url: 'mailto:someone@example.com' });
    expect(resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === r.id && s.layer === 'slide')!.hyperlink?.url).toBe('mailto:someone@example.com');
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    ed.run(setShapeLink(ed, SLIDE, [shape()], { url: 'https://example.com' }));
    expect(ed.pkg.relationships(SLIDE).length).toBe(before + 1);
    ed.run(setShapeLink(ed, SLIDE, [shape()], null));
    expect(shapeLink(ed, SLIDE, shape())).toBeUndefined();
    expect(ed.pkg.relationships(SLIDE).length).toBe(before);
  });
});
