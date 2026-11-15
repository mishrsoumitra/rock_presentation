// Sections: starting one at a slide, renaming, removing; and what adding, deleting and moving slides does to them.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, PresentationDocument } from '@pptx/model';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { addSection, addSlide, deleteSlide, duplicateSlide, moveSlide, removeAllSections, removeSection, renameSection, sections, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';

/** A deck of five slides; `n(part)` is a slide's number, so sections read as lists of slide numbers. */
function five(): { ed: Editor; slides: string[] } {
  const ed = openEditor(deckBytes());
  for (let i = 0; i < 4; i++) ed.run(addSlide(ed, LAYOUT).command);
  ed.log.clear();
  return { ed, slides: ed.doc.slideParts() };
}
/** The sections as "Name: 1 2 3" lines, slides by their original numbers. */
const shown = (ed: Editor, original: string[]) => sections(ed).map(s => `${s.name}: ${s.slides.map(p => original.indexOf(p) + 1 || 'new').join(' ')}`);

describe('adding sections', () => {
  test('a deck has no sections until one is added; the first puts the slides before it in "Default Section"', () => {
    const { ed, slides } = five();
    expect(sections(ed)).toEqual([]);
    expect(cycle(ed, () => addSection(ed, slides[2]!, 'Results'))).toEqual([]);
    ed.run(addSection(ed, slides[2]!, 'Results'));
    expect(shown(ed, slides)).toEqual(['Default Section: 1 2', 'Results: 3 4 5']);
    expect(sections(ed)[1]!.id).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
    expect(checkOrder(ed.presentationRoot)).toEqual([]);
    expect(ed.log.undoLabel).toBe('Add Section');
  });

  test('a section started at the first slide takes the whole deck; a second one splits the section it starts in', () => {
    const { ed, slides } = five();
    ed.run(addSection(ed, slides[0]!, 'Intro'));
    expect(shown(ed, slides)).toEqual(['Intro: 1 2 3 4 5']);
    ed.run(addSection(ed, slides[3]!, 'Close'));
    ed.run(addSection(ed, slides[1]!));
    expect(shown(ed, slides)).toEqual(['Intro: 1', 'Untitled Section: 2 3', 'Close: 4 5']);
    expect(new Set(sections(ed).map(s => s.id)).size).toBe(3);
    expect(() => addSection(ed, '/ppt/slides/slide99.xml').do()).toThrow(/not in this presentation/);
  });

  test('sections are in the saved file, where PowerPoint reads them', () => {
    const { ed, slides } = five();
    ed.run(addSection(ed, slides[2]!, 'Results & "more"'));
    const back = new PresentationDocument(OpcPackage.open(ed.save()));
    expect(back.presentation.sections.map(s => [s.name, s.slideIds.length])).toEqual([['Default Section', 2], ['Results & "more"', 3]]);
  });
});

describe('renaming and removing', () => {
  function three(): { ed: Editor; slides: string[] } {
    const t = five();
    t.ed.run(addSection(t.ed, t.slides[0]!, 'A'));
    t.ed.run(addSection(t.ed, t.slides[2]!, 'B'));
    t.ed.run(addSection(t.ed, t.slides[4]!, 'C'));
    t.ed.log.clear();
    return t;
  }

  test('a section is renamed by its place', () => {
    const { ed, slides } = three();
    expect(cycle(ed, () => renameSection(ed, 1, 'Middle'))).toEqual([]);
    ed.run(renameSection(ed, 1, 'Middle'));
    expect(shown(ed, slides)).toEqual(['A: 1 2', 'Middle: 3 4', 'C: 5']);
    expect(() => renameSection(ed, 9, 'x').do()).toThrow(/no such section/);
  });

  test('a removed section\'s slides join the section before it; the first section\'s join the one after', () => {
    const { ed, slides } = three();
    expect(cycle(ed, () => removeSection(ed, 1))).toEqual([]);
    ed.run(removeSection(ed, 1));
    expect(shown(ed, slides)).toEqual(['A: 1 2 3 4', 'C: 5']);
    ed.run(removeSection(ed, 0));
    expect(shown(ed, slides)).toEqual(['C: 1 2 3 4 5']);
    ed.run(removeSection(ed, 0));
    expect(sections(ed)).toEqual([]);
    expect(ed.doc.slideParts()).toEqual(slides);                          // the slides were never touched
    expect(checkOrder(ed.presentationRoot)).toEqual([]);
  });

  test('Remove All Sections leaves the slides and no sections, in one step', () => {
    const { ed, slides } = three();
    expect(cycle(ed, () => removeAllSections(ed))).toEqual([]);
    ed.run(removeAllSections(ed));
    expect(sections(ed)).toEqual([]);
    ed.undo();
    expect(shown(ed, slides)).toEqual(['A: 1 2', 'B: 3 4', 'C: 5']);
  });
});

describe('slides and their sections', () => {
  function two(): { ed: Editor; slides: string[] } {
    const t = five();
    t.ed.run(addSection(t.ed, t.slides[0]!, 'A'));
    t.ed.run(addSection(t.ed, t.slides[3]!, 'B'));
    t.ed.log.clear();
    return t;
  }

  test('a new or duplicated slide joins the section of the slide before it; a deleted slide leaves its section', () => {
    const { ed, slides } = two();
    ed.run(addSlide(ed, LAYOUT, 2).command);
    expect(shown(ed, slides)).toEqual(['A: 1 2 new 3', 'B: 4 5']);
    ed.run(duplicateSlide(ed, slides[4]!).command);
    expect(shown(ed, slides)).toEqual(['A: 1 2 new 3', 'B: 4 5 new']);
    ed.run(deleteSlide(ed, slides[3]!));
    expect(shown(ed, slides)).toEqual(['A: 1 2 new 3', 'B: 5 new']);
  });

  test('a moved slide joins the section where it lands', () => {
    const { ed, slides } = two();
    ed.run(moveSlide(ed, slides[0]!, 4));
    expect(shown(ed, slides)).toEqual(['A: 2 3', 'B: 4 5 1']);
    ed.run(moveSlide(ed, slides[4]!, 0));
    expect(shown(ed, slides)[0]).toBe('A: 5 2 3');
  });
});
