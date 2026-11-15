// Speaker notes: reading them, writing them on a deck that has no notes pages yet, and what slides do with them.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, CONTENT_TYPE, PresentationDocument } from '@pptx/model';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { addSlide, deleteSlide, duplicateSlide, notesText, setNotes, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';

const notesParts = (ed: Editor) => ed.pkg.partNames().filter(n => n.startsWith('/ppt/notesSlides/') && !n.endsWith('.rels'));
const masterParts = (ed: Editor) => ed.pkg.partNames().filter(n => n.startsWith('/ppt/notesMasters/') && !n.endsWith('.rels'));

describe('speaker notes', () => {
  test('a slide without a notes page reads as empty', () => {
    const ed = openEditor(deckBytes());
    expect(ed.doc.notesOf(SLIDE)).toBeUndefined();
    expect(notesText(ed, SLIDE)).toBe('');
  });

  test('the first notes of a deck: a notes page, a notes master with its own theme, all undone exactly', () => {
    const ed = openEditor(deckBytes());
    const themes = ed.pkg.partNames().filter(n => n.startsWith('/ppt/theme/')).length;
    expect(cycle(ed, () => setNotes(ed, SLIDE, 'Say hello'))).toEqual([]);
    expect(notesParts(ed)).toEqual([]);                                    // cycle leaves the deck as it was

    ed.run(setNotes(ed, SLIDE, 'Say hello'));
    expect(notesText(ed, SLIDE)).toBe('Say hello');
    const notes = ed.doc.notesOf(SLIDE)!;
    const master = ed.doc.notesMaster()!;
    expect(ed.pkg.getPart(notes)!.contentType).toBe(CONTENT_TYPE.notesSlide);
    expect(ed.pkg.getPart(master)!.contentType).toBe(CONTENT_TYPE.notesMaster);
    // the notes page points at the master and back at its slide; the master has a theme no slide master uses
    expect(ed.doc.targets(notes, 'notesMaster')).toEqual([master]);
    expect(ed.doc.targets(notes, 'slide')).toEqual([SLIDE]);
    const theme = ed.doc.targets(master, 'theme')[0]!;
    expect(ed.pkg.partNames().filter(n => n.startsWith('/ppt/theme/')).length).toBe(themes + 1);
    expect(ed.doc.masterParts().map(m => ed.doc.themeOf(m))).not.toContain(theme);
    expect(checkOrder(ed.presentationRoot)).toEqual([]);
    expect(checkOrder(ed.doc.root(master))).toEqual([]);
    expect(checkOrder(ed.doc.root(notes))).toEqual([]);
  });

  test('notes survive a save: the file opens again with the same text', () => {
    const ed = openEditor(deckBytes());
    ed.run(setNotes(ed, SLIDE, 'First point\nSecond point'));
    const back = openEditor(ed.save());
    expect(notesText(back, SLIDE)).toBe('First point\nSecond point');
    expect(new PresentationDocument(OpcPackage.open(ed.save())).notesMaster()).toBeDefined();
  });

  test('a second slide with notes shares the one notes master', () => {
    const ed = openEditor(deckBytes());
    const second = addSlide(ed, LAYOUT);
    ed.run(second.command);
    ed.run(setNotes(ed, SLIDE, 'one'));
    ed.run(setNotes(ed, second.part, 'two'));
    expect(notesParts(ed).length).toBe(2);
    expect(masterParts(ed).length).toBe(1);
    expect([notesText(ed, SLIDE), notesText(ed, second.part)]).toEqual(['one', 'two']);
  });

  test('changing notes that exist: lines become paragraphs, the page and master stay, undo is exact', () => {
    const ed = openEditor(deckBytes());
    ed.run(setNotes(ed, SLIDE, 'old'));
    ed.log.seal();
    expect(cycle(ed, () => setNotes(ed, SLIDE, 'new line 1\nnew line 2\n\nafter a gap'))).toEqual([]);
    ed.run(setNotes(ed, SLIDE, 'new line 1\nnew line 2\n\nafter a gap'));
    expect(notesText(ed, SLIDE)).toBe('new line 1\nnew line 2\n\nafter a gap');
    expect(notesParts(ed).length).toBe(1);
    ed.run(setNotes(ed, SLIDE, ''));
    expect(notesText(ed, SLIDE)).toBe('');
    expect(notesParts(ed).length).toBe(1);                                 // emptied notes keep their page
  });

  test('typing coalesces: keystrokes with one key are one undo step, after the step that made the page', () => {
    const ed = openEditor(deckBytes());
    const key = `notes:${SLIDE}`;
    ed.run(setNotes(ed, SLIDE, 'H', { coalesceKey: key }));
    ed.run(setNotes(ed, SLIDE, 'He', { coalesceKey: key }));
    ed.run(setNotes(ed, SLIDE, 'Hey', { coalesceKey: key }));
    expect(ed.log.undoLabel).toBe('Notes');
    ed.undo();
    expect(notesText(ed, SLIDE)).toBe('H');
    ed.undo();
    expect(ed.doc.notesOf(SLIDE)).toBeUndefined();
    expect(ed.log.canUndo).toBe(false);
  });

  test('text that XML cannot hold is written safely', () => {
    const ed = openEditor(deckBytes());
    ed.run(setNotes(ed, SLIDE, 'a < b & "c" > d'));
    expect(notesText(openEditor(ed.save()), SLIDE)).toBe('a < b & "c" > d');
  });

  test('a duplicated slide has its own copy of the notes; a deleted slide takes its notes with it', () => {
    const ed = openEditor(deckBytes());
    ed.run(setNotes(ed, SLIDE, 'shared start'));
    const dup = duplicateSlide(ed, SLIDE);
    ed.run(dup.command);
    const copy = dup.part!;
    expect(notesText(ed, copy)).toBe('shared start');
    expect(ed.doc.notesOf(copy)).not.toBe(ed.doc.notesOf(SLIDE));
    ed.run(setNotes(ed, copy, 'only the copy'));
    expect(notesText(ed, SLIDE)).toBe('shared start');
    ed.run(deleteSlide(ed, copy));
    expect(notesParts(ed).length).toBe(1);
    expect(masterParts(ed).length).toBe(1);                                // the master stays for the other page
  });

  test('a notes page that lost its text placeholder gets one back', () => {
    const ed = openEditor(deckBytes());
    ed.run(setNotes(ed, SLIDE, 'x'));
    const notes = ed.doc.notesOf(SLIDE)!;
    const tree = ed.doc.slide(notes).spTree!;
    const body = ed.doc.slide(notes).shapes.find(s => s.placeholder?.type === 'body')!.el;
    tree.children.splice(tree.children.indexOf(body), 1);
    ed.changed([notes]);
    expect(notesText(ed, SLIDE)).toBe('');
    ed.log.clear();
    expect(cycle(ed, () => setNotes(ed, SLIDE, 'back again'), { parse: false })).toEqual([]);
    ed.run(setNotes(ed, SLIDE, 'back again'));
    expect(notesText(ed, SLIDE)).toBe('back again');
    expect(checkOrder(ed.doc.root(notes))).toEqual([]);
  });
});
