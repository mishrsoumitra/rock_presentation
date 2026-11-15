// Comments: the first one of a deck and of a slide, authors and their numbering, changing, deleting, and what
// slides do with their comments.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, PresentationDocument } from '@pptx/model';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { addComment, addSlide, commentCount, deleteComment, deleteSlide, duplicateSlide, editComment, initialsOf, slideComments, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const WHEN = new Date('2026-11-15T09:30:00.000Z');
const ASHA = { name: 'Asha Rao' }, RAVI = { name: 'Ravi K. Menon', initials: 'RM' };

const parts = (ed: Editor, prefix: string) => ed.pkg.partNames().filter(n => n.startsWith(prefix) && !n.endsWith('.rels'));
const said = (ed: Editor, slide = SLIDE) => slideComments(ed, slide).map(c => `${c.initials}${c.index}: ${c.text}`);

describe('adding comments', () => {
  test('the first comment of a deck: its author list, the slide\'s comments part, both related, and undone exactly', () => {
    const ed = openEditor(deckBytes());
    expect(slideComments(ed, SLIDE)).toEqual([]);
    expect(cycle(ed, () => addComment(ed, SLIDE, 'Check this number', ASHA, undefined, WHEN))).toEqual([]);
    expect(parts(ed, '/ppt/comment')).toEqual([]);                        // cycle leaves the deck as it was

    ed.run(addComment(ed, SLIDE, 'Check this number', ASHA, { x: 914400, y: 457200 }, WHEN));
    expect(slideComments(ed, SLIDE)).toEqual([{ authorId: 0, index: 1, author: 'Asha Rao', initials: 'AR', date: '2026-11-15T09:30:00.000', text: 'Check this number' }]);
    expect(parts(ed, '/ppt/comment')).toEqual(['/ppt/commentAuthors.xml', '/ppt/comments/comment1.xml']);
    expect(ed.doc.targets(ed.doc.mainPart, 'commentAuthors')).toEqual(['/ppt/commentAuthors.xml']);
    expect(ed.doc.targets(SLIDE, 'comments')).toEqual(['/ppt/comments/comment1.xml']);
    expect(checkOrder(ed.doc.root('/ppt/comments/comment1.xml'))).toEqual([]);
    // the marker is placed in PowerPoint's units of 1/576 inch
    expect(new TextDecoder().decode(OpcPackage.open(ed.save()).getPart('/ppt/comments/comment1.xml')!.bytes)).toContain('<p:pos x="576" y="288"/>');
    expect(ed.log.undoLabel).toBe('New Comment');
  });

  test('each author\'s comments are numbered from 1; an author is found again by name; initials are given or made', () => {
    const ed = openEditor(deckBytes());
    ed.run(addComment(ed, SLIDE, 'one', ASHA, undefined, WHEN));
    ed.run(addComment(ed, SLIDE, 'two', RAVI, undefined, WHEN));
    ed.run(addComment(ed, SLIDE, 'three', { name: '  Asha Rao ' }, undefined, WHEN));
    expect(said(ed)).toEqual(['AR1: one', 'RM1: two', 'AR2: three']);
    expect(slideComments(ed, SLIDE).map(c => c.authorId)).toEqual([0, 1, 0]);
    expect(commentCount(ed, SLIDE)).toBe(3);
    expect([initialsOf('Asha Rao'), initialsOf('  madonna '), initialsOf('a b c d e'), initialsOf('')]).toEqual(['AR', 'M', 'ABC', '?']);
  });

  test('comments on two slides: a part each, one author list, each author\'s numbering running across the deck', () => {
    const ed = openEditor(deckBytes());
    const second = addSlide(ed, LAYOUT);
    ed.run(second.command);
    ed.run(addComment(ed, SLIDE, 'on one', ASHA, undefined, WHEN));
    ed.run(addComment(ed, second.part, 'on two', ASHA, undefined, WHEN));
    expect(parts(ed, '/ppt/comments/').length).toBe(2);
    expect(parts(ed, '/ppt/commentAuthors').length).toBe(1);
    expect([said(ed), said(ed, second.part)]).toEqual([['AR1: on one'], ['AR2: on two']]);
  });

  test('text XML must escape is kept as written, through a save', () => {
    const ed = openEditor(deckBytes());
    ed.run(addComment(ed, SLIDE, 'a < b & "c"\nsecond line', { name: 'O\'Brien & <Co>' }, undefined, WHEN));
    const back = openEditor(ed.save());
    expect(slideComments(back, SLIDE)[0]).toMatchObject({ text: 'a < b & "c"\nsecond line', author: 'O\'Brien & <Co>' });
    expect(new PresentationDocument(OpcPackage.open(ed.save())).targets(SLIDE, 'comments').length).toBe(1);
  });
});

describe('changing and deleting comments', () => {
  /** Three comments on the slide: Asha's first, Ravi's first, Asha's second. */
  function three(): Editor {
    const ed = openEditor(deckBytes());
    ed.run(addComment(ed, SLIDE, 'one', ASHA, undefined, WHEN));
    ed.run(addComment(ed, SLIDE, 'two', RAVI, undefined, WHEN));
    ed.run(addComment(ed, SLIDE, 'three', ASHA, undefined, WHEN));
    ed.log.clear();
    return ed;
  }

  test('editing changes the text only', () => {
    const ed = three();
    expect(cycle(ed, () => editComment(ed, SLIDE, 0, 2, 'three, revised'))).toEqual([]);
    ed.run(editComment(ed, SLIDE, 0, 2, 'three, revised'));
    expect(said(ed)).toEqual(['AR1: one', 'RM1: two', 'AR2: three, revised']);
    expect(slideComments(ed, SLIDE)[2]!.date).toBe('2026-11-15T09:30:00.000');
    expect(ed.log.undoLabel).toBe('Edit Comment');
    expect(() => editComment(ed, SLIDE, 0, 9, 'x')).toThrow(/not on this slide/);
  });

  test('deleting takes one comment; the last one takes the slide\'s comments part with it; a later comment keeps counting', () => {
    const ed = three();
    expect(cycle(ed, () => deleteComment(ed, SLIDE, 1, 1))).toEqual([]);
    ed.run(deleteComment(ed, SLIDE, 1, 1));
    expect(said(ed)).toEqual(['AR1: one', 'AR2: three']);
    ed.run(deleteComment(ed, SLIDE, 0, 1));
    ed.run(deleteComment(ed, SLIDE, 0, 2));
    expect(said(ed)).toEqual([]);
    expect(parts(ed, '/ppt/comments/')).toEqual([]);
    expect(ed.doc.targets(SLIDE, 'comments')).toEqual([]);
    // numbers are never reused: Asha's next comment is her third
    ed.run(addComment(ed, SLIDE, 'later', ASHA, undefined, WHEN));
    expect(said(ed)).toEqual(['AR3: later']);
    // undo back through the deletions brings each comment back
    for (let i = 0; i < 4; i++) ed.undo();
    expect(said(ed)).toEqual(['AR1: one', 'RM1: two', 'AR2: three']);
  });
});

describe('comments and slides', () => {
  test('a duplicated slide has its own copy of the comments; a deleted slide takes its comments with it', () => {
    const ed = openEditor(deckBytes());
    ed.run(addComment(ed, SLIDE, 'stays with its slide', ASHA, undefined, WHEN));
    const dup = duplicateSlide(ed, SLIDE);
    ed.run(dup.command);
    const copy = dup.part!;
    expect(said(ed, copy)).toEqual(['AR1: stays with its slide']);
    expect(ed.doc.targets(copy, 'comments')[0]).not.toBe(ed.doc.targets(SLIDE, 'comments')[0]);
    ed.run(editComment(ed, copy, 0, 1, 'changed on the copy'));
    expect(said(ed)).toEqual(['AR1: stays with its slide']);
    ed.run(deleteSlide(ed, copy));
    expect(parts(ed, '/ppt/comments/').length).toBe(1);
    expect(parts(ed, '/ppt/commentAuthors').length).toBe(1);              // the author list is the deck's
  });
});
