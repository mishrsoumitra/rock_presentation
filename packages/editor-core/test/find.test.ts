// Find and replace: where text is found (and where it must not be), in what order, and what replacing does to the
// text, its formatting and the undo history.
import { children, runsOf } from '@pptx/model';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  addSlide, bodyPlainText, compareMatches, findInText, findText, formatRuns, group, insertTextBox, replaceAll, replaceMatch, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const BOX = { x: 1000000, y: 1000000, w: 4000000, h: 800000 };

function box(ed: Editor, part: string, text: string): number {
  const r = insertTextBox(ed, part, BOX, text);
  ed.run(r.command);
  return r.id;
}
const textOf = (ed: Editor, part: string, id: number) => bodyPlainText(ed.shape(part, id)!.txBody!);

describe('finding in a string', () => {
  test('case is ignored unless asked for; matches never overlap', () => {
    expect(findInText('Zebra zebra ZEBRA', 'zebra')).toEqual([{ start: 0, end: 5 }, { start: 6, end: 11 }, { start: 12, end: 17 }]);
    expect(findInText('Zebra zebra ZEBRA', 'zebra', { matchCase: true })).toEqual([{ start: 6, end: 11 }]);
    expect(findInText('aaaa', 'aa')).toEqual([{ start: 0, end: 2 }, { start: 2, end: 4 }]);
    expect(findInText('anything', '')).toEqual([]);
  });

  test('whole words only: not inside a longer word, and digits and accented letters count as word characters', () => {
    expect(findInText('art party art.', 'art', { wholeWords: true })).toEqual([{ start: 0, end: 3 }, { start: 10, end: 13 }]);
    expect(findInText('art2 éart art', 'art', { wholeWords: true })).toEqual([{ start: 10, end: 13 }]);
  });

  test('the query is plain text, not a pattern', () => {
    expect(findInText('cost (a+b)*2 = $5.00', '(a+b)*2')).toEqual([{ start: 5, end: 12 }]);
    expect(findInText('a.c abc', 'a.c')).toEqual([{ start: 0, end: 3 }]);
    expect(findInText('50% [x] \\ ^', '[x] \\ ^')).toEqual([{ start: 4, end: 11 }]);
  });
});

describe('finding in a deck', () => {
  test('matches come in reading order: slide, then shape, then paragraph, then position', () => {
    const ed = openEditor(deckBytes());
    const second = addSlide(ed, LAYOUT);
    ed.run(second.command);
    const a = box(ed, SLIDE, 'quokka one\nand a quokka, quokka');
    const b = box(ed, SLIDE, 'no match here');
    const c = box(ed, second.part, 'Quokka on slide two');
    const found = findText(ed, 'quokka');
    expect(found.map(m => [m.slideIndex, ed.view(m.part, m.shape).id, m.p, m.start, m.end])).toEqual([
      [0, a, 0, 0, 6], [0, a, 1, 6, 12], [0, a, 1, 14, 20], [1, c, 0, 0, 6],
    ]);
    expect([...found].sort(compareMatches)).toEqual(found);
    expect(findText(ed, 'quokka', { matchCase: true }).length).toBe(3);
    expect(findText(ed, 'here').map(m => ed.view(m.part, m.shape).id)).toEqual([b]);
    expect(findText(ed, 'not in the deck')).toEqual([]);
  });

  test('text split across runs by formatting is still found; a match never spans two paragraphs', () => {
    const ed = openEditor(deckBytes());
    const id = box(ed, SLIDE, 'wombat\nwombat');
    const body = ed.shape(SLIDE, id)!.txBody!;
    ed.run(formatRuns(ed, SLIDE, body, { start: { p: 0, offset: 2 }, end: { p: 0, offset: 4 } }, { b: true }));
    expect(runsOf(children(body, 'a:p')[0]!).map(r => r.text)).toEqual(['wo', 'mb', 'at']);
    expect(findText(ed, 'wombat').map(m => [m.p, m.start, m.end])).toEqual([[0, 0, 6], [1, 0, 6]]);
    expect(findText(ed, 'wombat\nwombat')).toEqual([]);
    expect(findText(ed, 'atwo')).toEqual([]);
  });

  test('shapes inside a group are searched; layouts and masters are not', () => {
    const ed = openEditor(deckBytes());
    const a = box(ed, SLIDE, 'numbat left');
    const b = box(ed, SLIDE, 'numbat right');
    ed.run(group(ed, SLIDE, [ed.shape(SLIDE, a)!, ed.shape(SLIDE, b)!]));
    expect(findText(ed, 'numbat').map(m => ed.view(m.part, m.shape).id)).toEqual([a, b]);
    // "Click to edit" is the masters' and layouts' prompt text
    expect(findText(ed, 'Click to edit')).toEqual([]);
  });
});

describe('replacing', () => {
  test('one match: the text changes, the rest stays, and undo is exact', () => {
    const ed = openEditor(deckBytes());
    const id = box(ed, SLIDE, 'red fish, red boat');
    ed.log.clear();
    const second = () => findText(ed, 'red')[1]!;
    expect(cycle(ed, () => replaceMatch(ed, second(), 'blue'))).toEqual([]);
    ed.run(replaceMatch(ed, second(), 'blue'));
    expect(textOf(ed, SLIDE, id)).toBe('red fish, blue boat');
    expect(ed.log.undoLabel).toBe('Replace');
  });

  test('the new text takes the formatting of what it replaces', () => {
    const ed = openEditor(deckBytes());
    const id = box(ed, SLIDE, 'plain bold plain');
    const body = ed.shape(SLIDE, id)!.txBody!;
    ed.run(formatRuns(ed, SLIDE, body, { start: { p: 0, offset: 6 }, end: { p: 0, offset: 10 } }, { b: true }));
    ed.run(replaceMatch(ed, findText(ed, 'bold')[0]!, 'heavy'));
    const runs = runsOf(children(body, 'a:p')[0]!);
    expect(runs.map(r => r.text)).toEqual(['plain ', 'heavy', ' plain']);
    expect(runs.map(r => r.rPr?.attrs.some(a => a.local === 'b' && a.value === '1') ?? false)).toEqual([false, true, false]);
  });

  test('replace all: every slide, one undo step, and the count of what was replaced', () => {
    const ed = openEditor(deckBytes());
    const second = addSlide(ed, LAYOUT);
    ed.run(second.command);
    const a = box(ed, SLIDE, 'cat and cat\ncat');
    const b = box(ed, second.part, 'Cat again');
    ed.log.clear();
    expect(cycle(ed, () => replaceAll(ed, 'cat', 'tiger').command)).toEqual([]);
    const r = replaceAll(ed, 'cat', 'tiger');
    expect(r.count).toBe(4);
    ed.run(r.command);
    expect(textOf(ed, SLIDE, a)).toBe('tiger and tiger\ntiger');
    expect(textOf(ed, second.part, b)).toBe('tiger again');
    expect(ed.log.undoLabel).toBe('Replace All');
    ed.undo();
    expect(textOf(ed, SLIDE, a)).toBe('cat and cat\ncat');
    expect(textOf(ed, second.part, b)).toBe('Cat again');
    expect(ed.log.canUndo).toBe(false);
  });

  test('replace all with a longer, shorter, empty, or self-containing replacement', () => {
    const ed = openEditor(deckBytes());
    const id = box(ed, SLIDE, 'ab ab ab');
    ed.run(replaceAll(ed, 'ab', 'abab').command);
    expect(textOf(ed, SLIDE, id)).toBe('abab abab abab');                 // what was put in is not searched again
    ed.run(replaceAll(ed, 'abab', 'x').command);
    expect(textOf(ed, SLIDE, id)).toBe('x x x');
    ed.run(replaceAll(ed, 'x ', '').command);
    expect(textOf(ed, SLIDE, id)).toBe('x');
    const none = replaceAll(ed, 'zzz', 'y');
    expect(none.count).toBe(0);
    ed.run(none.command);
    expect(textOf(ed, SLIDE, id)).toBe('x');
  });

  test('replace all honours match case and whole words', () => {
    const ed = openEditor(deckBytes());
    const id = box(ed, SLIDE, 'Art art party');
    ed.run(replaceAll(ed, 'art', 'ART', { matchCase: true, wholeWords: true }).command);
    expect(textOf(ed, SLIDE, id)).toBe('Art ART party');
  });
});
