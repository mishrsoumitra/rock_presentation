// Tables: inserting one, cell text, rows and columns, merging and splitting. Each command is checked in the XML it
// leaves, in what the resolver then draws, and by exact undo and redo.
import { OpcPackage } from '@pptx/opc';
import { attr, checkOrder, children, PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  cellAt, cellBody, cellText, DEFAULT_TABLE_STYLE, formatParagraphs, formatRuns, deleteColumns, deleteRows, insertColumn, insertRow, insertTable, isMerged, mergeCells, regionOf, ROW_HEIGHT,
  moveBy, setCellText, setTransforms, splitCell, tableOf, tableSize, wholeRegions, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const BOX = { x: 1000000, y: 1000000, w: 6000000, h: 1500000 };

/** A deck with a table whose cells say where they are: "a1" is column a, row 1. */
function withTable(rows = 3, cols = 3): { ed: Editor; id: number } {
  const ed = openEditor(deckBytes());
  const t = insertTable(ed, SLIDE, rows, cols, BOX);
  ed.run(t.command);
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) ed.run(setCellText(ed, SLIDE, ed.shape(SLIDE, t.id)!, { r, c }, `${'abcdefgh'[c]}${r + 1}`));
  ed.log.clear();
  return { ed, id: t.id };
}
const frame = (ed: Editor, id: number) => ed.shape(SLIDE, id)!;
const tbl = (ed: Editor, id: number) => tableOf(frame(ed, id))!;

/**
 * The table as text, a row per line: each cell's text, then ":gN" and ":rN" for its spans, "<" when it is covered
 * from the left (hMerge) and "^" when covered from above (vMerge).
 */
function dump(ed: Editor, id: number): string[] {
  const t = tbl(ed, id);
  return children(t, 'a:tr').map((tr, r) => children(tr, 'a:tc').map((tc, c) => {
    const g = attr(tc, 'gridSpan'), rs = attr(tc, 'rowSpan');
    const own = [...tc.children].some(x => x.kind === 'element' && x.local === 'txBody') ? (attr(tc, 'hMerge') || attr(tc, 'vMerge') ? '' : cellText(t, r, c)) : '';
    return `${own}${g ? `:g${g}` : ''}${rs ? `:r${rs}` : ''}${attr(tc, 'hMerge') ? '<' : ''}${attr(tc, 'vMerge') ? '^' : ''}` || '.';
  }).join(' '));
}
/** What the resolver draws: every cell that is not covered, as "row,col spans text". */
function drawn(ed: Editor, id: number): string[] {
  const g = resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.graphic!;
  if (g.kind !== 'table') throw new Error('not a table');
  return g.table.cells.filter(c => !c.covered).map(c => `${c.row},${c.col} ${c.colSpan}x${c.rowSpan} ${c.text?.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('/') ?? ''}`);
}
const widths = (ed: Editor, id: number) => children(children(tbl(ed, id), 'a:tblGrid')[0]!, 'a:gridCol').map(c => Number(attr(c, 'w')));

describe('inserting a table', () => {
  test('rows by columns of empty cells, equal columns, the default style, and its definition in tableStyles.xml', () => {
    const ed = openEditor(deckBytes());
    expect(ed.doc.tableStylesPart()).toBeUndefined();
    expect(cycle(ed, () => insertTable(ed, SLIDE, 2, 3, BOX).command)).toEqual([]);
    expect(ed.doc.tableStylesPart()).toBeUndefined();                    // undone: the styles part is gone again
    const t = insertTable(ed, SLIDE, 2, 3, BOX);
    ed.run(t.command);
    expect(tableSize(tbl(ed, t.id))).toEqual({ rows: 2, cols: 3 });
    expect(widths(ed, t.id)).toEqual([2000000, 2000000, 2000000]);
    expect(frame(ed, t.id).xfrm).toMatchObject({ x: 1000000, y: 1000000, cx: 6000000, cy: 2 * ROW_HEIGHT });
    expect(frame(ed, t.id).name).toBe(`Table ${t.id - 1}`);
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    const styles = ed.doc.tableStylesPart()!;
    expect(children(ed.doc.root(styles), 'a:tblStyle').map(s => attr(s, 'styleId'))).toEqual([DEFAULT_TABLE_STYLE]);
    // the resolver finds the style: the header row is filled with accent 1, the body rows are banded
    const g = resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === t.id && s.layer === 'slide')!.graphic!;
    if (g.kind !== 'table') throw new Error('not a table');
    const fill = (r: number) => g.table.cells.find(c => c.row === r && c.col === 0)!.fill;
    expect(fill(0).kind).toBe('solid');
    expect(fill(0)).not.toEqual(fill(1));
  });

  test('a second table reuses the style definition; a saved deck opens with both tables', () => {
    const ed = openEditor(deckBytes());
    ed.run(insertTable(ed, SLIDE, 2, 2, BOX).command);
    ed.run(insertTable(ed, SLIDE, 4, 1, BOX).command);
    expect(children(ed.doc.root(ed.doc.tableStylesPart()!), 'a:tblStyle').length).toBe(1);
    const back = new PresentationDocument(OpcPackage.open(ed.save()));
    expect(back.slide(SLIDE).shapes.filter(s => s.kind === 'graphicFrame').map(s => tableSize(tableOf(s)!))).toEqual([{ rows: 2, cols: 2 }, { rows: 4, cols: 1 }]);
    expect(back.tableStylesPart()).toBeDefined();
  });

  test('sizes that make no table are refused', () => {
    const ed = openEditor(deckBytes());
    expect(() => insertTable(ed, SLIDE, 0, 3, BOX)).toThrow(/1 to 75/);
    expect(() => insertTable(ed, SLIDE, 3, 76, BOX)).toThrow(/1 to 75/);
  });
});

describe('cell text', () => {
  test('text goes into one cell, lines become paragraphs, and undo is exact', () => {
    const { ed, id } = withTable(2, 2);
    expect(cycle(ed, () => setCellText(ed, SLIDE, frame(ed, id), { r: 1, c: 0 }, 'two\nlines'))).toEqual([]);
    ed.run(setCellText(ed, SLIDE, frame(ed, id), { r: 1, c: 0 }, 'two\nlines'));
    expect(dump(ed, id)).toEqual(['a1 b1', 'two\nlines b2']);
    expect(drawn(ed, id)).toContain('1,0 1x1 two/lines');
    expect(ed.log.undoLabel).toBe('Typing');
    expect(() => setCellText(ed, SLIDE, frame(ed, id), { r: 5, c: 0 }, 'x')).toThrow(/no such cell/);
  });
});

describe('rows', () => {
  test('insert above, below and at the ends; the frame grows by the row height', () => {
    const { ed, id } = withTable(2, 2);
    const h0 = frame(ed, id).xfrm!.cy;
    expect(cycle(ed, () => insertRow(ed, SLIDE, frame(ed, id), 1))).toEqual([]);
    ed.run(insertRow(ed, SLIDE, frame(ed, id), 1));
    expect(dump(ed, id)).toEqual(['a1 b1', '. .', 'a2 b2']);
    expect(frame(ed, id).xfrm!.cy).toBe(h0 + ROW_HEIGHT);
    ed.run(insertRow(ed, SLIDE, frame(ed, id), 0));
    ed.run(insertRow(ed, SLIDE, frame(ed, id), 99));
    expect(dump(ed, id)).toEqual(['. .', 'a1 b1', '. .', 'a2 b2', '. .']);
    expect(tableSize(tbl(ed, id))).toEqual({ rows: 5, cols: 2 });
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Insert Row');
  });

  test('a new row looks like the row it is made from: its height, cell properties and text formatting', () => {
    const { ed, id } = withTable(2, 2);
    const first = children(tbl(ed, id), 'a:tr')[0]!;
    first.attrs.find(a => a.local === 'h')!.value = '500000';
    const tc = cellAt(tbl(ed, id), 0, 1)!;
    children(tc, 'a:tcPr')[0]!.attrs.push({ prefix: null, local: 'anchor', value: 'ctr', ns: '' });
    ed.changed([SLIDE]);
    ed.run(insertRow(ed, SLIDE, frame(ed, id), 1));
    const added = children(tbl(ed, id), 'a:tr')[1]!;
    expect(attr(added, 'h')).toBe('500000');
    expect(attr(children(children(added, 'a:tc')[1]!, 'a:tcPr')[0], 'anchor')).toBe('ctr');
    expect(attr(children(children(added, 'a:tc')[0]!, 'a:tcPr')[0], 'anchor')).toBeUndefined();
  });

  test('delete one row or several; the last row cannot go', () => {
    const { ed, id } = withTable(4, 2);
    const h0 = frame(ed, id).xfrm!.cy;
    expect(cycle(ed, () => deleteRows(ed, SLIDE, frame(ed, id), 1, 2))).toEqual([]);
    ed.run(deleteRows(ed, SLIDE, frame(ed, id), 1, 2));
    expect(dump(ed, id)).toEqual(['a1 b1', 'a4 b4']);
    expect(frame(ed, id).xfrm!.cy).toBe(h0 - 2 * ROW_HEIGHT);
    ed.run(deleteRows(ed, SLIDE, frame(ed, id), 0));
    expect(dump(ed, id)).toEqual(['a4 b4']);
    expect(() => ed.run(deleteRows(ed, SLIDE, frame(ed, id), 0))).toThrow(/at least one row/);
    expect(dump(ed, id)).toEqual(['a4 b4']);
  });
});

describe('columns', () => {
  test('insert left, right and at the ends; the new column is as wide as its neighbour and the frame grows', () => {
    const { ed, id } = withTable(2, 2);
    expect(cycle(ed, () => insertColumn(ed, SLIDE, frame(ed, id), 1))).toEqual([]);
    ed.run(insertColumn(ed, SLIDE, frame(ed, id), 1));
    expect(dump(ed, id)).toEqual(['a1 . b1', 'a2 . b2']);
    expect(widths(ed, id)).toEqual([3000000, 3000000, 3000000]);
    expect(frame(ed, id).xfrm!.cx).toBe(9000000);
    ed.run(insertColumn(ed, SLIDE, frame(ed, id), 0));
    ed.run(insertColumn(ed, SLIDE, frame(ed, id), 99));
    expect(dump(ed, id)).toEqual(['. a1 . b1 .', '. a2 . b2 .']);
    expect(tableSize(tbl(ed, id))).toEqual({ rows: 2, cols: 5 });
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
  });

  test('delete one column or several; the last column cannot go', () => {
    const { ed, id } = withTable(2, 4);
    expect(cycle(ed, () => deleteColumns(ed, SLIDE, frame(ed, id), 1, 2))).toEqual([]);
    ed.run(deleteColumns(ed, SLIDE, frame(ed, id), 2, 1));              // either way round
    expect(dump(ed, id)).toEqual(['a1 d1', 'a2 d2']);
    expect(widths(ed, id)).toEqual([1500000, 1500000]);
    expect(frame(ed, id).xfrm!.cx).toBe(3000000);
    ed.run(deleteColumns(ed, SLIDE, frame(ed, id), 1));
    expect(() => ed.run(deleteColumns(ed, SLIDE, frame(ed, id), 0))).toThrow(/at least one column/);
    expect(dump(ed, id)).toEqual(['a1', 'a2']);
  });
});

describe('merging and splitting', () => {
  test('a block of cells becomes one: spans on the first, the rest covered, the text gathered, and it draws as one cell', () => {
    const { ed, id } = withTable(3, 3);
    const range = { r0: 0, c0: 0, r1: 1, c1: 1 };
    expect(cycle(ed, () => mergeCells(ed, SLIDE, frame(ed, id), range))).toEqual([]);
    ed.run(mergeCells(ed, SLIDE, frame(ed, id), range));
    expect(dump(ed, id)).toEqual(['a1\nb1\na2\nb2:g2:r2 < c1', ':g2^ <^ c2', 'a3 b3 c3']);
    expect(drawn(ed, id)).toEqual(['0,0 2x2 a1/b1/a2/b2', '0,2 1x1 c1', '1,2 1x1 c2', '2,0 1x1 a3', '2,1 1x1 b3', '2,2 1x1 c3']);
    expect(regionOf(tbl(ed, id), 1, 1)).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(isMerged(tbl(ed, id), 1, 0)).toBe(true);
    expect(isMerged(tbl(ed, id), 2, 2)).toBe(false);
    expect(cellText(tbl(ed, id), 1, 1)).toBe('a1\nb1\na2\nb2');            // any cell of the merge reads as the merge
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
  });

  test('a row of cells and a column of cells; empty cells add no blank lines', () => {
    const { ed, id } = withTable(3, 3);
    ed.run(setCellText(ed, SLIDE, frame(ed, id), { r: 0, c: 1 }, ''));
    ed.run(mergeCells(ed, SLIDE, frame(ed, id), { r0: 0, c0: 2, r1: 0, c1: 0 }));       // given right to left
    ed.run(mergeCells(ed, SLIDE, frame(ed, id), { r0: 1, c0: 2, r1: 2, c1: 2 }));
    expect(dump(ed, id)).toEqual(['a1\nc1:g3 < <', 'a2 b2 c2\nc3:r2', 'a3 b3 ^']);
    expect(drawn(ed, id)).toEqual(['0,0 3x1 a1/c1', '1,0 1x1 a2', '1,1 1x1 b2', '1,2 1x2 c2/c3', '2,0 1x1 a3', '2,1 1x1 b3']);
  });

  test('a range that cuts through a merge grows to take it whole', () => {
    const { ed, id } = withTable(3, 3);
    ed.run(mergeCells(ed, SLIDE, frame(ed, id), { r0: 0, c0: 0, r1: 1, c1: 1 }));
    expect(wholeRegions(tbl(ed, id), { r0: 1, c0: 1, r1: 2, c1: 2 })).toEqual({ r0: 0, c0: 0, r1: 2, c1: 2 });
    ed.run(mergeCells(ed, SLIDE, frame(ed, id), { r0: 1, c0: 1, r1: 2, c1: 2 }));
    expect(drawn(ed, id)).toEqual(['0,0 3x3 a1/b1/a2/b2/c1/c2/a3/b3/c3']);
  });

  test('split gives the cells back, the text staying in the first; one cell alone does nothing to merge', () => {
    const { ed, id } = withTable(3, 3);
    ed.run(mergeCells(ed, SLIDE, frame(ed, id), { r0: 0, c0: 0, r1: 1, c1: 1 }));
    ed.log.seal();
    expect(cycle(ed, () => splitCell(ed, SLIDE, frame(ed, id), { r: 1, c: 1 }))).toEqual([]);
    ed.run(splitCell(ed, SLIDE, frame(ed, id), { r: 1, c: 1 }));
    expect(dump(ed, id)).toEqual(['a1\nb1\na2\nb2 . c1', '. . c2', 'a3 b3 c3']);
    expect(drawn(ed, id).length).toBe(9);
    const before = dump(ed, id);
    ed.run(mergeCells(ed, SLIDE, frame(ed, id), { r0: 2, c0: 2, r1: 2, c1: 2 }));
    expect(dump(ed, id)).toEqual(before);
  });
});

describe('rows and columns through merged cells', () => {
  /** 4 by 4 with a 2 by 2 merge in the middle (rows 2-3, columns b-c). */
  function merged(): { ed: Editor; id: number } {
    const t = withTable(4, 4);
    t.ed.run(mergeCells(t.ed, SLIDE, frame(t.ed, t.id), { r0: 1, c0: 1, r1: 2, c1: 2 }));
    t.ed.log.clear();
    expect(dump(t.ed, t.id)).toEqual(['a1 b1 c1 d1', 'a2 b2\nc2\nb3\nc3:g2:r2 < d2', 'a3 :g2^ <^ d3', 'a4 b4 c4 d4']);
    return t;
  }
  const big = (ed: Editor, id: number) => drawn(ed, id).find(d => d.includes('b2'))!;

  test('a row put through the middle of a merge makes it a row taller', () => {
    const { ed, id } = merged();
    expect(cycle(ed, () => insertRow(ed, SLIDE, frame(ed, id), 2))).toEqual([]);
    ed.run(insertRow(ed, SLIDE, frame(ed, id), 2));
    expect(dump(ed, id)).toEqual(['a1 b1 c1 d1', 'a2 b2\nc2\nb3\nc3:g2:r3 < d2', '. :g2^ <^ .', 'a3 :g2^ <^ d3', 'a4 b4 c4 d4']);
    expect(big(ed, id)).toBe('1,1 2x3 b2/c2/b3/c3');
  });

  test('a row put above or below a merge leaves it as it was', () => {
    const { ed, id } = merged();
    ed.run(insertRow(ed, SLIDE, frame(ed, id), 1));
    ed.run(insertRow(ed, SLIDE, frame(ed, id), 4));
    expect(big(ed, id)).toBe('2,1 2x2 b2/c2/b3/c3');
    expect(drawn(ed, id).length).toBe(21);                               // 6 rows of 4, less the 3 covered cells
  });

  test('a column put through the middle of a merge makes it a column wider', () => {
    const { ed, id } = merged();
    expect(cycle(ed, () => insertColumn(ed, SLIDE, frame(ed, id), 2))).toEqual([]);
    ed.run(insertColumn(ed, SLIDE, frame(ed, id), 2));
    expect(dump(ed, id)).toEqual(['a1 b1 . c1 d1', 'a2 b2\nc2\nb3\nc3:g3:r2 < < d2', 'a3 :g3^ <^ <^ d3', 'a4 b4 . c4 d4']);
    expect(big(ed, id)).toBe('1,1 3x2 b2/c2/b3/c3');
  });

  test('deleting the first row of a merge moves its text down; deleting its last row makes it shorter', () => {
    const { ed, id } = merged();
    expect(cycle(ed, () => deleteRows(ed, SLIDE, frame(ed, id), 1))).toEqual([]);
    ed.run(deleteRows(ed, SLIDE, frame(ed, id), 1));
    expect(dump(ed, id)).toEqual(['a1 b1 c1 d1', 'a3 b2\nc2\nb3\nc3:g2 < d3', 'a4 b4 c4 d4']);
    expect(big(ed, id)).toBe('1,1 2x1 b2/c2/b3/c3');
    const again = merged();
    again.ed.run(deleteRows(again.ed, SLIDE, frame(again.ed, again.id), 2));
    expect(dump(again.ed, again.id)).toEqual(['a1 b1 c1 d1', 'a2 b2\nc2\nb3\nc3:g2 < d2', 'a4 b4 c4 d4']);
  });

  test('deleting the first column of a merge moves its text right; deleting its last column makes it narrower', () => {
    const { ed, id } = merged();
    expect(cycle(ed, () => deleteColumns(ed, SLIDE, frame(ed, id), 1))).toEqual([]);
    ed.run(deleteColumns(ed, SLIDE, frame(ed, id), 1));
    expect(dump(ed, id)).toEqual(['a1 c1 d1', 'a2 b2\nc2\nb3\nc3:r2 d2', 'a3 ^ d3', 'a4 c4 d4']);
    expect(big(ed, id)).toBe('1,1 1x2 b2/c2/b3/c3');
    const again = merged();
    again.ed.run(deleteColumns(again.ed, SLIDE, frame(again.ed, again.id), 2));
    expect(dump(again.ed, again.id)).toEqual(['a1 b1 d1', 'a2 b2\nc2\nb3\nc3:r2 d2', 'a3 ^ d3', 'a4 b4 d4']);
  });

  test('deleting every row of a merge takes it away whole, and what is left is a plain grid', () => {
    const { ed, id } = merged();
    ed.run(deleteRows(ed, SLIDE, frame(ed, id), 1, 2));
    expect(dump(ed, id)).toEqual(['a1 b1 c1 d1', 'a4 b4 c4 d4']);
    expect(drawn(ed, id).length).toBe(8);
  });
});

describe('formatting cell text', () => {
  test('character and paragraph formatting work in a cell as in a shape, in schema order, and undo exactly', () => {
    const { ed, id } = withTable(2, 2);
    const body = () => cellBody(tbl(ed, id), 1, 0)!;
    const all = { start: { p: 0, offset: 0 }, end: { p: 0, offset: 2 } };
    expect(cycle(ed, () => formatParagraphs(ed, SLIDE, body(), all, { algn: 'ctr', lnSpc: { pct: 150000 }, bullet: { kind: 'char', char: '•' } }))).toEqual([]);
    ed.run(formatParagraphs(ed, SLIDE, body(), all, { algn: 'ctr', lnSpc: { pct: 150000 }, bullet: { kind: 'char', char: '•' } }));
    expect(cycle(ed, () => formatRuns(ed, SLIDE, body(), all, { b: true, fill: { kind: 'srgb', hex: 'FF0000' }, latin: 'Georgia' }))).toEqual([]);
    ed.run(formatRuns(ed, SLIDE, body(), all, { b: true, fill: { kind: 'srgb', hex: 'FF0000' }, latin: 'Georgia' }));
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    const g = resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.graphic!;
    if (g.kind !== 'table') throw new Error('not a table');
    const cell = g.table.cells.find(c => c.row === 1 && c.col === 0)!.text!.paragraphs[0]!;
    expect(cell.props.align).toBe('ctr');
    expect(cell.props.bullet?.char).toBe('•');
    expect(cell.runs[0]!.props).toMatchObject({ bold: true, latin: 'Georgia' });
    // the cell beside it is untouched
    expect(g.table.cells.find(c => c.row === 1 && c.col === 1)!.text!.paragraphs[0]!.props.align).not.toBe('ctr');
  });
});

describe('resizing a table', () => {
  const heights = (ed: Editor, id: number) => children(tbl(ed, id), 'a:tr').map(tr => Number(attr(tr, 'h')));

  test('the columns and rows scale with the frame and still add up to it; moving it changes neither', () => {
    const { ed, id } = withTable(3, 3);
    const x = frame(ed, id).xfrm!;
    expect(cycle(ed, () => setTransforms(ed, SLIDE, [{ shape: frame(ed, id), xfrm: { ...x, cx: 9000001, cy: x.cy * 2 } }], 'Resize'))).toEqual([]);
    ed.run(setTransforms(ed, SLIDE, [{ shape: frame(ed, id), xfrm: { ...x, cx: 9000001, cy: x.cy * 2 } }], 'Resize'));
    expect(widths(ed, id)).toEqual([3000000, 3000000, 3000001]);
    expect(widths(ed, id).reduce((a, b) => a + b, 0)).toBe(frame(ed, id).xfrm!.cx);
    expect(heights(ed, id)).toEqual([2 * ROW_HEIGHT, 2 * ROW_HEIGHT, 2 * ROW_HEIGHT]);
    ed.run(moveBy(ed, SLIDE, [frame(ed, id)], 500000, 250000));
    expect(widths(ed, id)).toEqual([3000000, 3000000, 3000001]);
    expect(frame(ed, id).xfrm).toMatchObject({ x: 1500000, y: 1250000, cx: 9000001 });
  });

  test('unequal columns keep their proportions', () => {
    const { ed, id } = withTable(1, 2);
    ed.run(insertColumn(ed, SLIDE, frame(ed, id), 2));
    const cols = children(children(tbl(ed, id), 'a:tblGrid')[0]!, 'a:gridCol');
    cols[0]!.attrs.find(a => a.local === 'w')!.value = '1000000';
    cols[1]!.attrs.find(a => a.local === 'w')!.value = '2000000';
    cols[2]!.attrs.find(a => a.local === 'w')!.value = '3000000';
    ed.changed([SLIDE]);
    const x = frame(ed, id).xfrm!;
    ed.run(setTransforms(ed, SLIDE, [{ shape: frame(ed, id), xfrm: { ...x, cx: 3000000 } }], 'Resize'));
    expect(widths(ed, id)).toEqual([333333, 666667, 1000000]);                // a third of each, the last taking the rounding
  });
});
