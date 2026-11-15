// apps/web/src/ui/tables.ts  -- working in a table: which cell a point is in, the selected cells, and their commands
import {
  cellAt, cellBody, cellText, compound, deleteColumns, deleteRows, insertColumn, insertRow, isMerged, mergeCells, regionOf, setCellText, splitCell, tableOf,
  tableSize, wholeRegions, type CellRange, type CellRef, type Command,
} from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { tableGrid, type RenderShape, type TableGrid } from '@pptx/render';
import { invert, type ResolvedTable } from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';
import type { XmlElement } from '@pptx/xml';
import { useEditor, type Cells, type EditorState } from '../state/store.ts';

const grids = new WeakMap<ResolvedTable, TableGrid>();

/** The resolved table a rendered shape draws, with where its rows and columns are (laid out once per table state). */
export function gridOf(rs: RenderShape | undefined, fonts: FontRegistry): { table: ResolvedTable; grid: TableGrid } | undefined {
  const g = rs?.source.graphic;
  if (!rs || g?.kind !== 'table') return undefined;
  let grid = grids.get(g.table);
  if (!grid) {
    grid = tableGrid(g.table, fonts);
    grids.set(g.table, grid);
  }
  return { table: g.table, grid };
}

/** The cell (row and column) under a point of the slide, or undefined outside the table. */
export function cellAtPoint(rs: RenderShape, fonts: FontRegistry, at: { x: number; y: number }): CellRef | undefined {
  const t = gridOf(rs, fonts);
  if (!t) return undefined;
  const m = invert(rs.transform);
  const x = m[0] * at.x + m[2] * at.y + m[4], y = m[1] * at.x + m[3] * at.y + m[5];
  const { colX, rowY } = t.grid;
  if (x < 0 || y < 0 || x > colX[colX.length - 1]! || y > rowY[rowY.length - 1]!) return undefined;
  const c = Math.min(colX.length - 2, Math.max(0, colX.findIndex((edge, i) => i > 0 && x <= edge) - 1));
  const r = Math.min(rowY.length - 2, Math.max(0, rowY.findIndex((edge, i) => i > 0 && y <= edge) - 1));
  return { r, c };
}

/** The corners of a block of cells on the slide (x, y pairs, clockwise from the top left), through the table's transform. */
export function cellCorners(rs: RenderShape, fonts: FontRegistry, range: CellRange): number[] | undefined {
  const t = gridOf(rs, fonts);
  if (!t) return undefined;
  const { colX, rowY } = t.grid;
  const x0 = colX[range.c0], x1 = colX[range.c1 + 1], y0 = rowY[range.r0], y1 = rowY[range.r1 + 1];
  if (x0 === undefined || x1 === undefined || y0 === undefined || y1 === undefined) return undefined;
  const m = rs.transform;
  return [[x0, y0], [x1, y0], [x1, y1], [x0, y1]].flatMap(([x, y]) => [m[0] * x! + m[2] * y! + m[4], m[1] * x! + m[3] * y! + m[5]]);
}

/** The cells in force: the store's, when the table they are in is the one selected shape and still has them. */
export function activeCells(s: Pick<EditorState, 'cells' | 'selection' | 'text'>): Cells | null {
  const c = s.cells;
  if (!c || s.text || s.selection.length !== 1 || s.selection[0] !== c.el) return null;
  const tbl = tableOf(c.el);
  if (!tbl) return null;
  const { rows, cols } = tableSize(tbl);
  const inside = (x: CellRef) => x.r >= 0 && x.c >= 0 && x.r < rows && x.c < cols;
  return inside(c.anchor) && inside(c.focus) ? c : null;
}

/** The block of cells selected: from the anchor to the focus, grown to hold whole any merged cell it touches. */
export function rangeOf(cells: Cells): CellRange {
  return wholeRegions(tableOf(cells.el)!, { r0: cells.anchor.r, c0: cells.anchor.c, r1: cells.focus.r, c1: cells.focus.c });
}

/** Whether the one selected shape is a table. */
export function selectedTable(s: Pick<EditorState, 'selection'>): XmlElement | undefined {
  return s.selection.length === 1 && tableOf(s.selection[0]!) ? s.selection[0] : undefined;
}

/** The text bodies formatting applies to in a table: the selected cells', or every cell's when none is picked. */
export function tableBodies(el: XmlElement, cells: Cells | null): XmlElement[] {
  const tbl = tableOf(el);
  if (!tbl) return [];
  const { rows, cols } = tableSize(tbl);
  const g = cells ? rangeOf(cells) : { r0: 0, c0: 0, r1: rows - 1, c1: cols - 1 };
  const out = new Set<XmlElement>();
  for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) { const b = cellBody(tbl, r, c); if (b) out.add(b); }
  return [...out];
}

function ctx() {
  const s = useEditor.getState();
  const el = selectedTable(s);
  if (!s.session || !s.slide || !el) return undefined;
  const frame = new ShapeView(el, s.session.ed.doc.owner(s.slide));
  const tbl = tableOf(el)!;
  const cells = activeCells(s);
  return { s, ed: s.session.ed, part: s.slide, el, frame, tbl, cells, range: cells ? rangeOf(cells) : undefined, size: tableSize(tbl) };
}

const pick = (el: XmlElement, r: number, c: number) => useEditor.setState({ cells: { el, anchor: { r, c }, focus: { r, c }, editing: false } });

/** What the table commands can do now, for the ribbon and the menu to enable. */
export function tableAbilities(): { table: boolean; cells: boolean; merge: boolean; split: boolean; deleteRow: boolean; deleteColumn: boolean } {
  const c = ctx();
  if (!c) return { table: false, cells: false, merge: false, split: false, deleteRow: false, deleteColumn: false };
  const g = c.range;
  return {
    table: true, cells: !!g,
    merge: !!g && (g.r1 > g.r0 || g.c1 > g.c0) && !(isMerged(c.tbl, g.r0, g.c0) && sameRange(regionOf(c.tbl, g.r0, g.c0), g)),
    split: !!g && isMerged(c.tbl, c.cells!.anchor.r, c.cells!.anchor.c),
    deleteRow: !!g && g.r1 - g.r0 + 1 < c.size.rows,
    deleteColumn: !!g && g.c1 - g.c0 + 1 < c.size.cols,
  };
}
const sameRange = (a: CellRange, b: CellRange) => a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;

/** Insert a row above or below the selected cells; the selection stays on the cells it was on. */
export function insertRowBeside(where: 'above' | 'below'): void {
  const c = ctx();
  if (!c?.range || !c.cells) return;
  const at = where === 'above' ? c.range.r0 : c.range.r1 + 1;
  c.s.run(insertRow(c.ed, c.part, c.frame, at, where === 'above' ? c.range.r0 : c.range.r1));
  if (where === 'above') useEditor.setState({ cells: { ...c.cells, anchor: { ...c.cells.anchor, r: c.cells.anchor.r + 1 }, focus: { ...c.cells.focus, r: c.cells.focus.r + 1 } } });
}

/** Insert a column left or right of the selected cells; the selection stays on the cells it was on. */
export function insertColumnBeside(where: 'left' | 'right'): void {
  const c = ctx();
  if (!c?.range || !c.cells) return;
  const at = where === 'left' ? c.range.c0 : c.range.c1 + 1;
  c.s.run(insertColumn(c.ed, c.part, c.frame, at, where === 'left' ? c.range.c0 : c.range.c1));
  if (where === 'left') useEditor.setState({ cells: { ...c.cells, anchor: { ...c.cells.anchor, c: c.cells.anchor.c + 1 }, focus: { ...c.cells.focus, c: c.cells.focus.c + 1 } } });
}

/** Delete the rows the selected cells are in; the cell that takes their place is selected. */
export function deleteSelectedRows(): void {
  const c = ctx();
  if (!c?.range || c.range.r1 - c.range.r0 + 1 >= c.size.rows) return;
  c.s.run(deleteRows(c.ed, c.part, c.frame, c.range.r0, c.range.r1));
  pick(c.el, Math.min(c.range.r0, tableSize(c.tbl).rows - 1), c.range.c0);
}

export function deleteSelectedColumns(): void {
  const c = ctx();
  if (!c?.range || c.range.c1 - c.range.c0 + 1 >= c.size.cols) return;
  c.s.run(deleteColumns(c.ed, c.part, c.frame, c.range.c0, c.range.c1));
  pick(c.el, c.range.r0, Math.min(c.range.c0, tableSize(c.tbl).cols - 1));
}

export function mergeSelectedCells(): void {
  const c = ctx();
  if (!c?.range) return;
  c.s.run(mergeCells(c.ed, c.part, c.frame, c.range));
  pick(c.el, c.range.r0, c.range.c0);
}

export function splitSelectedCell(): void {
  const c = ctx();
  if (!c?.cells) return;
  const g = regionOf(c.tbl, c.cells.anchor.r, c.cells.anchor.c);
  c.s.run(splitCell(c.ed, c.part, c.frame, c.cells.anchor));
  pick(c.el, g.r0, g.c0);
}

/** Delete on selected cells empties them (the table stays), as one undo step. */
export function clearSelectedCells(): void {
  const c = ctx();
  if (!c?.range) return;
  const cmds: Command[] = [];
  const seen = new Set<XmlElement>();
  for (let r = c.range.r0; r <= c.range.r1; r++) {
    for (let col = c.range.c0; col <= c.range.c1; col++) {
      const body = cellBody(c.tbl, r, col);
      if (!body || seen.has(body) || !cellText(c.tbl, r, col)) continue;
      seen.add(body);
      cmds.push(setCellText(c.ed, c.part, c.frame, { r, c: col }, ''));
    }
  }
  if (cmds.length) c.s.run(cmds.length === 1 ? { ...cmds[0]!, label: 'Clear Cells' } : compound('Clear Cells', cmds));
}

/** Move the selected cell with the arrow keys; Shift moves only the far corner, growing or shrinking the block. */
export function moveCell(dr: number, dc: number, extend: boolean): void {
  const c = ctx();
  if (!c?.cells) return;
  const from = c.cells.focus;
  // step over the rest of a merged cell, so one press always reaches the next cell
  const g = regionOf(c.tbl, from.r, from.c);
  const r = Math.max(0, Math.min(c.size.rows - 1, dr > 0 ? g.r1 + 1 : dr < 0 ? g.r0 - 1 : from.r));
  const col = Math.max(0, Math.min(c.size.cols - 1, dc > 0 ? g.c1 + 1 : dc < 0 ? g.c0 - 1 : from.c));
  useEditor.setState({ cells: { el: c.el, anchor: extend ? c.cells.anchor : { r, c: col }, focus: { r, c: col }, editing: false } });
}

/**
 * Tab goes to the next cell, row after row, and from the last cell adds a row and goes into it, as PowerPoint does;
 * Shift+Tab goes back. `editing` carries on typing in the cell arrived at.
 */
export function tabCell(back: boolean, editing: boolean): void {
  commitCellDraft();                                   // before a row may be added: the typing is the earlier undo step
  const c = ctx();
  if (!c?.cells) return;
  const g = regionOf(c.tbl, c.cells.focus.r, c.cells.focus.c);
  let r = g.r0, col = back ? g.c0 - 1 : g.c1 + 1;
  // the next cell that is the start of a cell (not the covered part of a merge)
  for (;;) {
    if (col < 0) { r--; col = c.size.cols - 1; }
    if (col >= c.size.cols) { r++; col = 0; }
    if (r < 0) return;
    if (r >= c.size.rows) {
      c.s.run(insertRow(c.ed, c.part, c.frame, c.size.rows));
      break;
    }
    const at = regionOf(c.tbl, r, col);
    if (at.r0 === r && at.c0 === col) break;
    col += back ? -1 : 1;
  }
  useEditor.setState({ cells: { el: c.el, anchor: { r, c: col }, focus: { r, c: col }, editing } });
}

/** Start typing in the selected cell; `seed` is a character already typed, added after the cell's text. */
export function editCell(seed = ''): boolean {
  const c = ctx();
  if (!c?.cells) return false;
  const g = regionOf(c.tbl, c.cells.focus.r, c.cells.focus.c);
  useEditor.setState({ cells: { el: c.el, anchor: { r: g.r0, c: g.c0 }, focus: { r: g.r0, c: g.c0 }, editing: true, ...(seed ? { seed } : {}) } });
  return true;
}

/**
 * What is being typed in a cell and not yet written to it. The cell editor keeps this up to date; it is written
 * whenever typing in that cell ends, however it ends (Escape, Tab, a click on another cell or shape, another
 * slide), by the subscription below, so nothing typed is ever lost.
 */
let draft: { el: XmlElement; part: string; at: CellRef; text: string } | null = null;

export function setCellDraft(el: XmlElement, at: CellRef, text: string): void {
  const part = useEditor.getState().slide;
  if (part) draft = { el, part, at, text };
}

/** Write the text being typed into its cell, when it differs from what the cell holds. */
export function commitCellDraft(): void {
  const d = draft;
  draft = null;
  const s = useEditor.getState();
  const tbl = d ? tableOf(d.el) : undefined;
  if (!d || !s.session || !tbl || !cellAt(tbl, d.at.r, d.at.c) || !s.session.ed.doc.slideParts().includes(d.part) || cellText(tbl, d.at.r, d.at.c) === d.text) return;
  s.run(setCellText(s.session.ed, d.part, new ShapeView(d.el, s.session.ed.doc.owner(d.part)), d.at, d.text));
}

const sameCell = (a: Cells | null, b: Cells | null) => !!a && !!b && a.el === b.el && a.anchor.r === b.anchor.r && a.anchor.c === b.anchor.c;
useEditor.subscribe((st, prev) => {
  if (draft && prev.cells?.editing && !(st.cells?.editing && sameCell(st.cells, prev.cells))) commitCellDraft();
});
