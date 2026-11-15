// packages/editor-core/src/tables.ts  -- tables: inserting one, cell text, rows and columns, merging and splitting
import { encodeXmlText, type XmlElement } from '@pptx/xml';
import { attr, child, children, clone, int, relType, setAttr, type ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { topSlot, type Rect } from './shapes.ts';
import { bodyPlainText, paragraphs, replaceText, type TextOptions } from './text.ts';
import { element, insertAt, nextShapeId } from './xml.ts';

/** PowerPoint's default for a new table: Medium Style 2 - Accent 1. */
export const DEFAULT_TABLE_STYLE = '{5C22544A-7EE6-4342-B048-85BB5DBAABE0}';
const TABLE_STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml';
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const TABLE_URI_STRICT = 'http://purl.oclc.org/ooxml/drawingml/table';
/** The height PowerPoint gives a new row of 18 pt text. */
export const ROW_HEIGHT = 370840;

const border = (edge: string, w = 12700) => `<a:${edge}><a:ln w="${w}" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:${edge}>`;
const accent = (tint?: number) => `<a:fill><a:solidFill>${tint ? `<a:schemeClr val="accent1"><a:tint val="${tint}"/></a:schemeClr>` : '<a:schemeClr val="accent1"/>'}</a:solidFill></a:fill>`;
const bold = '<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle>';

/** The definition of Medium Style 2 - Accent 1, as PowerPoint writes it into tableStyles.xml when a table uses it. */
const DEFAULT_STYLE_XML = `<a:tblStyle styleId="${DEFAULT_TABLE_STYLE}" styleName="Medium Style 2 - Accent 1">`
  + '<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle>'
  + `<a:tcStyle><a:tcBdr>${['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map(e => border(e)).join('')}</a:tcBdr>${accent(20000)}</a:tcStyle></a:wholeTbl>`
  + `<a:band1H><a:tcStyle><a:tcBdr/>${accent(40000)}</a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>`
  + `<a:band1V><a:tcStyle><a:tcBdr/>${accent(40000)}</a:tcStyle></a:band1V><a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V>`
  + `<a:lastCol>${bold}<a:tcStyle><a:tcBdr/>${accent()}</a:tcStyle></a:lastCol><a:firstCol>${bold}<a:tcStyle><a:tcBdr/>${accent()}</a:tcStyle></a:firstCol>`
  + `<a:lastRow>${bold}<a:tcStyle><a:tcBdr>${border('top', 38100)}</a:tcBdr>${accent()}</a:tcStyle></a:lastRow>`
  + `<a:firstRow>${bold}<a:tcStyle><a:tcBdr>${border('bottom', 38100)}</a:tcBdr>${accent()}</a:tcStyle></a:firstRow></a:tblStyle>`;

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const A_NS_STRICT = 'http://purl.oclc.org/ooxml/drawingml/main';

/** Make sure tableStyles.xml exists and defines the default style. Returns the part; runs inside a structural edit. */
function ensureDefaultStyle(ed: Editor): string {
  const existing = ed.doc.tableStylesPart();
  if (!existing) {
    const part = '/ppt/tableStyles.xml';
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
      + `<a:tblStyleLst xmlns:a="${ed.strict ? A_NS_STRICT : A_NS}" def="${DEFAULT_TABLE_STYLE}">${DEFAULT_STYLE_XML}</a:tblStyleLst>`;
    ed.pkg.addPart(part, TABLE_STYLES_TYPE, encodeXmlText(xml));
    ed.pkg.addRelationship(ed.doc.mainPart, relType('tableStyles', ed.strict), part);
    return part;
  }
  const root = ed.doc.root(existing);
  if (!children(root, 'a:tblStyle').some(s => attr(s, 'styleId') === DEFAULT_TABLE_STYLE)) {
    insertAt(root, root.children.length, element(root, DEFAULT_STYLE_XML, ed.strict));
    ed.pkg.markDirty(existing);
  }
  return existing;
}

const emptyCell = (extra = '') => `<a:tc${extra}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>`;

/**
 * Insert a table as PowerPoint does: a graphic frame of `rows` by `cols` empty cells in equal columns, with the
 * header row and banded rows on and the default table style (whose definition goes into tableStyles.xml).
 */
export function insertTable(ed: Editor, part: string, rows: number, cols: number, box: Rect): { command: Command; id: number } {
  if (!(rows >= 1 && cols >= 1 && rows <= 75 && cols <= 75)) throw new Error('a table has 1 to 75 rows and columns');
  const tree = ed.spTree(part);
  const id = nextShapeId(ed.doc.root(part));
  const r = Math.round;
  const colW = r(box.w / cols);
  const grid = Array.from({ length: cols }, () => `<a:gridCol w="${colW}"/>`).join('');
  const row = `<a:tr h="${ROW_HEIGHT}">${Array.from({ length: cols }, () => emptyCell()).join('')}</a:tr>`;
  const xml = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id - 1}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
    + `<p:xfrm><a:off x="${r(box.x)}" y="${r(box.y)}"/><a:ext cx="${colW * cols}" cy="${ROW_HEIGHT * rows}"/></p:xfrm>`
    + `<a:graphic><a:graphicData uri="${ed.strict ? TABLE_URI_STRICT : TABLE_URI}"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE}</a:tableStyleId></a:tblPr>`
    + `<a:tblGrid>${grid}</a:tblGrid>${Array.from({ length: rows }, () => row).join('')}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
  const styles = ed.doc.tableStylesPart() ?? '/ppt/tableStyles.xml';
  const command = ed.edit({
    label: 'Insert Table', parts: [part, ed.doc.mainPart, styles], structural: true,
    scope: () => [tree, ...(ed.doc.tableStylesPart() ? [ed.doc.root(ed.doc.tableStylesPart()!)] : [])],
    apply: () => {
      ensureDefaultStyle(ed);
      insertAt(tree, topSlot(tree), element(tree, xml, ed.strict));
    },
  });
  return { command, id };
}

/** The a:tbl of a graphic frame, when the frame holds a table. */
export function tableOf(shape: ShapeView | XmlElement): XmlElement | undefined {
  const el = 'el' in shape ? shape.el : shape;
  return child(child(child(el, 'a:graphic'), 'a:graphicData'), 'a:tbl');
}

export interface CellRef { r: number; c: number; }
/** A rectangle of cells, edges included. */
export interface CellRange { r0: number; c0: number; r1: number; c1: number; }

const rowsOf = (tbl: XmlElement) => children(tbl, 'a:tr');
const cellsOf = (tr: XmlElement) => children(tr, 'a:tc');
const span = (tc: XmlElement, name: 'gridSpan' | 'rowSpan') => Math.max(1, int(attr(tc, name), 1));
const flagged = (tc: XmlElement, name: 'hMerge' | 'vMerge') => attr(tc, name) === '1' || attr(tc, name) === 'true';

/** Rows and columns of a table. */
export function tableSize(tbl: XmlElement): { rows: number; cols: number } {
  return { rows: rowsOf(tbl).length, cols: children(child(tbl, 'a:tblGrid'), 'a:gridCol').length };
}

/** The cell element at a row and column (a covered cell of a merge included). */
export function cellAt(tbl: XmlElement, r: number, c: number): XmlElement | undefined {
  const tr = rowsOf(tbl)[r];
  return tr ? cellsOf(tr)[c] : undefined;
}

/**
 * The merged region a cell belongs to: the origin cell (top left, which holds the text and the spans) and how far
 * it reaches. A cell that is not merged is its own region of one.
 */
export function regionOf(tbl: XmlElement, r: number, c: number): CellRange {
  let r0 = r, c0 = c;
  // walk left, then up, to the cell that is covered by neither kind of merge
  while (c0 > 0 && flagged(cellAt(tbl, r0, c0)!, 'hMerge')) c0--;
  while (r0 > 0 && flagged(cellAt(tbl, r0, c0)!, 'vMerge')) r0--;
  const origin = cellAt(tbl, r0, c0);
  if (!origin) return { r0: r, c0: c, r1: r, c1: c };
  return { r0, c0, r1: r0 + span(origin, 'rowSpan') - 1, c1: c0 + span(origin, 'gridSpan') - 1 };
}

/** The text body of the cell that shows at a position (the origin of its merged region). */
export function cellBody(tbl: XmlElement, r: number, c: number): XmlElement | undefined {
  const g = regionOf(tbl, r, c);
  return child(cellAt(tbl, g.r0, g.c0), 'a:txBody');
}

export function cellText(tbl: XmlElement, r: number, c: number): string {
  const body = cellBody(tbl, r, c);
  return body ? bodyPlainText(body).replace(/\u000b/g, '\n') : '';
}

/** A range grown until it holds every merged region it touches whole. */
export function wholeRegions(tbl: XmlElement, range: CellRange): CellRange {
  let g = { r0: Math.min(range.r0, range.r1), c0: Math.min(range.c0, range.c1), r1: Math.max(range.r0, range.r1), c1: Math.max(range.c0, range.c1) };
  for (;;) {
    let grown = g;
    for (let r = g.r0; r <= g.r1; r++) {
      for (let c = g.c0; c <= g.c1; c++) {
        const x = regionOf(tbl, r, c);
        grown = { r0: Math.min(grown.r0, x.r0), c0: Math.min(grown.c0, x.c0), r1: Math.max(grown.r1, x.r1), c1: Math.max(grown.c1, x.c1) };
      }
    }
    if (grown.r0 === g.r0 && grown.c0 === g.c0 && grown.r1 === g.r1 && grown.c1 === g.c1) return g;
    g = grown;
  }
}

function tableFor(frame: ShapeView): XmlElement {
  const tbl = tableOf(frame);
  if (!tbl) throw new Error('this shape is not a table');
  return tbl;
}

/** Set a cell's text (plain; lines become paragraphs), keeping the formatting of its first run. */
export function setCellText(ed: Editor, part: string, frame: ShapeView, at: CellRef, text: string, options: TextOptions = {}): Command {
  const body = cellBody(tableFor(frame), at.r, at.c);
  if (!body) throw new Error('no such cell');
  return replaceText(ed, part, body, text, { label: 'Typing', ...options });
}

/** A new empty cell that looks like `like`: its cell properties and paragraph mark, none of its merging or text. */
function cellLike(parent: XmlElement, like: XmlElement | undefined, ed: Editor, merge = ''): XmlElement {
  const tc = element(parent, emptyCell(merge), ed.strict);
  if (!like) return tc;
  const tcPr = child(like, 'a:tcPr');
  if (tcPr) {
    const copy = clone(tcPr, tc);
    tc.children[tc.children.indexOf(child(tc, 'a:tcPr')!)] = copy;
  }
  const likeBody = child(like, 'a:txBody');
  const firstP = likeBody ? paragraphs(likeBody)[0] : undefined;
  const mark = firstP ? child(firstP, 'a:endParaRPr') ?? child(children(firstP, 'a:r')[0], 'a:rPr') : undefined;
  const pPr = firstP ? child(firstP, 'a:pPr') : undefined;
  const p = paragraphs(child(tc, 'a:txBody')!)[0]!;
  if (mark) {
    const m = clone(mark, p);
    m.local = 'endParaRPr';
    p.children[p.children.indexOf(child(p, 'a:endParaRPr')!)] = m;
  }
  if (pPr) insertAt(p, 0, clone(pPr, p));
  return tc;
}

/** Grow or shrink the frame with its table, so the selection box stays around what is drawn. */
function resizeFrame(frame: ShapeView, dw: number, dh: number): void {
  const x = frame.xfrm;
  if (x) frame.setXfrm({ ...x, cx: Math.max(0, x.cx + dw), cy: Math.max(0, x.cy + dh) });
}

/**
 * Insert a row at index `at` (0 puts it first; the row count puts it last). It looks like the row next to it:
 * `like` says which (the row above by default, the row below for the first). A merge that runs across the place
 * the row goes in is made one row taller, so it stays one cell.
 */
export function insertRow(ed: Editor, part: string, frame: ShapeView, at: number, like?: number): Command {
  const tbl = tableFor(frame);
  return ed.edit({
    label: 'Insert Row', parts: [part], scope: () => [frame.el],
    apply: () => {
      const rows = rowsOf(tbl);
      const { cols } = tableSize(tbl);
      const index = Math.max(0, Math.min(at, rows.length));
      const ref = rows[Math.max(0, Math.min(rows.length - 1, like ?? (index > 0 ? index - 1 : 0)))]!;
      const tr = element(tbl, `<a:tr h="${attr(ref, 'h') ?? ROW_HEIGHT}"/>`, ed.strict);
      tr.selfClosing = false;
      const grown = new Set<XmlElement>();
      for (let c = 0; c < cols; c++) {
        // a region that has rows on both sides of the new row passes through it
        const above = index > 0 ? regionOf(tbl, index - 1, c) : undefined;
        const through = above && above.r1 >= index ? above : undefined;
        let tc: XmlElement;
        if (through) {
          const origin = cellAt(tbl, through.r0, through.c0)!;
          if (!grown.has(origin)) {
            setAttr(origin, 'rowSpan', span(origin, 'rowSpan') + 1);
            grown.add(origin);
          }
          const w = through.c1 - through.c0 + 1;
          tc = cellLike(tr, cellsOf(ref)[c], ed, c === through.c0 ? `${w > 1 ? ` gridSpan="${w}"` : ''} vMerge="1"` : ' hMerge="1" vMerge="1"');
        } else tc = cellLike(tr, cellsOf(ref)[c], ed);
        tr.children.push(tc);
        tc.parent = tr;
      }
      const after = rows[index - 1];
      insertAt(tbl, after ? tbl.children.indexOf(after) + 1 : tbl.children.indexOf(rows[0]!), tr);
      resizeFrame(frame, 0, int(attr(ref, 'h'), ROW_HEIGHT));
    },
  });
}

/** Move a region's origin (its text and properties) one cell on when its first row or column is being removed. */
function moveOrigin(from: XmlElement, to: XmlElement): void {
  for (const name of ['a:txBody', 'a:tcPr']) {
    const mine = child(to, name), theirs = child(from, name);
    if (!theirs) continue;
    const copy = clone(theirs, to);
    if (mine) to.children[to.children.indexOf(mine)] = copy;
    else to.children.push(copy);
  }
}

/** Delete the rows `from` to `to` (inclusive). A merge that reaches outside them gets shorter and keeps its text. */
export function deleteRows(ed: Editor, part: string, frame: ShapeView, from: number, to = from): Command {
  const tbl = tableFor(frame);
  return ed.edit({
    label: 'Delete Row', parts: [part], scope: () => [frame.el],
    apply: () => {
      const { rows: n, cols } = tableSize(tbl);
      const a = Math.max(0, Math.min(from, to)), b = Math.min(n - 1, Math.max(from, to));
      if (b - a + 1 >= n) throw new Error('a table needs at least one row');
      // one row at a time from the bottom, so indexes above stay true
      for (let r = b; r >= a; r--) {
        const seen = new Set<XmlElement>();
        for (let c = 0; c < cols; c++) {
          const g = regionOf(tbl, r, c);
          const origin = cellAt(tbl, g.r0, g.c0)!;
          if (g.r1 === g.r0 || seen.has(origin)) continue;
          seen.add(origin);
          if (g.r0 === r) {
            // the region starts in this row: the cell below becomes its origin, one row shorter
            const next = cellAt(tbl, r + 1, g.c0)!;
            moveOrigin(origin, next);
            setAttr(next, 'vMerge', undefined);
            setAttr(next, 'rowSpan', g.r1 - g.r0 > 1 ? g.r1 - g.r0 : undefined);
            for (let cc = g.c0 + 1; cc <= g.c1; cc++) setAttr(cellAt(tbl, r + 1, cc)!, 'vMerge', undefined);
          } else setAttr(origin, 'rowSpan', g.r1 - g.r0 > 1 ? g.r1 - g.r0 : undefined);
        }
        const tr = rowsOf(tbl)[r]!;
        resizeFrame(frame, 0, -int(attr(tr, 'h'), ROW_HEIGHT));
        tbl.children.splice(tbl.children.indexOf(tr), 1);
      }
    },
  });
}

/** Insert a column at index `at`, as wide as the column next to it and looking like it; merges across it grow. */
export function insertColumn(ed: Editor, part: string, frame: ShapeView, at: number, like?: number): Command {
  const tbl = tableFor(frame);
  return ed.edit({
    label: 'Insert Column', parts: [part], scope: () => [frame.el],
    apply: () => {
      const grid = child(tbl, 'a:tblGrid');
      if (!grid) throw new Error('the table has no grid');
      const gridCols = children(grid, 'a:gridCol');
      const index = Math.max(0, Math.min(at, gridCols.length));
      const refIndex = Math.max(0, Math.min(gridCols.length - 1, like ?? (index > 0 ? index - 1 : 0)));
      const w = int(attr(gridCols[refIndex], 'w'), 914400);
      const rows = rowsOf(tbl);
      // decide every row's new cell before changing any, while the regions are still as they were
      const plans = rows.map((_, r) => {
        const left = index > 0 ? regionOf(tbl, r, index - 1) : undefined;
        return left && left.c1 >= index ? left : undefined;
      });
      const grown = new Set<XmlElement>();
      rows.forEach((tr, r) => {
        const through = plans[r];
        const cells = cellsOf(tr);
        let tc: XmlElement;
        if (through) {
          const origin = cellAt(tbl, through.r0, through.c0)!;
          if (!grown.has(origin)) {
            setAttr(origin, 'gridSpan', span(origin, 'gridSpan') + 1);
            grown.add(origin);
          }
          tc = cellLike(tr, cells[refIndex], ed, r === through.r0 ? ' hMerge="1"' : ' hMerge="1" vMerge="1"');
        } else tc = cellLike(tr, cells[refIndex], ed);
        const before = cells[index];
        insertAt(tr, before ? tr.children.indexOf(before) : (cells.length ? tr.children.indexOf(cells[cells.length - 1]!) + 1 : tr.children.length), tc);
      });
      // rows below the first of a region that grew carry its width too
      for (const origin of grown) {
        const g = regionOfCell(tbl, origin);
        for (let r = g.r0 + 1; r <= g.r1; r++) setAttr(cellAt(tbl, r, g.c0)!, 'gridSpan', span(origin, 'gridSpan'));
      }
      const col = element(grid, `<a:gridCol w="${w}"/>`, ed.strict);
      const beforeCol = gridCols[index];
      insertAt(grid, beforeCol ? grid.children.indexOf(beforeCol) : (gridCols.length ? grid.children.indexOf(gridCols[gridCols.length - 1]!) + 1 : grid.children.length), col);
      resizeFrame(frame, w, 0);
    },
  });
}

/** Where a cell element is in its table. */
function positionOf(tbl: XmlElement, tc: XmlElement): CellRef {
  const rows = rowsOf(tbl);
  const r = rows.indexOf(tc.parent!);
  return { r, c: cellsOf(rows[r]!).indexOf(tc) };
}
const regionOfCell = (tbl: XmlElement, tc: XmlElement): CellRange => { const p = positionOf(tbl, tc); return regionOf(tbl, p.r, p.c); };

/** Delete the columns `from` to `to` (inclusive). A merge that reaches outside them gets narrower and keeps its text. */
export function deleteColumns(ed: Editor, part: string, frame: ShapeView, from: number, to = from): Command {
  const tbl = tableFor(frame);
  return ed.edit({
    label: 'Delete Column', parts: [part], scope: () => [frame.el],
    apply: () => {
      const { rows: n, cols } = tableSize(tbl);
      const a = Math.max(0, Math.min(from, to)), b = Math.min(cols - 1, Math.max(from, to));
      if (b - a + 1 >= cols) throw new Error('a table needs at least one column');
      const grid = child(tbl, 'a:tblGrid')!;
      for (let c = b; c >= a; c--) {
        const seen = new Set<XmlElement>();
        for (let r = 0; r < n; r++) {
          const g = regionOf(tbl, r, c);
          const origin = cellAt(tbl, g.r0, g.c0)!;
          if (g.c1 === g.c0 || seen.has(origin)) continue;
          seen.add(origin);
          const width = g.c1 - g.c0;                    // the region's width once this column is gone
          if (g.c0 === c) {
            // the region starts in this column: the cell to the right becomes its origin
            const next = cellAt(tbl, g.r0, c + 1)!;
            moveOrigin(origin, next);
            setAttr(next, 'hMerge', undefined);
            setAttr(next, 'gridSpan', width > 1 ? width : undefined);
            setAttr(next, 'rowSpan', g.r1 > g.r0 ? g.r1 - g.r0 + 1 : undefined);
            for (let rr = g.r0 + 1; rr <= g.r1; rr++) {
              const below = cellAt(tbl, rr, c + 1)!;
              setAttr(below, 'hMerge', undefined);
              setAttr(below, 'gridSpan', width > 1 ? width : undefined);
            }
          } else {
            for (let rr = g.r0; rr <= g.r1; rr++) setAttr(cellAt(tbl, rr, g.c0)!, 'gridSpan', width > 1 ? width : undefined);
          }
        }
        for (const tr of rowsOf(tbl)) {
          const tc = cellsOf(tr)[c];
          if (tc) tr.children.splice(tr.children.indexOf(tc), 1);
        }
        const col = children(grid, 'a:gridCol')[c]!;
        resizeFrame(frame, -int(attr(col, 'w'), 0), 0);
        grid.children.splice(grid.children.indexOf(col), 1);
      }
    },
  });
}

const hasText = (body: XmlElement | undefined) => !!body && bodyPlainText(body).length > 0;

/**
 * Merge a rectangle of cells into one (grown first to hold whole any merge it touches). The top left cell keeps
 * its text and takes the others' as further paragraphs, as PowerPoint does; the others become its covered cells.
 */
export function mergeCells(ed: Editor, part: string, frame: ShapeView, range: CellRange): Command {
  const tbl = tableFor(frame);
  return ed.edit({
    label: 'Merge Cells', parts: [part], scope: () => [frame.el],
    apply: () => {
      const g = wholeRegions(tbl, range);
      if (g.r0 === g.r1 && g.c0 === g.c1) return;
      const origin = cellAt(tbl, g.r0, g.c0)!;
      const body = child(origin, 'a:txBody');
      const w = g.c1 - g.c0 + 1, h = g.r1 - g.r0 + 1;
      for (let r = g.r0; r <= g.r1; r++) {
        for (let c = g.c0; c <= g.c1; c++) {
          const tc = cellAt(tbl, r, c)!;
          if (tc === origin) continue;
          const other = child(tc, 'a:txBody');
          // text of cells that showed text goes to the merged cell, after its own
          if (body && other && !flagged(tc, 'hMerge') && !flagged(tc, 'vMerge') && hasText(other)) {
            if (!hasText(body)) for (const p of paragraphs(body)) body.children.splice(body.children.indexOf(p), 1);
            for (const p of paragraphs(other)) {
              const copy = clone(p, body);
              body.children.push(copy);
            }
          }
          if (other) {
            for (const p of paragraphs(other)) other.children.splice(other.children.indexOf(p), 1);
            insertAt(other, other.children.length, element(other, '<a:p><a:endParaRPr lang="en-US"/></a:p>', ed.strict));
          }
          setAttr(tc, 'rowSpan', undefined);
          setAttr(tc, 'gridSpan', c === g.c0 && w > 1 ? w : undefined);
          setAttr(tc, 'hMerge', c > g.c0 ? '1' : undefined);
          setAttr(tc, 'vMerge', r > g.r0 ? '1' : undefined);
        }
      }
      setAttr(origin, 'gridSpan', w > 1 ? w : undefined);
      setAttr(origin, 'rowSpan', h > 1 ? h : undefined);
      setAttr(origin, 'hMerge', undefined);
      setAttr(origin, 'vMerge', undefined);
    },
  });
}

/** Split a merged cell back into the cells it covers; the text stays in the top left one. */
export function splitCell(ed: Editor, part: string, frame: ShapeView, at: CellRef): Command {
  const tbl = tableFor(frame);
  return ed.edit({
    label: 'Split Cell', parts: [part], scope: () => [frame.el],
    apply: () => {
      const g = regionOf(tbl, at.r, at.c);
      for (let r = g.r0; r <= g.r1; r++) {
        for (let c = g.c0; c <= g.c1; c++) {
          const tc = cellAt(tbl, r, c)!;
          for (const name of ['gridSpan', 'rowSpan', 'hMerge', 'vMerge']) setAttr(tc, name, undefined);
        }
      }
    },
  });
}

/** Whether a cell is part of a merge (so Split Cell has something to do). */
export function isMerged(tbl: XmlElement, r: number, c: number): boolean {
  const g = regionOf(tbl, r, c);
  return g.r1 > g.r0 || g.c1 > g.c0;
}
