// packages/model/src/table.ts  -- a:tbl inside a graphic frame (ECMA-376 Part 1, 21.1.3)
import type { XmlElement } from '@pptx/xml';
import { attr, child, children, textOf } from './dom.ts';
import { bool, coord, int } from './units.ts';

export interface TableCell {
  el: XmlElement;                // a:tc
  gridSpan: number;              // columns this cell covers, from its own column
  rowSpan: number;
  hMerge: boolean;               // covered by a cell to the left (whose gridSpan reaches here)
  vMerge: boolean;               // covered by a cell above
  txBody?: XmlElement;
  tcPr?: XmlElement;
}
export interface TableRow { el: XmlElement; h: number; cells: TableCell[]; }
export interface TableFlags { firstRow: boolean; firstCol: boolean; lastRow: boolean; lastCol: boolean; bandRow: boolean; bandCol: boolean; rtl: boolean; }
export interface TableDef {
  el: XmlElement;
  tblPr?: XmlElement;
  styleId?: string;              // a:tableStyleId: a GUID naming a style in tableStyles.xml or a built-in one
  flags: TableFlags;
  columns: number[];             // widths in EMU
  rows: TableRow[];
}

export function parseTable(tbl: XmlElement): TableDef {
  const tblPr = child(tbl, 'a:tblPr');
  const styleId = textOf(child(tblPr, 'a:tableStyleId')).trim();
  const flag = (k: string) => bool(attr(tblPr, k), false);
  return {
    el: tbl,
    ...(tblPr ? { tblPr } : {}),
    ...(styleId ? { styleId } : {}),
    flags: { firstRow: flag('firstRow'), firstCol: flag('firstCol'), lastRow: flag('lastRow'), lastCol: flag('lastCol'), bandRow: flag('bandRow'), bandCol: flag('bandCol'), rtl: flag('rtl') },
    columns: children(child(tbl, 'a:tblGrid'), 'a:gridCol').map(c => coord(attr(c, 'w'), 0)),
    rows: children(tbl, 'a:tr').map(tr => ({
      el: tr,
      h: coord(attr(tr, 'h'), 0),
      cells: children(tr, 'a:tc').map(tc => {
        const txBody = child(tc, 'a:txBody');
        const tcPr = child(tc, 'a:tcPr');
        return {
          el: tc, gridSpan: Math.max(1, int(attr(tc, 'gridSpan'), 1)), rowSpan: Math.max(1, int(attr(tc, 'rowSpan'), 1)),
          hMerge: bool(attr(tc, 'hMerge'), false), vMerge: bool(attr(tc, 'vMerge'), false),
          ...(txBody ? { txBody } : {}), ...(tcPr ? { tcPr } : {}),
        };
      }),
    })),
  };
}
