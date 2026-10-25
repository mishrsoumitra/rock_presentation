// packages/resolve/src/tables.ts  -- tables with their table style applied (ECMA-376 Part 1, 20.1.4.2 and 21.1.3)
import type { XmlElement } from '@pptx/xml';
import {
  attr, child, children, colorIn, coord, fillIn, parseLine, parseTable, textOf,
  type ColorDef, type FillDef, type LineDef, type RunPropsDef, type TableDef,
} from '@pptx/model';
import type { SlideContext } from './context.ts';
import { scopeFor, themeScope } from './context.ts';
import {
  resolveFill, resolveLine, themeFill, themeLine, withPlaceholderColor, type Paint, type PaintScope, type Stroke,
} from './paint.ts';
import { resolveTextBody, type ResolvedTextBody } from './text.ts';

export interface ResolvedCell {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  /** Covered by a merged neighbour (hMerge or vMerge): drawn as part of that cell, not on its own. */
  covered: boolean;
  fill: Paint;
  borders: { l: Stroke | null; r: Stroke | null; t: Stroke | null; b: Stroke | null; tl2br: Stroke | null; bl2tr: Stroke | null };
  margins: { l: number; r: number; t: number; b: number };
  anchor: 't' | 'ctr' | 'b' | 'just' | 'dist';
  vert: string;
  text: ResolvedTextBody | null;
}

export interface ResolvedTable {
  columns: number[];
  /** Minimum row heights from the file; rows grow to fit their text when laid out. */
  rows: number[];
  cells: ResolvedCell[];
  background: Paint;
  rtl: boolean;
}

/** One part of a table style (wholeTbl, band1H, firstRow, ...), its colours unresolved. */
interface StylePart {
  fill?: FillDef | { ref: number; color?: ColorDef };
  borders: Partial<Record<'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV' | 'tl2br' | 'tr2bl', LineDef | { ref: number; color?: ColorDef } | 'none'>>;
  text: RunPropsDef;
}

function stylePart(el: XmlElement | undefined): StylePart | undefined {
  if (!el) return undefined;
  const out: StylePart = { borders: {}, text: {} };
  const tcStyle = child(el, 'a:tcStyle');
  const fillWrap = child(tcStyle, 'a:fill');
  const fillRef = child(tcStyle, 'a:fillRef');
  if (fillWrap) {
    const f = fillIn(fillWrap);
    if (f) out.fill = f;
  } else if (fillRef) {
    const color = colorIn(fillRef);
    out.fill = { ref: Number(attr(fillRef, 'idx') ?? 0), ...(color ? { color } : {}) };
  }
  for (const edge of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV', 'tl2br', 'tr2bl'] as const) {
    const b = child(child(tcStyle, 'a:tcBdr'), `a:${edge}`);
    if (!b) continue;
    const ln = child(b, 'a:ln');
    const lnRef = child(b, 'a:lnRef');
    if (ln) out.borders[edge] = parseLine(ln) ?? 'none';
    else if (lnRef) {
      const color = colorIn(lnRef);
      out.borders[edge] = { ref: Number(attr(lnRef, 'idx') ?? 0), ...(color ? { color } : {}) };
    }
  }
  const tx = child(el, 'a:tcTxStyle');
  if (tx) {
    const b = attr(tx, 'b');
    const i = attr(tx, 'i');
    if (b === 'on') out.text.b = true;
    else if (b === 'off') out.text.b = false;
    if (i === 'on') out.text.i = true;
    else if (i === 'off') out.text.i = false;
    const font = child(tx, 'a:font');
    const fontRef = child(tx, 'a:fontRef');
    if (font) {
      const latin = attr(child(font, 'a:latin'), 'typeface');
      if (latin) out.text.latin = { typeface: latin };
    } else if (fontRef) {
      const major = attr(fontRef, 'idx') === 'major';
      out.text.latin = { typeface: major ? '+mj-lt' : '+mn-lt' };
      out.text.ea = { typeface: major ? '+mj-ea' : '+mn-ea' };
      out.text.cs = { typeface: major ? '+mj-cs' : '+mn-cs' };
      const c = colorIn(fontRef);
      if (c) out.text.fill = { kind: 'solid', color: c };
    }
    const c = colorIn(tx);
    if (c) out.text.fill = { kind: 'solid', color: c };
  }
  return out;
}

/** The table style with this id from tableStyles.xml, as its parts in the order they apply. */
function tableStyle(ctx: SlideContext, styleId: string | undefined): Map<string, StylePart> {
  const parts = new Map<string, StylePart>();
  const stylesPart = ctx.doc.tableStylesPart();
  if (!stylesPart) return parts;
  const root = ctx.doc.root(stylesPart);
  const id = styleId ?? attr(root, 'def');
  const style = children(root, 'a:tblStyle').find(s => attr(s, 'styleId') === id);
  if (!style) return parts;
  for (const name of ['tblBg', 'wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'firstRow', 'seCell', 'swCell', 'neCell', 'nwCell']) {
    const p = stylePart(child(style, `a:${name}`));
    if (p) parts.set(name, p);
  }
  return parts;
}

function partFill(f: StylePart['fill'], ctx: SlideContext, scope: PaintScope): Paint | undefined {
  if (!f) return undefined;
  if ('ref' in f) return resolveFill(themeFill(ctx.theme, f.ref), withPlaceholderColor(themeScope(ctx), f));
  return resolveFill(f, scope);
}

function partLine(l: StylePart['borders'][keyof StylePart['borders']], ctx: SlideContext, scope: PaintScope): Stroke | null | undefined {
  if (l === undefined) return undefined;
  if (l === 'none') return null;
  if ('ref' in l) return resolveLine(themeLine(ctx.theme, l.ref), withPlaceholderColor(themeScope(ctx), l));
  return resolveLine(l, scope);
}

/** Cell edge lines from the cell's own properties (a:lnL and friends); undefined when not stated. */
function cellLine(tcPr: XmlElement | undefined, name: string, scope: PaintScope): Stroke | null | undefined {
  const el = child(tcPr, name);
  if (!el) return undefined;
  return resolveLine(parseLine(el), scope);
}

export function resolveTable(ctx: SlideContext, tbl: XmlElement, part: string): ResolvedTable {
  const t: TableDef = parseTable(tbl);
  const scope = scopeFor(ctx, part);
  const style = tableStyle(ctx, t.styleId);
  const nRows = t.rows.length;
  const nCols = t.columns.length;
  const f = t.flags;
  const cells: ResolvedCell[] = [];
  const bgPart = style.get('tblBg');

  t.rows.forEach((row, r) => {
    row.cells.forEach((cell, c) => {
      // Style parts that apply to this cell, in increasing precedence.
      const applied: StylePart[] = [];
      const add = (name: string, when: boolean) => {
        const p = style.get(name);
        if (when && p) applied.push(p);
      };
      add('wholeTbl', true);
      const bandRow = r - (f.firstRow ? 1 : 0);
      const bandCol = c - (f.firstCol ? 1 : 0);
      const inBodyRows = !(f.firstRow && r === 0) && !(f.lastRow && r === nRows - 1);
      const inBodyCols = !(f.firstCol && c === 0) && !(f.lastCol && c === nCols - 1);
      add(bandRow % 2 === 0 ? 'band1H' : 'band2H', f.bandRow && inBodyRows && bandRow >= 0);
      add(bandCol % 2 === 0 ? 'band1V' : 'band2V', f.bandCol && inBodyCols && bandCol >= 0);
      add('lastCol', f.lastCol && c === nCols - 1);
      add('firstCol', f.firstCol && c === 0);
      add('lastRow', f.lastRow && r === nRows - 1);
      add('firstRow', f.firstRow && r === 0);
      add('seCell', f.lastRow && f.lastCol && r === nRows - 1 && c === nCols - 1);
      add('swCell', f.lastRow && f.firstCol && r === nRows - 1 && c === 0);
      add('neCell', f.firstRow && f.lastCol && r === 0 && c === nCols - 1);
      add('nwCell', f.firstRow && f.firstCol && r === 0 && c === 0);

      let fill: Paint = { kind: 'none' };
      const text: RunPropsDef = {};
      const edges: Record<'l' | 'r' | 't' | 'b' | 'tl2br' | 'bl2tr', Stroke | null> = { l: null, r: null, t: null, b: null, tl2br: null, bl2tr: null };
      const lastRowOfCell = r + cell.rowSpan - 1;
      const lastColOfCell = c + cell.gridSpan - 1;
      for (const p of applied) {
        const pf = partFill(p.fill, ctx, scope);
        if (pf) fill = pf;
        Object.assign(text, p.text);
        const set = (edge: keyof typeof edges, outer: string, inner: string, isOuter: boolean) => {
          const v = partLine(p.borders[(isOuter ? outer : inner) as keyof StylePart['borders']], ctx, scope);
          if (v !== undefined) edges[edge] = v;
        };
        set('l', 'left', 'insideV', c === 0);
        set('r', 'right', 'insideV', lastColOfCell === nCols - 1);
        set('t', 'top', 'insideH', r === 0);
        set('b', 'bottom', 'insideH', lastRowOfCell === nRows - 1);
        const d1 = partLine(p.borders.tl2br, ctx, scope);
        if (d1 !== undefined) edges.tl2br = d1;
        const d2 = partLine(p.borders.tr2bl, ctx, scope);
        if (d2 !== undefined) edges.bl2tr = d2;
      }
      // The cell's own properties win.
      const tcPr = cell.tcPr;
      const own = fillIn(tcPr);
      if (own) fill = resolveFill(own, scope);
      for (const [edge, name] of [['l', 'a:lnL'], ['r', 'a:lnR'], ['t', 'a:lnT'], ['b', 'a:lnB'], ['tl2br', 'a:lnTlToBr'], ['bl2tr', 'a:lnBlToTr']] as const) {
        const v = cellLine(tcPr, name, scope);
        if (v !== undefined) edges[edge] = v;
      }
      const resolvedText = cell.txBody
        ? resolveTextBody({
            txBody: cell.txBody, lstStyles: [child(cell.txBody, 'a:lstStyle')], bodyPrs: [child(cell.txBody, 'a:bodyPr')],
            ...(ctx.master?.view.txStyles.other ? { bucket: ctx.master.view.txStyles.other } : {}),
            ...(ctx.doc.presentation.defaultTextStyle ? { defaultTextStyle: ctx.doc.presentation.defaultTextStyle } : {}),
            theme: ctx.theme, scope, slideNumber: ctx.slideNumber, extraRunProps: text,
          })
        : null;
      cells.push({
        row: r, col: c, rowSpan: cell.rowSpan, colSpan: cell.gridSpan, covered: cell.hMerge || cell.vMerge, fill, borders: edges,
        margins: {
          l: coord(attr(tcPr, 'marL'), 91440), r: coord(attr(tcPr, 'marR'), 91440),
          t: coord(attr(tcPr, 'marT'), 45720), b: coord(attr(tcPr, 'marB'), 45720),
        },
        anchor: (attr(tcPr, 'anchor') ?? 't') as ResolvedCell['anchor'], vert: attr(tcPr, 'vert') ?? 'horz', text: resolvedText,
      });
    });
  });

  let background: Paint = { kind: 'none' };
  if (bgPart?.fill) background = partFill(bgPart.fill, ctx, scope) ?? background;
  const ownBg = fillIn(child(t.tblPr, 'a:fill')) ?? fillIn(t.tblPr);
  if (ownBg) background = resolveFill(ownBg, scope);
  return { columns: t.columns, rows: t.rows.map(r => r.h), cells, background, rtl: t.flags.rtl };
}

/** The id of a table's style, for callers that need to know whether one applies. */
export const tableStyleId = (tbl: XmlElement) => textOf(child(child(tbl, 'a:tblPr'), 'a:tableStyleId')).trim() || undefined;
