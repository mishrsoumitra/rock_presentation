// apps/web/src/ui/CellEditor.tsx  -- typing in a table cell: a text box laid over the cell while it is edited
import type Konva from 'konva';
import { useLayoutEffect, useRef, useState, type RefObject } from 'react';
import { cellText, regionOf, tableOf } from '@pptx/editor-core';
import type { RenderShape } from '@pptx/render';
import type { FontRegistry } from '@pptx/text';
import { useEditor } from '../state/store.ts';
import { cellCorners, gridOf, setCellDraft, tabCell } from './tables.ts';

/**
 * The cell's text in a plain box over the cell. What is typed goes into the cell when typing there ends: by Escape
 * or Ctrl/Cmd+Enter, by a click elsewhere, or by Tab, which moves on to the next cell and keeps typing there (Enter
 * starts a new line, as in PowerPoint). The parent gives it a key per cell, so each cell gets a fresh box.
 */
export function CellEditor({ rs, stage, zoom, fonts }: { rs: RenderShape; stage: RefObject<Konva.Stage | null>; zoom: number; fonts: FontRegistry }) {
  const cells = useEditor(s => s.cells)!;
  const el = cells.el;
  const tbl = tableOf(el)!;
  const at = cells.anchor;
  const [draft, setDraft] = useState(() => cellText(tbl, at.r, at.c) + (cells.seed ?? ''));
  const ref = useRef<HTMLTextAreaElement>(null);

  useLayoutEffect(() => {
    const box = ref.current;
    if (!box) return;
    setCellDraft(el, at, box.value);
    box.focus({ preventScroll: true });
    box.setSelectionRange(box.value.length, box.value.length);
  }, []);

  const g = regionOf(tbl, at.r, at.c);
  const corners = cellCorners(rs, fonts, g);
  const st = stage.current;
  if (!corners || !st) return null;
  // the cell on the page: the box covers its bounds (a rotated table gets the box around the cell)
  const pts = [0, 2, 4, 6].map(i => st.getAbsoluteTransform().point({ x: corners[i]!, y: corners[i + 1]! }));
  const left = Math.min(...pts.map(p => p.x)), top = Math.min(...pts.map(p => p.y));
  const width = Math.max(...pts.map(p => p.x)) - left, height = Math.max(...pts.map(p => p.y)) - top;
  const resolved = gridOf(rs, fonts)?.table.cells.find(c => c.row === g.r0 && c.col === g.c0);
  const points = resolved?.text?.paragraphs[0]?.runs[0]?.props.size ?? resolved?.text?.paragraphs[0]?.endProps.size ?? 18;

  /** Typing in this cell ends; what was typed is written by the store's subscription (see `setCellDraft`). */
  const finish = () => {
    const cur = useEditor.getState().cells;
    if (cur && cur.el === el && cur.editing && cur.anchor.r === at.r && cur.anchor.c === at.c) useEditor.setState({ cells: { ...cur, editing: false } });
  };

  return (
    <textarea
      ref={ref} className="cell-editor" aria-label="Cell text" data-testid="cell-editor" value={draft}
      style={{ left, top, width, minHeight: height, fontSize: Math.max(9, points * 12700 * zoom) }}
      onChange={e => { setDraft(e.target.value); setCellDraft(el, at, e.target.value); }}
      onBlur={finish}
      onMouseDown={e => e.stopPropagation()}
      onKeyDown={e => {
        e.stopPropagation();
        if (e.key === 'Escape' || (e.key === 'Enter' && (e.metaKey || e.ctrlKey))) { e.preventDefault(); finish(); }
        else if (e.key === 'Tab') { e.preventDefault(); tabCell(e.shiftKey, true); }
      }}
    />
  );
}
