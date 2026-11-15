// apps/web/src/ui/clipboard.ts  -- shapes, pictures and table cells on the system clipboard (text has its own, in TextInput)
import { useEffect } from 'react';
import { bodyPlainText, cellAt, cellText, compound, copyShapes, setCellText, tableOf, tableSize, type Command } from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { drawShape, type RenderShape } from '@pptx/render';
import type { XmlElement } from '@pptx/xml';
import { useEditor } from '../state/store.ts';
import * as act from './actions.ts';
import { boundsOf } from './placement.ts';
import { htmlToParagraphs, paragraphsToHtml, plainOf, type Para } from './richText.ts';
import { activeCells, clearSelectedCells, rangeOf } from './tables.ts';
import { typeRich } from './textActions.ts';

/** The attribute that marks HTML this editor put on the clipboard: its value names the shapes kept in memory. */
const MARK = 'data-pptx-clip';
const idIn = (html: string) => new RegExp(`${MARK}="([a-z0-9-]+)"`).exec(html)?.[1];

/** The text of shapes, for apps that take only text: each shape's paragraphs, a table's rows with tabs between cells. */
function textOf(els: XmlElement[]): Para[] {
  const s = useEditor.getState();
  const owner = s.session!.ed.doc.owner(s.slide!);
  const out: Para[] = [];
  const visit = (el: XmlElement) => {
    const v = new ShapeView(el, owner);
    if (v.kind === 'grpSp') { for (const c of v.children) visit(c.el); return; }
    const tbl = tableOf(el);
    if (tbl) {
      const { rows, cols } = tableSize(tbl);
      for (let r = 0; r < rows; r++) out.push([{ text: Array.from({ length: cols }, (_, c) => cellText(tbl, r, c).replace(/\n/g, ' ')).join('\t') }]);
      return;
    }
    if (v.txBody) for (const line of bodyPlainText(v.txBody).split('\n')) out.push([{ text: line }]);
  };
  els.forEach(visit);
  return out.filter((p, i) => p[0]!.text || i < out.length - 1);
}

/**
 * A picture of shapes, for apps that take pictures: the shapes drawn on a clear background at screen resolution
 * (twice that for small selections), cropped to their bounds.
 */
export async function shapesPng(els: XmlElement[]): Promise<Blob | undefined> {
  const s = useEditor.getState();
  if (!s.session || !s.slide || typeof OffscreenCanvas === 'undefined') return undefined;
  const scene = s.session.scenes.get(s.slide);
  const find = (list: RenderShape[], el: XmlElement): RenderShape | undefined => {
    for (const r of list) {
      if (r.source?.view.el === el && r.layer === 'slide') return r;
      const inner = find(r.children, el);
      if (inner) return inner;
    }
    return undefined;
  };
  const shapes = els.map(el => find(scene.slide.shapes, el)).filter((r): r is RenderShape => !!r);
  if (!shapes.length) return undefined;
  const bs = shapes.map(boundsOf);
  const x = Math.min(...bs.map(b => b.x)), y = Math.min(...bs.map(b => b.y));
  const w = Math.max(...bs.map(b => b.x + b.w)) - x, h = Math.max(...bs.map(b => b.y + b.h)) - y;
  if (!(w > 0 && h > 0)) return undefined;
  const scale = Math.min(2 / 9525, 4096 / w, 4096 / h);          // two device pixels per CSS pixel, within 4096
  const canvas = new OffscreenCanvas(Math.max(1, Math.ceil(w * scale)), Math.max(1, Math.ceil(h * scale)));
  const c = canvas.getContext('2d');
  if (!c) return undefined;
  for (const rs of shapes) {
    drawShape(c as unknown as CanvasRenderingContext2D, rs, {
      scale, base: [1, 0, 0, 1, -x * scale, -y * scale], image: p => s.session!.images.get(p), createCanvas: (cw, ch) => new OffscreenCanvas(cw, ch),
    });
  }
  return canvas.convertToBlob({ type: 'image/png' });
}

/**
 * Put the selected shapes on the clipboard: kept whole in memory for pasting here (with pictures and all), and as
 * text and HTML for other apps. Where the browser allows it a picture of them goes on too, a moment later.
 * Returns false when nothing is selected.
 */
function copyShapesTo(data: DataTransfer): boolean {
  const s = useEditor.getState();
  if (!s.session || !s.slide || !s.selection.length) return false;
  const views = s.selection.map(el => new ShapeView(el, s.session!.ed.doc.owner(s.slide!)));
  const id = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  const paras = textOf(s.selection);
  const plain = plainOf(paras);
  const html = `<div ${MARK}="${id}">${paragraphsToHtml(paras)}</div>`;
  data.setData('text/plain', plain);
  data.setData('text/html', html);
  useEditor.setState({ clipboard: copyShapes(s.session.ed, s.slide, views), clipId: id });
  // a picture too, where the asynchronous clipboard lets a page write one (never needed for pasting here)
  const selection = s.selection;
  if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
    void shapesPng(selection).then(png => (png ? navigator.clipboard.write([new ClipboardItem({
      'text/plain': new Blob([plain], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }), 'image/png': png,
    })]) : undefined)).catch(() => undefined);
  }
  return true;
}

/** The picked cells as rows of tab-separated text, the way spreadsheets put cells on the clipboard. */
function copyCellsTo(data: DataTransfer): boolean {
  const s = useEditor.getState();
  const cells = activeCells(s);
  if (!cells) return false;
  const tbl = tableOf(cells.el)!;
  const g = rangeOf(cells);
  const rows: string[] = [];
  for (let r = g.r0; r <= g.r1; r++) rows.push(Array.from({ length: g.c1 - g.c0 + 1 }, (_, i) => cellText(tbl, r, g.c0 + i).replace(/[\n\t]/g, ' ')).join('\t'));
  data.setData('text/plain', rows.join('\n'));
  data.setData('text/html', `<table>${rows.map(r => `<tr>${r.split('\t').map(c => `<td>${c.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</td>`).join('')}</tr>`).join('')}</table>`);
  return true;
}

/** Rows of tab-separated text go into the table from the picked cell, right and down, as far as the table reaches. */
function pasteCells(text: string): boolean {
  const s = useEditor.getState();
  const cells = activeCells(s);
  if (!cells || !s.session || !s.slide) return false;
  const tbl = tableOf(cells.el)!;
  const { rows, cols } = tableSize(tbl);
  const g = rangeOf(cells);
  const frame = new ShapeView(cells.el, s.session.ed.doc.owner(s.slide));
  const lines = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
  const cmds: Command[] = [];
  lines.forEach((line, i) => line.forEach((value, j) => {
    const r = g.r0 + i, c = g.c0 + j;
    if (r < rows && c < cols && cellAt(tbl, r, c) && cellText(tbl, r, c) !== value) cmds.push(setCellText(s.session!.ed, s.slide!, frame, { r, c }, value));
  }));
  if (cmds.length) s.run(cmds.length === 1 ? { ...cmds[0]!, label: 'Paste' } : compound('Paste', cmds));
  return true;
}

/**
 * Paste what the clipboard holds onto the slide: shapes copied here (matched by the mark in the HTML) come back
 * whole; a picture becomes a picture; text becomes a text box, with bold, italic and underline when it has them.
 */
export async function pasteFrom(data: { html: string; plain: string; files: File[] }): Promise<void> {
  const s = useEditor.getState();
  if (!s.session || !s.slide) return;
  if (s.clipboard && s.clipId && idIn(data.html) === s.clipId) { await act.paste(); return; }
  if (activeCells(s) && data.plain && pasteCells(data.plain)) return;
  const pictures = data.files.filter(f => f.type.startsWith('image/'));
  if (pictures.length) {
    for (const f of pictures) await act.insertPictureFile(f);
    return;
  }
  const paras = data.html ? htmlToParagraphs(data.html) : data.plain.replace(/\r\n?/g, '\n').split('\n').map(line => [{ text: line }]);
  if (!paras.some(p => p.some(seg => seg.text.trim()))) return;                 // nothing but white space is nothing to paste
  act.insertTextBoxHere();
  s.session.ed.log.seal();
  typeRich(paras);
  s.session.ed.log.seal();
}

/** Whether a clipboard event belongs to the slide: not to a text field somewhere, the notes, or the box being edited. */
function forSlide(e: Event): boolean {
  const s = useEditor.getState();
  const t = e.target as HTMLElement | null;
  if (!s.session || s.mode !== 'edit' || s.text || s.cells?.editing) return false;
  return !(t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable || t.closest?.('.slides')));
}

/**
 * Ctrl/Cmd+C, X and V on the slide, through the browser's clipboard events (which carry the real clipboard, with no
 * permission asked). Mounted once by the app.
 */
export function useClipboard(): void {
  useEffect(() => {
    const copy = (e: ClipboardEvent) => {
      if (!forSlide(e) || !e.clipboardData) return;
      if (!(copyCellsTo(e.clipboardData) || copyShapesTo(e.clipboardData))) return;
      e.preventDefault();
    };
    const cut = (e: ClipboardEvent) => {
      if (!forSlide(e) || !e.clipboardData) return;
      const cells = !!activeCells(useEditor.getState());
      if (!(copyCellsTo(e.clipboardData) || copyShapesTo(e.clipboardData))) return;
      e.preventDefault();
      if (cells) clearSelectedCells();
      else act.deleteSelection();
    };
    const paste = (e: ClipboardEvent) => {
      if (!forSlide(e) || !e.clipboardData) return;
      e.preventDefault();
      void pasteFrom({ html: e.clipboardData.getData('text/html'), plain: e.clipboardData.getData('text/plain'), files: [...e.clipboardData.files] });
    };
    document.addEventListener('copy', copy);
    document.addEventListener('cut', cut);
    document.addEventListener('paste', paste);
    return () => {
      document.removeEventListener('copy', copy);
      document.removeEventListener('cut', cut);
      document.removeEventListener('paste', paste);
    };
  }, []);
}
