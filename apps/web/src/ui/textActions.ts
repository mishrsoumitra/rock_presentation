// apps/web/src/ui/textActions.ts  -- typing and formatting, for the keyboard and the ribbon alike
import {
  deleteText, formatParagraphs, formatRuns, paragraphLength, paragraphs, typeText, type ParaPatch, type RunPatch, type TextPos, type TextRange,
} from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import type { ResolvedParagraph, ResolvedRunProps } from '@pptx/resolve';
import type { XmlElement } from '@pptx/xml';
import { refit } from '../engine/textfit.ts';
import { session, useEditor } from '../state/store.ts';
import { keyOf } from './SlideCanvas.tsx';

const ordered = (a: TextPos, b: TextPos): TextRange => (a.p < b.p || (a.p === b.p && a.offset <= b.offset) ? { start: a, end: b } : { start: b, end: a });

export function wholeBody(txBody: XmlElement): TextRange {
  const ps = paragraphs(txBody);
  const last = ps.length - 1;
  return { start: { p: 0, offset: 0 }, end: { p: Math.max(0, last), offset: last >= 0 ? paragraphLength(ps[last]!) : 0 } };
}

/** Options that keep a text edit and its autofit refit in one undo step. */
function opts(part: string, el: XmlElement, coalesce: string | undefined) {
  const s = session()!;
  return { scope: el, after: () => refit(s.ed, part, el, s.fonts), ...(coalesce ? { coalesceKey: coalesce } : {}) };
}

/** Type at the caret (replacing the selection). Typing in one box coalesces into one undo step. */
export function typeChars(chars: string): void {
  const st = useEditor.getState();
  const s = st.session, t = st.text;
  if (!s || !t) return;
  const body = new ShapeView(t.el, s.ed.doc.owner(t.part)).txBody;
  if (!body) return;
  const range = ordered(t.anchor, t.caret);
  const edit = typeText(s.ed, t.part, body, range, chars, { ...opts(t.part, t.el, `type:${keyOf(t.el)}`), ...(st.pending ? { patch: st.pending } : {}) });
  st.run(edit.command);
  useEditor.setState({ text: { ...t, caret: edit.caret, anchor: edit.caret } });
}

export function deleteChars(direction: 'back' | 'forward'): void {
  const st = useEditor.getState();
  const s = st.session, t = st.text;
  if (!s || !t) return;
  const body = new ShapeView(t.el, s.ed.doc.owner(t.part)).txBody;
  if (!body) return;
  const at = t.anchor.p === t.caret.p && t.anchor.offset === t.caret.offset ? t.caret : ordered(t.anchor, t.caret);
  const edit = deleteText(s.ed, t.part, body, at, direction, opts(t.part, t.el, `type:${keyOf(t.el)}`));
  st.run(edit.command);
  useEditor.setState({ text: { ...t, caret: edit.caret, anchor: edit.caret } });
}

/** The bodies a formatting command applies to: the text selection, or all text of each selected shape. */
function targets(): { part: string; el: XmlElement; body: XmlElement; range: TextRange }[] {
  const st = useEditor.getState();
  const s = st.session;
  if (!s || !st.slide) return [];
  if (st.text) {
    const body = new ShapeView(st.text.el, s.ed.doc.owner(st.text.part)).txBody;
    return body ? [{ part: st.text.part, el: st.text.el, body, range: ordered(st.text.anchor, st.text.caret) }] : [];
  }
  return st.selection.flatMap(el => {
    const body = new ShapeView(el, s.ed.doc.owner(st.slide!)).txBody;
    return body ? [{ part: st.slide!, el, body, range: wholeBody(body) }] : [];
  });
}

/** Character formatting. With only a caret, it applies to what is typed next (and to an empty paragraph's mark). */
export function formatText(patch: RunPatch, label = 'Font'): void {
  const st = useEditor.getState();
  const s = st.session;
  if (!s) return;
  const t = st.text;
  if (t && t.anchor.p === t.caret.p && t.anchor.offset === t.caret.offset) {
    useEditor.setState({ pending: { ...st.pending, ...patch } });
    const body = new ShapeView(t.el, s.ed.doc.owner(t.part)).txBody;
    const p = body ? paragraphs(body)[t.caret.p] : undefined;
    if (!p || paragraphLength(p) > 0) return;
  }
  const ts = targets();
  if (!ts.length) return;
  for (const x of ts) st.run(formatRuns(s.ed, x.part, x.body, x.range, patch, { ...opts(x.part, x.el, undefined), label }));
}

export function formatPara(patch: ParaPatch, label = 'Paragraph'): void {
  const st = useEditor.getState();
  const s = st.session;
  if (!s) return;
  for (const x of targets()) st.run(formatParagraphs(s.ed, x.part, x.body, x.range, patch, { ...opts(x.part, x.el, undefined), label }));
}

/** The resolved formatting at the caret or of the selected shape's first run (for the ribbon's toggles). */
export function formattingHere(): { run?: ResolvedRunProps; para?: ResolvedParagraph } {
  const st = useEditor.getState();
  const s = st.session;
  if (!s || !st.slide) return {};
  const el = st.text?.el ?? st.selection[0];
  if (!el) return {};
  const scene = s.scenes.get(st.slide);
  const walk = (list: typeof scene.slide.shapes): ResolvedParagraph[] | undefined => {
    for (const r of list) {
      if (r.source?.view.el === el) return r.source.text?.paragraphs;
      const inner = walk(r.children);
      if (inner) return inner;
    }
    return undefined;
  };
  const paras = walk(scene.slide.shapes);
  const p = paras?.[st.text?.caret.p ?? 0];
  if (!p) return {};
  let run = p.runs[0]?.props ?? p.endProps;
  if (st.text) {
    let at = 0;
    for (const r of p.runs) {
      const len = r.kind === 'break' ? 1 : r.text.length;
      if (st.text.caret.offset > at && st.text.caret.offset <= at + len) { run = r.props; break; }
      at += len;
    }
  }
  return { run: { ...run, ...(st.pending?.b !== undefined ? { bold: !!st.pending.b } : {}), ...(st.pending?.i !== undefined ? { italic: !!st.pending.i } : {}) }, para: p };
}
