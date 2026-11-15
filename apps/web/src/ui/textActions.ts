// apps/web/src/ui/textActions.ts  -- typing and formatting, for the keyboard and the ribbon alike
import {
  clearRunFormatting, deleteText, formatBody, formatParagraphs, formatRuns, paragraphLength, paragraphs, replaceAll, replaceMatch, typeText,
  type BodyPatch, type Command, type FindOptions, type ParaPatch, type RunPatch, type TextMatch, type TextPos, type TextRange,
} from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import type { ResolvedBodyProps, ResolvedParagraph, ResolvedRunProps } from '@pptx/resolve';
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

/**
 * Select all the text of the box being edited (Cmd+A, a triple click): every paragraph, from the start of the first
 * to the end of the last. Typing then replaces it all, Backspace clears the box, and formatting applies to all of it.
 */
export function selectAll(): void {
  const st = useEditor.getState();
  const s = st.session, t = st.text;
  if (!s || !t) return;
  const body = new ShapeView(t.el, s.ed.doc.owner(t.part)).txBody;
  if (!body) return;
  const all = wholeBody(body);
  s.ed.log.seal();
  useEditor.setState({ text: { ...t, anchor: all.start, caret: all.end }, pending: null });
}

/** Options that keep a text edit and its autofit refit in one undo step. */
function opts(part: string, el: XmlElement, coalesce: string | undefined) {
  const s = session()!;
  return { scope: el, after: () => refit(s.ed, part, el, s.fonts), ...(coalesce ? { coalesceKey: coalesce } : {}) };
}

/** Run a command that changes one shape's text: the slide's render tree keeps its other shapes. */
function runOn(part: string, el: XmlElement, cmd: Command): void {
  const st = useEditor.getState();
  const s = st.session;
  if (!s) return;
  const before = s.ed.versionOf(part);
  s.scenes.expectChange(part, el);
  st.run(cmd);
  if (s.ed.versionOf(part) === before) s.scenes.forgetExpected(part);    // the command failed and changed nothing
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
  runOn(t.part, t.el, edit.command);
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
  runOn(t.part, t.el, edit.command);
  useEditor.setState({ text: { ...t, caret: edit.caret, anchor: edit.caret } });
}

/** Replace one found match; autofit text is refitted in the same undo step, as when typing. */
export function replaceOne(m: TextMatch, replacement: string): void {
  const s = session();
  if (s) runOn(m.part, m.shape, replaceMatch(s.ed, m, replacement, opts(m.part, m.shape, undefined)));
}

/** Replace every match in the deck as one undo step; returns how many there were. Text editing ends first. */
export function replaceEverywhere(query: string, replacement: string, find: FindOptions): number {
  const st = useEditor.getState();
  const s = st.session;
  if (!s) return 0;
  const r = replaceAll(s.ed, query, replacement, find, m => opts(m.part, m.shape, undefined));
  if (!r.count) return 0;
  useEditor.setState({ text: null, pending: null });
  s.ed.log.seal();
  st.run(r.command);
  s.ed.log.seal();
  return r.count;
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
  for (const x of ts) runOn(x.part, x.el, formatRuns(s.ed, x.part, x.body, x.range, patch, { ...opts(x.part, x.el, undefined), label }));
}

export function formatPara(patch: ParaPatch, label = 'Paragraph'): void {
  const st = useEditor.getState();
  const s = st.session;
  if (!s) return;
  for (const x of targets()) runOn(x.part, x.el, formatParagraphs(s.ed, x.part, x.body, x.range, patch, { ...opts(x.part, x.el, undefined), label }));
}

/** Text box options (anchor, margins, columns) for the box being edited or every selected shape with text. */
export function formatTextBox(patch: BodyPatch, label = 'Text Options'): void {
  const s = useEditor.getState().session;
  if (!s) return;
  for (const x of targets()) runOn(x.part, x.el, formatBody(s.ed, x.part, x.body, patch, { ...opts(x.part, x.el, undefined), label }));
}

/**
 * Clear All Formatting: the selected text, or all text of the selected shapes, goes back to what its placeholder or
 * the deck's defaults give. With only a caret, the caret's whole paragraph is cleared, and so is pending formatting.
 */
export function clearFormatting(): void {
  const st = useEditor.getState();
  const s = st.session;
  if (!s) return;
  useEditor.setState({ pending: null });
  for (const x of targets()) {
    const caret = x.range.start.p === x.range.end.p && x.range.start.offset === x.range.end.offset;
    const p = caret ? paragraphs(x.body)[x.range.start.p] : undefined;
    const range = p ? { start: { p: x.range.start.p, offset: 0 }, end: { p: x.range.start.p, offset: paragraphLength(p) } } : x.range;
    runOn(x.part, x.el, clearRunFormatting(s.ed, x.part, x.body, range, opts(x.part, x.el, undefined)));
  }
}

/** The resolved formatting at the caret or of the selected shape's first run (for the ribbon's toggles). */
export function formattingHere(): { run?: ResolvedRunProps; para?: ResolvedParagraph; body?: ResolvedBodyProps } {
  const st = useEditor.getState();
  const s = st.session;
  if (!s || !st.slide) return {};
  const el = st.text?.el ?? st.selection[0];
  if (!el) return {};
  const scene = s.scenes.get(st.slide);
  const walk = (list: typeof scene.slide.shapes): { paragraphs: ResolvedParagraph[]; body: ResolvedBodyProps } | undefined => {
    for (const r of list) {
      if (r.source?.view.el === el) return r.source.text ?? undefined;
      const inner = walk(r.children);
      if (inner) return inner;
    }
    return undefined;
  };
  const text = walk(scene.slide.shapes);
  const p = text?.paragraphs[st.text?.caret.p ?? 0];
  if (!p || !text) return {};
  let run = p.runs[0]?.props ?? p.endProps;
  if (st.text) {
    let at = 0;
    for (const r of p.runs) {
      const len = r.kind === 'break' ? 1 : r.text.length;
      if (st.text.caret.offset > at && st.text.caret.offset <= at + len) { run = r.props; break; }
      at += len;
    }
  }
  const pend = st.pending;
  return {
    run: {
      ...run, ...(pend?.b !== undefined ? { bold: !!pend.b } : {}), ...(pend?.i !== undefined ? { italic: !!pend.i } : {}),
      ...(pend?.baseline !== undefined ? { baseline: (pend.baseline ?? 0) / 100000 } : {}),
    },
    para: p, body: text.body,
  };
}
