// apps/web/src/ui/TextInput.tsx  -- keyboard and IME input while editing text, through a hidden textarea
import type Konva from 'konva';
import { useLayoutEffect, useRef, type RefObject } from 'react';
import { useEditor } from '../state/store.ts';
import { caretRect, move, type Move, type TextView } from './caret.ts';
import { keystrokeStart } from './perf.ts';
import { deleteChars, formatPara, formatText, formattingHere, selectAll, typeChars } from './textActions.ts';

/**
 * Text input goes to a textarea kept under the caret (how browser editors take IME composition and mobile
 * keyboards): what it receives is typed into the deck and the textarea is emptied again.
 */
export function TextInput({ view, stage, zoom }: { view: TextView & { transform: number[] }; stage: RefObject<Konva.Stage | null>; zoom: number }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const text = useEditor(s => s.text)!;
  const goalX = useRef<number | undefined>(undefined);
  const composing = useRef(false);

  // focused as soon as it is in the page, before the browser paints: keys typed right after the click land here
  useLayoutEffect(() => { ref.current?.focus({ preventScroll: true }); }, [text.el]);

  // keep the textarea under the caret, so the IME's candidate window opens there
  const r = caretRect(view, text.caret);
  let left = 0, top = 0;
  const st = stage.current;
  if (r && st) {
    const m = view.transform;
    const sx = m[0]! * r.x + m[2]! * r.top + m[4]!, sy = m[1]! * r.x + m[3]! * r.top + m[5]!;
    const abs = st.getAbsoluteTransform().point({ x: sx, y: sy });
    left = abs.x;
    top = abs.y;
  }

  const setCaret = (caret: { p: number; offset: number }, extend: boolean) => {
    const t = useEditor.getState().text;
    if (!t) return;
    useEditor.getState().session?.ed.log.seal();
    useEditor.setState({ text: { ...t, caret, anchor: extend ? t.anchor : caret }, pending: null });
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (composing.current) return;
    const mod = e.metaKey || e.ctrlKey;
    const t = useEditor.getState().text!;
    const moves: Record<string, Move> = {
      ArrowLeft: mod || e.altKey ? 'wordLeft' : 'left', ArrowRight: mod || e.altKey ? 'wordRight' : 'right', ArrowUp: 'up', ArrowDown: 'down',
      Home: mod ? 'start' : 'home', End: mod ? 'finish' : 'end',
    };
    const how = moves[e.key];
    if (how) {
      e.preventDefault();
      const collapsed = t.anchor.p === t.caret.p && t.anchor.offset === t.caret.offset;
      if (!e.shiftKey && !collapsed && (how === 'left' || how === 'right')) {
        const a = t.anchor, c = t.caret;
        const first = a.p < c.p || (a.p === c.p && a.offset < c.offset) ? a : c;
        const last = first === a ? c : a;
        setCaret(how === 'left' ? first : last, false);
        return;
      }
      const res = move(view, t.caret, how, how === 'up' || how === 'down' ? goalX.current : undefined);
      goalX.current = how === 'up' || how === 'down' ? res.goalX : undefined;
      setCaret(res.pos, e.shiftKey);
      return;
    }
    goalX.current = undefined;
    if (e.key === 'Backspace' || e.key === 'Delete') {
      e.preventDefault();
      keystrokeStart();
      deleteChars(e.key === 'Backspace' ? 'back' : 'forward');
    } else if (e.key === 'Enter') {
      e.preventDefault();
      keystrokeStart();
      typeChars(e.shiftKey ? '\u000b' : '\n');
    } else if (e.key === 'Tab') {
      e.preventDefault();
      keystrokeStart();
      const lvl = formattingHere().para?.props.level ?? 0;
      const bulleted = !!formattingHere().para?.props.bullet;
      if (bulleted || e.shiftKey) formatPara({ lvl: Math.max(0, Math.min(8, lvl + (e.shiftKey ? -1 : 1))) }, e.shiftKey ? 'Decrease Indent' : 'Increase Indent');
      else typeChars('\t');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      useEditor.getState().session?.ed.log.seal();
      useEditor.setState({ text: null, selection: [t.el] });
    } else if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      selectAll();
    } else if (mod && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      const run = formattingHere().run;
      const k = e.key.toLowerCase();
      if (k === 'b') formatText({ b: !run?.bold }, 'Bold');
      else if (k === 'i') formatText({ i: !run?.italic }, 'Italic');
      else formatText({ u: run?.underline && run.underline !== 'none' ? 'none' : 'sng' }, 'Underline');
    } else if (mod && (e.key === 'z' || e.key === 'y')) {
      e.preventDefault();
      if (e.key === 'y' || e.shiftKey) useEditor.getState().redo();
      else useEditor.getState().undo();
    } else if (mod && e.key === ']') {
      e.preventDefault();
      const size = formattingHere().run?.size ?? 18;
      formatText({ sz: Math.round((size + 2) * 100) }, 'Increase Font Size');
    } else if (mod && e.key === '[') {
      e.preventDefault();
      const size = formattingHere().run?.size ?? 18;
      formatText({ sz: Math.max(100, Math.round((size - 2) * 100)) }, 'Decrease Font Size');
    }
  };

  const flush = (el: HTMLTextAreaElement) => {
    const v = el.value;
    el.value = '';
    if (!v) return;
    keystrokeStart();
    typeChars(v.replace(/\r\n|\r/g, '\n'));
  };

  return (
    <textarea
      ref={ref}
      className="text-input"
      aria-label="Text"
      data-testid="text-input"
      style={{ left, top: top, height: r ? r.height * zoom : 16 }}
      onKeyDown={onKeyDown}
      onCompositionStart={() => { composing.current = true; }}
      onCompositionEnd={e => { composing.current = false; flush(e.currentTarget); }}
      onInput={e => { if (!composing.current) flush(e.currentTarget); }}
      onPaste={e => {
        e.preventDefault();
        const s = e.clipboardData.getData('text/plain');
        if (s) typeChars(s.replace(/\r\n|\r/g, '\n'));
      }}
      onBlur={() => useEditor.getState().session?.ed.log.seal()}
    />
  );
}
