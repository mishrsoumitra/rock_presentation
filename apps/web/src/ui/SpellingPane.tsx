// apps/web/src/ui/SpellingPane.tsx  -- the spelling pane: the slide's text where the browser's spell checker can see it
import { bodyPlainText, retypeText } from '@pptx/editor-core';
import type { ShapeView } from '@pptx/model';
import { refit } from '../engine/textfit.ts';
import { useEditor } from '../state/store.ts';

/** A line break inside a paragraph, as the pane shows it (a text box cannot show the character the deck uses). */
const BREAK = '↵';
const shownText = (s: string) => s.replace(/\u000b/g, BREAK);
const deckText = (s: string) => s.replace(/↵/g, '\u000b');

/** The shapes of a slide that hold text, in reading order, groups opened. */
function withText(shapes: ShapeView[]): ShapeView[] {
  return shapes.flatMap(s => (s.kind === 'grpSp' ? withText(s.children) : s.txBody ? [s] : []));
}

/**
 * The text of each shape of the slide in a plain box of its own, with the browser's spell checker on: it underlines
 * what it does not know and offers corrections from its right-click menu. Whatever is changed in a box goes into
 * the shape, retyping only what changed, so the formatting around a corrected word stays. The slide's own text is
 * drawn, not typed into a box, which is why the checker cannot see it there.
 */
export function SpellingPane() {
  const s = useEditor(st => st.session);
  const part = useEditor(st => st.slide);
  useEditor(st => st.version);
  const selection = useEditor(st => st.selection);
  if (!s || !part) return null;
  const ed = s.ed;
  const slides = ed.doc.slideParts();
  const at = slides.indexOf(part);
  const shapes = withText(ed.slide(part).shapes).filter(v => bodyPlainText(v.txBody!).length > 0);
  const go = (to: number) => { const next = slides[to]; if (next) useEditor.getState().setSlide(next); };
  return (
    <aside className="pane spelling-pane" aria-label="Spelling" data-testid="spelling-pane">
      <header><strong>Spelling</strong><button type="button" aria-label="Close" onClick={() => useEditor.setState({ pane: 'none' })}>×</button></header>
      <p className="hint">Your browser underlines words it does not know. Right-click one for corrections; what you change here changes the slide.</p>
      {shapes.length === 0 && <p className="hint" data-testid="spelling-empty">This slide has no text.</p>}
      <ol>
        {shapes.map(v => (
          <li key={v.id} className={selection.includes(v.el) ? 'selected' : ''}>
            <label htmlFor={`spell-${v.id}`}>{v.name || 'Text'}</label>
            <textarea
              id={`spell-${v.id}`} spellCheck data-testid="spelling-text" data-shape={v.name} value={shownText(bodyPlainText(v.txBody!))}
              rows={Math.min(8, Math.max(1, bodyPlainText(v.txBody!).split('\n').length))}
              onFocus={() => { if (!useEditor.getState().selection.includes(v.el)) useEditor.getState().select([v.el]); }}
              onBlur={() => ed.log.seal()}
              onKeyDown={e => {
                e.stopPropagation();
                const mod = e.metaKey || e.ctrlKey;
                if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) useEditor.getState().redo(); else useEditor.getState().undo(); }
                else if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); useEditor.getState().redo(); }
              }}
              onChange={e => {
                const cmd = retypeText(ed, part, v.txBody!, deckText(e.target.value), { scope: v.el, after: () => refit(ed, part, v.el, s.fonts), coalesceKey: `spell:${part}:${v.id}` });
                if (cmd) useEditor.getState().run(cmd);
              }}
            />
          </li>
        ))}
      </ol>
      <div className="pane-tools spelling-nav">
        <button type="button" disabled={at <= 0} onClick={() => go(at - 1)} data-testid="spelling-prev">Previous Slide</button>
        <span>Slide {at + 1} of {slides.length}</span>
        <button type="button" disabled={at >= slides.length - 1} onClick={() => go(at + 1)} data-testid="spelling-next">Next Slide</button>
      </div>
    </aside>
  );
}
