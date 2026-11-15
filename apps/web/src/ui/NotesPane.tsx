// apps/web/src/ui/NotesPane.tsx  -- speaker notes under the slide, as PowerPoint and Google Slides show them
import { notesText, setNotes } from '@pptx/editor-core';
import { useEditor } from '../state/store.ts';

/**
 * The notes of the slide on screen, in a plain text box. Each change is a command, so undo and redo cover notes
 * like everything else; a run of typing is one step. The text shown always comes from the deck, never from the box.
 */
export function NotesPane() {
  const session = useEditor(s => s.session);
  const slide = useEditor(s => s.slide);
  // the notes text itself is what this subscribes to: typing on the slide does not redraw the pane
  const text = useEditor(s => (s.session && s.slide ? notesText(s.session.ed, s.slide) : ''));
  if (!session || !slide) return null;
  return (
    <section className="notes" aria-label="Speaker notes">
      <textarea
        className="notes-text" data-testid="notes" placeholder="Click to add speaker notes" value={text} spellCheck
        onFocus={() => { if (useEditor.getState().text) useEditor.setState({ text: null }); session.ed.log.seal(); }}
        onBlur={() => session.ed.log.seal()}
        onChange={e => {
          if (e.target.value === text) return;
          useEditor.getState().run(setNotes(session.ed, slide, e.target.value, { coalesceKey: `notes:${slide}` }));
        }}
        onKeyDown={e => {
          // the editor's undo, not the browser's: the box holds no history of its own
          const mod = e.metaKey || e.ctrlKey;
          const k = e.key.toLowerCase();
          if (mod && k === 'z') { e.preventDefault(); if (e.shiftKey) useEditor.getState().redo(); else useEditor.getState().undo(); }
          else if (mod && k === 'y') { e.preventDefault(); useEditor.getState().redo(); }
          else if (e.key === 'Escape') e.currentTarget.blur();
        }}
      />
    </section>
  );
}
