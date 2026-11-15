// apps/web/src/ui/SelectionPane.tsx  -- the selection pane: every shape of the slide, to pick, rename, hide; alt text
import { useEffect, useRef, useState } from 'react';
import { rename, setAltText, setShapesHidden } from '@pptx/editor-core';
import type { ShapeView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import { useEditor } from '../state/store.ts';
import * as act from './actions.ts';

/** The group a shape sits in, if any: selecting a member from the pane enters its group, as a double click does. */
const groupOf = (el: XmlElement) => (el.parent && el.parent.local === 'grpSp' ? el.parent : null);

function Row({ shape, depth, part }: { shape: ShapeView; depth: number; part: string }) {
  const selection = useEditor(s => s.selection);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState('');
  const st = useEditor.getState;
  const ed = st().session!.ed;
  const selected = selection.includes(shape.el);
  const commit = () => {
    setEditing(false);
    const name = draft.trim();
    if (name && name !== shape.name) st().run(rename(ed, part, shape, name));
  };
  return (
    <>
      <li
        className={`${selected ? 'selected' : ''}${shape.hidden ? ' hidden-shape' : ''}`} style={{ paddingLeft: 4 + depth * 14 }} data-testid="pane-row" data-name={shape.name}
        onClick={e => {
          const add = e.shiftKey || e.metaKey || e.ctrlKey;
          const cur = st().selection;
          // shapes selected together share a container: adding one from elsewhere starts a new selection
          const together = add && cur.length > 0 && cur[0]!.parent === shape.el.parent;
          st().select(together ? (cur.includes(shape.el) ? cur.filter(x => x !== shape.el) : [...cur, shape.el]) : [shape.el]);
          st().setState({ entered: groupOf(shape.el) });
        }}
        onDoubleClick={() => { setDraft(shape.name); setEditing(true); }}
      >
        {editing
          ? <input className="pane-name" autoFocus aria-label="Name" value={draft} data-testid="pane-rename" onChange={e => setDraft(e.target.value)} onBlur={commit}
              onClick={e => e.stopPropagation()}
              onKeyDown={e => { e.stopPropagation(); if (e.key === 'Enter') commit(); else if (e.key === 'Escape') setEditing(false); }} />
          : <span className="what" title="Double-click to rename">{shape.name || '(no name)'}</span>}
        <button type="button" className="eye" aria-label={shape.hidden ? `Show ${shape.name}` : `Hide ${shape.name}`} aria-pressed={!shape.hidden} title={shape.hidden ? 'Show' : 'Hide'} data-testid="pane-eye"
          onClick={e => { e.stopPropagation(); st().run(setShapesHidden(ed, part, [shape], !shape.hidden)); }}>{shape.hidden ? '○' : '●'}</button>
      </li>
      {shape.kind === 'grpSp' && [...shape.children].reverse().map(c => <Row key={c.id} shape={c} depth={depth + 1} part={part} />)}
    </>
  );
}

/** Alt text of the one selected shape: what a screen reader says for it. Written when the box loses the focus. */
function AltText({ shape, part }: { shape: ShapeView; part: string }) {
  const [draft, setDraft] = useState(shape.descr ?? '');
  const ref = useRef<HTMLTextAreaElement>(null);
  // another shape selected, or the text changed by undo: the box shows what the deck holds
  useEffect(() => { setDraft(shape.descr ?? ''); }, [shape.el, shape.descr]);
  useEffect(() => {
    const focus = () => ref.current?.focus();
    window.addEventListener('focus-alt-text', focus);
    return () => window.removeEventListener('focus-alt-text', focus);
  }, []);
  // Escape gives the change up: the blur it causes must not write it
  const givenUp = useRef(false);
  const commit = () => {
    if (givenUp.current) { givenUp.current = false; return; }
    const s = useEditor.getState();
    if (draft !== (shape.descr ?? '')) s.run(setAltText(s.session!.ed, part, shape, draft));
  };
  return (
    <section className="alt-text">
      <label htmlFor="alt-text"><strong>Alt text</strong> for {shape.name || 'this shape'}</label>
      <textarea id="alt-text" ref={ref} data-testid="alt-text" value={draft} placeholder="Describe this for someone who cannot see it"
        onChange={e => setDraft(e.target.value)} onBlur={commit}
        onKeyDown={e => { if (e.key === 'Escape') { givenUp.current = true; setDraft(shape.descr ?? ''); e.currentTarget.blur(); } }} />
    </section>
  );
}

/**
 * The shapes of the slide, topmost first, a group's members under it. A click selects (Shift, Ctrl or Cmd adds), a
 * double click renames, the dot hides or shows. Hidden shapes stay listed, which is the only place they can be
 * found again.
 */
export function SelectionPane() {
  const s = useEditor(st => st.session);
  const part = useEditor(st => st.slide);
  useEditor(st => st.version);
  const selection = useEditor(st => st.selection);
  if (!s || !part) return null;
  const ed = s.ed;
  const shapes = [...ed.slide(part).shapes].reverse();
  const all = ed.slide(part).shapes;
  const one = selection.length === 1 ? ed.view(part, selection[0]!) : undefined;
  const run = useEditor.getState().run;
  return (
    <aside className="pane selection-pane" aria-label="Selection Pane" data-testid="selection-pane">
      <header><strong>Selection</strong><button type="button" aria-label="Close" onClick={() => useEditor.setState({ pane: 'none' })}>×</button></header>
      <div className="pane-tools">
        <button type="button" disabled={!all.some(x => x.hidden)} onClick={() => run(setShapesHidden(ed, part, all, false))} data-testid="pane-show-all">Show All</button>
        <button type="button" disabled={!all.some(x => !x.hidden)} onClick={() => run(setShapesHidden(ed, part, all, true))} data-testid="pane-hide-all">Hide All</button>
        <button type="button" aria-label="Bring Forward" title="Bring Forward" disabled={!selection.length} onClick={() => act.order('forward')} data-testid="pane-forward">▲</button>
        <button type="button" aria-label="Send Backward" title="Send Backward" disabled={!selection.length} onClick={() => act.order('backward')} data-testid="pane-backward">▼</button>
      </div>
      {shapes.length === 0 && <p className="hint">This slide has no shapes.</p>}
      <ol>{shapes.map(x => <Row key={x.id} shape={x} depth={0} part={part} />)}</ol>
      {one && <AltText key={one.id} shape={one} part={part} />}
    </aside>
  );
}
