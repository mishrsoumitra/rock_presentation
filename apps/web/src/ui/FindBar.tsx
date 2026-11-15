// apps/web/src/ui/FindBar.tsx  -- find and replace: a small panel over the slide, as in Google Slides
import { useEffect, useMemo, useRef, useState } from 'react';
import { compareMatches, findText, type TextMatch } from '@pptx/editor-core';
import { useEditor } from '../state/store.ts';
import { replaceEverywhere, replaceOne } from './textActions.ts';
import { keepFocus } from './widgets.tsx';

type Place = Pick<TextMatch, 'slideIndex' | 'shapeIndex' | 'p' | 'start'>;

/** Go to a match: its slide, its shape selected, and the found text selected in it, as a text selection. */
function show(m: TextMatch): void {
  const st = useEditor.getState();
  if (st.slide !== m.part) st.setSlide(m.part);
  st.session?.ed.log.seal();
  const parent = m.shape.parent;
  useEditor.setState({
    selection: [m.shape], entered: parent && parent.local === 'grpSp' ? parent : null, pending: null,
    text: { part: m.part, el: m.shape, anchor: { p: m.p, offset: m.start }, caret: { p: m.p, offset: m.end } },
  });
}

/**
 * Ctrl/Cmd+F opens it (Ctrl+H or Cmd+Shift+H with the replace box focused). Enter finds the next match and
 * Shift+Enter the one before, going round the deck; the count says which match of how many is shown. Replace
 * changes the match shown and moves on; Replace All changes every match in one undo step.
 */
export function FindBar() {
  const open = useEditor(s => s.finding);
  const session = useEditor(s => s.session);
  const version = useEditor(s => s.version);
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWords, setWholeWords] = useState(false);
  const [place, setPlace] = useState<Place | null>(null);
  const findRef = useRef<HTMLInputElement>(null);
  const replaceRef = useRef<HTMLInputElement>(null);
  const focusOn = useRef<'find' | 'replace' | null>(null);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const s = useEditor.getState();
      if (!s.session || s.mode !== 'edit') return;
      const mod = e.metaKey || e.ctrlKey;
      const k = e.key.toLowerCase();
      const find = mod && !e.shiftKey && k === 'f';
      const replace = (e.ctrlKey && !e.shiftKey && k === 'h') || (e.metaKey && e.shiftKey && k === 'h');
      if (!find && !replace) return;
      e.preventDefault();
      e.stopPropagation();
      focusOn.current = replace ? 'replace' : 'find';
      if (s.finding) (replace ? replaceRef : findRef).current?.select();
      else useEditor.setState({ finding: true });
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, []);

  useEffect(() => {
    if (!open) return;
    (focusOn.current === 'replace' ? replaceRef : findRef).current?.select();
    focusOn.current = null;
  }, [open]);

  const options = useMemo(() => ({ matchCase, wholeWords }), [matchCase, wholeWords]);
  // found again whenever the deck changes (`version`), so the count is always true
  const matches = useMemo(() => (open && session && query ? findText(session.ed, query, options) : []), [open, session, query, options, version]);
  const at = place ? matches.findIndex(m => compareMatches(m, place) === 0) : -1;

  if (!open || !session) return null;

  const go = (m: TextMatch | undefined) => {
    if (!m) { setPlace(null); return; }
    setPlace({ slideIndex: m.slideIndex, shapeIndex: m.shapeIndex, p: m.p, start: m.start });
    show(m);
  };
  const step = (dir: 1 | -1) => {
    const n = matches.length;
    if (!n) return;
    if (at >= 0) { go(matches[(at + dir + n) % n]); return; }
    // nothing shown yet: start from the slide on screen
    const here = Math.max(0, session.ed.doc.slideParts().indexOf(useEditor.getState().slide ?? ''));
    const from = dir === 1 ? matches.find(m => m.slideIndex >= here) ?? matches[0] : [...matches].reverse().find(m => m.slideIndex <= here) ?? matches[n - 1];
    go(from);
  };
  const replace = () => {
    const m = matches[at];
    if (!m) { step(1); return; }
    replaceOne(m, replacement);
    // the next match after the text just put in (which is not searched again), going round to the first
    const fresh = findText(session.ed, query, options);
    const after = { ...m, start: m.start + replacement.length };
    go(fresh.find(x => compareMatches(x, after) >= 0) ?? fresh[0]);
  };
  const replaceAll = () => {
    const n = replaceEverywhere(query, replacement, options);
    setPlace(null);
    useEditor.getState().notify('info', n ? `Replaced ${n} occurrence${n === 1 ? '' : 's'}.` : `"${query}" was not found.`);
  };
  const close = () => useEditor.setState({ finding: false });
  const keys = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey ? -1 : 1); }
    else if (e.key === 'Escape') { e.preventDefault(); close(); }
    else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) useEditor.getState().redo(); else useEditor.getState().undo(); }
  };
  const count = !query ? '' : !matches.length ? 'No results' : at >= 0 ? `${at + 1} of ${matches.length}` : `${matches.length} found`;

  return (
    <div className="find-bar" role="search" aria-label="Find and replace" data-testid="find-bar">
      <div className="find-row">
        <input ref={findRef} aria-label="Find" placeholder="Find" value={query} data-testid="find-input" onKeyDown={keys}
          onChange={e => { setQuery(e.target.value); setPlace(null); }} />
        <span className="find-count" data-testid="find-count" aria-live="polite">{count}</span>
        <button type="button" title="Previous (Shift+Enter)" aria-label="Previous" disabled={!matches.length} onClick={() => step(-1)} onMouseDown={keepFocus} data-testid="find-prev">↑</button>
        <button type="button" title="Next (Enter)" aria-label="Next" disabled={!matches.length} onClick={() => step(1)} onMouseDown={keepFocus} data-testid="find-next">↓</button>
        <button type="button" title="Close (Esc)" aria-label="Close" onClick={close} onMouseDown={keepFocus} data-testid="find-close">×</button>
      </div>
      <div className="find-row">
        <input ref={replaceRef} aria-label="Replace with" placeholder="Replace with" value={replacement} data-testid="replace-input" onKeyDown={keys}
          onChange={e => setReplacement(e.target.value)} />
        <button type="button" disabled={!matches.length} onClick={replace} onMouseDown={keepFocus} data-testid="replace-one">Replace</button>
        <button type="button" disabled={!matches.length} onClick={replaceAll} onMouseDown={keepFocus} data-testid="replace-all">Replace All</button>
      </div>
      <div className="find-row options">
        <label><input type="checkbox" checked={matchCase} onChange={e => { setMatchCase(e.target.checked); setPlace(null); }} data-testid="find-case" /> Match case</label>
        <label><input type="checkbox" checked={wholeWords} onChange={e => { setWholeWords(e.target.checked); setPlace(null); }} data-testid="find-words" /> Whole words</label>
      </div>
    </div>
  );
}
