// apps/web/src/ui/CommentsPane.tsx  -- the comments pane: the current slide's comments, to read, add, change and delete
import { useEffect, useRef, useState } from 'react';
import { addComment, deleteComment, editComment, slideComments, type SlideComment } from '@pptx/editor-core';
import { useEditor } from '../state/store.ts';

const NAME_KEY = 'presentations.commentName';

/** The name comments are written under: kept in this browser, asked for the first time. */
function storedName(): string {
  try { return localStorage.getItem(NAME_KEY) ?? ''; } catch { return ''; }
}
function keepName(name: string): void {
  try { localStorage.setItem(NAME_KEY, name); } catch { /* a browser that keeps nothing: the name lasts for this visit */ }
}

/** When a comment was written, as a reader wants it: the date and time in their own words. */
function when(date: string): string {
  const d = new Date(/[zZ]|[+-]\d\d:?\d\d$/.test(date) ? date : `${date}Z`);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
}

function Comment({ c, part }: { c: SlideComment; part: string }) {
  const [draft, setDraft] = useState<string | null>(null);
  const st = useEditor.getState;
  const ed = st().session!.ed;
  const save = () => {
    const text = (draft ?? '').trim();
    if (text && text !== c.text) st().run(editComment(ed, part, c.authorId, c.index, text));
    setDraft(null);
  };
  return (
    <li className="comment" data-testid="comment">
      <div className="comment-head"><span className="comment-initials" aria-hidden="true">{c.initials}</span><strong>{c.author}</strong><span className="comment-when">{when(c.date)}</span></div>
      {draft === null
        ? <p className="comment-text">{c.text}</p>
        : <textarea autoFocus aria-label="Comment" value={draft} data-testid="comment-edit" onChange={e => setDraft(e.target.value)}
            onKeyDown={e => { e.stopPropagation(); if (e.key === 'Escape') setDraft(null); else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) save(); }} />}
      <div className="comment-tools">
        {draft === null
          ? <><button type="button" onClick={() => setDraft(c.text)} data-testid="comment-edit-btn">Edit</button>
              <button type="button" onClick={() => st().run(deleteComment(ed, part, c.authorId, c.index))} data-testid="comment-delete">Delete</button></>
          : <><button type="button" onClick={save} data-testid="comment-save">Save</button><button type="button" onClick={() => setDraft(null)}>Cancel</button></>}
      </div>
    </li>
  );
}

/**
 * Comments on the slide on screen, oldest first, with a box to add one. They are written under a name the pane
 * asks for once and keeps in this browser. Adding, changing and deleting are commands, so undo covers them.
 */
export function CommentsPane() {
  const s = useEditor(st => st.session);
  const part = useEditor(st => st.slide);
  useEditor(st => st.version);
  const [name, setName] = useState(storedName);
  const [text, setText] = useState('');
  const box = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    const focus = () => box.current?.focus();
    window.addEventListener('focus-new-comment', focus);
    return () => window.removeEventListener('focus-new-comment', focus);
  }, []);
  if (!s || !part) return null;
  const comments = slideComments(s.ed, part);
  const add = () => {
    const body = text.trim(), who = name.trim();
    if (!body || !who) return;
    keepName(who);
    useEditor.getState().run(addComment(s.ed, part, body, { name: who }));
    setText('');
  };
  return (
    <aside className="pane comments-pane" aria-label="Comments" data-testid="comments-pane">
      <header><strong>Comments</strong><button type="button" aria-label="Close" onClick={() => useEditor.setState({ pane: 'none' })}>×</button></header>
      {comments.length === 0 && <p className="hint">No comments on this slide.</p>}
      <ol>{comments.map(c => <Comment key={`${c.authorId}:${c.index}`} c={c} part={part} />)}</ol>
      <div className="comment-new">
        <input type="text" aria-label="Your name" placeholder="Your name" value={name} onChange={e => setName(e.target.value)} onKeyDown={e => e.stopPropagation()} data-testid="comment-name" />
        <textarea ref={box} aria-label="New comment" placeholder="Add a comment" value={text} data-testid="comment-new" onChange={e => setText(e.target.value)}
          onKeyDown={e => { e.stopPropagation(); if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) add(); }} />
        <button type="button" disabled={!text.trim() || !name.trim()} onClick={add} data-testid="comment-add">Comment</button>
      </div>
    </aside>
  );
}
