// apps/web/src/ui/SlidePanel.tsx  -- slide thumbnails, in the side panel and in grid view: pick, reorder, delete, duplicate
import { memo, useEffect, useRef, useState, type ClipboardEvent } from 'react';
import { ANIM_UNDERSTOOD } from '@pptx/anim';
import { commentCount, removeAllSections, removeSection, renameSection, sections, type Section } from '@pptx/editor-core';
import { useEditor } from '../state/store.ts';
import { deleteCurrentSlide, duplicateCurrentSlide, masterPages, moveCurrentSlide, copyCurrentSlide, cutCurrentSlide, pasteSlideClip } from './actions.ts';

const THUMB = 168;

/**
 * One thumbnail, rendered by the worker when it scrolls into view and again when its slide changes. It re-renders
 * when its own scene key moves (its slide, layout, master, theme or the presentation), not on every change anywhere.
 */
const Thumb = memo(function Thumb({ part, index, width = THUMB, prefix = 'thumb', onOpen }: { part: string; index: number; width?: number; prefix?: string; onOpen?: () => void }) {
  const session = useEditor(s => s.session)!;
  const current = useEditor(s => s.slide === part);
  const key = useEditor(s => (s.session ? s.session.scenes.key(part) : ''));
  const synced = useEditor(s => s.session?.thumbs.synced ?? 0);
  const generation = useEditor(s => s.session?.thumbs.generation ?? 0);
  useEditor(s => (s.pane === 'comments' ? s.version : 0));        // comment counts follow while comments are being worked on
  const ref = useRef<HTMLCanvasElement>(null);
  const [visible, setVisible] = useState(false);
  const host = useRef<HTMLDivElement>(null);
  const size = session.ed.doc.presentation.slideSize;
  const h = Math.round((width * size.cy) / size.cx);
  // (a slide just deleted may still be asked before the panel drops its thumbnail)
  const v = session.ed.pkg.getPart(part) ? session.ed.doc.slide(part) : undefined;
  const hidden = v?.hidden ?? false, animated = !!v && (!!v.timing || !!v.transition(ANIM_UNDERSTOOD));
  // (the presentation's version, which this re-renders on, moves when a slide gets its first comment or loses its last)
  const comments = v ? commentCount(session.ed, part) : 0;

  useEffect(() => {
    const el = host.current;
    if (!el) return;
    const io = new IntersectionObserver(es => setVisible(es.some(e => e.isIntersecting)), { rootMargin: '200px' });
    io.observe(el);
    return () => io.disconnect();
  }, []);

  // the current slide stays in view however it became current: arrow keys, Page Down on the slide, a new slide
  useEffect(() => {
    if (current) host.current?.scrollIntoView({ block: 'nearest' });
  }, [current]);

  useEffect(() => {
    // the slide being typed in is sent to the worker after a pause: until then its thumbnail would be stale
    if (!visible || session.thumbs.isPending(part)) return;
    let live = true;
    const ratio = window.devicePixelRatio || 1;
    session.thumbs.get(part, key, Math.round(width * ratio)).then(bmp => {
      const c = ref.current;
      if (!live || !c || !bmp) return;
      c.width = bmp.width;
      c.height = bmp.height;
      c.getContext('2d')!.drawImage(bmp, 0, 0);
    });
    return () => { live = false; };
  }, [visible, key, part, session, generation, synced, width]);

  return (
    <div
      ref={host}
      className={`thumb${current ? ' current' : ''}${hidden ? ' hidden-slide' : ''}`}
      draggable
      data-testid={`${prefix}-${index + 1}`}
      onDragStart={e => { e.dataTransfer.setData('text/x-slide', part); e.dataTransfer.effectAllowed = 'move'; }}
      onDragOver={e => { if (e.dataTransfer.types.includes('text/x-slide')) e.preventDefault(); }}
      onDrop={e => {
        const from = e.dataTransfer.getData('text/x-slide');
        if (from && from !== part) moveCurrentSlide(index, from);
      }}
      onClick={() => useEditor.getState().setSlide(part)}
      onDoubleClick={onOpen}
      onContextMenu={e => {
        e.preventDefault();
        if (useEditor.getState().slide !== part) useEditor.getState().setSlide(part);
        useEditor.setState({ menu: { x: e.clientX, y: e.clientY, kind: 'thumb' } });
      }}
    >
      <span className="thumb-num">{index + 1}{animated && <span className="thumb-anim" title="Animations or a transition">★</span>}
        {comments > 0 && <span className="thumb-comments" title={`${comments} comment${comments === 1 ? '' : 's'}`} data-testid={`${prefix}-comments-${index + 1}`}>{comments}</span>}</span>
      <canvas ref={ref} style={{ width, height: h }} aria-label={`Slide ${index + 1}`} />
    </div>
  );
});

/** Ctrl/Cmd+Z and Ctrl/Cmd+Y (or Shift+Z) where slides are listed: the editor's undo and redo, as on the slide. True when it was one of them. */
function undoKeys(e: { key: string; metaKey: boolean; ctrlKey: boolean; shiftKey: boolean; preventDefault(): void }): boolean {
  if (!(e.metaKey || e.ctrlKey)) return false;
  const k = e.key.toLowerCase();
  if (k !== 'z' && k !== 'y') return false;
  e.preventDefault();
  if (k === 'y' || e.shiftKey) useEditor.getState().redo(); else useEditor.getState().undo();
  return true;
}

/**
 * Ctrl/Cmd+C, X and V where slides are listed (the panel, the grid) copy, cut and paste the current slide; the
 * clipboard carries a mark that names the slide held in memory, so only a paste of that same clipboard brings it back.
 */
function slideClipboard(parts: string[]) {
  const number = () => parts.indexOf(useEditor.getState().slide ?? '') + 1;
  const put = (e: ClipboardEvent, n: number, id: string) => {
    e.preventDefault();
    e.clipboardData.setData('text/plain', `Slide ${n}`);
    e.clipboardData.setData('text/html', `<div data-pptx-slide="${id}"></div>`);
  };
  return {
    onCopy: (e: ClipboardEvent) => { const id = copyCurrentSlide(); if (id) put(e, number(), id); },
    onCut: (e: ClipboardEvent) => { const n = number(); const id = cutCurrentSlide(); if (id) put(e, n, id); },
    onPaste: (e: ClipboardEvent) => {
      const held = useEditor.getState().slideClip;
      if (!held || !e.clipboardData.getData('text/html').includes(`data-pptx-slide="${held.id}"`)) return;
      e.preventDefault();
      pasteSlideClip();
    },
  };
}

/** A master or a layout in master view's list: its picture and its name; a click makes it the page being edited. */
function MasterThumb({ part, kind, name, index }: { part: string; kind: 'master' | 'layout'; name: string; index: number }) {
  const session = useEditor(s => s.session)!;
  const current = useEditor(s => s.slide === part);
  const key = useEditor(s => (s.session ? s.session.scenes.key(part) : ''));
  const synced = useEditor(s => s.session?.thumbs.synced ?? 0);
  const ref = useRef<HTMLCanvasElement>(null);
  const size = session.ed.doc.presentation.slideSize;
  const width = kind === 'master' ? THUMB : THUMB - 24;
  useEffect(() => {
    let live = true;
    void session.thumbs.get(part, key, Math.round(width * (window.devicePixelRatio || 1))).then(bmp => {
      const c = ref.current;
      if (!live || !c || !bmp) return;
      c.width = bmp.width;
      c.height = bmp.height;
      c.getContext('2d')!.drawImage(bmp, 0, 0);
    });
    return () => { live = false; };
  }, [key, part, session, synced, width]);
  return (
    <div className={`thumb master-thumb ${kind}${current ? ' current' : ''}`} data-testid={`master-thumb-${index + 1}`} data-kind={kind} title={name}
      onClick={() => useEditor.getState().setSlide(part)}>
      <canvas ref={ref} style={{ width, height: Math.round((width * size.cy) / size.cx) }} aria-label={name} />
      <span className="master-name">{name}</span>
    </div>
  );
}

/** Master view's list: each slide master, and under it its layouts. */
export function MasterPanel() {
  const session = useEditor(s => s.session);
  useEditor(s => (s.session ? s.session.ed.versionOf(s.session.ed.doc.mainPart) : -1));
  if (!session) return <aside className="slides" />;
  return (
    <aside className="slides masters" tabIndex={0} aria-label="Slide masters and layouts" data-testid="master-panel" onKeyDown={e => { undoKeys(e); }}>
      {masterPages().map((p, i) => <MasterThumb key={p.part} index={i} {...p} />)}
    </aside>
  );
}

/** The thumbnails as wide as grid view shows them. */
const GRID_THUMB = 240;

/**
 * Grid view (PowerPoint's slide sorter): every slide as a larger thumbnail, in rows. A click picks a slide, a drag
 * moves it, the arrow keys go between slides in all four directions, and a double click or Enter opens the slide
 * in the normal view. Delete, Ctrl+D, the clipboard and the right-click menu work as in the side panel.
 */
export function SlideGrid() {
  const session = useEditor(s => s.session);
  useEditor(s => (s.session ? s.session.ed.versionOf(s.session.ed.doc.mainPart) : -1));
  const host = useRef<HTMLDivElement>(null);
  useEffect(() => { host.current?.focus({ preventScroll: true }); }, []);
  if (!session) return null;
  const parts = session.ed.doc.slideParts();
  const open = () => useEditor.setState({ view: 'normal' });
  /** How many thumbnails fit in a row, from where they have been laid out. */
  const perRow = () => {
    const items = host.current ? [...host.current.querySelectorAll<HTMLElement>('.thumb')] : [];
    const top = items[0]?.offsetTop;
    const n = items.findIndex(i => i.offsetTop !== top);
    return n < 0 ? Math.max(1, items.length) : n;
  };
  return (
    <div
      ref={host} className="slide-grid" tabIndex={0} aria-label="Slides, grid view" data-testid="slide-grid" {...slideClipboard(parts)}
      onKeyDown={e => {
        const i = parts.indexOf(useEditor.getState().slide ?? '');
        const go = (to: number) => { e.preventDefault(); const next = parts[Math.max(0, Math.min(parts.length - 1, to))]; if (next) useEditor.getState().setSlide(next); };
        if (undoKeys(e)) return;
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteCurrentSlide(); }
        else if ((e.metaKey || e.ctrlKey) && e.key === 'd') { e.preventDefault(); duplicateCurrentSlide(); }
        else if (e.key === 'Enter' || e.key === 'Escape') { e.preventDefault(); open(); }
        else if (e.key === 'ArrowRight') go(i + 1);
        else if (e.key === 'ArrowLeft') go(i - 1);
        else if (e.key === 'ArrowDown') go(i + perRow());
        else if (e.key === 'ArrowUp') go(i - perRow());
        else if (e.key === 'Home') go(0);
        else if (e.key === 'End') go(parts.length - 1);
      }}
    >
      {parts.map((p, i) => <Thumb key={p} part={p} index={i} width={GRID_THUMB} prefix="grid-thumb" onOpen={open} />)}
    </div>
  );
}

export function SlidePanel() {
  const session = useEditor(s => s.session);
  const folded = useEditor(s => s.folded);
  // the list of slides changes with the presentation part (slides added, removed, moved), not with their content
  useEditor(s => (s.session ? s.session.ed.versionOf(s.session.ed.doc.mainPart) : -1));
  if (!session) return <aside className="slides" />;
  const doc = session.ed.doc;
  const parts = doc.slideParts();
  const secs = sections(session.ed);
  return (
    <aside
      className="slides" tabIndex={0} aria-label="Slides" {...slideClipboard(parts)}
      onKeyDown={e => {
        if (undoKeys(e)) return;
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteCurrentSlide(); }
        else if ((e.metaKey || e.ctrlKey) && e.key === 'd') { e.preventDefault(); duplicateCurrentSlide(); }
        else if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Home' || e.key === 'End') {
          e.preventDefault();
          const i = parts.indexOf(useEditor.getState().slide ?? '');
          const to = e.key === 'Home' ? 0 : e.key === 'End' ? parts.length - 1 : i + (e.key === 'ArrowDown' ? 1 : -1);
          const next = parts[Math.max(0, Math.min(parts.length - 1, to))];
          if (next) useEditor.getState().setSlide(next);
        }
      }}
    >
      {secs.length === 0
        ? parts.map((p, i) => <Thumb key={p} part={p} index={i} />)
        : secs.map(sec => (
          <div key={sec.id || sec.index} className="section" data-testid="section" data-name={sec.name}>
            <SectionHead section={sec} folded={folded.includes(sec.id)} />
            {!folded.includes(sec.id) && sec.slides.map(p => <Thumb key={p} part={p} index={parts.indexOf(p)} />)}
          </div>
        ))}
    </aside>
  );
}

/**
 * A section's heading in the slide panel: its name and how many slides it has. The arrow folds the section shut
 * (its slides are only hidden from the panel); a double click on the name renames it; the cross removes the
 * section and keeps its slides.
 */
function SectionHead({ section, folded }: { section: Section; folded: boolean }) {
  const [draft, setDraft] = useState<string | null>(null);
  const st = useEditor.getState;
  const ed = st().session!.ed;
  const commit = () => {
    const name = (draft ?? '').trim();
    if (name && name !== section.name) st().run(renameSection(ed, section.index, name));
    setDraft(null);
  };
  const fold = () => useEditor.setState(s => ({ folded: folded ? s.folded.filter(id => id !== section.id) : [...s.folded, section.id] }));
  return (
    <div className="section-head" data-testid="section-head">
      <button type="button" className="section-fold" aria-expanded={!folded} aria-label={`${folded ? 'Expand' : 'Collapse'} ${section.name}`} onClick={fold} data-testid="section-fold">{folded ? '▸' : '▾'}</button>
      {draft === null
        ? <span className="section-name" title="Double-click to rename" onDoubleClick={() => setDraft(section.name)}>{section.name || 'Untitled Section'}</span>
        : <input className="section-rename" autoFocus aria-label="Section name" value={draft} data-testid="section-rename" onChange={e => setDraft(e.target.value)} onBlur={commit}
            onKeyDown={e => { e.stopPropagation(); if (e.key === 'Enter') commit(); else if (e.key === 'Escape') setDraft(null); }} />}
      <span className="section-count">{section.slides.length}</span>
      <button type="button" className="section-remove" aria-label={`Remove section ${section.name}`} title="Remove section (its slides stay)" onClick={() => st().run(removeSection(ed, section.index))} data-testid="section-remove">×</button>
    </div>
  );
}

/** Remove every section of the deck (the menu's Remove All Sections). */
export function removeSections(): void {
  const s = useEditor.getState();
  if (s.session) s.run(removeAllSections(s.session.ed));
}
