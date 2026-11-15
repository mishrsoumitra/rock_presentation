// apps/web/src/ui/Presenter.tsx  -- the slide show, alone or in presenter view, played by the Show engine
import { useEffect, useRef, useState } from 'react';
import { notesText } from '@pptx/editor-core';
import { useEditor } from '../state/store.ts';
import { clock } from './clock.ts';
import { Show } from './show.ts';

/** The next slide, small: the same picture the slide panel shows of it. */
function NextSlide({ part }: { part: string | undefined }) {
  const session = useEditor(s => s.session)!;
  const ref = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const c = ref.current;
    if (!c || !part) return undefined;
    let live = true;
    void session.thumbs.get(part, session.scenes.key(part), 640).then(bmp => {
      if (!live || !bmp) return;
      c.width = bmp.width;
      c.height = bmp.height;
      c.getContext('2d')!.drawImage(bmp, 0, 0);
    });
    return () => { live = false; };
  }, [part, session]);
  if (!part) return <div className="presenter-end" data-testid="presenter-next-end">End of slide show</div>;
  return <canvas ref={ref} className="presenter-next-canvas" data-testid="presenter-next" aria-label="Next slide" />;
}

/**
 * The slide show. Alone, the slide fills the window. In presenter view the slide shares the window with what the
 * speaker needs: the next slide, the notes of the current one, a timer (click it to start again) and the clock.
 * L turns the laser pointer on and off in both: a red dot that follows the mouse, while clicks leave the slide alone.
 */
export function Presenter() {
  const session = useEditor(s => s.session)!;
  const withNotes = useEditor(s => s.presenterView);
  const ref = useRef<HTMLCanvasElement>(null);
  const stage = useRef<HTMLDivElement>(null);
  const show = useRef<Show | null>(null);
  const [parts] = useState(() => { const doc = session.ed.doc; return doc.slideParts().filter(p => !doc.slide(p).hidden); });
  const [index, setIndex] = useState(() => Math.max(0, parts.indexOf(useEditor.getState().slide ?? '')));
  const [laser, setLaser] = useState(false);
  const [dot, setDot] = useState<{ x: number; y: number } | null>(null);
  const [started, setStarted] = useState(() => Date.now());
  const [now, setNow] = useState(() => Date.now());

  useEffect(() => {
    const canvas = ref.current, box = stage.current;
    if (!canvas || !box) return undefined;
    const doc = session.ed.doc;
    let last: string | undefined;
    const exit = () => useEditor.setState({ mode: 'edit', slide: last ?? useEditor.getState().slide });
    // the slide is as large as fits where it is shown: the window, or its part of presenter view
    const size = () => {
      const ratio = window.devicePixelRatio || 1;
      const s = doc.presentation.slideSize;
      const scale = Math.min(box.clientWidth / s.cx, box.clientHeight / s.cy);
      const w = Math.max(1, Math.round(s.cx * scale)), h = Math.max(1, Math.round(s.cy * scale));
      canvas.width = Math.round(w * ratio);
      canvas.height = Math.round(h * ratio);
      canvas.style.width = `${w}px`;
      canvas.style.height = `${h}px`;
      show.current?.start();
    };
    size();
    const s = new Show(session, canvas, parts, Math.max(0, parts.indexOf(useEditor.getState().slide ?? '')), exit);
    s.onSlide = part => { if (part) last = part; setIndex(s.index); };
    show.current = s;
    // pictures arrive after the first frames: draw again when they do
    const unsubscribe = useEditor.subscribe((st, prev) => { if (st.version !== prev.version) s.start(); });
    const onKey = (e: KeyboardEvent) => {
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(e.key)) { e.preventDefault(); s.next(); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(e.key)) { e.preventDefault(); s.previous(); }
      else if (e.key === 'Escape') { e.preventDefault(); exit(); }
      else if (e.key === 'Home') s.goTo(0);
      else if (e.key === 'End') s.goTo(parts.length - 1);
      else if (e.key === 'l' || e.key === 'L') { e.preventDefault(); setLaser(on => !on); }
    };
    window.addEventListener('keydown', onKey);
    window.addEventListener('resize', size);
    const ro = new ResizeObserver(size);
    ro.observe(box);
    document.documentElement.requestFullscreen?.().catch(() => undefined);
    return () => {
      s.stop();
      unsubscribe();
      ro.disconnect();
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('resize', size);
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    };
  }, [session, parts]);

  // the timer ticks only in presenter view, where it is shown
  useEffect(() => {
    if (!withNotes) return undefined;
    const t = setInterval(() => setNow(Date.now()), 250);
    return () => clearInterval(t);
  }, [withNotes]);

  const slideProps = {
    ref: stage,
    onClick: (e: React.MouseEvent) => { if (!laser) show.current?.click(e.clientX, e.clientY); },
    onContextMenu: (e: React.MouseEvent) => { e.preventDefault(); if (!laser) show.current?.previous(); },
    onMouseMove: (e: React.MouseEvent) => { if (laser) setDot({ x: e.clientX, y: e.clientY }); },
    onMouseLeave: () => setDot(null),
  };
  const pointer = laser && dot && <div className="laser" data-testid="laser" style={{ left: dot.x, top: dot.y }} />;

  if (!withNotes) {
    return (
      <div className={`presenter${laser ? ' lasering' : ''}`} data-testid="presenter" {...slideProps}>
        <canvas ref={ref} />
        {pointer}
      </div>
    );
  }
  const current = parts[index];
  const ended = index >= parts.length;
  const notes = current ? notesText(session.ed, current) : '';
  return (
    <div className="presenter presenter-view" data-testid="presenter">
      <div className={`presenter-slide${laser ? ' lasering' : ''}`} data-testid="presenter-slide" {...slideProps}>
        <canvas ref={ref} />
        {pointer}
      </div>
      <aside className="presenter-side">
        <div className="presenter-label">Next</div>
        <div className="presenter-next"><NextSlide part={parts[index + 1]} /></div>
        <div className="presenter-label">Notes</div>
        <div className="presenter-notes" data-testid="presenter-notes">{notes || <span className="hint">No notes for this slide.</span>}</div>
      </aside>
      <footer className="presenter-bar">
        <button type="button" className="presenter-timer" title="Time since the show started. Click to start again." data-testid="presenter-timer" onClick={() => { setStarted(Date.now()); setNow(Date.now()); }}>{clock(now - started)}</button>
        <span data-testid="presenter-clock">{new Date(now).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}</span>
        <span className="presenter-count" data-testid="presenter-count">{ended ? 'End of slide show' : `Slide ${index + 1} of ${parts.length}`}</span>
        <button type="button" aria-pressed={laser} onClick={() => setLaser(on => !on)} data-testid="presenter-laser" title="Laser pointer (L)">Laser</button>
        <button type="button" onClick={() => show.current?.previous()} data-testid="presenter-prev">Previous</button>
        <button type="button" onClick={() => show.current?.next()} data-testid="presenter-nextbtn">Next</button>
        <button type="button" onClick={() => useEditor.setState({ mode: 'edit' })} data-testid="presenter-end">End Show</button>
      </footer>
    </div>
  );
}
