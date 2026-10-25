// apps/web/src/ui/Presenter.tsx  -- the slide show: full window, played by the Show engine
import { useEffect, useRef } from 'react';
import { useEditor } from '../state/store.ts';
import { Show } from './show.ts';

export function Presenter() {
  const session = useEditor(s => s.session)!;
  const ref = useRef<HTMLCanvasElement>(null);
  const show = useRef<Show | null>(null);

  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const doc = session.ed.doc;
    const parts = doc.slideParts().filter(p => !doc.slide(p).hidden);
    const start = Math.max(0, parts.indexOf(useEditor.getState().slide ?? ''));
    let last: string | undefined;
    const exit = () => useEditor.setState({ mode: 'edit', slide: last ?? useEditor.getState().slide });
    const size = () => {
      const ratio = window.devicePixelRatio || 1;
      const s = doc.presentation.slideSize;
      const scale = Math.min(window.innerWidth / s.cx, window.innerHeight / s.cy);
      const w = Math.round(s.cx * scale), h = Math.round(s.cy * scale);
      canvas.width = Math.round(w * ratio);
      canvas.height = Math.round(h * ratio);
      canvas.style.width = `${w}px`;
      canvas.style.height = `${h}px`;
      show.current?.start();
    };
    size();
    const s = new Show(session, canvas, parts, start, exit);
    s.onSlide = part => { if (part) last = part; };
    show.current = s;
    // pictures arrive after the first frames: draw again when they do
    const unsubscribe = useEditor.subscribe((st, prev) => { if (st.version !== prev.version) s.start(); });
    const onKey = (e: KeyboardEvent) => {
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(e.key)) { e.preventDefault(); s.next(); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(e.key)) { e.preventDefault(); s.previous(); }
      else if (e.key === 'Escape') { e.preventDefault(); exit(); }
      else if (e.key === 'Home') s.goTo(0);
      else if (e.key === 'End') s.goTo(parts.length - 1);
    };
    window.addEventListener('keydown', onKey);
    window.addEventListener('resize', size);
    document.documentElement.requestFullscreen?.().catch(() => undefined);
    return () => {
      s.stop();
      unsubscribe();
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('resize', size);
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    };
  }, [session]);

  return (
    <div className="presenter" data-testid="presenter"
      onClick={e => show.current?.click(e.clientX, e.clientY)}
      onContextMenu={e => { e.preventDefault(); show.current?.previous(); }}>
      <canvas ref={ref} />
    </div>
  );
}
