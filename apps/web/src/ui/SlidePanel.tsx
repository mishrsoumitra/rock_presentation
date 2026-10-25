// apps/web/src/ui/SlidePanel.tsx  -- slide thumbnails: pick, reorder by dragging, delete and duplicate
import { useEffect, useRef, useState } from 'react';
import { ANIM_UNDERSTOOD } from '@pptx/anim';
import { useEditor } from '../state/store.ts';
import { deleteCurrentSlide, duplicateCurrentSlide, moveCurrentSlide } from './actions.ts';

const THUMB = 168;

/** One thumbnail, rendered by the worker when it scrolls into view and again when its slide changes. */
function Thumb({ part, index, hidden, animated }: { part: string; index: number; hidden: boolean; animated: boolean }) {
  const session = useEditor(s => s.session)!;
  const current = useEditor(s => s.slide === part);
  useEditor(s => s.version);
  const ref = useRef<HTMLCanvasElement>(null);
  const [visible, setVisible] = useState(false);
  const host = useRef<HTMLDivElement>(null);
  const key = session.scenes.key(part);
  const size = session.ed.doc.presentation.slideSize;
  const h = Math.round((THUMB * size.cy) / size.cx);

  useEffect(() => {
    const el = host.current;
    if (!el) return;
    const io = new IntersectionObserver(es => setVisible(es.some(e => e.isIntersecting)), { rootMargin: '200px' });
    io.observe(el);
    return () => io.disconnect();
  }, []);

  useEffect(() => {
    if (!visible) return;
    let live = true;
    const ratio = window.devicePixelRatio || 1;
    session.thumbs.get(part, key, Math.round(THUMB * ratio)).then(bmp => {
      const c = ref.current;
      if (!live || !c || !bmp) return;
      c.width = bmp.width;
      c.height = bmp.height;
      c.getContext('2d')!.drawImage(bmp, 0, 0);
    });
    return () => { live = false; };
  }, [visible, key, part, session, session.thumbs.generation]);

  return (
    <div
      ref={host}
      className={`thumb${current ? ' current' : ''}${hidden ? ' hidden-slide' : ''}`}
      draggable
      data-testid={`thumb-${index + 1}`}
      onDragStart={e => { e.dataTransfer.setData('text/x-slide', part); e.dataTransfer.effectAllowed = 'move'; }}
      onDragOver={e => { if (e.dataTransfer.types.includes('text/x-slide')) e.preventDefault(); }}
      onDrop={e => {
        const from = e.dataTransfer.getData('text/x-slide');
        if (from && from !== part) moveCurrentSlide(index, from);
      }}
      onClick={() => useEditor.getState().setSlide(part)}
    >
      <span className="thumb-num">{index + 1}{animated && <span className="thumb-anim" title="Animations or a transition">★</span>}</span>
      <canvas ref={ref} style={{ width: THUMB, height: h }} aria-label={`Slide ${index + 1}`} />
    </div>
  );
}

export function SlidePanel() {
  const session = useEditor(s => s.session);
  useEditor(s => s.version);
  if (!session) return <aside className="slides" />;
  const doc = session.ed.doc;
  const parts = doc.slideParts();
  return (
    <aside
      className="slides" tabIndex={0} aria-label="Slides"
      onKeyDown={e => {
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteCurrentSlide(); }
        else if ((e.metaKey || e.ctrlKey) && e.key === 'd') { e.preventDefault(); duplicateCurrentSlide(); }
        else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
          e.preventDefault();
          const i = parts.indexOf(useEditor.getState().slide ?? '');
          const next = parts[Math.max(0, Math.min(parts.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)))];
          if (next) useEditor.getState().setSlide(next);
        }
      }}
    >
      {parts.map((p, i) => {
        const v = doc.slide(p);
        return <Thumb key={p} part={p} index={i} hidden={v.hidden} animated={!!v.timing || !!v.transition(ANIM_UNDERSTOOD)} />;
      })}
    </aside>
  );
}
