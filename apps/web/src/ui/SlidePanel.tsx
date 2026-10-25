// apps/web/src/ui/SlidePanel.tsx  -- slide thumbnails: pick, reorder by dragging, delete and duplicate
import { memo, useEffect, useRef, useState } from 'react';
import { ANIM_UNDERSTOOD } from '@pptx/anim';
import { useEditor } from '../state/store.ts';
import { deleteCurrentSlide, duplicateCurrentSlide, moveCurrentSlide } from './actions.ts';

const THUMB = 168;

/**
 * One thumbnail, rendered by the worker when it scrolls into view and again when its slide changes. It re-renders
 * when its own scene key moves (its slide, layout, master, theme or the presentation), not on every change anywhere.
 */
const Thumb = memo(function Thumb({ part, index }: { part: string; index: number }) {
  const session = useEditor(s => s.session)!;
  const current = useEditor(s => s.slide === part);
  const key = useEditor(s => (s.session ? s.session.scenes.key(part) : ''));
  const synced = useEditor(s => s.session?.thumbs.synced ?? 0);
  const generation = useEditor(s => s.session?.thumbs.generation ?? 0);
  const ref = useRef<HTMLCanvasElement>(null);
  const [visible, setVisible] = useState(false);
  const host = useRef<HTMLDivElement>(null);
  const size = session.ed.doc.presentation.slideSize;
  const h = Math.round((THUMB * size.cy) / size.cx);
  // (a slide just deleted may still be asked before the panel drops its thumbnail)
  const v = session.ed.pkg.getPart(part) ? session.ed.doc.slide(part) : undefined;
  const hidden = v?.hidden ?? false, animated = !!v && (!!v.timing || !!v.transition(ANIM_UNDERSTOOD));

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
    session.thumbs.get(part, key, Math.round(THUMB * ratio)).then(bmp => {
      const c = ref.current;
      if (!live || !c || !bmp) return;
      c.width = bmp.width;
      c.height = bmp.height;
      c.getContext('2d')!.drawImage(bmp, 0, 0);
    });
    return () => { live = false; };
  }, [visible, key, part, session, generation, synced]);

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
});

export function SlidePanel() {
  const session = useEditor(s => s.session);
  // the list of slides changes with the presentation part (slides added, removed, moved), not with their content
  useEditor(s => (s.session ? s.session.ed.versionOf(s.session.ed.doc.mainPart) : -1));
  if (!session) return <aside className="slides" />;
  const doc = session.ed.doc;
  const parts = doc.slideParts();
  return (
    <aside
      className="slides" tabIndex={0} aria-label="Slides"
      onKeyDown={e => {
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
      {parts.map((p, i) => <Thumb key={p} part={p} index={i} />)}
    </aside>
  );
}
