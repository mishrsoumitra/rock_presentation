// apps/web/src/ui/Rulers.tsx  -- the rulers along the top and left of the slide area
import { useEffect, useState, type RefObject } from 'react';
import { rulerTicks, type Tick } from './rulers.ts';

const THICK = 18;
const LENGTH: Record<Tick['size'], number> = { inch: 9, half: 6, eighth: 3 };

/**
 * Two rulers in inches, zero at the slide's top left corner, that stay along the edges of the slide area as it
 * scrolls and follow the zoom. Pressing the mouse on one pulls a guide out of it: a horizontal guide from the top
 * ruler, a vertical one from the left.
 */
export function Rulers({ host, ox, oy, zoom, onPull }: {
  host: RefObject<HTMLDivElement | null>; ox: number; oy: number; zoom: number; onPull: (axis: 'x' | 'y', e: MouseEvent) => void;
}) {
  const [view, setView] = useState({ left: 0, top: 0, w: 0, h: 0 });
  useEffect(() => {
    const el = host.current;
    if (!el) return undefined;
    const read = () => setView({ left: el.scrollLeft, top: el.scrollTop, w: el.clientWidth, h: el.clientHeight });
    read();
    el.addEventListener('scroll', read, { passive: true });
    const ro = new ResizeObserver(read);
    ro.observe(el);
    return () => { el.removeEventListener('scroll', read); ro.disconnect(); };
  }, [host]);
  const ppi = 914400 * zoom;
  const across = rulerTicks(ox - view.left, ppi, view.w), down = rulerTicks(oy - view.top, ppi, view.h);
  return (
    <div className="rulers" aria-hidden="true">
      <svg className="ruler ruler-h" width={view.w} height={THICK} data-testid="ruler-h" onMouseDown={e => { e.preventDefault(); onPull('y', e.nativeEvent); }}>
        {across.map((t, i) => <line key={i} x1={t.at} x2={t.at} y1={THICK - LENGTH[t.size]} y2={THICK} />)}
        {across.filter(t => t.label !== undefined).map((t, i) => <text key={`l${i}`} x={t.at + 3} y={9}>{t.label}</text>)}
      </svg>
      <svg className="ruler ruler-v" width={THICK} height={view.h} data-testid="ruler-v" onMouseDown={e => { e.preventDefault(); onPull('x', e.nativeEvent); }}>
        {down.map((t, i) => <line key={i} y1={t.at} y2={t.at} x1={THICK - LENGTH[t.size]} x2={THICK} />)}
        {down.filter(t => t.label !== undefined).map((t, i) => <text key={`l${i}`} x={2} y={t.at + 10}>{t.label}</text>)}
      </svg>
      <div className="ruler-corner" />
    </div>
  );
}
