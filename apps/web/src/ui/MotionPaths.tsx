// apps/web/src/ui/MotionPaths.tsx  -- motion paths on the slide: the ones a shape has, and one being drawn
import { useRef, useState } from 'react';
import { addMotionPath, compound, motionPathsOf, removeAnimation, type Command, type PathPoint } from '@pptx/editor-core';
import { listEffects } from '@pptx/anim';
import { ShapeView } from '@pptx/model';
import { useEditor } from '../state/store.ts';

/** Where the slide is on the page: its corner (px, in the slide area's own coordinates) and px per EMU. */
export interface SlidePlace { ox: number; oy: number; zoom: number; }

const line = (pts: { x: number; y: number }[]) => pts.map(p => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');

/**
 * The motion paths of the selected shape, drawn from its middle: a dashed line with a green start and a red end,
 * as PowerPoint shows them. They are only shown; a path is changed by drawing it again.
 */
export function MotionPathOverlay({ place, center, spid }: { place: SlidePlace; center: PathPoint; spid: number }) {
  const s = useEditor(st => st.session);
  const part = useEditor(st => st.slide);
  useEditor(st => st.version);
  if (!s || !part) return null;
  const paths = motionPathsOf(s.ed, part, spid);
  if (!paths.length) return null;
  const px = (p: PathPoint) => ({ x: place.ox + (center.x + p.x) * place.zoom, y: place.oy + (center.y + p.y) * place.zoom });
  return (
    <svg className="motion-paths" data-testid="motion-paths" aria-hidden="true">
      {paths.map(p => {
        const pts = p.points.map(px);
        const a = pts[0]!, b = pts[pts.length - 1]!;
        return (
          <g key={p.index} data-path={p.path}>
            <polyline points={line(pts)} className="motion-path" />
            <circle cx={a.x} cy={a.y} r={4} className="motion-start" />
            <circle cx={b.x} cy={b.y} r={4} className="motion-end" />
          </g>
        );
      })}
    </svg>
  );
}

/** Finish a drawn path: animate the shape along it (in place of its animations, or as one more). False when it goes nowhere. */
export function finishPath(points: PathPoint[]): boolean {
  const st = useEditor.getState();
  const want = st.pathDrawing, s = st.session, part = st.slide;
  if (!want || !s || !part) return false;
  const view = new ShapeView(want.el, s.ed.doc.owner(part));
  let add: Command;
  try {
    add = addMotionPath(s.ed, part, view, points);
  } catch {
    return false;
  }
  const old = want.replace ? listEffects(s.ed.slide(part)).filter(e => e.spid === view.id).map(e => e.index).sort((a, b) => b - a) : [];
  const steps = [...old.map(i => removeAnimation(s.ed, part, i)), add];
  useEditor.setState({ pathDrawing: null });
  st.run(steps.length === 1 ? add : compound('Draw Motion Path', steps));
  return true;
}

/**
 * Drawing a path: the slide area takes the mouse; a drag draws the line, and letting go animates the shape along
 * it. A click that goes nowhere leaves the drawing on, to try again; Escape puts it down.
 */
export function PathDrawer({ place }: { place: SlidePlace }) {
  const [pts, setPts] = useState<{ x: number; y: number }[] | null>(null);
  const drawn = useRef<{ x: number; y: number }[]>([]);
  const at = (e: React.PointerEvent) => { const r = e.currentTarget.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; };
  return (
    <div
      className="path-drawer" data-testid="path-drawer"
      onPointerDown={e => { if (e.button !== 0) return; e.currentTarget.setPointerCapture(e.pointerId); drawn.current = [at(e)]; setPts(drawn.current); }}
      onPointerMove={e => { if (!pts) return; drawn.current = [...drawn.current, at(e)]; setPts(drawn.current); }}
      onPointerUp={e => {
        if (!pts) return;
        const all = [...drawn.current, at(e)];
        setPts(null);
        drawn.current = [];
        const ok = finishPath(all.map(p => ({ x: (p.x - place.ox) / place.zoom, y: (p.y - place.oy) / place.zoom })));
        if (!ok) useEditor.getState().notify('info', 'Drag on the slide to draw the path. Press Esc to stop.');
      }}
      onContextMenu={e => e.preventDefault()}
    >
      {pts && pts.length > 1 && <svg aria-hidden="true"><polyline points={line(pts)} className="motion-path drawing" /></svg>}
    </div>
  );
}
