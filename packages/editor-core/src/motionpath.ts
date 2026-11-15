// packages/editor-core/src/motionpath.ts  -- motion paths drawn by hand: from the points of a drag to an animation
import { addEffect, findPreset, listEffects, parseMotionPath } from '@pptx/anim';
import type { XmlElement } from '@pptx/xml';
import type { ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';

export interface PathPoint { x: number; y: number; }

/** Distance from a point to the segment a-b. */
function distance(p: PathPoint, a: PathPoint, b: PathPoint): number {
  const dx = b.x - a.x, dy = b.y - a.y;
  const len = dx * dx + dy * dy;
  const t = len ? Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / len)) : 0;
  return Math.hypot(p.x - (a.x + t * dx), p.y - (a.y + t * dy));
}

/**
 * The points that carry a drawn line's shape: those further than `tolerance` from the line through their kept
 * neighbours (Ramer, Douglas and Peucker). A drag gives hundreds of points; a path needs the corners.
 */
export function simplifyPath(points: PathPoint[], tolerance: number): PathPoint[] {
  if (points.length < 3) return points.slice();
  const keep = new Uint8Array(points.length);
  keep[0] = keep[points.length - 1] = 1;
  const stack: [number, number][] = [[0, points.length - 1]];
  while (stack.length) {
    const [a, b] = stack.pop()!;
    let far = -1, most = tolerance;
    for (let i = a + 1; i < b; i++) {
      const d = distance(points[i]!, points[a]!, points[b]!);
      if (d > most) { most = d; far = i; }
    }
    if (far < 0) continue;
    keep[far] = 1;
    stack.push([a, far], [far, b]);
  }
  return points.filter((_, i) => keep[i]);
}

const num = (v: number) => { const s = v.toFixed(5).replace(/\.?0+$/, ''); return s === '-0' || s === '' ? '0' : s; };

/**
 * A motion path for points drawn on the slide (EMU): the offsets from the first point, as fractions of the slide,
 * joined by straight lines. The shape starts where it is and moves by what was drawn, wherever on the slide the
 * drawing was made, as in PowerPoint. Undefined when the drawing goes nowhere (a click, or a jitter).
 */
export function pathFromPoints(points: PathPoint[], slide: { cx: number; cy: number }, tolerance = slide.cx / 500): string | undefined {
  const kept = simplifyPath(points, tolerance);
  const first = kept[0];
  if (!first) return undefined;
  const segs: string[] = [];
  let last = first;
  for (const p of kept.slice(1)) {
    if (Math.hypot(p.x - last.x, p.y - last.y) < tolerance) continue;
    segs.push(`L ${num((p.x - first.x) / slide.cx)} ${num((p.y - first.y) / slide.cy)}`);
    last = p;
  }
  return segs.length ? `M 0 0 ${segs.join(' ')} E` : undefined;
}

/**
 * Animate a shape along a path drawn by hand (slide EMU points): a Custom Path effect, on a click, two seconds
 * long as PowerPoint makes one. Throws when the drawing goes nowhere.
 */
export function addMotionPath(ed: Editor, part: string, shape: ShapeView, points: PathPoint[]): Command {
  const path = pathFromPoints(points, ed.doc.presentation.slideSize);
  if (!path) throw new Error('the path drawn goes nowhere');
  const preset = findPreset('path', 0);
  if (!preset) throw new Error('no Custom Path preset');
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Draw Motion Path', parts: [part], scope: () => [root],
    apply: () => { addEffect(ed.slide(part), shape.id, preset, { trigger: 'click', path }); },
  });
}

const motionOf = (el: XmlElement): XmlElement | undefined => {
  if (el.local === 'animMotion') return el;
  for (const c of el.children) if (c.kind === 'element') { const m = motionOf(c); if (m) return m; }
  return undefined;
};

/**
 * The motion paths a shape follows: each effect's place in the animation pane, its path as written, and the path
 * as slide offsets (EMU) from where the shape is, flattened to points, for drawing it on the slide.
 */
export function motionPathsOf(ed: Editor, part: string, spid: number): { index: number; path: string; points: PathPoint[] }[] {
  const size = ed.doc.presentation.slideSize;
  const out: { index: number; path: string; points: PathPoint[] }[] = [];
  for (const e of listEffects(ed.slide(part))) {
    if (e.spid !== spid) continue;
    const path = motionOf(e.el)?.attrs.find(a => a.local === 'path')?.value;
    const parsed = path ? parseMotionPath(path) : undefined;
    if (path && parsed) out.push({ index: e.index, path, points: parsed.points.map(p => ({ x: p.x * size.cx, y: p.y * size.cy })) });
  }
  return out;
}
