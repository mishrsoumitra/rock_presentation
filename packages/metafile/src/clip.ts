// packages/metafile/src/clip.ts  -- clipping regions: rectangle lists combined exactly, paths by intersection
//
// A canvas can only narrow its clip, so a region is kept as a list of terms that are all intersected, and the
// canvas clip is rebuilt from scratch whenever the region changes. Rectangle regions (what GDI regions are) combine
// exactly in every mode; a path combines exactly by intersection or difference, approximately otherwise.
import type { Path } from './geom.ts';
import type { Rect } from './types.ts';

export type ClipTerm =
  | { kind: 'rects'; rects: Rect[] }                       // the union of the rectangles
  | { kind: 'path'; path: Path; rule: CanvasFillRule }     // inside the path
  | { kind: 'outside'; path: Path };                       // outside the path

/** The intersection of its terms, in output coordinates; no terms is no clipping. Never mutated. */
export type Clip = readonly ClipTerm[];

export const RGN_AND = 1;
export const RGN_OR = 2;
export const RGN_XOR = 3;
export const RGN_DIFF = 4;
export const RGN_COPY = 5;

/** Most rectangles a region keeps; beyond that it is replaced by its bounding box. */
const MAX_RECTS = 4096;
/** Stands in for "everywhere" when a rectangle is taken out of an unclipped region. */
export const EVERYWHERE: Rect = { l: -1e9, t: -1e9, r: 1e9, b: 1e9 };

const empty = (r: Rect) => r.r <= r.l || r.b <= r.t;

/** The bounding box, as a one-rectangle region. */
function box(rects: Rect[]): Rect[] {
  let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
  for (const x of rects) { l = Math.min(l, x.l); t = Math.min(t, x.t); r = Math.max(r, x.r); b = Math.max(b, x.b); }
  return l < r && t < b ? [{ l, t, r, b }] : [];
}

const bound = (rects: Rect[]) => (rects.length <= MAX_RECTS ? rects : box(rects));

export function intersectRects(a: Rect[], b: Rect[]): Rect[] {
  const out: Rect[] = [];
  // too many pairs: the larger list stands in as its bounding box
  if (a.length * b.length > MAX_RECTS * 16) {
    if (a.length > b.length) a = box(a); else b = box(b);
  }
  for (const x of a) {
    for (const y of b) {
      const r = { l: Math.max(x.l, y.l), t: Math.max(x.t, y.t), r: Math.min(x.r, y.r), b: Math.min(x.b, y.b) };
      if (!empty(r)) out.push(r);
    }
  }
  return bound(out);
}

/** a minus b, as up to four rectangles. */
function cut(a: Rect, b: Rect): Rect[] {
  if (b.l >= a.r || b.r <= a.l || b.t >= a.b || b.b <= a.t) return [a];
  const out: Rect[] = [];
  if (b.t > a.t) out.push({ l: a.l, t: a.t, r: a.r, b: b.t });
  if (b.b < a.b) out.push({ l: a.l, t: b.b, r: a.r, b: a.b });
  const t = Math.max(a.t, b.t), bb = Math.min(a.b, b.b);
  if (b.l > a.l) out.push({ l: a.l, t, r: b.l, b: bb });
  if (b.r < a.r) out.push({ l: b.r, t, r: a.r, b: bb });
  return out.filter(r => !empty(r));
}

export function subtractRects(a: Rect[], b: Rect[]): Rect[] {
  let out = a;
  for (const y of b) {
    out = out.flatMap(x => cut(x, y));
    if (out.length > MAX_RECTS) return bound(out);
  }
  return out;
}

const rectsOf = (clip: Clip) => clip.find(t => t.kind === 'rects') as { kind: 'rects'; rects: Rect[] } | undefined;
const withRects = (clip: Clip, rects: Rect[]): Clip => [...clip.filter(t => t.kind !== 'rects'), { kind: 'rects', rects }];

/** Combine a region (rectangles in output coordinates; null for RGN_COPY: the default, unclipped) with the clip. */
export function combineRects(clip: Clip, mode: number, region: Rect[] | null): Clip {
  const rects = (region ?? []).filter(r => !empty(r));
  const current = rectsOf(clip);
  switch (mode) {
    case RGN_COPY: return region === null ? [] : [{ kind: 'rects', rects: bound(rects) }];
    case RGN_AND: return withRects(clip, current ? intersectRects(current.rects, rects) : bound(rects));
    case RGN_DIFF: return withRects(clip, subtractRects(current ? current.rects : [EVERYWHERE], rects));
    case RGN_OR:
      // a union cannot narrow a path term: the paths are dropped, which clips less than asked
      if (clip.length === 0) return [];
      return [{ kind: 'rects', rects: bound([...(current ? current.rects : [EVERYWHERE]), ...rects]) }];
    case RGN_XOR: {
      const a = current ? current.rects : [EVERYWHERE];
      return withRects(clip, bound([...subtractRects(a, rects), ...subtractRects(rects, a)]));
    }
    default: return clip;
  }
}

/** Combine a path (output coordinates) with the clip. */
export function combinePath(clip: Clip, mode: number, path: Path, rule: CanvasFillRule): Clip {
  switch (mode) {
    case RGN_COPY: return [{ kind: 'path', path, rule }];
    case RGN_AND: return [...clip, { kind: 'path', path, rule }];
    case RGN_DIFF: return [...clip, { kind: 'outside', path }];
    case RGN_OR: return [];                                      // approximated: no clipping
    case RGN_XOR: return [{ kind: 'path', path, rule: 'evenodd' }];
    default: return clip;
  }
}

/** The clip moved by (dx, dy) output units. */
export function translateClip(clip: Clip, dx: number, dy: number): Clip {
  const m = [1, 0, 0, 1, dx, dy] as [number, number, number, number, number, number];
  return clip.map(t => t.kind === 'rects'
    ? { kind: 'rects', rects: t.rects.map(r => r === EVERYWHERE ? r : { l: r.l + dx, t: r.t + dy, r: r.r + dx, b: r.b + dy }) }
    : { ...t, path: t.path.map(m) });
}
