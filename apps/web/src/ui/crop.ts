// apps/web/src/ui/crop.ts  -- cropping a picture with handles: the geometry, apart from the mouse and the canvas
import type { Xfrm } from '@pptx/model';

/** How much of each side of a picture is cut off, as fractions of its width or height (0 to 1). */
export interface Crop { l: number; t: number; r: number; b: number; }
/** A picture's frame and what of the picture shows in it. */
export interface Cropped { xfrm: Xfrm; crop: Crop; }

/** Which edges a handle moves: -1 the left or top edge, 1 the right or bottom edge, 0 neither, in the picture's own (unflipped, unrotated) space. */
export interface CropHandle { x: -1 | 0 | 1; y: -1 | 0 | 1; }
export const CROP_HANDLES: CropHandle[] = [
  { x: -1, y: -1 }, { x: 0, y: -1 }, { x: 1, y: -1 }, { x: -1, y: 0 }, { x: 1, y: 0 }, { x: -1, y: 1 }, { x: 0, y: 1 }, { x: 1, y: 1 },
];

/** The least of a picture that can be left showing, as a fraction of its width or height. */
const MIN = 0.02;
const rad = (rot: number) => (rot / 60000) * Math.PI / 180;

/** A vector of the slide in the picture's own space: turned back by its rotation, and mirrored where it is flipped. */
function toLocal(x: Xfrm, dx: number, dy: number): [number, number] {
  const a = -rad(x.rot);
  const u = dx * Math.cos(a) - dy * Math.sin(a), v = dx * Math.sin(a) + dy * Math.cos(a);
  return [x.flipH ? -u : u, x.flipV ? -v : v];
}
function toSlide(x: Xfrm, u: number, v: number): [number, number] {
  const a = rad(x.rot);
  const fu = x.flipH ? -u : u, fv = x.flipV ? -v : v;
  return [fu * Math.cos(a) - fv * Math.sin(a), fu * Math.sin(a) + fv * Math.cos(a)];
}

/**
 * The whole picture's rectangle, uncropped, in the picture's own space with its origin at the frame's centre:
 * what would show if nothing were cut off. Cropping changes the frame and the crop, never this.
 */
export function fullPicture(c: Cropped): { x0: number; y0: number; w: number; h: number } {
  const w = c.xfrm.cx / Math.max(MIN, 1 - c.crop.l - c.crop.r), h = c.xfrm.cy / Math.max(MIN, 1 - c.crop.t - c.crop.b);
  return { x0: -c.xfrm.cx / 2 - c.crop.l * w, y0: -c.xfrm.cy / 2 - c.crop.t * h, w, h };
}

/**
 * Drag a crop handle by a displacement on the slide (EMU), from the state the drag started in. The picture stays
 * exactly where it is: the frame's edge moves over it, between the picture's own edge and the least that may show,
 * and the frame is placed so that what still shows has not moved (rotated and flipped pictures included).
 */
export function dragCrop(start: Cropped, handle: CropHandle, dx: number, dy: number): Cropped {
  const { xfrm, crop } = start;
  const [du, dv] = toLocal(xfrm, dx, dy);
  const full = fullPicture(start);
  // the showing part, as distances from the picture's left and top edges
  let x0 = crop.l * full.w, x1 = (1 - crop.r) * full.w, y0 = crop.t * full.h, y1 = (1 - crop.b) * full.h;
  const ox0 = x0, ox1 = x1, oy0 = y0, oy1 = y1;
  const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));
  if (handle.x < 0) x0 = clamp(x0 + du, 0, x1 - MIN * full.w);
  if (handle.x > 0) x1 = clamp(x1 + du, x0 + MIN * full.w, full.w);
  if (handle.y < 0) y0 = clamp(y0 + dv, 0, y1 - MIN * full.h);
  if (handle.y > 0) y1 = clamp(y1 + dv, y0 + MIN * full.h, full.h);
  // the frame's centre moves by half of what each edge moved, in the picture's space, then on the slide
  const [sx, sy] = toSlide(xfrm, ((x0 + x1) - (ox0 + ox1)) / 2, ((y0 + y1) - (oy0 + oy1)) / 2);
  const cx = x1 - x0, cy = y1 - y0;
  const centre = { x: xfrm.x + xfrm.cx / 2 + sx, y: xfrm.y + xfrm.cy / 2 + sy };
  const snap = (v: number) => (Math.abs(v) < 1e-9 ? 0 : v);
  return {
    xfrm: { ...xfrm, x: centre.x - cx / 2, y: centre.y - cy / 2, cx, cy },
    crop: { l: snap(x0 / full.w), r: snap(1 - x1 / full.w), t: snap(y0 / full.h), b: snap(1 - y1 / full.h) },
  };
}

/** A point of the picture's own space (origin at the frame's centre) on the slide. */
export function onSlide(xfrm: Xfrm, u: number, v: number): { x: number; y: number } {
  const [sx, sy] = toSlide(xfrm, u, v);
  return { x: xfrm.x + xfrm.cx / 2 + sx, y: xfrm.y + xfrm.cy / 2 + sy };
}

/** Where a crop handle sits on the slide: on the frame's edge or corner it moves. */
export function handlePoint(xfrm: Xfrm, handle: CropHandle): { x: number; y: number } {
  return onSlide(xfrm, (handle.x * xfrm.cx) / 2, (handle.y * xfrm.cy) / 2);
}

/** The corners of the whole picture on the slide (x, y pairs, clockwise in its own space), for the outline behind the frame. */
export function fullCorners(c: Cropped): number[] {
  const f = fullPicture(c);
  return [[f.x0, f.y0], [f.x0 + f.w, f.y0], [f.x0 + f.w, f.y0 + f.h], [f.x0, f.y0 + f.h]].flatMap(([u, v]) => { const p = onSlide(c.xfrm, u!, v!); return [p.x, p.y]; });
}
