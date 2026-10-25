// apps/web/src/ui/placement.ts  -- where a shape's Konva node sits, and back from a node to an a:xfrm
import type { Xfrm } from '@pptx/model';
import type { RenderShape } from '@pptx/render';
import { boxMatrix, invert, multiply, type Matrix } from '@pptx/resolve';

/**
 * A node for a shape: its box in slide EMU, centred on (x, y) and turned by `rotation` degrees about that centre,
 * which is how Konva's Transformer moves, resizes and rotates it. Flips are not part of the node: the renderer
 * draws them inside the box.
 */
export interface Placement { x: number; y: number; w: number; h: number; rotation: number; }

/** From a shape's parent space (its group's child space) to the slide. */
export function parentOf(s: RenderShape): Matrix {
  const x = s.source?.xfrm;
  return x ? multiply(s.transform, invert(boxMatrix(x))) : [1, 0, 0, 1, 0, 0];
}

/**
 * The node placement of a shape: the centre and size of its box on the slide, and its rotation (its own, plus its
 * group's). Taken from the xfrm rather than the drawn matrix, which cannot tell a shape flipped both ways from one
 * turned half round.
 */
export function placementOf(s: RenderShape): Placement {
  const x = s.source?.xfrm;
  if (!x) return { x: s.box.w / 2, y: s.box.h / 2, w: s.box.w, h: s.box.h, rotation: 0 };
  const m = parentOf(s);
  const cx = x.x + x.cx / 2, cy = x.y + x.cy / 2;
  return {
    x: m[0] * cx + m[2] * cy + m[4], y: m[1] * cx + m[3] * cy + m[5],
    w: x.cx * Math.hypot(m[0], m[1]), h: x.cy * Math.hypot(m[2], m[3]),
    rotation: x.rot / 60000 + (Math.atan2(m[1], m[0]) * 180) / Math.PI,
  };
}

/** The node's own transform at rest: box coordinates (0..w, 0..h) to slide coordinates. */
export function restMatrix(p: Placement): Matrix {
  const r = (p.rotation * Math.PI) / 180;
  const c = Math.cos(r), s = Math.sin(r);
  // T(x, y) * R * T(-w/2, -h/2)
  return [c, s, -s, c, p.x - (c * p.w) / 2 + (s * p.h) / 2, p.y - (s * p.w) / 2 - (c * p.h) / 2];
}

/**
 * What a node's sceneFunc draws under: the context's transform when the function runs (pixel ratio, stage zoom,
 * the node's live position while dragged) undoing the node's resting transform, and undoing the device scale the
 * renderer applies itself. The renderer then draws the shape at its slide coordinates and follows the node.
 */
export function drawBase(current: DOMMatrix, rest: Placement, deviceScale: number): Matrix {
  const ctm: Matrix = [current.a, current.b, current.c, current.d, current.e, current.f];
  return multiply(multiply(ctm, invert(restMatrix(rest))), [1 / deviceScale, 0, 0, 1 / deviceScale, 0, 0]);
}

/**
 * The xfrm for a shape whose node now has this box (centre, size, rotation in slide space), in the shape's parent
 * space: `parent` maps the parent space to the slide (a group's child space). Flips stay as they were.
 */
export function xfrmFor(p: Placement, parent: Matrix, old: Xfrm): Xfrm {
  const inv = invert(parent);
  const cx = inv[0] * p.x + inv[2] * p.y + inv[4];
  const cy = inv[1] * p.x + inv[3] * p.y + inv[5];
  const sx = Math.hypot(parent[0], parent[1]) || 1, sy = Math.hypot(parent[2], parent[3]) || 1;
  const groupRot = (Math.atan2(parent[1], parent[0]) * 180) / Math.PI;
  const w = p.w / sx, h = p.h / sy;
  const rot = ((Math.round((p.rotation - groupRot) * 60000) % 21600000) + 21600000) % 21600000;
  return { ...old, x: cx - w / 2, y: cy - h / 2, cx: w, cy: h, rot };
}

/** The axis-aligned bounds of a shape on the slide. */
export function boundsOf(s: RenderShape): { x: number; y: number; w: number; h: number } {
  const m = s.transform;
  const pts = [[0, 0], [s.box.w, 0], [s.box.w, s.box.h], [0, s.box.h]].map(([u, v]) => [m[0] * u! + m[2] * v! + m[4], m[1] * u! + m[3] * v! + m[5]] as const);
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  const x = Math.min(...xs), y = Math.min(...ys);
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

export const toMatrix = (m: DOMMatrix): Matrix => [m.a, m.b, m.c, m.d, m.e, m.f];
