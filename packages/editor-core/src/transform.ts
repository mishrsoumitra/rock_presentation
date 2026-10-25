// packages/editor-core/src/transform.ts  -- moving, resizing, rotating, flipping, aligning and distributing shapes
import { is, ShapeView, type Xfrm } from '@pptx/model';
import { boxMatrix, childSpaceMatrix, IDENTITY, invert, matchLayoutPlaceholder, matchMasterPlaceholder, multiply, type Matrix } from '@pptx/resolve';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { twins, unitOf } from './xml.ts';

export interface Box { x: number; y: number; w: number; h: number; }

const ZERO: Xfrm = { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };

/** The group element a shape sits in, through alternate content, or undefined at the top of the shape tree. */
export function parentGroup(el: XmlElement): XmlElement | undefined {
  const container = unitOf(el).parent;
  return container && (is(container, 'p:grpSp') || is(container, 'dsp:grpSp')) ? container : undefined;
}

/**
 * The transform a shape is drawn with: its own xfrm, else (a placeholder) its layout's or master's placeholder's,
 * as the renderer resolves it.
 */
export function effectiveXfrm(ed: Editor, part: string, shape: ShapeView): Xfrm {
  const own = shape.xfrm;
  if (own) return own;
  const ph = shape.placeholder;
  if (!ph) return ZERO;
  const doc = ed.doc;
  const kind = doc.slide(part).kind;
  const layout = kind === 'slide' ? doc.layoutOf(part) : undefined;
  const master = kind === 'layout' ? doc.masterOf(part) : layout ? doc.masterOf(layout) : undefined;
  const fromLayout = layout ? matchLayoutPlaceholder(ph, doc.slide(layout).shapes) : undefined;
  if (fromLayout?.xfrm) return fromLayout.xfrm;
  const masterPh = master ? matchMasterPlaceholder((fromLayout?.placeholder ?? ph), doc.slide(master).shapes) : undefined;
  return masterPh?.xfrm ?? ZERO;
}

/** From a shape's parent coordinate space (its group's child space) to slide coordinates. */
export function parentMatrix(ed: Editor, part: string, el: XmlElement): Matrix {
  const g = parentGroup(el);
  if (!g) return IDENTITY;
  const group = new ShapeView(g, ed.doc.owner(part));
  const x = group.xfrm ?? ZERO;
  return multiply(parentMatrix(ed, part, g), multiply(boxMatrix(x), childSpaceMatrix(x, group.childSpace)));
}

/** The axis-aligned bounds of a transformed box: what PowerPoint aligns, distributes and groups by. */
export function visualBounds(x: Xfrm, m: Matrix = IDENTITY): Box {
  const full = multiply(m, boxMatrix(x));
  const pts = [[0, 0], [x.cx, 0], [x.cx, x.cy], [0, x.cy]].map(([u, v]) => [full[0] * u! + full[2] * v! + full[4], full[1] * u! + full[3] * v! + full[5]]);
  const xs = pts.map(p => p[0]!), ys = pts.map(p => p[1]!);
  const x0 = Math.min(...xs), y0 = Math.min(...ys);
  return { x: x0, y: y0, w: Math.max(...xs) - x0, h: Math.max(...ys) - y0 };
}

export function unionBounds(boxes: Box[]): Box {
  const x0 = Math.min(...boxes.map(b => b.x)), y0 = Math.min(...boxes.map(b => b.y));
  const x1 = Math.max(...boxes.map(b => b.x + b.w)), y1 = Math.max(...boxes.map(b => b.y + b.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/** A slide-space displacement in a shape's parent space (groups scale and rotate their children). */
function toParentVector(m: Matrix, dx: number, dy: number): [number, number] {
  const inv = invert([m[0], m[1], m[2], m[3], 0, 0]);
  return [inv[0] * dx + inv[2] * dy, inv[1] * dx + inv[3] * dy];
}

export interface TransformChange { shape: ShapeView; xfrm: Xfrm; }

/**
 * Set transforms: the one command behind moving, resizing, rotating and flipping. Each xfrm is in the shape's
 * parent space. A placeholder that inherited its position gets its own a:xfrm, as PowerPoint writes it; a shape in
 * an alternate content block is changed in every branch.
 */
export function setTransforms(ed: Editor, part: string, changes: TransformChange[], label = 'Move', coalesceKey?: string): Command {
  return ed.edit({
    label, parts: [part], ...(coalesceKey ? { coalesceKey } : {}),
    scope: () => changes.map(c => unitOf(c.shape.el)),
    apply: () => {
      for (const { shape, xfrm } of changes) {
        for (const el of twins(shape.el)) new ShapeView(el, ed.doc.owner(part)).setXfrm(xfrm);
      }
    },
  });
}

/** Move shapes by a displacement in slide EMU (nudging with the arrow keys coalesces into one step per run). */
export function moveBy(ed: Editor, part: string, shapes: ShapeView[], dx: number, dy: number, label = 'Move'): Command {
  const changes = shapes.map(shape => {
    const x = effectiveXfrm(ed, part, shape);
    const [px, py] = toParentVector(parentMatrix(ed, part, shape.el), dx, dy);
    return { shape, xfrm: { ...x, x: x.x + px, y: x.y + py } };
  });
  return setTransforms(ed, part, changes, label, `move:${part}:${shapes.map(s => s.id).join(',')}`);
}

/** Rotate each shape in place by a quarter turn. */
export function rotateQuarter(ed: Editor, part: string, shapes: ShapeView[], direction: 1 | -1): Command {
  return setTransforms(ed, part, shapes.map(shape => {
    const x = effectiveXfrm(ed, part, shape);
    return { shape, xfrm: { ...x, rot: (((x.rot + direction * 5400000) % 21600000) + 21600000) % 21600000 } };
  }), direction > 0 ? 'Rotate Right 90°' : 'Rotate Left 90°');
}

/** Flip each shape in place, as PowerPoint's Flip Horizontal and Flip Vertical do for one shape. */
export function flip(ed: Editor, part: string, shapes: ShapeView[], axis: 'h' | 'v'): Command {
  return setTransforms(ed, part, shapes.map(shape => {
    const x = effectiveXfrm(ed, part, shape);
    // flipping a rotated shape mirrors its rotation too, so it looks mirrored on the slide
    const rot = x.rot ? 21600000 - x.rot : 0;
    return { shape, xfrm: axis === 'h' ? { ...x, flipH: !x.flipH, rot } : { ...x, flipV: !x.flipV, rot } };
  }), axis === 'h' ? 'Flip Horizontal' : 'Flip Vertical');
}

export type Alignment = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';

/** Align shapes to each other's bounds, or to the slide when there is one shape or `toSlide` is set. */
export function align(ed: Editor, part: string, shapes: ShapeView[], how: Alignment, toSlide = shapes.length < 2): Command {
  const bounds = shapes.map(s => visualBounds(effectiveXfrm(ed, part, s), parentMatrix(ed, part, s.el)));
  const size = ed.doc.presentation.slideSize;
  const ref = toSlide ? { x: 0, y: 0, w: size.cx, h: size.cy } : unionBounds(bounds);
  return setTransforms(ed, part, shapes.map((shape, i) => {
    const b = bounds[i]!;
    const dx = how === 'left' ? ref.x - b.x : how === 'center' ? ref.x + ref.w / 2 - (b.x + b.w / 2) : how === 'right' ? ref.x + ref.w - (b.x + b.w) : 0;
    const dy = how === 'top' ? ref.y - b.y : how === 'middle' ? ref.y + ref.h / 2 - (b.y + b.h / 2) : how === 'bottom' ? ref.y + ref.h - (b.y + b.h) : 0;
    const x = effectiveXfrm(ed, part, shape);
    const [px, py] = toParentVector(parentMatrix(ed, part, shape.el), dx, dy);
    return { shape, xfrm: { ...x, x: x.x + px, y: x.y + py } };
  }), `Align ${how[0]!.toUpperCase()}${how.slice(1)}`);
}

/** Spread three or more shapes so the gaps between them are equal; the outermost two stay where they are. */
export function distribute(ed: Editor, part: string, shapes: ShapeView[], axis: 'h' | 'v'): Command {
  const items = shapes.map(shape => ({ shape, b: visualBounds(effectiveXfrm(ed, part, shape), parentMatrix(ed, part, shape.el)) }));
  const pos = (b: Box) => (axis === 'h' ? b.x : b.y);
  const len = (b: Box) => (axis === 'h' ? b.w : b.h);
  items.sort((p, q) => pos(p.b) - pos(q.b));
  const first = items[0], last = items[items.length - 1];
  const changes: TransformChange[] = [];
  if (first && last && items.length > 2) {
    const total = items.reduce((a, it) => a + len(it.b), 0);
    const gap = (pos(last.b) + len(last.b) - pos(first.b) - total) / (items.length - 1);
    let at = pos(first.b);
    for (const it of items) {
      const d = at - pos(it.b);
      const x = effectiveXfrm(ed, part, it.shape);
      const [px, py] = toParentVector(parentMatrix(ed, part, it.shape.el), axis === 'h' ? d : 0, axis === 'v' ? d : 0);
      changes.push({ shape: it.shape, xfrm: { ...x, x: x.x + px, y: x.y + py } });
      at += len(it.b) + gap;
    }
  }
  return setTransforms(ed, part, changes, axis === 'h' ? 'Distribute Horizontally' : 'Distribute Vertically');
}
