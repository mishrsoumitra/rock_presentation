// packages/editor-core/src/arrange.ts  -- z-order, grouping and ungrouping
import { removeEffectsOf } from '@pptx/anim';
import { child, ShapeView, type ChildSpace, type Xfrm } from '@pptx/model';
import { boxMatrix, childSpaceMatrix, multiply } from '@pptx/resolve';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { effectiveXfrm, unionBounds, visualBounds } from './transform.ts';
import { element, insertAt, nextShapeId, shapesOfUnit, unitOf, unitsIn } from './xml.ts';

export type ZOrder = 'front' | 'back' | 'forward' | 'backward';

/** The container (shape tree or group) of shapes, which must all share it. */
function containerOf(shapes: ShapeView[]): XmlElement {
  const containers = new Set(shapes.map(s => unitOf(s.el).parent));
  const c = [...containers][0];
  if (containers.size !== 1 || !c) throw new Error('the shapes are not in the same group');
  return c;
}

/** Put units back into the slots units occupied, in a new order; text, nvGrpSpPr and extLst stay where they are. */
function reorderUnits(container: XmlElement, order: XmlElement[]): void {
  const slots: number[] = [];
  container.children.forEach((c, i) => { if (c.kind === 'element' && order.includes(c)) slots.push(i); });
  slots.forEach((slot, i) => { container.children[slot] = order[i]!; });
}

/**
 * Bring to front, send to back, or move one step. With several shapes selected their order among themselves is
 * kept, and a step moves each past the nearest unselected neighbour, as PowerPoint does.
 */
export function zOrder(ed: Editor, part: string, shapes: ShapeView[], op: ZOrder): Command {
  const container = containerOf(shapes);
  const label = { front: 'Bring to Front', back: 'Send to Back', forward: 'Bring Forward', backward: 'Send Backward' }[op];
  return ed.edit({
    label, parts: [part], scope: () => [container],
    apply: () => {
      const units = unitsIn(container);
      const chosen = new Set(shapes.map(s => unitOf(s.el)));
      let order = units.slice();
      if (op === 'front') order = [...units.filter(u => !chosen.has(u)), ...units.filter(u => chosen.has(u))];
      else if (op === 'back') order = [...units.filter(u => chosen.has(u)), ...units.filter(u => !chosen.has(u))];
      else if (op === 'forward') {
        for (let i = order.length - 2; i >= 0; i--) {
          if (chosen.has(order[i]!) && !chosen.has(order[i + 1]!)) [order[i], order[i + 1]] = [order[i + 1]!, order[i]!];
        }
      } else {
        for (let i = 1; i < order.length; i++) {
          if (chosen.has(order[i]!) && !chosen.has(order[i - 1]!)) [order[i], order[i - 1]] = [order[i - 1]!, order[i]!];
        }
      }
      reorderUnits(container, order);
    },
  });
}

/** Remove the animations of shapes, as PowerPoint does when they are grouped or ungrouped. */
function dropEffects(ed: Editor, part: string, ids: number[]): void {
  if (!child(ed.doc.root(part), 'p:timing')) return;
  for (const id of ids) removeEffectsOf(ed.slide(part), id);
}

/**
 * Group shapes: a p:grpSp whose off and ext are the union of the shapes' bounds and whose chOff and chExt equal
 * them, so the shapes keep their coordinates. It takes the place of the topmost shape in the z-order. The shapes'
 * own animations are removed, as PowerPoint does: effects play on whole groups.
 */
export function group(ed: Editor, part: string, shapes: ShapeView[]): Command {
  const container = containerOf(shapes);
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Group', parts: [part], scope: () => [root],
    apply: () => {
      const chosen = new Set(shapes.map(s => unitOf(s.el)));
      const units = unitsIn(container).filter(u => chosen.has(u));
      if (units.length < 2) throw new Error('grouping needs two or more shapes');
      const b = unionBounds(shapes.map(s => visualBounds(effectiveXfrm(ed, part, s))));
      const r = (n: number) => Math.round(n);
      const id = nextShapeId(ed.doc.root(part));
      const g = element(container, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id - 1}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
        + `<p:grpSpPr><a:xfrm><a:off x="${r(b.x)}" y="${r(b.y)}"/><a:ext cx="${r(b.w)}" cy="${r(b.h)}"/>`
        + `<a:chOff x="${r(b.x)}" y="${r(b.y)}"/><a:chExt cx="${r(b.w)}" cy="${r(b.h)}"/></a:xfrm></p:grpSpPr></p:grpSp>`, ed.strict);
      const at = container.children.indexOf(units[units.length - 1]!);
      insertAt(container, at + 1, g);
      for (const u of units) {
        container.children.splice(container.children.indexOf(u), 1);
        g.children.push(u);
        u.parent = g;
      }
      dropEffects(ed, part, shapes.map(s => s.id));
    },
  });
}

/** A child's transform carried out of its group into the group's parent space. */
function ungroupedXfrm(groupX: Xfrm, space: ChildSpace | null, c: Xfrm): Xfrm {
  const m = multiply(boxMatrix(groupX), childSpaceMatrix(groupX, space));
  const cx = c.x + c.cx / 2, cy = c.y + c.cy / 2;
  const centre = [m[0] * cx + m[2] * cy + m[4], m[1] * cx + m[3] * cy + m[5]] as const;
  // the child's own axes, through the group's scale and rotation, give its new width and height
  const rad = (c.rot / 60000) * Math.PI / 180;
  const ux = [Math.cos(rad), Math.sin(rad)], uy = [-Math.sin(rad), Math.cos(rad)];
  const len = (v: number[]) => Math.hypot(m[0] * v[0]! + m[2] * v[1]!, m[1] * v[0]! + m[3] * v[1]!);
  const w = c.cx * len(ux), h = c.cy * len(uy);
  const oneFlip = groupX.flipH !== groupX.flipV;
  const rot = ((((oneFlip ? -c.rot : c.rot) + groupX.rot) % 21600000) + 21600000) % 21600000;
  return {
    x: centre[0] - w / 2, y: centre[1] - h / 2, cx: w, cy: h, rot,
    flipH: groupX.flipH !== c.flipH, flipV: groupX.flipV !== c.flipV,
  };
}

/**
 * Ungroup: the members take the group's place in the z-order, with their transforms carried out of it. The group's
 * animations go with it (their target no longer exists).
 */
export function ungroup(ed: Editor, part: string, groupShape: ShapeView): Command {
  if (groupShape.kind !== 'grpSp') throw new Error('not a group');
  const unit = unitOf(groupShape.el);
  const container = unit.parent;
  if (!container || unit !== groupShape.el) throw new Error('this group cannot be ungrouped');
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Ungroup', parts: [part], scope: () => [root],
    apply: () => {
      const gx = groupShape.xfrm ?? { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };
      const space = groupShape.childSpace;
      const members = unitsIn(groupShape.el);
      for (const u of members) {
        for (const s of shapesOfUnit(u)) {
          const v = new ShapeView(s, ed.doc.owner(part));
          const x = v.xfrm;
          if (x) v.setXfrm(ungroupedXfrm(gx, space, x));
        }
      }
      const at = container.children.indexOf(unit);
      container.children.splice(at, 1, ...members);
      for (const u of members) u.parent = container;
      unit.parent = null;
      dropEffects(ed, part, [groupShape.id]);
    },
  });
}

/** Rename a shape (the selection pane). */
export function rename(ed: Editor, part: string, shape: ShapeView, name: string): Command {
  return ed.edit({
    label: 'Rename', parts: [part], scope: () => [unitOf(shape.el)],
    apply: () => {
      const cNvPr = shape.cNvPr;
      if (!cNvPr) return;
      const a = cNvPr.attrs.find(x => x.local === 'name' && x.ns === '');
      if (a) a.value = name;
      else cNvPr.attrs.push({ prefix: null, local: 'name', value: name, ns: '' });
    },
  });
}

/** Alt text (cNvPr descr), which PowerPoint shows in the Alt Text pane. */
export function setAltText(ed: Editor, part: string, shape: ShapeView, descr: string): Command {
  return ed.edit({
    label: 'Alt Text', parts: [part], scope: () => [unitOf(shape.el)],
    apply: () => {
      const cNvPr = shape.cNvPr;
      if (!cNvPr) return;
      cNvPr.attrs = cNvPr.attrs.filter(x => !(x.local === 'descr' && x.ns === ''));
      if (descr) cNvPr.attrs.push({ prefix: null, local: 'descr', value: descr, ns: '' });
    },
  });
}

