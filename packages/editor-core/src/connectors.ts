// packages/editor-core/src/connectors.ts  -- connectors glued to shapes: connection points, attaching, following
import { evaluateGeometry, presetGeometry } from '@pptx/geometry';
import { attr, child, insertInOrder, int, removeChildren, ShapeView, type Xfrm } from '@pptx/model';
import { boxMatrix } from '@pptx/resolve';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { element, unitsIn, shapesOfUnit } from './xml.ts';

export interface Point { x: number; y: number; }
/** A connection point of a shape: its index (what a connector stores) and where it is, in the shape's container. */
export interface Site extends Point { index: number; }
/** What an end of a connector is attached to: a shape, by id, and one of its connection points. */
export interface Attachment { id: number; index: number; }

/** A shape's own frame; shapes placed by their placeholder (no frame of their own) have no connection points here. */
const frameOf = (shape: ShapeView): Xfrm | undefined => shape.xfrm ?? undefined;

/**
 * The connection points of a shape, as its geometry defines them (a rectangle's four edge middles, an ellipse's
 * eight points...), where they are now: through the shape's size, rotation and flips. Connectors, groups and
 * shapes without a frame of their own have none.
 */
export function connectionSites(shape: ShapeView): Site[] {
  const x = frameOf(shape);
  if (!x || shape.kind === 'cxnSp' || shape.kind === 'grpSp') return [];
  const g = shape.geometry;
  const geom = g?.kind === 'preset' ? presetGeometry(g.prst) : undefined;
  const m = boxMatrix(x);
  const at = (u: number, v: number) => ({ x: m[0] * u + m[2] * v + m[4], y: m[1] * u + m[3] * v + m[5] });
  let local: Point[];
  try {
    local = geom ? evaluateGeometry(geom, x.cx, x.cy, g?.kind === 'preset' ? g.av : []).connectionSites : [];
  } catch {
    local = [];
  }
  // a picture, a table or a custom shape: the middles of its four sides, in a rectangle's order
  if (!local.length) local = [{ x: x.cx / 2, y: 0 }, { x: 0, y: x.cy / 2 }, { x: x.cx / 2, y: x.cy }, { x: x.cx, y: x.cy / 2 }];
  return local.map((p, index) => ({ index, ...at(p.x, p.y) }));
}

/** Whether a shape is a connector this module handles: a p:cxnSp, unrotated (its ends are then its frame's corners). */
export function isConnector(shape: ShapeView): boolean {
  return shape.kind === 'cxnSp' && !!shape.xfrm && shape.xfrm.rot === 0;
}

/** Where a connector starts and ends: its frame's corners, which corner being said by its flips. */
export function connectorEnds(shape: ShapeView): { start: Point; end: Point } {
  const x = shape.xfrm!;
  return {
    start: { x: x.flipH ? x.x + x.cx : x.x, y: x.flipV ? x.y + x.cy : x.y },
    end: { x: x.flipH ? x.x : x.x + x.cx, y: x.flipV ? x.y : x.y + x.cy },
  };
}

/** The frame of a connector that runs from one point to another. */
export function frameBetween(start: Point, end: Point): Xfrm {
  return {
    x: Math.round(Math.min(start.x, end.x)), y: Math.round(Math.min(start.y, end.y)), cx: Math.round(Math.abs(end.x - start.x)), cy: Math.round(Math.abs(end.y - start.y)),
    rot: 0, flipH: start.x > end.x, flipV: start.y > end.y,
  };
}

const NV = (cxn: XmlElement) => child(child(cxn, 'p:nvCxnSpPr'), 'p:cNvCxnSpPr');

/** What each end of a connector is attached to, from a:stCxn and a:endCxn. */
export function attachments(shape: ShapeView): { start?: Attachment; end?: Attachment } {
  const nv = NV(shape.el);
  const read = (name: string): Attachment | undefined => {
    const e = child(nv, name);
    return e ? { id: int(attr(e, 'id'), 0), index: int(attr(e, 'idx'), 0) } : undefined;
  };
  const start = read('a:stCxn'), end = read('a:endCxn');
  return { ...(start ? { start } : {}), ...(end ? { end } : {}) };
}

function writeAttachment(ed: Editor, cxn: XmlElement, name: 'a:stCxn' | 'a:endCxn', to: Attachment | null): void {
  const nv = NV(cxn);
  if (!nv) return;
  removeChildren(nv, name);
  if (to) insertInOrder(nv, element(nv, `<${name} id="${to.id}" idx="${to.index}"/>`, ed.strict));
  nv.selfClosing = nv.children.length === 0;
}

/** The shapes that sit in the same container as a shape (its slide's shape tree or its group), by id. */
function siblings(shape: XmlElement, owner: ShapeView['owner']): Map<number, ShapeView> {
  const out = new Map<number, ShapeView>();
  const container = shape.parent;
  if (!container) return out;
  for (const u of unitsIn(container)) for (const s of shapesOfUnit(u)) { const v = new ShapeView(s, owner); if (!out.has(v.id)) out.set(v.id, v); }
  return out;
}

/**
 * Put a connector's ends at two points and say what each is attached to (null: to nothing; undefined: as it was).
 * This is what dragging an end does: every move of a drag passes one `coalesceKey`, making the drag one undo step.
 */
export function setConnectorEnds(
  ed: Editor, part: string, connector: ShapeView, start: Point, end: Point, attach: { start?: Attachment | null; end?: Attachment | null } = {}, coalesceKey?: string,
): Command {
  return ed.edit({
    label: 'Move End Point', parts: [part], scope: () => [connector.el], ...(coalesceKey ? { coalesceKey } : {}),
    apply: () => {
      connector.setXfrm(frameBetween(start, end));
      if (attach.start !== undefined) writeAttachment(ed, connector.el, 'a:stCxn', attach.start);
      if (attach.end !== undefined) writeAttachment(ed, connector.el, 'a:endCxn', attach.end);
    },
  });
}

/**
 * The connectors of a container that are attached to any of some shapes, with where their ends should now be: an
 * attached end at its connection point, a free end where it is. Used after shapes move, so connectors follow.
 */
export function rerouted(ed: Editor, part: string, moved: ShapeView[]): { connector: ShapeView; xfrm: Xfrm }[] {
  const out: { connector: ShapeView; xfrm: Xfrm }[] = [];
  const ids = new Set(moved.map(s => s.id));
  const seen = new Set<XmlElement>();
  for (const s of moved) {
    const all = siblings(s.el, ed.doc.owner(part));
    for (const c of all.values()) {
      if (seen.has(c.el) || !isConnector(c) || ids.has(c.id)) continue;
      const a = attachments(c);
      if (!((a.start && ids.has(a.start.id)) || (a.end && ids.has(a.end.id)))) continue;
      seen.add(c.el);
      const ends = connectorEnds(c);
      const point = (at: Attachment | undefined, fallback: Point): Point => {
        const target = at ? all.get(at.id) : undefined;
        const site = target ? connectionSites(target)[at!.index] : undefined;
        return site ?? fallback;
      };
      out.push({ connector: c, xfrm: frameBetween(point(a.start, ends.start), point(a.end, ends.end)) });
    }
  }
  return out;
}

/** The connection point nearest a point among the shapes of a container, within a distance; undefined when none is. */
export function nearestSite(shapes: ShapeView[], at: Point, within: number, except?: XmlElement): { shape: ShapeView; site: Site } | undefined {
  let best: { shape: ShapeView; site: Site; d: number } | undefined;
  for (const shape of shapes) {
    if (shape.el === except) continue;
    for (const site of connectionSites(shape)) {
      const d = Math.hypot(site.x - at.x, site.y - at.y);
      if (d <= within && (!best || d < best.d)) best = { shape, site, d };
    }
  }
  return best ? { shape: best.shape, site: best.site } : undefined;
}
