// packages/editor-core/src/shapes.ts  -- inserting, deleting, duplicating, copying and pasting shapes
import { normalizeName, resolveTarget, type OpcPackage } from '@pptx/opc';
import { removeEffectsOf } from '@pptx/anim';
import { child, clone, declareUsedNamespaces, is, relKind, ShapeView } from '@pptx/model';
import { serializeElement, type XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { copyPart, dropUnreferenced, mediaPart, relRefs, renameRelRefs } from './parts.ts';
import { convertFlavour, element, esc, fragment, idOf, insertAt, nextShapeId, shapeIds, shapesOfUnit, unitOf, unitsIn } from './xml.ts';

export interface Rect { x: number; y: number; w: number; h: number; }

const CONNECTORS = /^(line|straightConnector1|bentConnector[2-5]|curvedConnector[2-5])$/;

/** PowerPoint's names for new shapes ("Oval 4"); others are the preset name in words. */
const NAMES: Record<string, string> = {
  rect: 'Rectangle', roundRect: 'Rectangle: Rounded Corners', ellipse: 'Oval', triangle: 'Isosceles Triangle', rtTriangle: 'Right Triangle',
  line: 'Straight Connector', straightConnector1: 'Straight Arrow Connector', rightArrow: 'Arrow: Right', leftArrow: 'Arrow: Left',
  upArrow: 'Arrow: Up', downArrow: 'Arrow: Down', star5: 'Star: 5 Points', diamond: 'Diamond', pentagon: 'Arrow: Pentagon',
  hexagon: 'Hexagon', octagon: 'Octagon', parallelogram: 'Parallelogram', trapezoid: 'Trapezoid', cloud: 'Cloud', heart: 'Heart',
  wedgeRectCallout: 'Speech Bubble: Rectangle', wedgeRoundRectCallout: 'Speech Bubble: Rectangle with Corners Rounded',
  wedgeEllipseCallout: 'Speech Bubble: Oval', flowChartProcess: 'Flowchart: Process', flowChartDecision: 'Flowchart: Decision',
};

export function shapeName(prst: string): string {
  return NAMES[prst] ?? prst.replace(/([a-z])([A-Z0-9])/g, '$1 $2').replace(/^./, c => c.toUpperCase());
}

/** Where new shapes go: on top of everything, before the shape tree's extLst. */
function topSlot(tree: XmlElement): number {
  const units = unitsIn(tree);
  const last = units[units.length - 1];
  if (last) return tree.children.indexOf(last) + 1;
  const pr = child(tree, 'p:grpSpPr') ?? child(tree, 'p:nvGrpSpPr');
  return pr ? tree.children.indexOf(pr) + 1 : tree.children.length;
}

const r = Math.round;
const xfrmXml = (b: Rect) => `<a:xfrm><a:off x="${r(b.x)}" y="${r(b.y)}"/><a:ext cx="${r(Math.max(0, b.w))}" cy="${r(Math.max(0, b.h))}"/></a:xfrm>`;

/**
 * Insert a preset shape as PowerPoint writes a newly drawn one: no fill or line of its own, a p:style that takes
 * them from the theme (accent 1, line darker), and an empty centred paragraph ready for text. Lines and connectors
 * are p:cxnSp with the theme's connector style. Returns the command and the new shape's id.
 */
export function insertShape(ed: Editor, part: string, prst: string, box: Rect): { command: Command; id: number } {
  const tree = ed.spTree(part);
  const id = nextShapeId(ed.doc.root(part));
  const name = `${esc(shapeName(prst))} ${id - 1}`;
  const xml = CONNECTORS.test(prst)
    ? `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>`
      + `<p:spPr>${xfrmXml(box)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${prst === 'straightConnector1' ? '<a:ln><a:tailEnd type="triangle"/></a:ln>' : ''}</p:spPr>`
      + '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
      + '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>'
    : `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
      + `<p:spPr>${xfrmXml(box)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></p:spPr>`
      + '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
      + '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'
      + '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
  const command = ed.edit({
    label: 'Insert Shape', parts: [part], scope: () => [tree],
    apply: () => insertAt(tree, topSlot(tree), element(tree, xml, ed.strict)),
  });
  return { command, id };
}

/** Paragraph XML for plain text: one a:p per line, U+000B as a line break inside a paragraph. */
export function paragraphsXml(text: string, lang = 'en-US', pPr = ''): string {
  return text.split(/\r\n|\r|\n/).map(line => {
    const runs = line.split('\u000b').map(t => (t ? `<a:r><a:rPr lang="${lang}" dirty="0"/><a:t>${esc(t)}</a:t></a:r>` : '')).join('<a:br><a:rPr lang="' + lang + '" dirty="0"/></a:br>');
    return `<a:p>${pPr}${runs}<a:endParaRPr lang="${lang}" dirty="0"/></a:p>`;
  }).join('');
}

/** Insert a text box as PowerPoint's Text Box tool does: no fill, square wrap, and the box grows with its text. */
export function insertTextBox(ed: Editor, part: string, box: Rect, text = ''): { command: Command; id: number } {
  const tree = ed.spTree(part);
  const id = nextShapeId(ed.doc.root(part));
  const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xfrmXml(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
    + `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphsXml(text)}</p:txBody></p:sp>`;
  const command = ed.edit({
    label: 'Insert Text Box', parts: [part], scope: () => [tree],
    apply: () => insertAt(tree, topSlot(tree), element(tree, xml, ed.strict)),
  });
  return { command, id };
}

/**
 * Delete shapes: their units leave the tree, a group left empty goes too, connectors attached to them are
 * detached, and relationships and parts only they used are removed (a picture's media part stays while any part
 * still shows it), and so do their animations. `onDelete` lets the caller remove anything else that refers to them
 * inside the same step.
 */
export function deleteShapes(ed: Editor, part: string, shapes: ShapeView[], onDelete?: (ids: Set<number>) => void): Command {
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Delete', parts: [part], structural: true, scope: () => [root],
    apply: () => {
      const units = [...new Set(shapes.map(s => unitOf(s.el)))];
      const ids = new Set<number>();
      const refs: string[] = [];
      for (const u of units) {
        for (const id of shapeIds(u)) ids.add(id);
        refs.push(...relRefs(u).map(x => x.value));
      }
      for (const u of units) {
        let parent = u.parent;
        if (!parent) continue;
        parent.children.splice(parent.children.indexOf(u), 1);
        u.parent = null;
        // an emptied group goes as well, up the chain
        while (parent && is(parent, 'p:grpSp') && unitsIn(parent).length === 0) {
          const up: XmlElement | null = unitOf(parent).parent;
          if (!up) break;
          const unit = unitOf(parent);
          for (const id of shapeIds(unit)) ids.add(id);
          up.children.splice(up.children.indexOf(unit), 1);
          unit.parent = null;
          parent = up;
        }
      }
      detachConnectors(root, ids);
      // their animations go too: an effect aimed at a shape that is gone makes PowerPoint repair the file
      if (child(root, 'p:timing')) for (const id of ids) removeEffectsOf(ed.slide(part), id);
      onDelete?.(ids);
      dropUnreferenced(ed.pkg, part, refs);
    },
  });
}

/** Connectors glued to shapes that are gone keep their geometry but lose the glue (a:stCxn, a:endCxn). */
function detachConnectors(root: XmlElement, ids: Set<number>): void {
  const walk = (e: XmlElement) => {
    if (is(e, 'p:cNvCxnSpPr')) {
      e.children = e.children.filter(c => !(c.kind === 'element' && (is(c, 'a:stCxn') || is(c, 'a:endCxn')) && ids.has(Number(c.attrs.find(a => a.local === 'id')?.value))));
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(root);
}

/** Give every shape in copied units a fresh id, keeping branches of an alternate content block on one id. */
function renumber(units: XmlElement[], taken: Set<number>, start: number): Map<number, number> {
  const map = new Map<number, number>();
  let next = start;
  const walk = (e: XmlElement) => {
    if (e.local === 'cNvPr') {
      const a = e.attrs.find(x => x.local === 'id' && x.ns === '');
      const old = Number(a?.value);
      if (a) {
        let fresh = map.get(old);
        if (fresh === undefined) {
          while (taken.has(next)) next++;
          fresh = next++;
          map.set(old, fresh);
          taken.add(fresh);
        }
        a.value = String(fresh);
      }
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  for (const u of units) walk(u);
  // connectors glued to shapes copied along follow the copies
  const glue = (e: XmlElement) => {
    if (is(e, 'a:stCxn') || is(e, 'a:endCxn')) {
      const a = e.attrs.find(x => x.local === 'id' && x.ns === '');
      if (a && map.has(Number(a.value))) a.value = String(map.get(Number(a.value)));
    }
    for (const c of e.children) if (c.kind === 'element') glue(c);
  };
  for (const u of units) glue(u);
  return map;
}

/** Shift a unit's shapes by (dx, dy) in their parent space. */
function offsetUnit(ed: Editor, part: string, unit: XmlElement, dx: number, dy: number): void {
  if (!dx && !dy) return;
  for (const s of shapesOfUnit(unit)) {
    const v = new ShapeView(s, ed.doc.owner(part));
    const x = v.xfrm;
    if (x) v.setXfrm({ ...x, x: x.x + dx, y: x.y + dy });
  }
}

/** Relationship kinds a copy of a shape gets its own copy of, so the copy can change independently. */
const OWNED = new Set(['chart', 'chartUserShapes', 'diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors', 'diagramDrawing',
  'oleObject', 'package', 'control', 'vmlDrawing']);

/**
 * Re-point relationship references of copied XML that will live in `part`: shared targets (pictures, links) get
 * a relationship from `part` (the same one when there is one), owned targets (charts, diagrams, embedded objects)
 * are copied. `from` and `source` say where the XML came from; `mapShared` handles another package's targets.
 */
function relink(
  ed: Editor, part: string, from: OpcPackage, source: string, units: XmlElement[],
  mapShared?: (kind: string, target: string) => string | null | undefined,
): void {
  const pkg = ed.pkg;
  const srcRels = from.relationships(source);
  const mine = pkg.relationships(part);
  const map = new Map<string, string>();
  const dropped = new Set<string>();
  const copied = new Map<string, string>();
  for (const id of new Set(units.flatMap(u => relRefs(u).map(x => x.value)))) {
    const rel = srcRels.find(x => x.id === id);
    if (!rel) { dropped.add(id); continue; }
    const kind = relKind(rel.type);
    if (rel.external) {
      const same = mine.find(x => x.external && x.type === rel.type && x.target === rel.target);
      map.set(id, same ? same.id : pkg.addRelationship(part, rel.type, rel.target, { external: true }));
      continue;
    }
    const t = resolveTarget(normalizeName(source), rel.target);
    let target: string | null | undefined;
    if (OWNED.has(kind)) target = from.getPart(t) ? copyPart(pkg, from, t, copied, mapShared) : null;
    else target = from === pkg ? (pkg.getPart(t) ? t : null) : mapShared?.(kind, t);
    if (target === undefined && from.getPart(t)) target = copyPart(pkg, from, t, copied, mapShared);
    if (!target) { dropped.add(id); continue; }
    const same = mine.find(x => !x.external && x.type === rel.type && resolveTarget(normalizeName(part), x.target) === target);
    map.set(id, same && !OWNED.has(kind) ? same.id : pkg.addRelationship(part, rel.type, target));
  }
  for (const u of units) {
    renameRelRefs(u, map);
    // a link whose target could not come along is removed with the element that holds it
    if (dropped.size) removeDangling(u, dropped);
  }
}

function removeDangling(root: XmlElement, ids: Set<string>): void {
  for (const ref of relRefs(root)) {
    if (!ids.has(ref.value)) continue;
    const el = ref.el;
    if (is(el, 'a:hlinkClick') || is(el, 'a:hlinkMouseOver') || is(el, 'a:hlinkHover')) {
      const p = el.parent;
      if (p) p.children.splice(p.children.indexOf(el), 1);
    }
  }
}

/** Duplicate shapes (Ctrl+D): copies on top, offset a little, with fresh ids and their own charts and diagrams. */
export function duplicateShapes(ed: Editor, part: string, shapes: ShapeView[], offset = { dx: 152400, dy: 152400 }): { command: Command; ids: number[] } {
  const root = ed.doc.root(part);
  const ids: number[] = [];
  const command = ed.edit({
    label: 'Duplicate', parts: [part], structural: true, scope: () => [root],
    apply: () => {
      const units = [...new Set(shapes.map(s => unitOf(s.el)))];
      const taken = shapeIds(root);
      const copies = units.map(u => clone(u));
      const map = renumber(copies, taken, Math.max(0, ...taken) + 1);
      relink(ed, part, ed.pkg, part, copies);
      for (let i = 0; i < units.length; i++) {
        const container = units[i]!.parent!;
        offsetUnit(ed, part, copies[i]!, offset.dx, offset.dy);
        insertAt(container, topSlot(container), copies[i]!);
        const first = shapesOfUnit(units[i]!)[0];
        if (first) ids.push(map.get(idOf(first)) ?? 0);
      }
    },
  });
  return { command, ids };
}

/** What Copy puts on the clipboard: the units as XML, and where they came from (to bring their parts along). */
export interface ShapeClip {
  xml: string[];
  from: OpcPackage;
  source: string;
  slideSize: { cx: number; cy: number };
}

/** A unit as standalone XML: the namespace declarations it relies on are copied onto it. */
function serializeUnit(u: XmlElement): string {
  const c = clone(u);
  declareUsedNamespaces(c);
  return serializeElement(c);
}

export function copyShapes(ed: Editor, part: string, shapes: ShapeView[]): ShapeClip {
  const units = [...new Set(shapes.map(s => unitOf(s.el)))];
  return {
    xml: units.map(serializeUnit),
    from: ed.pkg, source: part, slideSize: ed.doc.presentation.slideSize,
  };
}

/**
 * Paste copied shapes onto a slide: fresh ids, relationships re-created from this slide (pictures from another
 * deck become media parts here, reusing identical ones; slide links to another deck are dropped), and an offset
 * when pasting onto the slide they came from, as PowerPoint does.
 */
export function pasteShapes(ed: Editor, part: string, clip: ShapeClip, hashes?: ReadonlyMap<string, string>): { command: Command; ids: number[] } {
  const root = ed.doc.root(part);
  const tree = ed.spTree(part);
  const ids: number[] = [];
  const command = ed.edit({
    label: 'Paste', parts: [part], structural: true, scope: () => [root],
    apply: () => {
      const units = clip.xml.flatMap(x => fragment(tree, x, ed.strict));
      for (const u of units) convertFlavour(u, ed.strict);
      const taken = shapeIds(root);
      const map = renumber(units, taken, Math.max(0, ...taken) + 1);
      relink(ed, part, clip.from, clip.source, units, clip.from === ed.pkg ? undefined : (kind, target) => {
        if (kind === 'image' || kind === 'media' || kind === 'video' || kind === 'audio') {
          const p = clip.from.getPart(target);
          const hash = hashes?.get(target);
          if (!p || !hash) return null;
          const ext = target.slice(target.lastIndexOf('.') + 1).toLowerCase();
          return mediaPart(ed.pkg, p.bytes, ext, p.contentType, hash, kind === 'image' ? 'image' : 'media');
        }
        return null;
      });
      const same = clip.from === ed.pkg && clip.source === part;
      for (const u of units) {
        if (same) offsetUnit(ed, part, u, 152400, 152400);
        insertAt(tree, topSlot(tree), u);
        const first = shapesOfUnit(u)[0];
        if (first) ids.push(map.get(idOf(first)) ?? idOf(first));
      }
    },
  });
  return { command, ids };
}

