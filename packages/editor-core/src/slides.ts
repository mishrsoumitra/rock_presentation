// packages/editor-core/src/slides.ts  -- adding, deleting, duplicating and moving slides; layouts; sections
import { encodeXmlText, type XmlElement } from '@pptx/xml';
import { normalizeName, resolveTarget } from '@pptx/opc';
import {
  attr, child, children, CONTENT_TYPE, insertInOrder, int, is, relKind, relType, removeChildren, setAttr, ShapeView, uriFor,
} from '@pptx/model';
import { matchLayoutPlaceholder } from '@pptx/resolve';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { collectOrphans, copyPart, freeName, relRefs } from './parts.ts';
import { retarget } from './pictures.ts';
import { effectiveXfrm } from './transform.ts';
import { element, insertAt, nextShapeId, unitsIn } from './xml.ts';

const PH_NAMES: Record<string, string> = {
  title: 'Title', ctrTitle: 'Title', subTitle: 'Subtitle', body: 'Text Placeholder', obj: 'Content Placeholder',
  pic: 'Picture Placeholder', tbl: 'Table Placeholder', chart: 'Chart Placeholder', dgm: 'SmartArt Placeholder',
  media: 'Media Placeholder', clipArt: 'Clip Art Placeholder',
};
/** Placeholders a new slide does not get: PowerPoint adds these only through Header & Footer. */
const NOT_ON_NEW_SLIDES = new Set(['dt', 'ftr', 'sldNum', 'hdr', 'sldImg']);

const SECTIONS_EXT = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';

/** An empty slide placeholder for a layout placeholder: same type and index, no position of its own. */
function placeholderXml(layoutPh: ShapeView, id: number): string {
  const ph = layoutPh.placeholder!;
  const phEl = child(layoutPh.nvPr, 'p:ph');
  const attrs = ['type', 'orient', 'sz', 'idx'].map(k => [k, attr(phEl, k)] as const).filter(([, v]) => v !== undefined).map(([k, v]) => ` ${k}="${v}"`).join('');
  const name = `${PH_NAMES[ph.type] ?? 'Placeholder'} ${id - 1}`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${attrs}/></p:nvPr></p:nvSpPr>`
    + '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
}

function layoutPlaceholders(ed: Editor, layout: string): ShapeView[] {
  return ed.doc.slide(layout).shapes.filter(s => s.placeholder && !NOT_ON_NEW_SLIDES.has(s.placeholder.type));
}

function slideXml(ed: Editor, layout: string): string {
  const s = ed.strict;
  let id = 1;
  const phs = layoutPlaceholders(ed, layout).map(ph => placeholderXml(ph, ++id)).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<p:sld xmlns:a="${uriFor('a', s)}" xmlns:r="${uriFor('r', s)}" xmlns:p="${uriFor('p', s)}"><p:cSld><p:spTree>`
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    + `${phs}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

/** The p14:sldIdLst elements of the presentation's sections, in order. */
function sectionLists(pres: XmlElement): XmlElement[] {
  for (const ext of children(child(pres, 'p:extLst'), 'p:ext')) {
    if (attr(ext, 'uri') !== SECTIONS_EXT) continue;
    return children(child(ext, 'p14:sectionLst'), 'p14:section').map(s => child(s, 'p14:sldIdLst')).filter((x): x is XmlElement => !!x);
  }
  return [];
}

/** Take a slide id out of its section. */
function leaveSection(pres: XmlElement, id: number): void {
  for (const lst of sectionLists(pres)) {
    lst.children = lst.children.filter(c => !(c.kind === 'element' && is(c, 'p14:sldId') && int(attr(c, 'id'), 0) === id));
  }
}

/**
 * Put a slide id into the section that holds the slide before it in presentation order (the first section when
 * it is first), right after that slide, so sections stay contiguous as PowerPoint keeps them.
 */
function joinSection(ed: Editor, pres: XmlElement, id: number): void {
  const lists = sectionLists(pres);
  if (!lists.length) return;
  const order = ed.doc.presentation.slideIds.map(e => e.id);
  const before = order[order.indexOf(id) - 1];
  for (const lst of lists) {
    const ids = children(lst, 'p14:sldId');
    const at = before === undefined ? -1 : ids.findIndex(e => int(attr(e, 'id'), 0) === before);
    if (before !== undefined && at < 0) continue;
    const el = element(lst, `<p14:sldId id="${id}"/>`, ed.strict);
    insertAt(lst, at >= 0 ? lst.children.indexOf(ids[at]!) + 1 : ids[0] ? lst.children.indexOf(ids[0]) : 0, el);
    return;
  }
  const el = element(lists[0]!, `<p14:sldId id="${id}"/>`, ed.strict);
  insertAt(lists[0]!, 0, el);
}

/**
 * A new slide from a layout, as PowerPoint makes one: one empty placeholder per layout placeholder (not date,
 * footer or slide number), a slideLayout relationship, the content type, a relationship from the presentation and
 * a p:sldId after the slide at `index - 1` (at the end by default).
 */
export function addSlide(ed: Editor, layout: string, index?: number): { command: Command; part: string } {
  const pres = ed.presentationRoot;
  const main = ed.doc.mainPart;
  const part = freeName(ed.pkg, '/ppt/slides/slide#.xml');
  const command = ed.edit({
    label: 'New Slide', parts: [main, part], structural: true, scope: () => [pres],
    apply: () => {
      ed.pkg.addPart(part, CONTENT_TYPE.slide, encodeXmlText(slideXml(ed, layout)));
      ed.pkg.addRelationship(part, relType('slideLayout', ed.strict), layout);
      const rId = ed.pkg.addRelationship(main, relType('slide', ed.strict), part);
      const id = ed.doc.presentation.addSlideId(rId, index);
      joinSection(ed, pres, id);
    },
  });
  return { command, part };
}

/** Elements that link to a slide through relationship `ids` of `source` (hyperlinks to it) are removed. */
function unlinkIn(root: XmlElement, ids: Set<string>): void {
  for (const ref of relRefs(root)) {
    if (!ids.has(ref.value)) continue;
    const el = ref.el;
    if (is(el, 'a:hlinkClick') || is(el, 'a:hlinkMouseOver') || is(el, 'a:hlinkHover') || is(el, 'p:sld')) {
      const p = el.parent;
      if (p) p.children.splice(p.children.indexOf(el), 1);
    }
  }
}

/**
 * Delete a slide: its p:sldId, its place in sections and custom shows, links to it from other slides, its
 * relationship from the presentation, the slide part, and whatever only it used (notes, comments, pictures, charts).
 */
export function deleteSlide(ed: Editor, slide: string): Command {
  const pkg = ed.pkg;
  const main = ed.doc.mainPart;
  const pres = ed.presentationRoot;
  // other slides (and layouts) that link to this one; its own notes page goes with it as an orphan
  const linking = [...new Set(pkg.incomingRelationships(slide).map(r => r.source))]
    .filter(src => src !== slide && src !== main && pkg.getPart(src)?.contentType !== CONTENT_TYPE.notesSlide);
  return ed.edit({
    label: 'Delete Slide', parts: [main, ...linking], structural: true,
    scope: () => [pres, ...linking.map(src => ed.doc.root(src))],
    apply: () => {
      const entry = ed.doc.presentation.slideIds.find(e => ed.doc.target(main, e.rId) === slide);
      const presRels = pkg.relationships(main).filter(r => !r.external && resolveTarget(main, r.target) === normalizeName(slide)).map(r => r.id);
      if (entry) {
        ed.doc.presentation.removeSlideId(entry.id);
        leaveSection(pres, entry.id);
      }
      unlinkIn(pres, new Set(presRels));                  // custom shows
      for (const id of presRels) pkg.removeRelationship(main, id);
      for (const src of linking) {
        const ids = new Set(pkg.relationships(src).filter(r => !r.external && resolveTarget(normalizeName(src), r.target) === normalizeName(slide)).map(r => r.id));
        unlinkIn(ed.doc.root(src), ids);
        for (const id of ids) pkg.removeRelationship(src, id);
      }
      const targets = pkg.relationships(slide).filter(r => !r.external).map(r => resolveTarget(normalizeName(slide), r.target));
      pkg.removePart(slide, { force: true });
      collectOrphans(pkg, targets);
    },
  });
}

/** Duplicate a slide right after it: its own copies of charts, diagrams and notes; pictures and layout shared. */
export function duplicateSlide(ed: Editor, slide: string): { command: Command; part: string | undefined } {
  const pres = ed.presentationRoot;
  const main = ed.doc.mainPart;
  let part: string | undefined;
  const command = ed.edit({
    label: 'Duplicate Slide', parts: [main], structural: true, scope: () => [pres],
    apply: () => {
      const copy = copyPart(ed.pkg, ed.pkg, slide, new Map());
      part = copy;
      // the copied notes page must point back at the copy, not the original
      for (const r of ed.pkg.relationships(copy)) {
        if (relKind(r.type) !== 'notesSlide' || r.external) continue;
        const notes = resolveTarget(normalizeName(copy), r.target);
        for (const back of ed.pkg.relationships(notes)) if (relKind(back.type) === 'slide') retarget(ed, notes, back.id, copy);
      }
      const at = ed.doc.slideParts().indexOf(slide);
      const rId = ed.pkg.addRelationship(main, relType('slide', ed.strict), copy);
      const id = ed.doc.presentation.addSlideId(rId, at + 1);
      joinSection(ed, pres, id);
    },
  });
  return { command, get part() { return part; } };
}

/** Move a slide to a position in the presentation (and into the section there). */
export function moveSlide(ed: Editor, slide: string, to: number): Command {
  const pres = ed.presentationRoot;
  const main = ed.doc.mainPart;
  return ed.edit({
    label: 'Move Slide', parts: [main], scope: () => [pres],
    apply: () => {
      const entry = ed.doc.presentation.slideIds.find(e => ed.doc.target(main, e.rId) === slide);
      if (!entry) return;
      ed.doc.presentation.moveSlideId(entry.id, to);
      if (sectionLists(pres).length) {
        leaveSection(pres, entry.id);
        joinSection(ed, pres, entry.id);
      }
    },
  });
}

/** Hide or show a slide in the slide show (p:sld show="0"). */
export function setSlideHidden(ed: Editor, slide: string, hidden: boolean): Command {
  const root = ed.doc.root(slide);
  return ed.edit({
    label: hidden ? 'Hide Slide' : 'Show Slide', parts: [slide], scope: () => [root],
    apply: () => setAttr(root, 'show', hidden ? '0' : undefined),
  });
}

/**
 * Apply another layout to a slide: the slideLayout relationship is re-pointed; placeholders that match one of the
 * new layout's (by type and index) take its position; ones that match nothing keep where they were drawn; the new
 * layout's placeholders the slide lacks are added empty.
 */
export function applyLayout(ed: Editor, slide: string, layout: string): Command {
  const root = ed.doc.root(slide);
  return ed.edit({
    label: 'Layout', parts: [slide], scope: () => [root],
    apply: () => {
      const view = ed.doc.slide(slide);
      const newPhs = ed.doc.slide(layout).shapes.filter(s => s.placeholder);
      const matched = new Set<ShapeView>();
      const keep: { shape: ShapeView; xfrm: ReturnType<typeof effectiveXfrm> }[] = [];
      for (const s of view.shapes) {
        const ph = s.placeholder;
        if (!ph) continue;
        const m = matchLayoutPlaceholder(ph, newPhs);
        if (m) {
          matched.add(m);
          const spPr = s.spPr;
          if (spPr) removeChildren(spPr, 'a:xfrm');
        } else if (!s.xfrm) keep.push({ shape: s, xfrm: effectiveXfrm(ed, slide, s) });
      }
      for (const k of keep) k.shape.setXfrm(k.xfrm);   // before the relationship moves: the old layout's position
      const rel = ed.pkg.relationships(slide).find(r => relKind(r.type) === 'slideLayout');
      if (rel) retarget(ed, slide, rel.id, layout);
      else ed.pkg.addRelationship(slide, relType('slideLayout', ed.strict), layout);
      const tree = view.spTree!;
      let id = nextShapeId(root);
      for (const ph of layoutPlaceholders(ed, layout)) {
        if (matched.has(ph)) continue;
        const units = unitsIn(tree);
        const last = units[units.length - 1];
        insertAt(tree, last ? tree.children.indexOf(last) + 1 : tree.children.length, element(tree, placeholderXml(ph, id++), ed.strict));
      }
    },
  });
}

/** A slide's background: a solid colour, or back to the layout's (no p:bg of its own). */
export function setBackground(ed: Editor, slide: string, color: { hex: string } | null): Command {
  const root = ed.doc.root(slide);
  return ed.edit({
    label: 'Format Background', parts: [slide], scope: () => [root],
    apply: () => {
      const cSld = child(root, 'p:cSld');
      if (!cSld) return;
      removeChildren(cSld, 'p:bg');
      if (color) insertInOrder(cSld, element(cSld, `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${color.hex.replace(/^#/, '').toUpperCase()}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`, ed.strict));
    },
  });
}

