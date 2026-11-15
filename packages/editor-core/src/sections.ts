// packages/editor-core/src/sections.ts  -- sections: named runs of slides in the slide panel
import { attr, child, children, insertInOrder, int, setAttr, uriFor } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { element, esc, insertAt } from './xml.ts';

const SECTIONS_EXT = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';

/** A section as the slide panel shows it: its name, and the slides in it (part names, in order). */
export interface Section { index: number; name: string; id: string; slides: string[]; }

/** The p14:sectionLst of a presentation, when it has sections. */
function sectionList(pres: XmlElement): XmlElement | undefined {
  for (const ext of children(child(pres, 'p:extLst'), 'p:ext')) if (attr(ext, 'uri') === SECTIONS_EXT) return child(ext, 'p14:sectionLst');
  return undefined;
}
const sectionsOf = (lst: XmlElement | undefined) => children(lst, 'p14:section');
const idsOf = (section: XmlElement) => children(child(section, 'p14:sldIdLst'), 'p14:sldId').map(e => int(attr(e, 'id'), 0));

/**
 * The deck's sections with their slides, in presentation order. A deck without sections has none. Slides that no
 * section lists (a file another program left that way) are shown with the section before them.
 */
export function sections(ed: Editor): Section[] {
  const lst = sectionList(ed.presentationRoot);
  const els = sectionsOf(lst);
  if (!els.length) return [];
  const partOf = new Map(ed.doc.presentation.slideIds.map(e => [e.id, ed.doc.target(ed.doc.mainPart, e.rId)] as const));
  const owner = new Map<number, number>();
  els.forEach((s, i) => { for (const id of idsOf(s)) owner.set(id, i); });
  const out: Section[] = els.map((s, index) => ({ index, name: attr(s, 'name') ?? '', id: attr(s, 'id') ?? '', slides: [] }));
  let last = 0;
  for (const e of ed.doc.presentation.slideIds) {
    const at = owner.get(e.id) ?? last;
    last = at;
    const part = partOf.get(e.id);
    if (part) out[at]!.slides.push(part);
  }
  return out;
}

/** A new section id: a GUID in braces, as PowerPoint writes them. */
function guid(): string {
  const hex = (n: number) => Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16)).join('').toUpperCase();
  return `{${hex(8)}-${hex(4)}-4${hex(3)}-${'89AB'[Math.floor(Math.random() * 4)]}${hex(3)}-${hex(12)}}`;
}

const sectionXml = (name: string, ids: number[]) =>
  `<p14:section name="${esc(name)}" id="${guid()}"><p14:sldIdLst>${ids.map(id => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst></p14:section>`;

/** The section list, made (with the extension that holds it) when the deck has none yet. */
function ensureList(ed: Editor): XmlElement {
  const pres = ed.presentationRoot;
  const have = sectionList(pres);
  if (have) return have;
  let extLst = child(pres, 'p:extLst');
  if (!extLst) {
    extLst = element(pres, '<p:extLst/>', ed.strict);
    insertInOrder(pres, extLst);
  }
  const ext = element(extLst, `<p:ext uri="${SECTIONS_EXT}"><p14:sectionLst xmlns:p14="${uriFor('p14', ed.strict)}"/></p:ext>`, ed.strict);
  insertAt(extLst, 0, ext);
  extLst.selfClosing = false;
  return child(ext, 'p14:sectionLst')!;
}

/** Write a section's slide ids. */
function setIds(ed: Editor, section: XmlElement, ids: number[]): void {
  let lst = child(section, 'p14:sldIdLst');
  if (!lst) {
    lst = element(section, '<p14:sldIdLst/>', ed.strict);
    insertAt(section, 0, lst);
  }
  lst.children = ids.map(id => { const e = element(lst!, `<p14:sldId id="${id}"/>`, ed.strict); e.parent = lst!; return e; });
  lst.selfClosing = ids.length === 0;
  section.selfClosing = false;
}

/**
 * Start a section at a slide: that slide and those after it, up to the next section, are the new section. In a
 * deck with no sections yet, the slides before it become "Default Section", as PowerPoint names it.
 */
export function addSection(ed: Editor, slide: string, name = 'Untitled Section'): Command {
  const pres = ed.presentationRoot;
  return ed.edit({
    label: 'Add Section', parts: [ed.doc.mainPart], scope: () => [pres],
    apply: () => {
      const order = ed.doc.presentation.slideIds;
      const at = order.findIndex(e => ed.doc.target(ed.doc.mainPart, e.rId) === slide);
      if (at < 0) throw new Error('that slide is not in this presentation');
      const lst = ensureList(ed);
      const els = sectionsOf(lst);
      const id = order[at]!.id;
      if (!els.length) {
        const before = order.slice(0, at).map(e => e.id), from = order.slice(at).map(e => e.id);
        if (before.length) insertAt(lst, lst.children.length, element(lst, sectionXml('Default Section', before), ed.strict));
        insertAt(lst, lst.children.length, element(lst, sectionXml(name, from), ed.strict));
        return;
      }
      // split the section that holds the slide, at the slide (a section that starts there already gets one before it)
      const holder = els.find(s => idsOf(s).includes(id)) ?? els[els.length - 1]!;
      const ids = idsOf(holder);
      const cut = Math.max(0, ids.indexOf(id));
      setIds(ed, holder, ids.slice(0, cut));
      insertAt(lst, lst.children.indexOf(holder) + 1, element(lst, sectionXml(name, ids.slice(cut)), ed.strict));
      // a section left with no slides at the very start would only be an empty heading: it goes
      if (cut === 0 && els.indexOf(holder) === 0 && !(attr(holder, 'name') ?? '').trim()) lst.children.splice(lst.children.indexOf(holder), 1);
    },
  });
}

/** Rename a section (by its place in the list). */
export function renameSection(ed: Editor, index: number, name: string): Command {
  const pres = ed.presentationRoot;
  return ed.edit({
    label: 'Rename Section', parts: [ed.doc.mainPart], scope: () => [pres],
    apply: () => {
      const s = sectionsOf(sectionList(pres))[index];
      if (!s) throw new Error('no such section');
      setAttr(s, 'name', name);
    },
  });
}

/**
 * Remove a section, keeping its slides: they join the section before it (the one after, for the first). Removing
 * the only section leaves a deck with no sections.
 */
export function removeSection(ed: Editor, index: number): Command {
  const pres = ed.presentationRoot;
  return ed.edit({
    label: 'Remove Section', parts: [ed.doc.mainPart], scope: () => [pres],
    apply: () => {
      const lst = sectionList(pres);
      const els = sectionsOf(lst);
      const s = els[index];
      if (!lst || !s) throw new Error('no such section');
      const into = els[index - 1] ?? els[index + 1];
      if (into) setIds(ed, into, index > 0 ? [...idsOf(into), ...idsOf(s)] : [...idsOf(s), ...idsOf(into)]);
      lst.children.splice(lst.children.indexOf(s), 1);
      if (!sectionsOf(lst).length) removeAllSectionsIn(pres);
    },
  });
}

function removeAllSectionsIn(pres: XmlElement): void {
  const extLst = child(pres, 'p:extLst');
  if (!extLst) return;
  extLst.children = extLst.children.filter(c => !(c.kind === 'element' && attr(c, 'uri') === SECTIONS_EXT));
  if (!children(extLst, 'p:ext').length) pres.children.splice(pres.children.indexOf(extLst), 1);
}

/** Remove every section; the slides stay as they are. */
export function removeAllSections(ed: Editor): Command {
  const pres = ed.presentationRoot;
  return ed.edit({ label: 'Remove All Sections', parts: [ed.doc.mainPart], scope: () => [pres], apply: () => removeAllSectionsIn(pres) });
}
