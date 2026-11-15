// packages/editor-core/src/links.ts  -- hyperlinks on text and on shapes: to a web address, or to a slide of the deck
import { normalizeName, resolveTarget } from '@pptx/opc';
import { attr, attrNs, child, insertInOrder, relKind, relType, runsOf, type ShapeView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { dropUnreferenced } from './parts.ts';
import { paragraphs, runElementsIn, type TextOptions, type TextPos, type TextRange } from './text.ts';
import { element, esc, twins, unitOf } from './xml.ts';

/** Where a link goes: a web or mail address (kept outside the file, as PowerPoint does), or a slide of this deck. */
export type LinkTarget = { url: string; tooltip?: string } | { slide: string; tooltip?: string };

const SLIDE_JUMP = 'ppaction://hlinksldjump';

/** The a:hlinkClick for a target, with the relationship it needs added to `part`. */
function linkXml(ed: Editor, part: string, target: LinkTarget): string {
  const tip = target.tooltip ? ` tooltip="${esc(target.tooltip)}"` : '';
  if ('url' in target) {
    const id = ed.pkg.addRelationship(part, relType('hyperlink', ed.strict), target.url, { external: true });
    return `<a:hlinkClick r:id="${id}"${tip}/>`;
  }
  if (!ed.doc.slideParts().includes(target.slide)) throw new Error('that slide is not in this presentation');
  const id = ed.pkg.addRelationship(part, relType('slide', ed.strict), target.slide);
  return `<a:hlinkClick r:id="${id}" action="${SLIDE_JUMP}"${tip}/>`;
}

/** What an a:hlinkClick of `part` points at, when it is a link this editor edits (an address or a slide). */
function targetOf(ed: Editor, part: string, link: XmlElement | undefined): LinkTarget | undefined {
  const id = link ? attrNs(link, 'r', 'id') : undefined;
  if (!link || !id) return undefined;
  const rel = ed.pkg.relationships(part).find(r => r.id === id);
  if (!rel) return undefined;
  const tooltip = attr(link, 'tooltip');
  const tip = tooltip ? { tooltip } : {};
  if (rel.external) return { url: rel.target, ...tip };
  return relKind(rel.type) === 'slide' ? { slide: resolveTarget(normalizeName(part), rel.target), ...tip } : undefined;
}

const linkOfRun = (run: XmlElement) => child(child(run, 'a:rPr'), 'a:hlinkClick');

/**
 * Link the text of a range (null removes its links). The range's runs all point at one relationship; a link they
 * had before is replaced, and its relationship removed when nothing else uses it.
 */
export function setTextLink(ed: Editor, part: string, txBody: XmlElement, range: TextRange, target: LinkTarget | null, options: TextOptions = {}): Command {
  return ed.edit({
    label: options.label ?? (target ? 'Hyperlink' : 'Remove Hyperlink'), parts: [part], structural: true, scope: () => [options.scope ?? txBody],
    apply: () => {
      const runs = runElementsIn(txBody, range).filter(r => r.local !== 'br');
      const old = new Set<string>();
      for (const r of runs) {
        const rPr = child(r, 'a:rPr');
        const link = rPr ? child(rPr, 'a:hlinkClick') : undefined;
        if (!rPr || !link) continue;
        const id = attrNs(link, 'r', 'id');
        if (id) old.add(id);
        rPr.children.splice(rPr.children.indexOf(link), 1);
        rPr.selfClosing = rPr.children.length === 0;
      }
      if (target && runs.length) {
        const xml = linkXml(ed, part, target);
        for (const r of runs) {
          let rPr = child(r, 'a:rPr');
          if (!rPr) {
            rPr = element(r, '<a:rPr lang="en-US"/>', ed.strict);
            insertInOrder(r, rPr);
          }
          insertInOrder(rPr, element(rPr, xml, ed.strict));
        }
      }
      dropUnreferenced(ed.pkg, part, old);
      options.after?.();
    },
  });
}

/**
 * The link at a caret: where it goes and the whole stretch of text that carries it (runs next to each other that
 * point at the same place), so editing or removing it takes all of it. Undefined when the caret is not in a link.
 */
export function textLinkAt(ed: Editor, part: string, txBody: XmlElement, pos: TextPos): { range: TextRange; target: LinkTarget } | undefined {
  const p = paragraphs(txBody)[pos.p];
  if (!p) return undefined;
  const runs = runsOf(p);
  const spans: { start: number; end: number; id: string | undefined; link: XmlElement | undefined }[] = [];
  let at = 0;
  for (const r of runs) {
    const len = r.kind === 'br' ? 1 : r.text.length;
    const link = r.kind === 'br' ? undefined : linkOfRun(r.el);
    spans.push({ start: at, end: at + len, id: link ? attrNs(link, 'r', 'id') : undefined, link });
    at += len;
  }
  // the run the caret is in; at a boundary, the run before it, else the one after
  let i = spans.findIndex(s => pos.offset > s.start && pos.offset <= s.end && s.id);
  if (i < 0) i = spans.findIndex(s => pos.offset === s.start && s.id);
  const here = spans[i];
  if (!here?.id) return undefined;
  const target = targetOf(ed, part, here.link);
  if (!target) return undefined;
  let a = i, b = i;
  while (a > 0 && spans[a - 1]!.id === here.id) a--;
  while (b < spans.length - 1 && spans[b + 1]!.id === here.id) b++;
  return { range: { start: { p: pos.p, offset: spans[a]!.start }, end: { p: pos.p, offset: spans[b]!.end } }, target };
}

/** Link whole shapes: a click anywhere on one follows it in the slide show (null removes the link). */
export function setShapeLink(ed: Editor, part: string, shapes: ShapeView[], target: LinkTarget | null): Command {
  return ed.edit({
    label: target ? 'Hyperlink' : 'Remove Hyperlink', parts: [part], structural: true, scope: () => shapes.map(s => unitOf(s.el)),
    apply: () => {
      const old = new Set<string>();
      const xml = target ? linkXml(ed, part, target) : undefined;
      for (const s of shapes) {
        for (const el of twins(s.el)) {
          const cNvPr = ed.view(part, el).cNvPr;
          if (!cNvPr) continue;
          const link = child(cNvPr, 'a:hlinkClick');
          if (link) {
            const id = attrNs(link, 'r', 'id');
            if (id) old.add(id);
            cNvPr.children.splice(cNvPr.children.indexOf(link), 1);
          }
          if (xml) insertInOrder(cNvPr, element(cNvPr, xml, ed.strict));
          cNvPr.selfClosing = cNvPr.children.length === 0;
        }
      }
      dropUnreferenced(ed.pkg, part, old);
    },
  });
}

/** Where a shape's own link goes, if it has one this editor edits. */
export function shapeLink(ed: Editor, part: string, shape: ShapeView): LinkTarget | undefined {
  return targetOf(ed, part, shape.hlinkClick);
}
