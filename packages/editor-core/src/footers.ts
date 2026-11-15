// packages/editor-core/src/footers.ts  -- Header & Footer: the date, the footer text and the slide number on a slide
import type { XmlElement } from '@pptx/xml';
import { attr, child, children, runsOf, type ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import { compound, type Command } from './history.ts';
import { element, esc, insertAt, nextShapeId, unitsIn } from './xml.ts';

/** The three things Header & Footer puts on a slide, by placeholder type. */
export type FooterKind = 'dt' | 'ftr' | 'sldNum';
/**
 * What a slide shows at its foot. The date is either kept up to date (a field, with the text it showed last) or
 * fixed text; the footer is text; the number is the slide's own.
 */
export interface HeaderFooter {
  date: { auto: boolean; text: string } | null;
  footer: string | null;
  slideNumber: boolean;
}

const KINDS: FooterKind[] = ['dt', 'ftr', 'sldNum'];
const NAMES: Record<FooterKind, string> = { dt: 'Date Placeholder', ftr: 'Footer Placeholder', sldNum: 'Slide Number Placeholder' };
/** The ids PowerPoint's own fields carry; a field's id only has to be a GUID. */
const FIELD_ID: Record<'slidenum' | 'datetime1', string> = { slidenum: '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}', datetime1: '{5BCAD085-E8A6-4E2C-8A1B-8B4D6F0C3A11}' };
/** Where each sits on a slide whose layout and master have no place for it: along the bottom, as PowerPoint's masters have them. */
const PLACE: Record<FooterKind, { x: number; w: number; algn: string }> = {
  dt: { x: 0.06875, w: 0.225, algn: 'l' }, ftr: { x: 0.33125, w: 0.3375, algn: 'ctr' }, sldNum: { x: 0.70625, w: 0.225, algn: 'r' },
};

const walk = (shapes: ShapeView[], out: ShapeView[] = []): ShapeView[] => {
  for (const s of shapes) { out.push(s); if (s.kind === 'grpSp') walk(s.children, out); }
  return out;
};
const placeholder = (ed: Editor, part: string, kind: FooterKind): ShapeView | undefined =>
  walk(ed.doc.slide(part).shapes).find(s => s.placeholder?.type === kind);

/** The layout's (or else the master's) placeholder of a kind: where a slide's own takes its place and look from. */
function inherited(ed: Editor, part: string, kind: FooterKind): ShapeView | undefined {
  const layout = ed.doc.layoutOf(part);
  const master = layout ? ed.doc.masterOf(layout) : undefined;
  return (layout ? placeholder(ed, layout, kind) : undefined) ?? (master ? placeholder(ed, master, kind) : undefined);
}

const parasOf = (s: ShapeView): XmlElement[] => children(s.txBody, 'a:p');
const textOfShape = (s: ShapeView): string => parasOf(s).map(p => runsOf(p).map(r => r.text).join('')).join('\n');

/** What a slide shows at its foot now. */
export function headerFooter(ed: Editor, part: string): HeaderFooter {
  const dt = placeholder(ed, part, 'dt'), ftr = placeholder(ed, part, 'ftr');
  const field = dt ? parasOf(dt).some(p => runsOf(p).some(r => r.fieldType?.startsWith('datetime'))) : false;
  return {
    date: dt ? { auto: !!field, text: textOfShape(dt) } : null,
    footer: ftr ? textOfShape(ftr) : null,
    slideNumber: !!placeholder(ed, part, 'sldNum'),
  };
}

/** The text of one of the three, as a paragraph's content. */
function runsXml(kind: FooterKind, spec: HeaderFooter, size: string): string {
  const rPr = `<a:rPr lang="en-US"${size}/>`;
  if (kind === 'sldNum') return `<a:fld id="${FIELD_ID.slidenum}" type="slidenum">${rPr}<a:t>‹#›</a:t></a:fld>`;
  if (kind === 'dt') return spec.date!.auto ? `<a:fld id="${FIELD_ID.datetime1}" type="datetime1">${rPr}<a:t>${esc(spec.date!.text)}</a:t></a:fld>` : `<a:r>${rPr}<a:t>${esc(spec.date!.text)}</a:t></a:r>`;
  return `<a:r>${rPr}<a:t>${esc(spec.footer ?? '')}</a:t></a:r>`;
}

function shapeXml(ed: Editor, part: string, kind: FooterKind, spec: HeaderFooter, id: number): string {
  const from = inherited(ed, part, kind);
  const phEl = from ? child(from.nvPr, 'p:ph') : undefined;
  const attrs = phEl
    ? ['type', 'orient', 'sz', 'idx'].map(k => [k, attr(phEl, k)] as const).filter(([, v]) => v !== undefined).map(([k, v]) => ` ${k}="${v}"`).join('')
    : ` type="${kind}" sz="${kind === 'ftr' ? 'quarter' : 'half'}" idx="${10 + KINDS.indexOf(kind)}"`;
  // with nothing to inherit from, the placeholder carries its own place and a footer's small text
  const size = ed.doc.presentation.slideSize;
  const own = PLACE[kind], r = Math.round;
  const spPr = from ? '<p:spPr/>' : `<p:spPr><a:xfrm><a:off x="${r(size.cx * own.x)}" y="${r(size.cy * 0.927)}"/><a:ext cx="${r(size.cx * own.w)}" cy="${r(size.cy * 0.0532)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`;
  const bodyPr = from ? '<a:bodyPr/>' : '<a:bodyPr anchor="ctr"/>';
  const pPr = from ? '' : `<a:pPr algn="${own.algn}"/>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${NAMES[kind]} ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${attrs}/></p:nvPr></p:nvSpPr>`
    + `${spPr}<p:txBody>${bodyPr}<a:lstStyle/><a:p>${pPr}${runsXml(kind, spec, from ? '' : ' sz="1200"')}</a:p></p:txBody></p:sp>`;
}

const wanted = (kind: FooterKind, spec: HeaderFooter): boolean => (kind === 'dt' ? spec.date !== null : kind === 'ftr' ? spec.footer !== null : spec.slideNumber);

/** Does a slide already show exactly this at its foot? Then there is nothing to change on it. */
function same(a: HeaderFooter, b: HeaderFooter): boolean {
  return a.slideNumber === b.slideNumber && a.footer === b.footer && (a.date === null) === (b.date === null) && (!a.date || !b.date || (a.date.auto === b.date.auto && a.date.text === b.date.text));
}

/**
 * Set what one slide shows at its foot. Each of the three is a placeholder that takes its place and look from the
 * slide's layout (or master); one that is wanted and missing is added, one that is there and unwanted is removed,
 * and one that stays has its text replaced when it differs (keeping the look of its first run).
 */
export function setHeaderFooter(ed: Editor, part: string, spec: HeaderFooter): Command {
  const tree = ed.spTree(part);
  return ed.edit({
    label: 'Header & Footer', parts: [part], scope: () => [tree],
    apply: () => {
      const now = headerFooter(ed, part);
      if (same(now, spec)) return;
      let id = nextShapeId(ed.doc.root(part));
      for (const kind of KINDS) {
        const have = placeholder(ed, part, kind);
        const want = wanted(kind, spec);
        if (have && !want) { const p = have.el.parent!; p.children.splice(p.children.indexOf(have.el), 1); continue; }
        if (!want) continue;
        if (!have) {
          const units = unitsIn(tree);
          const last = units[units.length - 1];
          insertAt(tree, last ? tree.children.indexOf(last) + 1 : tree.children.length, element(tree, shapeXml(ed, part, kind, spec, id++), ed.strict));
          continue;
        }
        if (kind === 'sldNum') continue;                                  // a number is a number
        const changed = kind === 'dt' ? !now.date || now.date.auto !== spec.date!.auto || now.date.text !== spec.date!.text : now.footer !== spec.footer;
        if (!changed) continue;
        // new text in the first paragraph, looking like what was there
        const body = child(have.el, 'p:txBody');
        const paras = body ? body.children.filter((c): c is XmlElement => c.kind === 'element' && c.local === 'p') : [];
        const first = paras[0];
        if (!body || !first) continue;
        const oldRun = first.children.find((c): c is XmlElement => c.kind === 'element' && (c.local === 'r' || c.local === 'fld'));
        const oldProps = oldRun ? child(oldRun, 'a:rPr') : undefined;
        const fresh = element(first, runsXml(kind, spec, ''), ed.strict);
        if (oldProps) { const mine = child(fresh, 'a:rPr')!; const kept = oldProps; fresh.children[fresh.children.indexOf(mine)] = kept; kept.parent = fresh; }
        first.children = first.children.filter(c => !(c.kind === 'element' && ['r', 'fld', 'br'].includes(c.local)));
        const end = first.children.findIndex(c => c.kind === 'element' && c.local === 'endParaRPr');
        first.children.splice(end < 0 ? first.children.length : end, 0, fresh);
        first.selfClosing = false;
        for (const extra of paras.slice(1)) body.children.splice(body.children.indexOf(extra), 1);
      }
    },
  });
}

/**
 * Apply to All: the same on every slide, as one undo step. With `skipTitle`, slides on a title layout are left
 * showing nothing ("Don't show on title slide").
 */
export function setHeaderFooterAll(ed: Editor, spec: HeaderFooter, skipTitle = false): Command {
  const none: HeaderFooter = { date: null, footer: null, slideNumber: false };
  const isTitle = (part: string) => { const l = ed.doc.layoutOf(part); return !!l && attr(ed.doc.root(l), 'type') === 'title'; };
  return compound('Header & Footer', ed.doc.slideParts().map(part => setHeaderFooter(ed, part, skipTitle && isTitle(part) ? none : spec)));
}
