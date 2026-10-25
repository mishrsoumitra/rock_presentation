// packages/charts/src/xml.ts  -- reading chart XML: CT_Boolean and friends, and small synthetic DrawingML elements
import type { XmlElement, XmlNode } from '@pptx/xml';
import { attr, child, children, effectiveChildren, is, UNDERSTOOD } from '@pptx/model';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** Element children with each mc:AlternateContent replaced by the branch we read (c14:style comes as c:style). */
export const kids = (el: XmlElement | undefined): XmlElement[] => (el ? effectiveChildren(el, UNDERSTOOD) : []);

const hasAlternates = (el: XmlElement) => el.children.some(c => c.kind === 'element' && is(c, 'mc:AlternateContent'));

/** The first effective child named `name` ('c:legend'), looking through mc:AlternateContent. */
export function kid(el: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!el) return undefined;
  return child(el, name) ?? (hasAlternates(el) ? kids(el).find(c => is(c, name)) : undefined);
}

/** All effective children named `name`. */
export const kidsNamed = (el: XmlElement | undefined, name: string): XmlElement[] =>
  el && hasAlternates(el) ? kids(el).filter(c => is(c, name)) : children(el, name);

/** The val attribute of a child element (CT_Boolean, CT_Double, CT_UnsignedInt...). */
export const valOf = (el: XmlElement | undefined, name: string): string | undefined => attr(kid(el, name), 'val');

/** A CT_Boolean child: absent gives `dflt`, present without val means true (the schema default). */
export function flag(el: XmlElement | undefined, name: string, dflt: boolean): boolean {
  const c = kid(el, name);
  if (!c) return dflt;
  const v = attr(c, 'val');
  return v === undefined || v === '1' || v === 'true';
}

/** A numeric child value; Strict and some writers add a '%' to percentages ("150%" for 150). */
export function numOf(el: XmlElement | undefined, name: string): number | undefined {
  const v = valOf(el, name);
  if (v === undefined) return undefined;
  const n = Number.parseFloat(v.endsWith('%') ? v.slice(0, -1) : v);
  return Number.isFinite(n) ? n : undefined;
}

/** A finite number from element text (c:v), or undefined. */
export function numText(s: string): number | undefined {
  const t = s.trim();
  if (!t) return undefined;
  const n = Number(t);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * A DrawingML element built in memory, for text the chart makes itself (tick labels, legend entries). Elements of
 * the part can be passed as content to be read through it; their own parent links are left alone.
 */
export function aElement(local: string, content: (XmlNode | string)[] = []): XmlElement {
  return {
    kind: 'element', prefix: 'a', local, ns: A_NS, selfClosing: content.length === 0, parent: null, attrs: [],
    children: content.map(c => (typeof c === 'string' ? { kind: 'text' as const, value: c } : c)),
  };
}
