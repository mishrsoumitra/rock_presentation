// XML snippets for the model tests, written the way PowerPoint writes them.
import { parseXml, type XmlElement } from '@pptx/xml';

export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NSDECL = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;

/** A p:sld whose spTree holds `shapes`. */
export const slideXml = (shapes: string, extra = '') =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${NSDECL}><p:cSld><p:spTree>` +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  `${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${extra}</p:sld>`;

export const parse = (xml: string): XmlElement => parseXml(xml).root;

/** The first descendant element with this local name. */
export function find(el: XmlElement, local: string): XmlElement {
  const stack: XmlElement[] = [el];
  while (stack.length) {
    const e = stack.shift()!;
    if (e.local === local) return e;
    for (const c of e.children) if (c.kind === 'element') stack.push(c);
  }
  throw new Error(`no <${local}>`);
}
