// packages/model/src/textbody.ts  -- paragraphs and runs of a text body (p:txBody, a:txBody, c:rich, dsp:txBody)
import type { XmlElement } from '@pptx/xml';
import { attr, child, children, is, textOf } from './dom.ts';

export interface RunView {
  el: XmlElement;                      // a:r, a:br or a:fld
  kind: 'r' | 'br' | 'fld';
  rPr?: XmlElement;
  text: string;                        // '' for a:br; a field's cached text for a:fld
  fieldType?: string;                  // slidenum, datetime1, ...
  fieldId?: string;
}

export const paragraphsOf = (txBody: XmlElement | undefined) => children(txBody, 'a:p');

export function runsOf(p: XmlElement): RunView[] {
  const out: RunView[] = [];
  for (const c of p.children) {
    if (c.kind !== 'element') continue;
    const rPr = child(c, 'a:rPr');
    if (is(c, 'a:r')) out.push({ el: c, kind: 'r', text: textOf(child(c, 'a:t')), ...(rPr ? { rPr } : {}) });
    else if (is(c, 'a:br')) out.push({ el: c, kind: 'br', text: '', ...(rPr ? { rPr } : {}) });
    else if (is(c, 'a:fld')) {
      const type = attr(c, 'type');
      const id = attr(c, 'id');
      out.push({ el: c, kind: 'fld', text: textOf(child(c, 'a:t')), ...(rPr ? { rPr } : {}), ...(type ? { fieldType: type } : {}), ...(id ? { fieldId: id } : {}) });
    }
  }
  return out;
}

/** Plain text of a paragraph; a line break (a:br) is U+000B, as PowerPoint's object model reports it. */
export const paragraphText = (p: XmlElement) => runsOf(p).map(r => (r.kind === 'br' ? '\u000b' : r.text)).join('');

/** Plain text of a text body, paragraphs separated by U+000D as in PowerPoint's TextRange.Text. */
export const bodyText = (txBody: XmlElement | undefined) => paragraphsOf(txBody).map(paragraphText).join('\r');
