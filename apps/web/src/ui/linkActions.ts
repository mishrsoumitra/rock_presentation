// apps/web/src/ui/linkActions.ts  -- hyperlinks from the ribbon, the menu and Ctrl/Cmd+K: what is linked now, and linking it
import { paragraphLength, paragraphs, setShapeLink, setTextLink, shapeLink, tableOf, textLinkAt, typeText, type LinkTarget, type TextRange } from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { refit } from '../engine/textfit.ts';
import { useEditor } from '../state/store.ts';

/** What a link would be put on: the text selected (or the link the caret is in), or the selected shapes. */
export interface LinkContext {
  on: 'text' | 'shape';
  /** The link there now, when there is one. */
  target?: LinkTarget;
  /** For text: there is text to link (a selection, or a link at the caret). Without, the address itself is typed in. */
  hasText: boolean;
}

const same = (a: { p: number; offset: number }, b: { p: number; offset: number }) => a.p === b.p && a.offset === b.offset;
const ordered = (a: { p: number; offset: number }, b: { p: number; offset: number }): TextRange => (a.p < b.p || (a.p === b.p && a.offset <= b.offset) ? { start: a, end: b } : { start: b, end: a });

/** The text being edited, its range and the link at its caret. */
function textContext() {
  const st = useEditor.getState();
  const s = st.session, t = st.text;
  if (!s || !t) return undefined;
  const body = new ShapeView(t.el, s.ed.doc.owner(t.part)).txBody;
  if (!body) return undefined;
  const caret = same(t.anchor, t.caret);
  const range = ordered(t.anchor, t.caret);
  // with a selection, the link its first character has; with a caret, the link it is in
  const at = caret ? t.caret : { p: range.start.p, offset: Math.min(range.start.offset + 1, paragraphLength(paragraphs(body)[range.start.p]!)) };
  return { s, t, body, caret, range, link: textLinkAt(s.ed, t.part, body, at) };
}

export function linkContext(): LinkContext | undefined {
  const st = useEditor.getState();
  if (!st.session || !st.slide) return undefined;
  const tc = textContext();
  if (tc) return { on: 'text', hasText: !tc.caret || !!tc.link, ...(tc.link ? { target: tc.link.target } : {}) };
  const shapes = st.selection.filter(el => !tableOf(el));
  if (!shapes.length) return undefined;
  const target = shapeLink(st.session.ed, st.slide, new ShapeView(shapes[0]!, st.session.ed.doc.owner(st.slide)));
  return { on: 'shape', hasText: false, ...(target ? { target } : {}) };
}

/**
 * Put a link where the context says (null removes it). On a text selection: on that text. With the caret in a
 * link: on all of that link. With a bare caret: the address is typed in and linked, as PowerPoint does.
 */
export function applyLink(target: LinkTarget | null): void {
  const st = useEditor.getState();
  const s = st.session;
  if (!s || !st.slide) return;
  const tc = textContext();
  if (tc) {
    const { t, body } = tc;
    const opts = { scope: t.el, after: () => refit(s.ed, t.part, t.el, s.fonts) };
    s.ed.log.seal();
    let range = tc.caret ? tc.link?.range : tc.range;
    if (!range) {
      if (!target) return;
      // nothing to link: the address (or the slide's number) becomes the linked text
      const text = 'url' in target ? target.url : `Slide ${s.ed.doc.slideParts().indexOf(target.slide) + 1}`;
      const typed = typeText(s.ed, t.part, body, t.caret, text, opts);
      st.run(typed.command);
      range = { start: t.caret, end: typed.caret };
      useEditor.setState({ text: { ...t, anchor: typed.caret, caret: typed.caret } });
    }
    st.run(setTextLink(s.ed, t.part, body, range, target, opts));
    s.ed.log.seal();
    return;
  }
  const shapes = st.selection.filter(el => !tableOf(el)).map(el => new ShapeView(el, s.ed.doc.owner(st.slide!)));
  if (shapes.length) st.run(setShapeLink(s.ed, st.slide, shapes, target));
}

/** The slides a link can go to: each one's part, number and title (for the dialog's list). */
export function slideChoices(): { part: string; label: string }[] {
  const ed = useEditor.getState().session?.ed;
  if (!ed) return [];
  return ed.doc.slideParts().map((part, i) => {
    const title = ed.doc.slide(part).shapes.find(sh => sh.placeholder?.type === 'title' || sh.placeholder?.type === 'ctrTitle');
    const text = title?.txBody ? paragraphs(title.txBody).map(p => p.children.map(c => (c.kind === 'element' && c.local === 'r' ? c.children.map(x => (x.kind === 'element' && x.local === 't' ? x.children.map(y => ('value' in y ? y.value : '')).join('') : '')).join('') : '')).join('')).join(' ').trim() : '';
    return { part, label: text ? `Slide ${i + 1}: ${text}` : `Slide ${i + 1}` };
  });
}
