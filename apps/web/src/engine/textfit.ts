// apps/web/src/engine/textfit.ts  -- refitting autofit text after an edit, as PowerPoint does while you type
import { createElement, insertInOrder, ShapeView, child, removeChildren } from '@pptx/model';
import type { Editor } from '@pptx/editor-core';
import { shapeTextLayout } from '@pptx/render';
import { resolveShapeIn, type ResolvedShape } from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';
import type { XmlElement } from '@pptx/xml';

/**
 * The font scale and line spacing reduction pairs PowerPoint steps through to shrink text on overflow, in order;
 * these are the values it writes in a:normAutofit.
 */
const STEPS: [number, number][] = [
  [100000, 0], [100000, 10000], [92500, 10000], [92500, 20000], [85000, 20000], [77500, 20000], [70000, 20000],
  [62500, 20000], [55000, 20000], [47500, 20000], [40000, 20000], [32500, 20000], [25000, 20000],
];

/**
 * Refit one shape's text in place (call it inside the text command, so it is part of the same undo step):
 * a:normAutofit gets the first step at which the text fits; a:spAutoFit makes the shape exactly as tall as its
 * text (and, unwrapped, as wide).
 *
 * Only those two kinds of body do any work: a plain text box returns before anything is resolved or laid out, and
 * the others resolve this shape alone, not the slide. Shrinking tries the step the text is at first and searches
 * from there, so a keystroke usually costs one layout rather than thirteen.
 */
export function refit(ed: Editor, part: string, shapeEl: XmlElement, fonts: FontRegistry): void {
  const view = new ShapeView(shapeEl, ed.doc.owner(part));
  const bodyPr = view.txBody ? child(view.txBody, 'a:bodyPr') : undefined;
  if (!bodyPr) return;
  // the body's own autofit decides; without one, only a placeholder can inherit one from its layout or master
  const own = child(bodyPr, 'a:normAutofit') ? 'norm' : child(bodyPr, 'a:spAutoFit') ? 'shape' : child(bodyPr, 'a:noAutofit') ? 'none' : undefined;
  if (own === 'none' || (own === undefined && !view.placeholder)) return;
  const s = resolveShapeIn(ed.doc, part, shapeEl);
  if (!s?.text) return;
  const auto = s.text.body.autofit;
  if (auto.kind === 'norm') {
    const fits = (step: [number, number]) => {
      const text = { ...s.text!, body: { ...s.text!.body, autofit: { kind: 'norm' as const, fontScale: step[0] / 100000, lnSpcReduction: step[1] / 100000 } } };
      const t = shapeTextLayout(s, fonts, text);
      return !t || t.layout.contentHeight <= t.area.h + 1;
    };
    const pick = STEPS[fitFrom(stepIndex(auto.fontScale, auto.lnSpcReduction), fits)]!;
    const norm = child(bodyPr, 'a:normAutofit');
    const want = { fontScale: pick[0] === 100000 ? undefined : String(pick[0]), lnSpcReduction: pick[1] ? String(pick[1]) : undefined };
    const has = { fontScale: norm?.attrs.find(a => a.local === 'fontScale')?.value, lnSpcReduction: norm?.attrs.find(a => a.local === 'lnSpcReduction')?.value };
    if (norm && has.fontScale === want.fontScale && has.lnSpcReduction === want.lnSpcReduction) return;
    removeChildren(bodyPr, 'a:normAutofit');
    const el = createElement(bodyPr, 'a:normAutofit');
    if (want.fontScale) el.attrs.push({ prefix: null, local: 'fontScale', value: want.fontScale, ns: '' });
    if (want.lnSpcReduction) el.attrs.push({ prefix: null, local: 'lnSpcReduction', value: want.lnSpcReduction, ns: '' });
    insertInOrder(bodyPr, el);
    bodyPr.selfClosing = false;
  } else if (auto.kind === 'shape') {
    resizeToText(s, view, fonts);
  }
}

/** The step the body is at now (what a:normAutofit says), or the first when its values are not a step. */
function stepIndex(fontScale: number, lnSpcReduction: number): number {
  const i = STEPS.findIndex(st => Math.abs(st[0] / 100000 - fontScale) < 1e-6 && Math.abs(st[1] / 100000 - lnSpcReduction) < 1e-6);
  return i < 0 ? 0 : i;
}

/**
 * The first step at which the text fits, found from `start`: when it fits there, earlier (larger) steps are tried
 * while they fit too (deleting text lets it grow back); when it does not, later steps are tried until one fits.
 */
function fitFrom(start: number, fits: (step: [number, number]) => boolean): number {
  let i = start;
  if (fits(STEPS[i]!)) {
    while (i > 0 && fits(STEPS[i - 1]!)) i--;
  } else {
    while (i < STEPS.length - 1 && !fits(STEPS[i]!)) i++;
  }
  return i;
}

function resizeToText(s: ResolvedShape, view: ShapeView, fonts: FontRegistry): void {
  const t = shapeTextLayout(s, fonts);
  const x = view.xfrm ?? s.xfrm;
  if (!t) return;
  const cy = Math.round(x.cy - t.area.h + t.layout.contentHeight);
  const widest = Math.max(0, ...t.layout.lines.map(l => l.width));
  const cx = s.text!.body.wrap ? x.cx : Math.round(x.cx - t.area.w + widest);
  if (Math.abs(cy - x.cy) < 2 && Math.abs(cx - x.cx) < 2) return;
  view.setXfrm({ ...x, cx, cy });
}
