// apps/web/src/engine/textfit.ts  -- refitting autofit text after an edit, as PowerPoint does while you type
import { createElement, insertInOrder, ShapeView, child, removeChildren } from '@pptx/model';
import type { Editor } from '@pptx/editor-core';
import { shapeTextLayout } from '@pptx/render';
import { resolveSlide, type ResolvedShape } from '@pptx/resolve';
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

function findShape(shapes: ResolvedShape[], el: XmlElement): ResolvedShape | undefined {
  for (const s of shapes) {
    if (s.view.el === el && s.layer === 'slide') return s;
    const inner = findShape(s.children, el);
    if (inner) return inner;
  }
  return undefined;
}

/**
 * Refit one shape's text in place (call it inside the text command, so it is part of the same undo step):
 * a:normAutofit gets the first step at which the text fits; a:spAutoFit makes the shape exactly as tall as its
 * text (and, unwrapped, as wide).
 */
export function refit(ed: Editor, part: string, shapeEl: XmlElement, fonts: FontRegistry): void {
  const s = findShape(resolveSlide(ed.doc, part).shapes, shapeEl);
  if (!s?.text) return;
  const auto = s.text.body.autofit;
  const view = new ShapeView(shapeEl, ed.doc.owner(part));
  const bodyPr = child(view.txBody, 'a:bodyPr');
  if (!bodyPr) return;
  if (auto.kind === 'norm') {
    let pick = STEPS[STEPS.length - 1]!;
    for (const step of STEPS) {
      const text = { ...s.text, body: { ...s.text.body, autofit: { kind: 'norm' as const, fontScale: step[0] / 100000, lnSpcReduction: step[1] / 100000 } } };
      const t = shapeTextLayout(s, fonts, text);
      if (!t || t.layout.contentHeight <= t.area.h + 1) {
        pick = step;
        break;
      }
    }
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
    const t = shapeTextLayout(s, fonts);
    const x = view.xfrm ?? s.xfrm;
    if (!t) return;
    const cy = Math.round(x.cy - t.area.h + t.layout.contentHeight);
    const widest = Math.max(0, ...t.layout.lines.map(l => l.width));
    const cx = s.text.body.wrap ? x.cx : Math.round(x.cx - t.area.w + widest);
    if (Math.abs(cy - x.cy) < 2 && Math.abs(cx - x.cx) < 2) return;
    view.setXfrm({ ...x, cx, cy });
  }
}
