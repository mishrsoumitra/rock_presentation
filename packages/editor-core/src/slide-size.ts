// packages/editor-core/src/slide-size.ts  -- changing the slide size and scaling content to it
import { child, children, is, ShapeView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { shapesOfUnit, unitsIn } from './xml.ts';

/** p:sldSz type values PowerPoint writes for its standard sizes; a size it does not name is "custom". */
const TYPES: [number, number, string | undefined][] = [
  [12192000, 6858000, undefined], [9144000, 6858000, 'screen4x3'], [9144000, 5143500, 'screen16x9'], [9144000, 5715000, 'screen16x10'],
  [9906000, 6858000, 'A4'],
];

export type ScaleMode = 'maximize' | 'ensureFit' | 'none';

/**
 * Change the slide size. With 'maximize' content scales by the larger of the two ratios (it fills the slide and may
 * run off it), with 'ensureFit' by the smaller (it all stays on the slide), centred either way, as PowerPoint's two
 * choices do; 'none' leaves content alone. Scaling covers every master, layout and slide: top-level transforms,
 * table grids and font sizes. Group members follow their group.
 */
export function setSlideSize(ed: Editor, cx: number, cy: number, mode: ScaleMode = 'ensureFit'): Command {
  const old = ed.doc.presentation.slideSize;
  const sx = cx / old.cx, sy = cy / old.cy;
  const s = mode === 'none' ? 1 : mode === 'maximize' ? Math.max(sx, sy) : Math.min(sx, sy);
  const dx = mode === 'none' ? 0 : (cx - old.cx * s) / 2, dy = mode === 'none' ? 0 : (cy - old.cy * s) / 2;
  const parts = [...ed.doc.masterParts().flatMap(m => [m, ...ed.doc.layoutsOf(m)]), ...ed.doc.slideParts()];
  const main = ed.doc.mainPart;
  return ed.edit({
    label: 'Slide Size', parts: [main, ...parts],
    scope: () => [ed.presentationRoot, ...parts.map(p => ed.doc.root(p))],
    apply: () => {
      const known = TYPES.find(([w, h]) => w === Math.round(cx) && h === Math.round(cy));
      const type = known ? known[2] : 'custom';
      ed.doc.presentation.setSlideSize(cx, cy, type);
      if (s === 1 && !dx && !dy) return;
      for (const p of parts) {
        const tree = ed.doc.slide(p).spTree;
        if (!tree) continue;
        for (const u of unitsIn(tree)) {
          for (const el of shapesOfUnit(u)) {
            const v = new ShapeView(el, ed.doc.owner(p));
            const x = v.xfrm;
            if (x) v.setXfrm({ ...x, x: x.x * s + dx, y: x.y * s + dy, cx: x.cx * s, cy: x.cy * s });
          }
        }
        if (s !== 1) scaleInside(ed.doc.root(p), s);
      }
      if (s !== 1) {
        const dts = ed.doc.presentation.defaultTextStyle;
        if (dts) scaleInside(dts, s);
      }
    },
  });
}

/** Font sizes and table grids under an element, scaled. */
function scaleInside(root: XmlElement, s: number): void {
  const walk = (e: XmlElement) => {
    if (is(e, 'a:rPr') || is(e, 'a:endParaRPr') || is(e, 'a:defRPr')) {
      const a = e.attrs.find(x => x.local === 'sz' && x.ns === '');
      if (a) a.value = String(Math.max(100, Math.round(Number(a.value) * s)));
    }
    if (is(e, 'a:tbl')) {
      for (const g of children(child(e, 'a:tblGrid'), 'a:gridCol')) {
        const w = g.attrs.find(x => x.local === 'w');
        if (w) w.value = String(Math.round(Number(w.value) * s));
      }
      for (const tr of children(e, 'a:tr')) {
        const h = tr.attrs.find(x => x.local === 'h');
        if (h) h.value = String(Math.round(Number(h.value) * s));
      }
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(root);
}
