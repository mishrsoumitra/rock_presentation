// packages/editor-core/src/format.ts  -- shape fill and outline
import { ensureChild, insertInOrder, removeChildren, setAttr, ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import type { ColorSpec } from './text.ts';
import { element, esc, twins, unitOf } from './xml.ts';

const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

export function colorXml(c: ColorSpec): string {
  if (c.kind === 'srgb') return `<a:srgbClr val="${esc(c.hex.replace(/^#/, '').toUpperCase())}"/>`;
  const mods = `${c.lumMod !== undefined ? `<a:lumMod val="${c.lumMod}"/>` : ''}${c.lumOff !== undefined ? `<a:lumOff val="${c.lumOff}"/>` : ''}`;
  return mods ? `<a:schemeClr val="${esc(c.val)}">${mods}</a:schemeClr>` : `<a:schemeClr val="${esc(c.val)}"/>`;
}

/** The shape's properties element, created in its place when the shape has none (a placeholder's empty p:spPr). */
function spPrOf(view: ShapeView) {
  const name = view.kind === 'grpSp' ? 'grpSpPr' : 'spPr';
  return ensureChild(view.el, `${view.el.prefix === 'dsp' ? 'dsp' : 'p'}:${name}`);
}

/**
 * Fill shapes with a colour, with nothing ('none': a:noFill), or with what the theme style gives (null: the fill
 * element is removed). Pictures and graphic frames have no fill of their own and are left alone.
 */
export function setFill(ed: Editor, part: string, shapes: ShapeView[], fill: ColorSpec | 'none' | null): Command {
  return ed.edit({
    label: 'Shape Fill', parts: [part], scope: () => shapes.map(s => unitOf(s.el)),
    apply: () => {
      for (const s of shapes) {
        // each branch of an alternate content block is checked on its own: a picture's twin may be a graphic frame
        for (const el of twins(s.el)) {
          const v = new ShapeView(el, s.owner);
          if (v.kind === 'graphicFrame' || v.kind === 'pic' || v.kind === 'contentPart') continue;
          const spPr = spPrOf(v);
          removeChildren(spPr, ...FILLS);
          if (fill === 'none') insertInOrder(spPr, element(spPr, '<a:noFill/>', ed.strict));
          else if (fill) insertInOrder(spPr, element(spPr, `<a:solidFill>${colorXml(fill)}</a:solidFill>`, ed.strict));
          spPr.selfClosing = spPr.children.length === 0;
        }
      }
    },
  });
}

export interface LineSpec {
  color?: ColorSpec | 'none' | null;
  /** Width in EMU (12700 per point). */
  width?: number | null;
  /** a:prstDash value: solid, dash, sysDot, lgDashDot... */
  dash?: string | null;
}

/** Outline colour, weight and dash. Each property set separately; null removes the shape's own value. */
export function setLine(ed: Editor, part: string, shapes: ShapeView[], line: LineSpec): Command {
  return ed.edit({
    label: 'Shape Outline', parts: [part], scope: () => shapes.map(s => unitOf(s.el)),
    apply: () => {
      for (const s of shapes) {
        for (const el of twins(s.el)) {
          const v = new ShapeView(el, s.owner);
          if (v.kind === 'graphicFrame' || v.kind === 'grpSp' || v.kind === 'contentPart') continue;
          const spPr = spPrOf(v);
          const ln = ensureChild(spPr, 'a:ln');
          if (line.width !== undefined) setAttr(ln, 'w', line.width === null ? undefined : Math.round(line.width));
          if (line.color !== undefined) {
            removeChildren(ln, 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill');
            if (line.color === 'none') insertInOrder(ln, element(ln, '<a:noFill/>', ed.strict));
            else if (line.color) insertInOrder(ln, element(ln, `<a:solidFill>${colorXml(line.color)}</a:solidFill>`, ed.strict));
          }
          if (line.dash !== undefined) {
            removeChildren(ln, 'a:prstDash', 'a:custDash');
            if (line.dash) insertInOrder(ln, element(ln, `<a:prstDash val="${esc(line.dash)}"/>`, ed.strict));
          }
          ln.selfClosing = ln.children.length === 0;
          if (ln.children.length === 0 && ln.attrs.length === 0) spPr.children.splice(spPr.children.indexOf(ln), 1);
          spPr.selfClosing = spPr.children.length === 0;
        }
      }
    },
  });
}
