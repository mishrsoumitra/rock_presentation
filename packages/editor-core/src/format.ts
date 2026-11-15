// packages/editor-core/src/format.ts  -- shape fill and outline
import { attr, child, children, clone, ensureChild, insertInOrder, int, removeChildren, setAttr, ShapeView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
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
  /** What the line begins and ends with: none, triangle, stealth, diamond, oval or arrow (ST_LineEndType). */
  headEnd?: string | null;
  tailEnd?: string | null;
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
          for (const end of ['headEnd', 'tailEnd'] as const) {
            const type = line[end];
            if (type === undefined) continue;
            removeChildren(ln, `a:${end}`);
            if (type) insertInOrder(ln, element(ln, `<a:${end} type="${esc(type)}"/>`, ed.strict));
          }
          ln.selfClosing = ln.children.length === 0;
          if (ln.children.length === 0 && ln.attrs.length === 0) spPr.children.splice(spPr.children.indexOf(ln), 1);
          spPr.selfClosing = spPr.children.length === 0;
        }
      }
    },
  });
}

const COLORS = ['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr', 'a:scrgbClr', 'a:hslClr'];
const colorIn = (el: XmlElement | undefined): XmlElement | undefined => (el ? COLORS.map(n => child(el, n)).find(x => x) : undefined);

/** Whether fills apply to this kind of shape (pictures and graphic frames have none of their own). */
const fillable = (v: ShapeView) => v.kind !== 'graphicFrame' && v.kind !== 'pic' && v.kind !== 'contentPart';

/**
 * The colour a shape is filled with, as a colour the editor can write again: its own solid fill's, else the one its
 * theme style names (p:style/a:fillRef). Undefined for no fill, a gradient or a picture.
 */
export function fillColorOf(view: ShapeView): ColorSpec | undefined {
  const spPr = view.spPr;
  const own = spPr ? ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'].map(n => child(spPr, n)).find(x => x) : undefined;
  if (own && own.local !== 'solidFill') return undefined;
  const c = colorIn(own ?? child(child(view.el, 'p:style'), 'a:fillRef'));
  if (!c) return undefined;
  if (c.local === 'srgbClr') return { kind: 'srgb', hex: attr(c, 'val') ?? '000000' };
  if (c.local !== 'schemeClr') return undefined;
  const mod = child(c, 'a:lumMod'), off = child(c, 'a:lumOff');
  return { kind: 'scheme', val: attr(c, 'val') ?? 'accent1', ...(mod ? { lumMod: int(attr(mod, 'val'), 100000) } : {}), ...(off ? { lumOff: int(attr(off, 'val'), 0) } : {}) };
}

/** A two-colour gradient: from one colour to another along an angle (degrees clockwise from left-to-right). */
export interface GradientSpec { from: ColorSpec; to: ColorSpec; angle: number; }

/** Fill shapes with a linear gradient that turns with the shape. */
export function setGradientFill(ed: Editor, part: string, shapes: ShapeView[], g: GradientSpec): Command {
  const ang = Math.round((((g.angle % 360) + 360) % 360) * 60000);
  const xml = `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${colorXml(g.from)}</a:gs><a:gs pos="100000">${colorXml(g.to)}</a:gs></a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
  return ed.edit({
    label: 'Shape Fill', parts: [part], scope: () => shapes.map(s => unitOf(s.el)),
    apply: () => {
      for (const s of shapes) {
        for (const el of twins(s.el)) {
          const v = new ShapeView(el, s.owner);
          if (!fillable(v)) continue;
          const spPr = spPrOf(v);
          removeChildren(spPr, ...FILLS);
          insertInOrder(spPr, element(spPr, xml, ed.strict));
        }
      }
    },
  });
}

/**
 * How see-through a shape's fill is, from 0 (solid) to 100 (invisible): a:alpha on its fill colour, or on every
 * stop of its gradient. A shape filled by its theme style gets that colour as a fill of its own first; a shape
 * with no fill, or a picture fill, is left alone.
 */
export function setFillTransparency(ed: Editor, part: string, shapes: ShapeView[], percent: number): Command {
  const alpha = Math.round((100 - Math.max(0, Math.min(100, percent))) * 1000);
  const setAlpha = (color: XmlElement) => {
    removeChildren(color, 'a:alpha');
    color.selfClosing = color.children.length === 0;
    if (alpha < 100000) {
      const a = element(color, `<a:alpha val="${alpha}"/>`, ed.strict);
      a.parent = color;
      color.children.push(a);
      color.selfClosing = false;
    }
  };
  return ed.edit({
    label: 'Transparency', parts: [part], scope: () => shapes.map(s => unitOf(s.el)),
    apply: () => {
      for (const s of shapes) {
        for (const el of twins(s.el)) {
          const v = new ShapeView(el, s.owner);
          if (!fillable(v)) continue;
          const spPr = spPrOf(v);
          let solid = child(spPr, 'a:solidFill');
          const grad = child(spPr, 'a:gradFill');
          if (!solid && !grad && !FILLS.some(n => child(spPr, n))) {
            // filled by the theme style: that colour becomes the shape's own fill, so it can carry the alpha
            const styled = colorIn(child(child(el, 'p:style'), 'a:fillRef'));
            if (!styled) continue;
            solid = element(spPr, '<a:solidFill/>', ed.strict);
            solid.selfClosing = false;
            const copy = clone(styled, solid);
            solid.children.push(copy);
            insertInOrder(spPr, solid);
          }
          const colors = solid ? [colorIn(solid)] : grad ? children(child(grad, 'a:gsLst'), 'a:gs').map(gs => colorIn(gs)) : [];
          for (const c of colors) if (c) setAlpha(c);
        }
      }
    },
  });
}

/** The shadows the ribbon offers, as PowerPoint's Outer presets write them. */
export const SHADOWS: Record<string, string> = {
  'bottom-right': '<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw>',
  bottom: '<a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw>',
  center: '<a:outerShdw blurRad="63500" sx="102000" sy="102000" algn="ctr" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw>',
};

/**
 * Give shapes one of the preset shadows, or none (null: an empty a:effectLst, which also turns off a shadow the
 * theme style would give). Other effects the shape has (glow, reflection) are kept.
 */
export function setShadow(ed: Editor, part: string, shapes: ShapeView[], preset: keyof typeof SHADOWS | null): Command {
  if (preset !== null && !SHADOWS[preset]) throw new Error(`no such shadow: ${preset}`);
  return ed.edit({
    label: 'Shadow', parts: [part], scope: () => shapes.map(s => unitOf(s.el)),
    apply: () => {
      for (const s of shapes) {
        for (const el of twins(s.el)) {
          const v = new ShapeView(el, s.owner);
          if (v.kind === 'graphicFrame' || v.kind === 'contentPart') continue;
          const spPr = spPrOf(v);
          const fx = ensureChild(spPr, 'a:effectLst');
          removeChildren(fx, 'a:outerShdw', 'a:innerShdw', 'a:prstShdw');
          if (preset) insertInOrder(fx, element(fx, SHADOWS[preset]!, ed.strict));
          fx.selfClosing = fx.children.length === 0;
        }
      }
    },
  });
}

