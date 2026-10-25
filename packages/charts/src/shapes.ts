// packages/charts/src/shapes.ts  -- shapes drawn over a chart (c:userShapes, ECMA-376 Part 1, 21.3)
//
// A chart's drawing part anchors shapes to fractions of the chart area (cdr:relSizeAnchor) or to a fraction and a
// size in EMU (cdr:absSizeAnchor). Rectangles, ellipses, rounded rectangles and lines are drawn as such; other
// preset geometries are drawn as their box, since the chart package does not evaluate shape geometry.
import type { XmlElement } from '@pptx/xml';
import { attr, child, children, effectsIn, fillIn, is, parseLine, parseStyle, textOf } from '@pptx/model';
import { resolveShapeStyle, resolveTextBody } from '@pptx/resolve';
import type { Primitive } from '@pptx/render';
import { layoutText } from '@pptx/text';
import { ellipse, line, pathPrim, rect, roundRect, type Cmd, type Rect } from './draw.ts';
import type { Env } from './style.ts';
import { placeAt } from './text.ts';

const num = (el: XmlElement | undefined): number => {
  const v = Number(textOf(el).trim());
  return Number.isFinite(v) ? v : 0;
};

/** Where an anchor puts its shape, in the chart's box; undefined for anything but the two anchors. */
function anchorRect(a: XmlElement, box: { w: number; h: number }): Rect | undefined {
  const from = child(a, 'cdr:from');
  const fx = num(child(from, 'cdr:x')) * box.w;
  const fy = num(child(from, 'cdr:y')) * box.h;
  let r: Rect;
  if (is(a, 'cdr:relSizeAnchor')) {
    const to = child(a, 'cdr:to');
    r = { x: fx, y: fy, w: num(child(to, 'cdr:x')) * box.w - fx, h: num(child(to, 'cdr:y')) * box.h - fy };
  } else if (is(a, 'cdr:absSizeAnchor')) {
    const ext = child(a, 'cdr:ext');
    r = { x: fx, y: fy, w: Number(attr(ext, 'cx') ?? 0), h: Number(attr(ext, 'cy') ?? 0) };
  } else return undefined;
  if (![r.x, r.y, r.w, r.h].every(Number.isFinite)) return undefined;
  // a "to" corner above or left of "from" still means the box between them
  return { x: Math.min(r.x, r.x + r.w), y: Math.min(r.y, r.y + r.h), w: Math.abs(r.w), h: Math.abs(r.h) };
}

/** Turn path points about the centre of `r` (xfrm rot, in degrees). */
function rotated(cmds: Cmd[], r: Rect, deg: number): Cmd[] {
  if (!deg) return cmds;
  const a = (deg * Math.PI) / 180;
  const cx = r.x + r.w / 2;
  const cy = r.y + r.h / 2;
  const t = (x: number, y: number) => ({
    x: cx + (x - cx) * Math.cos(a) - (y - cy) * Math.sin(a),
    y: cy + (x - cx) * Math.sin(a) + (y - cy) * Math.cos(a),
  });
  return cmds.map(c => {
    if (c.op === 'Z') return c;
    const p = t(c.x, c.y);
    if (c.op === 'C') {
      const p1 = t(c.x1, c.y1);
      const p2 = t(c.x2, c.y2);
      return { op: 'C', x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y, x: p.x, y: p.y };
    }
    if (c.op === 'Q') {
      const p1 = t(c.x1, c.y1);
      return { op: 'Q', x1: p1.x, y1: p1.y, x: p.x, y: p.y };
    }
    return { op: c.op, x: p.x, y: p.y };
  });
}

function drawShape(env: Env, sp: XmlElement, r: Rect, paths: Primitive[], texts: Primitive[]): void {
  const spPr = child(sp, 'cdr:spPr');
  const xfrm = child(spPr, 'a:xfrm');
  const flipH = attr(xfrm, 'flipH') === '1';
  const flipV = attr(xfrm, 'flipV') === '1';
  const rot = Number(attr(xfrm, 'rot') ?? 0) / 60000 || 0;
  const prst = attr(child(spPr, 'a:prstGeom'), 'prst') ?? 'rect';
  const isLine = prst === 'line' || prst.startsWith('straightConnector') || is(sp, 'cdr:cxnSp');
  let cmds: Cmd[];
  if (isLine) {
    const x0 = flipH ? r.x + r.w : r.x;
    const x1 = flipH ? r.x : r.x + r.w;
    const y0 = flipV ? r.y + r.h : r.y;
    const y1 = flipV ? r.y : r.y + r.h;
    cmds = line({ x: x0, y: y0 }, { x: x1, y: y1 });
  } else if (prst === 'ellipse') cmds = ellipse(r.x + r.w / 2, r.y + r.h / 2, r.w / 2, r.h / 2);
  else if (prst === 'roundRect') cmds = roundRect(r, Math.min(r.w, r.h) / 6);
  else cmds = rect(r);
  const fill = fillIn(spPr);
  const ln = parseLine(child(spPr, 'a:ln'));
  const effects = effectsIn(spPr);
  const own = { ...(fill ? { fill } : {}), ...(ln ? { line: ln } : {}), ...(effects ? { effects } : {}) };
  const refs = parseStyle(child(sp, 'cdr:style'));
  const style = resolveShapeStyle(own, refs, env.theme, env.scope, env.scope);
  const prim = pathPrim(rotated(cmds, r, rot), isLine ? { kind: 'none' } : style.fill, style.line, style.effects);
  if (prim) paths.push(prim);

  const txBody = child(sp, 'cdr:txBody');
  if (!txBody || isLine) return;
  const fontRef = refs?.font;
  const body = resolveTextBody({
    txBody, lstStyles: [child(txBody, 'a:lstStyle')], bodyPrs: [child(txBody, 'a:bodyPr')],
    theme: env.theme, scope: env.scope, slideNumber: 0, ...(fontRef ? { fontRef } : {}),
  });
  if (!body.paragraphs.some(p => p.runs.some(run => run.text.trim()))) return;
  const layout = layoutText(body, { w: r.w, h: r.h }, env.fonts);
  // the text turns with its shape, about the shape's centre
  texts.push({ kind: 'text', transform: placeAt(r.x + r.w / 2, r.y + r.h / 2, r.w / 2, r.h / 2, rot), layout });
}

/** Draw the shapes of the chart's drawing part, if it has one, over the chart. */
export function drawUserShapes(env: Env, box: { w: number; h: number }, paths: Primitive[], texts: Primitive[]): void {
  const part = env.doc.targets(env.part, 'chartUserShapes')[0];
  if (!part) return;
  const root = env.doc.root(part);
  for (const a of children(root)) {
    const r = anchorRect(a, box);
    if (!r) continue;
    for (const sp of children(a)) if (is(sp, 'cdr:sp') || is(sp, 'cdr:cxnSp')) drawShape(env, sp, r, paths, texts);
  }
}
