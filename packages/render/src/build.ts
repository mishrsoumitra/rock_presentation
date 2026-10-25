// packages/render/src/build.ts  -- a resolved slide to the render tree (plan sections 3.1, 3.5 and 3.6)
import { evaluateGeometry, parseCustomGeometry, presetGeometry, type EvaluatedGeometry, type ShapeGeometry } from '@pptx/geometry';
import {
  multiply, resolveDrawing, type Matrix, type Paint, type ResolvedCell, type ResolvedShape, type ResolvedSlide, type ResolvedTable,
  type ResolvedRunProps, type ResolvedTextBody, type SlideContext, type Stroke,
} from '@pptx/resolve';
import { layoutText, type FontRegistry, type TextLayout } from '@pptx/text';
import type { PathSpec, Primitive, RenderShape, RenderSlide } from './tree.ts';

export interface BuildOptions {
  fonts: FontRegistry;
  /** Draws a chart part into a box (EMU); @pptx/charts provides one. Without it charts show as placeholders. */
  chart?: (chartPart: string, box: { w: number; h: number }, ctx: SlideContext) => Primitive[] | undefined;
}

const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
const rotateAbout = (deg: number, cx: number, cy: number): Matrix => {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  return [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy];
};

const RECT: ShapeGeometry = presetGeometry('rect')!;

function geometryOf(s: ResolvedShape): ShapeGeometry {
  const g = s.geometry;
  if (!g) return RECT;
  if (g.kind === 'preset') return presetGeometry(g.prst) ?? RECT;
  try {
    return parseCustomGeometry(g.el);
  } catch {
    return RECT;
  }
}

function evaluate(s: ResolvedShape, w: number, h: number): EvaluatedGeometry {
  const geom = geometryOf(s);
  try {
    return evaluateGeometry(geom, w, h, s.geometry?.kind === 'preset' ? s.geometry.av : []);
  } catch {
    // a custom geometry from a file can name guides that do not exist: draw its box rather than nothing
    return evaluateGeometry(RECT, w, h, []);
  }
}

const pathsOf = (e: EvaluatedGeometry): PathSpec[] => e.paths.map(p => ({ commands: p.commands, fill: p.fill, stroke: p.stroke }));

/** Lay out a text body in a rectangle of the shape's box and place it: text box coordinates to slide EMU. */
function textPrim(text: ResolvedTextBody, area: { x: number; y: number; w: number; h: number }, base: Matrix, fonts: FontRegistry): Primitive | undefined {
  if (!text.paragraphs.length) return undefined;
  const placed = placeText(text, area, base, fonts);
  if (!placed.layout.lines.some(l => l.runs.length)) return undefined;
  return { kind: 'text', transform: placed.transform, layout: placed.layout };
}

function placeText(text: ResolvedTextBody, area: { x: number; y: number; w: number; h: number }, base: Matrix, fonts: FontRegistry): { layout: TextLayout<ResolvedRunProps>; transform: Matrix } {
  const layout = layoutText(text, { w: area.w, h: area.h }, fonts);
  let m: Matrix = multiply(base, translate(area.x, area.y));
  // a:bodyPr rot turns the text within its box, about the box centre
  if (text.body.rot) m = multiply(m, rotateAbout(text.body.rot, area.w / 2, area.h / 2));
  // vertical text is laid out in the rotated box and turned into place
  if (layout.rotation === 90) m = multiply(m, [0, 1, -1, 0, area.w, 0]);
  else if (layout.rotation === 270) m = multiply(m, [0, -1, 1, 0, 0, area.h]);
  return { layout, transform: m };
}

/** The rectangle of a shape's box its text is laid out in (the geometry's text rectangle, or SmartArt's own). */
export function textAreaOf(s: ResolvedShape): { x: number; y: number; w: number; h: number } {
  if (s.textXfrm) return { x: 0, y: 0, w: s.textXfrm.cx, h: s.textXfrm.cy };   // textTransform places it
  return evaluate(s, s.xfrm.cx, s.xfrm.cy).textRect;
}

/**
 * A shape's text laid out and placed as the renderer draws it, even when it is empty: what text editing places
 * the caret in and what autofit measures. `text` replaces the shape's own resolved text (to try a font scale).
 */
export function shapeTextLayout(s: ResolvedShape, fonts: FontRegistry, text: ResolvedTextBody | null = s.text): { layout: TextLayout<ResolvedRunProps>; transform: Matrix; area: { x: number; y: number; w: number; h: number } } | undefined {
  if (!text) return undefined;
  const area = textAreaOf(s);
  return { ...placeText(text, area, s.textTransform, fonts), area };
}

const rectPath = (x: number, y: number, w: number, h: number): PathSpec => ({
  commands: [{ op: 'M', x, y }, { op: 'L', x: x + w, y }, { op: 'L', x: x + w, y: y + h }, { op: 'L', x, y: y + h }, { op: 'Z' }],
  fill: 'norm', stroke: false,
});
const linePath = (x1: number, y1: number, x2: number, y2: number): PathSpec => ({
  commands: [{ op: 'M', x: x1, y: y1 }, { op: 'L', x: x2, y: y2 }], fill: 'none', stroke: true,
});

/** A table: cells with their fills, borders and text; rows grow to fit their text, as PowerPoint does. */
function tablePrims(table: ResolvedTable, base: Matrix, fonts: FontRegistry, frameW: number): Primitive[] {
  const cols = table.columns;
  const rows = [...table.rows];
  const colX = [0];
  for (const w of cols) colX.push(colX.at(-1)! + w);
  const spanW = (c: ResolvedCell) => colX[Math.min(cols.length, c.col + c.colSpan)]! - colX[c.col]!;
  const layouts = new Map<ResolvedCell, TextLayout<ResolvedRunProps>>();
  for (const cell of table.cells) {
    if (cell.covered || !cell.text) continue;
    const w = Math.max(0, spanW(cell) - cell.margins.l - cell.margins.r);
    // laid out top-anchored in an unbounded box: the cell applies its own anchoring once row heights are known
    const body = { ...cell.text, body: { ...cell.text.body, insets: { l: 0, t: 0, r: 0, b: 0 }, anchor: 't' as const, wrap: true } };
    const layout = layoutText(body, { w, h: 1e12 }, fonts);
    layouts.set(cell, layout);
    if (cell.rowSpan === 1) rows[cell.row] = Math.max(rows[cell.row] ?? 0, layout.contentHeight + cell.margins.t + cell.margins.b);
  }
  const rowY = [0];
  for (const h of rows) rowY.push(rowY.at(-1)! + h);
  const prims: Primitive[] = [];
  const totalH = rowY.at(-1)!;
  if (table.background.kind !== 'none') {
    prims.push({ kind: 'path', transform: base, box: { w: frameW, h: totalH }, paths: [rectPath(0, 0, colX.at(-1)!, totalH)], fill: table.background, stroke: null, effects: {} });
  }
  const borders: { stroke: Stroke; path: PathSpec }[] = [];
  for (const cell of table.cells) {
    if (cell.covered) continue;
    const x = colX[cell.col] ?? 0;
    const y = rowY[cell.row] ?? 0;
    const w = spanW(cell);
    const h = (rowY[Math.min(rows.length, cell.row + cell.rowSpan)] ?? y) - y;
    const box = { w, h };
    const at = multiply(base, translate(x, y));
    if (cell.fill.kind !== 'none') prims.push({ kind: 'path', transform: at, box, paths: [rectPath(0, 0, w, h)], fill: cell.fill, stroke: null, effects: {} });
    const b = cell.borders;
    if (b.l) borders.push({ stroke: b.l, path: linePath(x, y, x, y + h) });
    if (b.r) borders.push({ stroke: b.r, path: linePath(x + w, y, x + w, y + h) });
    if (b.t) borders.push({ stroke: b.t, path: linePath(x, y, x + w, y) });
    if (b.b) borders.push({ stroke: b.b, path: linePath(x, y + h, x + w, y + h) });
    if (b.tl2br) borders.push({ stroke: b.tl2br, path: linePath(x, y, x + w, y + h) });
    if (b.bl2tr) borders.push({ stroke: b.bl2tr, path: linePath(x, y + h, x + w, y) });
    const layout = layouts.get(cell);
    if (layout && cell.text) {
      const innerH = h - cell.margins.t - cell.margins.b;
      const shift = cell.anchor === 'ctr' ? (innerH - layout.contentHeight) / 2 : cell.anchor === 'b' ? innerH - layout.contentHeight : 0;
      prims.push({ kind: 'text', transform: multiply(at, translate(cell.margins.l, cell.margins.t + shift - layout.contentTop)), layout });
    }
  }
  // borders go on top of every fill, so a neighbour's fill cannot cover a shared edge
  for (const { stroke, path } of borders) prims.push({ kind: 'path', transform: base, box: { w: frameW, h: totalH }, paths: [path], fill: { kind: 'none' }, stroke, effects: {} });
  return prims;
}

function collectImages(p: Paint, into: Set<string>): void {
  if (p.kind === 'image') {
    if (p.image.svgPart) into.add(p.image.svgPart);
    if (p.image.part) into.add(p.image.part);
  }
}

/** A text body scaled down in steps, as shrink-on-overflow would, until it fits the box (the SmartArt outline). */
function shrunkToFit(text: ResolvedTextBody, w: number, h: number, fonts: FontRegistry): ResolvedTextBody {
  for (const scale of [1, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3]) {
    const t = { ...text, body: { ...text.body, autofit: { kind: 'norm' as const, fontScale: scale, lnSpcReduction: scale < 1 ? 0.2 : 0 } } };
    if (layoutText(t, { w, h }, fonts).contentHeight <= h) return t;
  }
  return { ...text, body: { ...text.body, autofit: { kind: 'norm', fontScale: 0.25, lnSpcReduction: 0.2 } } };
}

function buildShape(slide: ResolvedSlide, s: ResolvedShape, opts: BuildOptions, images: Set<string>): RenderShape {
  const w = s.xfrm.cx;
  const h = s.xfrm.cy;
  const geom = evaluate(s, w, h);
  const outline = pathsOf(geom);
  const prims: Primitive[] = [];
  const box = { w, h };
  collectImages(s.fill, images);

  if (s.kind !== 'grpSp' && s.kind !== 'graphicFrame') {
    if (s.fill.kind !== 'none' || s.line || s.effects.outerShadow || s.effects.glow) {
      prims.push({ kind: 'path', transform: s.transform, box, paths: outline, fill: s.fill, stroke: s.line, effects: s.effects });
    }
  }

  const g = s.graphic;
  if (g?.kind === 'table') prims.push(...tablePrims(g.table, s.transform, opts.fonts, w));
  else if (g?.kind === 'chart') {
    const chart = opts.chart?.(g.part, box, slide.context);
    // pictures inside a chart (picture and texture fills) load like any other
    for (const p of chart ?? []) if (p.kind === 'path') collectImages(p.fill, images);
    if (chart) prims.push(...chart.map(p => ({ ...p, transform: multiply(s.transform, p.transform) })));
    else prims.push({ kind: 'missing', transform: s.transform, box, label: 'Chart' });
  } else if (g?.kind === 'ole') {
    collectImages(g.preview, images);
    if (g.preview.kind === 'image') prims.push({ kind: 'path', transform: s.transform, box, paths: [rectPath(0, 0, w, h)], fill: g.preview, stroke: null, effects: {} });
    else prims.push({ kind: 'missing', transform: s.transform, box, label: 'Embedded object' });
  } else if (g?.kind === 'unknown') {
    prims.push({ kind: 'missing', transform: s.transform, box, label: 'Unsupported content' });
  }

  let children: RenderShape[] = [];
  if (g?.kind === 'diagram') {
    const drawn = g.drawingPart ? resolveDrawing(slide, g.drawingPart, s) : [];
    if (drawn.length) children = drawn.map(c => buildShape(slide, c, opts, images));
    else {
      const t = g.fallback ? textPrim(shrunkToFit(g.fallback, w, h, opts.fonts), { x: 0, y: 0, w, h }, s.textTransform, opts.fonts) : undefined;
      prims.push(t ?? { kind: 'missing', transform: s.transform, box, label: 'SmartArt' });
    }
  }

  if (s.text) {
    // SmartArt shapes carry their own text rectangle (placed by textTransform); other shapes use their geometry's
    const area = s.textXfrm ? { x: 0, y: 0, w: s.textXfrm.cx, h: s.textXfrm.cy } : geom.textRect;
    const t = textPrim(s.text, area, s.textTransform, opts.fonts);
    if (t) prims.push(t);
  }

  if (s.kind === 'grpSp') children = s.children.map(c => buildShape(slide, c, opts, images));
  return { id: s.id, name: s.name, layer: s.layer, transform: s.transform, box, outline, prims, children, source: s };
}

/**
 * One shape of a resolved slide built on its own (`slide` gives charts and SmartArt their context). The editor builds
 * the shape an edit touched this way and keeps the rest of the slide's tree; `images` collects the parts it draws.
 */
export function buildSlideShape(slide: ResolvedSlide, s: ResolvedShape, opts: BuildOptions, images: Set<string>): RenderShape {
  return buildShape(slide, s, opts, images);
}

/** Build the render tree of a resolved slide. Pure: the same slide and fonts give the same tree. */
export function buildSlide(slide: ResolvedSlide, opts: BuildOptions): RenderSlide {
  const images = new Set<string>();
  collectImages(slide.background, images);
  const shapes = slide.shapes.map(s => buildShape(slide, s, opts, images));
  return { part: slide.part, size: slide.size, background: slide.background, shapes, images };
}
