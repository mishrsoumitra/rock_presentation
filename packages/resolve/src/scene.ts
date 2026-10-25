// packages/resolve/src/scene.ts  -- a slide as a list of shapes with every effective value resolved
import { parseXml, type XmlElement } from '@pptx/xml';
import {
  attr, attrNs, child, children, colorIn, effectiveChildren, effectsIn, fillIn, is, parseFill, parseLine, parseStyle, ShapeView, textOf, UNDERSTOOD,
  type ChildSpace, type Placeholder, type PresentationDocument, type ShapeKind, type SlideKind, type SlideView, type Xfrm,
} from '@pptx/model';
import type { ListStyleDef } from '@pptx/model';
import { slideContext, scopeFor, themeScope, type SlideContext } from './context.ts';
import {
  resolveEffects, resolveFill, resolveShapeStyle, themeFill, withPlaceholderColor,
  type Paint, type ResolvedEffects, type Stroke,
} from './paint.ts';
import { matchLayoutPlaceholder, matchMasterPlaceholder, textStyleBucket } from './placeholders.ts';
import { resolveTable, type ResolvedTable } from './tables.ts';
import { resolveTextBody, type ResolvedTextBody } from './text.ts';

/** [a, b, c, d, e, f]: x' = a*x + c*y + e, y' = b*x + d*y + f. Shape box coordinates (EMU) to slide EMU. */
export type Matrix = [number, number, number, number, number, number];

export type GeometryRef = { kind: 'preset'; prst: string; av: { name: string; fmla: string }[] } | { kind: 'custom'; el: XmlElement };

export type ResolvedGraphic =
  | { kind: 'table'; table: ResolvedTable }
  | { kind: 'chart'; part: string }
  | { kind: 'diagram'; drawingPart?: string; dataPart?: string; fallback?: ResolvedTextBody }
  | { kind: 'ole'; preview: Paint }
  | { kind: 'unknown'; uri: string };

export interface ResolvedShape {
  id: number;
  name: string;
  kind: ShapeKind;
  layer: 'master' | 'layout' | 'slide';
  part: string;
  view: ShapeView;
  placeholder: Placeholder | null;
  xfrm: Xfrm;                                  // effective, in the parent's coordinate space
  transform: Matrix;                           // the shape's box (0..cx, 0..cy) to slide coordinates
  textTransform: Matrix;                       // the same without mirroring: text stays readable (SmartArt: textXfrm's box)
  geometry: GeometryRef | null;
  fill: Paint;
  line: Stroke | null;
  effects: ResolvedEffects;
  text: ResolvedTextBody | null;
  /** SmartArt drawing shapes lay out their text in their own rectangle (in the shape's parent space). */
  textXfrm?: Xfrm;
  graphic?: ResolvedGraphic;
  media?: { kind: 'video' | 'audio'; part?: string; url?: string };
  /** Where a click goes in the slide show: an external URL, a slide of this deck (a jump), or an action. */
  hyperlink?: { url?: string; action?: string; slide?: string };
  children: ResolvedShape[];
}

export interface ResolvedSlide {
  part: string;
  kind: SlideKind;
  size: { cx: number; cy: number };
  background: Paint;
  /** In drawing order: the master's shapes, the layout's, then the slide's own. */
  shapes: ResolvedShape[];
  context: SlideContext;
}

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

export function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

export function applyMatrix(m: Matrix, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

export function invert(m: Matrix): Matrix {
  const det = m[0] * m[3] - m[1] * m[2];
  if (!det) return IDENTITY;
  const [a, b, c, d, e, f] = m;
  return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det];
}

/** A box placed by an xfrm: rotate about its centre, mirror for flips. `text` leaves out mirroring, and turns text
 *  upside down for a vertical flip, as PowerPoint does. */
export function boxMatrix(x: Xfrm, text = false): Matrix {
  const hw = x.cx / 2;
  const hh = x.cy / 2;
  const deg = x.rot / 60000 + (text && x.flipV ? 180 : 0);
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const fx = !text && x.flipH ? -1 : 1;
  const fy = !text && x.flipV ? -1 : 1;
  // T(x+hw, y+hh) * R * S(fx, fy) * T(-hw, -hh)
  const a = cos * fx;
  const b = sin * fx;
  const c = -sin * fy;
  const d = cos * fy;
  return [a, b, c, d, x.x + hw - (a * hw + c * hh), x.y + hh - (b * hw + d * hh)];
}

/**
 * The box a SmartArt drawing shape's text is laid out in (dsp:txXfrm): its own rectangle, in the shape's
 * coordinates, turned about its own centre. Its rot adds to the shape's: PowerPoint writes rot="-5400000" on a shape
 * turned by 5400000 for upright text, and rot="10800000" on one turned by 10800000. Without rot, files disagree (a
 * rectangle the size of the shape's box that turns with it; one the size of the turned shape that stays upright),
 * so the text goes whichever way its rectangle fits in the shape.
 */
export function smartArtTextBox(shape: Xfrm, text: Xfrm, hasRot: boolean): Xfrm {
  const box = { ...text, flipH: false, flipV: shape.flipV };
  if (hasRot) return { ...box, rot: shape.rot + text.rot };
  const turned = { ...box, rot: shape.rot };
  const upright = { ...box, rot: 0 };
  return outside(shape, upright) + 1 < outside(shape, turned) ? upright : turned;
}

/** How far a (turned) rectangle's corners reach outside a shape's (turned) box, in EMU. */
function outside(shape: Xfrm, box: Xfrm): number {
  const m = multiply(invert(boxMatrix({ ...shape, flipH: false, flipV: false })), boxMatrix(box, true));
  let out = 0;
  for (const [u, v] of [[0, 0], [box.cx, 0], [box.cx, box.cy], [0, box.cy]] as const) {
    const x = m[0] * u + m[2] * v + m[4];
    const y = m[1] * u + m[3] * v + m[5];
    out += Math.max(0, -x, x - shape.cx) + Math.max(0, -y, y - shape.cy);
  }
  return out;
}

/** Child coordinates of a group (its chOff/chExt space) into the group's own box. */
export function childSpaceMatrix(group: Xfrm, space: ChildSpace | null): Matrix {
  if (!space) return IDENTITY;
  const sx = space.chExt.cx ? group.cx / space.chExt.cx : 1;
  const sy = space.chExt.cy ? group.cy / space.chExt.cy : 1;
  return [sx, 0, 0, sy, -space.chOff.x * sx, -space.chOff.y * sy];
}

const ZERO_XFRM: Xfrm = { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };

interface Inherited { layoutPh?: ShapeView; masterPh?: ShapeView; }

function inheritedFor(ctx: SlideContext, layer: 'master' | 'layout' | 'slide', s: ShapeView): Inherited {
  const ph = s.placeholder;
  if (!ph) return {};
  if (layer === 'slide') {
    if (ctx.kind === 'notes') {
      const masterPh = ctx.master ? matchMasterPlaceholder(ph, ctx.master.view.shapes) : undefined;
      return masterPh ? { masterPh } : {};
    }
    const layoutPh = ctx.layout ? matchLayoutPlaceholder(ph, ctx.layout.view.shapes) : undefined;
    const masterPh = ctx.master ? matchMasterPlaceholder(layoutPh?.placeholder ?? ph, ctx.master.view.shapes) : undefined;
    return { ...(layoutPh ? { layoutPh } : {}), ...(masterPh ? { masterPh } : {}) };
  }
  if (layer === 'layout') {
    const masterPh = ctx.master ? matchMasterPlaceholder(ph, ctx.master.view.shapes) : undefined;
    return masterPh ? { masterPh } : {};
  }
  return {};
}

interface ResolveState {
  ctx: SlideContext;
  listCache: Map<XmlElement, ListStyleDef>;
  background: Paint;
}

/**
 * The text of a SmartArt data model as an outline: its points in tree order (dgm:cxn parOf connections from the
 * document point down), each point's text a paragraph at its depth, resolved in the master's body style.
 */
function diagramOutline(state: ResolveState, dataPart: string, part: string): ResolvedTextBody | undefined {
  const ctx = state.ctx;
  let root: XmlElement;
  try {
    root = ctx.doc.root(dataPart);
  } catch {
    return undefined;
  }
  const pts = new Map<string, XmlElement>();
  for (const pt of children(child(root, 'dgm:ptLst'), 'dgm:pt')) pts.set(attr(pt, 'modelId') ?? '', pt);
  const kids = new Map<string, { id: string; ord: number }[]>();
  for (const c of children(child(root, 'dgm:cxnLst'), 'dgm:cxn')) {
    const type = attr(c, 'type') ?? 'parOf';
    if (type !== 'parOf') continue;
    const src = attr(c, 'srcId') ?? '';
    const list = kids.get(src) ?? [];
    list.push({ id: attr(c, 'destId') ?? '', ord: Number(attr(c, 'srcOrd') ?? 0) });
    kids.set(src, list);
  }
  const doc = [...pts.values()].find(p => attr(p, 'type') === 'doc');
  if (!doc) return undefined;
  const lines: { level: number; text: string }[] = [];
  const seen = new Set<string>();
  const walk = (id: string, depth: number) => {
    if (seen.has(id) || depth > 9) return;                  // cycles and absurd depth in hand-made files
    seen.add(id);
    for (const k of (kids.get(id) ?? []).sort((a, b) => a.ord - b.ord)) {
      const pt = pts.get(k.id);
      const type = attr(pt, 'type') ?? 'node';
      if (!pt || (type !== 'node' && type !== 'asst')) continue;
      const text = children(child(pt, 'dgm:t'), 'a:p').map(p => children(p, 'a:r').map(r => textOf(child(r, 'a:t'))).join('')).join(' ').trim();
      if (text) lines.push({ level: Math.min(8, depth), text });
      walk(k.id, depth + 1);
    }
  };
  walk(attr(doc, 'modelId') ?? '', 0);
  if (!lines.length) return undefined;
  const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const body = parseXml(`<a:txBody xmlns:a="${a}"><a:bodyPr anchor="ctr"/><a:lstStyle/>${lines.map(l =>
    `<a:p>${l.level ? `<a:pPr lvl="${l.level}"/>` : ''}<a:r><a:rPr lang="en-US"/><a:t>${esc(l.text)}</a:t></a:r></a:p>`).join('')}</a:txBody>`).root;
  const bucket = ctx.master?.view.txStyles.body;
  return resolveTextBody({
    txBody: body, lstStyles: [], bodyPrs: [child(body, 'a:bodyPr')],
    ...(bucket ? { bucket } : {}),
    ...(ctx.doc.presentation.defaultTextStyle ? { defaultTextStyle: ctx.doc.presentation.defaultTextStyle } : {}),
    theme: ctx.theme, scope: scopeFor(ctx, part), slideNumber: 0,
  }, state.listCache);
}

const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';

/** The preview picture of a legacy OLE object: v:shape (o:spid) > v:imagedata (o:relid) in a VML drawing. */
function vmlPreview(ctx: SlideContext, part: string, spid: string | undefined): { part: string } | undefined {
  if (!spid) return undefined;
  for (const vml of ctx.doc.targets(part, 'vmlDrawing')) {
    let root: XmlElement;
    try {
      root = ctx.doc.root(vml);
    } catch {
      continue;                                   // old VML is not always well-formed XML; then there is no preview
    }
    const named = (e: XmlElement) => e.ns === VML_NS && e.local === 'shape'
      && e.attrs.some(a => a.value === spid && ((a.ns === OFFICE_NS && a.local === 'spid') || (a.ns === '' && a.local === 'id')));
    const find = (e: XmlElement): XmlElement | undefined => {
      if (named(e)) return e;
      for (const c of e.children) {
        if (c.kind !== 'element') continue;
        const hit = find(c);
        if (hit) return hit;
      }
      return undefined;
    };
    const shape = find(root);
    const data = shape?.children.find((c): c is XmlElement => c.kind === 'element' && c.ns === VML_NS && c.local === 'imagedata');
    const relId = data?.attrs.find(a => a.local === 'relid' || (a.local === 'id' && a.ns.endsWith('relationships')))?.value;
    const target = relId ? ctx.doc.target(vml, relId) : undefined;
    if (target) return { part: target };
  }
  return undefined;
}

function resolveGraphic(state: ResolveState, s: ShapeView, part: string): ResolvedGraphic | undefined {
  const gd = s.graphicData;
  if (!gd) return undefined;
  const ctx = state.ctx;
  const kind = gd.uri.slice(gd.uri.lastIndexOf('/') + 1);
  if (kind === 'table') {
    const tbl = child(gd.el, 'a:tbl');
    return tbl ? { kind: 'table', table: resolveTable(ctx, tbl, part) } : { kind: 'unknown', uri: gd.uri };
  }
  if (kind === 'chart') {
    const rId = attrNs(child(gd.el, 'c:chart'), 'r', 'id');
    const target = rId ? ctx.doc.target(part, rId) : undefined;
    return target ? { kind: 'chart', part: target } : { kind: 'unknown', uri: gd.uri };
  }
  if (kind === 'diagram') {
    const relIds = child(gd.el, 'dgm:relIds');
    const dataPart = ctx.doc.target(part, attrNs(relIds, 'r', 'dm') ?? '');
    let drawingPart: string | undefined;
    // the cached drawing is named by a relationship of the slide, recorded in the data model's extension
    if (dataPart) {
      for (const ext of children(child(ctx.doc.root(dataPart), 'dgm:extLst'), 'a:ext')) {
        const relId = attr(child(ext, 'dsp:dataModelExt'), 'relId');
        if (relId) drawingPart = ctx.doc.target(part, relId);
      }
    }
    drawingPart ??= ctx.doc.targets(part, 'diagramDrawing').length === 1 ? ctx.doc.targets(part, 'diagramDrawing')[0] : undefined;
    // without PowerPoint's cached drawing (files from other producers), or with an empty one (PowerPoint lays the
    // diagram out again from its data; we cannot), the data model's text as a bulleted list
    const drawn = drawingPart ? drawnShapes(ctx.doc.root(drawingPart)).length > 0 : false;
    const fallback = !drawn && dataPart ? diagramOutline(state, dataPart, part) : undefined;
    return { kind: 'diagram', ...(drawingPart ? { drawingPart } : {}), ...(dataPart ? { dataPart } : {}), ...(fallback ? { fallback } : {}) };
  }
  if (kind === 'ole') {
    const ole = effectiveChildren(gd.el, UNDERSTOOD).find(c => is(c, 'p:oleObj'));
    let preview = resolveFill(parseFill(child(child(ole, 'p:pic'), 'p:blipFill')), scopeFor(ctx, part));
    if (preview.kind === 'none') {
      // PowerPoint 2007 kept the preview in the slide's VML drawing, on the v:shape named by the object's spid
      const image = vmlPreview(ctx, part, attr(ole, 'spid'));
      if (image) {
        preview = {
          kind: 'image', image, srcRect: { l: 0, t: 0, r: 0, b: 0 }, rotWithShape: true, effects: [],
          mode: { kind: 'stretch', fillRect: { l: 0, t: 0, r: 0, b: 0 } },
        };
      }
    }
    return { kind: 'ole', preview };
  }
  return { kind: 'unknown', uri: gd.uri };
}

function resolveShape(state: ResolveState, s: ShapeView, layer: 'master' | 'layout' | 'slide', part: string, parent: Matrix): ResolvedShape {
  const ctx = state.ctx;
  const inh = inheritedFor(ctx, layer, s);
  const chain = [s, inh.layoutPh, inh.masterPh].filter((v): v is ShapeView => v !== undefined);
  const xfrm = s.xfrm ?? inh.layoutPh?.xfrm ?? inh.masterPh?.xfrm ?? ZERO_XFRM;
  const transform = multiply(parent, boxMatrix(xfrm));
  const txXfrm = s.txXfrm;
  const textTransform = multiply(parent, boxMatrix(txXfrm ? smartArtTextBox(xfrm, txXfrm, s.txXfrmHasRot) : xfrm, true));

  let geometry: GeometryRef | null = null;
  for (const v of chain) {
    const g = v.geometry;
    if (g) {
      geometry = g;
      break;
    }
  }
  if (!geometry && (s.kind === 'sp' || s.kind === 'pic')) geometry = { kind: 'preset', prst: 'rect', av: [] };

  // spPr properties are inherited from placeholders piece by piece; style references fill what is left
  const pick = <T>(f: (v: ShapeView) => T | undefined): { value: T; from: ShapeView } | undefined => {
    for (const v of chain) {
      const x = f(v);
      if (x !== undefined) return { value: x, from: v };
    }
    return undefined;
  };
  const partOf = (v: ShapeView) => (v === s ? part : v.owner?.part ?? part);
  const fillPick = pick(v => fillIn(v.spPr));
  const linePick = pick(v => parseLine(child(v.spPr, 'a:ln')));
  const effectsPick = pick(v => effectsIn(v.spPr));
  const stylePick = pick(v => parseStyle(v.style));
  const styleScope = scopeFor(ctx, part);
  const resolvedStyle = resolveShapeStyle(
    {}, stylePick?.value, ctx.theme, styleScope, themeScope(ctx),
  );
  let fill: Paint = fillPick ? resolveFill(fillPick.value, scopeFor(ctx, partOf(fillPick.from))) : resolvedStyle.fill;
  if (s.kind === 'pic' && s.blipFill) fill = resolveFill(parseFill(s.blipFill), scopeFor(ctx, part));
  if (s.useBgFill) fill = state.background;
  let line = resolvedStyle.line;
  if (linePick) {
    const own = resolveShapeStyle({ line: linePick.value }, stylePick?.value, ctx.theme, scopeFor(ctx, partOf(linePick.from)), themeScope(ctx));
    line = own.line;
  }
  const effects = effectsPick ? resolveEffects(effectsPick.value, scopeFor(ctx, partOf(effectsPick.from))) : resolvedStyle.effects;

  let text: ResolvedTextBody | null = null;
  const txBody = s.txBody;
  if (txBody) {
    const bucketName = textStyleBucket(s.placeholder);
    const bucket = ctx.kind === 'notes' && s.placeholder?.type === 'body'
      ? ctx.master?.view.notesStyle
      : ctx.master?.view.txStyles[bucketName];
    const fontRef = stylePick?.value.font;
    text = resolveTextBody({
      txBody,
      lstStyles: chain.map(v => child(v.txBody, 'a:lstStyle')),
      bodyPrs: chain.map(v => child(v.txBody, 'a:bodyPr')),
      ...(fontRef ? { fontRef } : {}),
      ...(bucket ? { bucket } : {}),
      ...(ctx.doc.presentation.defaultTextStyle ? { defaultTextStyle: ctx.doc.presentation.defaultTextStyle } : {}),
      theme: ctx.theme,
      scope: scopeFor(ctx, part),
      slideNumber: layer === 'slide' && ctx.kind === 'slide' ? ctx.slideNumber : 0,
    }, state.listCache);
  }

  const out: ResolvedShape = {
    id: s.id, name: s.name, kind: s.kind, layer, part, view: s, placeholder: s.placeholder, xfrm, transform, textTransform,
    geometry, fill, line, effects, text, children: [],
  };
  if (txXfrm) out.textXfrm = txXfrm;
  const graphic = resolveGraphic(state, s, part);
  if (graphic) out.graphic = graphic;
  const media = s.media;
  if (media) {
    const rId = media.embed ?? media.link;
    const target = rId ? ctx.doc.target(part, rId) : undefined;
    const rel = rId ? ctx.doc.pkg.relationships(part).find(r => r.id === rId) : undefined;
    out.media = { kind: media.kind, ...(target ? { part: target } : {}), ...(rel?.external ? { url: rel.target } : {}) };
  }
  const link = s.hlinkClick;
  if (link) {
    const rId = attrNs(link, 'r', 'id');
    const rel = rId ? ctx.doc.pkg.relationships(part).find(r => r.id === rId) : undefined;
    const action = attr(link, 'action');
    const slide = rel && !rel.external && rId ? ctx.doc.target(part, rId) : undefined;
    out.hyperlink = { ...(rel?.external ? { url: rel.target } : {}), ...(action ? { action } : {}), ...(slide ? { slide } : {}) };
  }
  if (s.kind === 'grpSp') {
    const inner = multiply(transform, childSpaceMatrix(xfrm, s.childSpace));
    // a:grpFill in a member means "the group's fill"
    out.children = s.children.filter(c => !c.hidden).map(c => {
      const r = resolveShape(state, c, layer, part, inner);
      return r.fill.kind === 'group' ? { ...r, fill } : r;
    });
  }
  return out;
}

/** The background of a slide: its own p:bg, else its layout's, else its master's. p:bgPr is a fill of its own;
 *  p:bgRef points into the theme's fills (1001 and up: background fills) with its colour as phClr. */
function resolveBackground(ctx: SlideContext): Paint {
  const layers = [ctx.slide, ctx.layout, ctx.master].filter((l): l is { part: string; view: SlideView } => l !== undefined);
  for (const l of layers) {
    const bg = l.view.bg;
    if (!bg) continue;
    const bgPr = child(bg, 'p:bgPr');
    if (bgPr) return resolveFill(fillIn(bgPr), scopeFor(ctx, l.part));
    const bgRef = child(bg, 'p:bgRef');
    if (bgRef) {
      const color = colorIn(bgRef);
      return resolveFill(themeFill(ctx.theme, Number(attr(bgRef, 'idx') ?? 0)), withPlaceholderColor(themeScope(ctx), color ? { color } : undefined));
    }
  }
  return { kind: 'solid', color: { r: 1, g: 1, b: 1, a: 1 } };
}

function shapesOf(state: ResolveState, layer: 'master' | 'layout' | 'slide', part: string, view: SlideView, drawPlaceholders: boolean): ResolvedShape[] {
  return [...view.shapes, ...view.controlPictures]
    .filter(s => !s.hidden && (drawPlaceholders || !s.placeholder))
    .map(s => resolveShape(state, s, layer, part, IDENTITY));
}

/** Resolve a slide (or a layout, master or notes page on its own) into shapes in drawing order. */
export function resolveSlide(doc: PresentationDocument, part: string): ResolvedSlide {
  const ctx = slideContext(doc, part);
  const state: ResolveState = { ctx, listCache: new Map(), background: { kind: 'none' } };
  state.background = resolveBackground(ctx);
  const shapes: ResolvedShape[] = [];
  const own = doc.slide(part);
  const size = ctx.kind === 'notes' || ctx.kind === 'notesMaster' ? doc.presentation.notesSize : doc.presentation.slideSize;
  if (ctx.kind === 'slide') {
    const showLayout = own.showMasterSp;
    const showMaster = showLayout && (ctx.layout?.view.showMasterSp ?? true);
    if (showMaster && ctx.master) shapes.push(...shapesOf(state, 'master', ctx.master.part, ctx.master.view, false));
    if (showLayout && ctx.layout) shapes.push(...shapesOf(state, 'layout', ctx.layout.part, ctx.layout.view, false));
    shapes.push(...shapesOf(state, 'slide', part, own, true));
  } else if (ctx.kind === 'layout') {
    if (own.showMasterSp && ctx.master) shapes.push(...shapesOf(state, 'master', ctx.master.part, ctx.master.view, false));
    shapes.push(...shapesOf(state, 'layout', part, own, true));
  } else {
    shapes.push(...shapesOf(state, 'slide', part, own, true));
  }
  return { part, kind: ctx.kind, size: { cx: size.cx, cy: size.cy }, background: state.background, shapes, context: ctx };
}

/**
 * The shapes of a SmartArt diagram's cached drawing (the dsp:drawing part PowerPoint writes beside the data model).
 * They are plain shapes positioned inside the diagram's frame, so they resolve like slide shapes, in the slide's
 * colours and theme, with the frame's transform as their parent. Empty when the part is not a drawing.
 */
/** The shapes of a cached SmartArt drawing (dsp:drawing). */
function drawnShapes(root: XmlElement): XmlElement[] {
  const tree = child(root, 'dsp:spTree');
  if (!tree) return [];
  return effectiveChildren(tree, UNDERSTOOD).filter(c => is(c, 'dsp:sp') || is(c, 'dsp:grpSp') || is(c, 'dsp:pic') || is(c, 'dsp:cxnSp'));
}

export function resolveDrawing(slide: ResolvedSlide, drawingPart: string, frame: ResolvedShape): ResolvedShape[] {
  const ctx = slide.context;
  const state: ResolveState = { ctx, listCache: new Map(), background: slide.background };
  return drawnShapes(ctx.doc.root(drawingPart))
    .map(c => new ShapeView(c, ctx.doc.owner(drawingPart)))
    .filter(s => !s.hidden)
    .map(s => resolveShape(state, s, 'slide', drawingPart, frame.transform));
}
