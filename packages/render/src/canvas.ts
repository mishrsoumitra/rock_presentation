// packages/render/src/canvas.ts  -- the Canvas 2D backend (plan sections 3.3 and 3.6)
//
// The same code draws in the editor (inside Konva's sceneFunc), in workers (OffscreenCanvas) and in headless
// Chromium for tests and exports, so the pixels the tests compare are the product's pixels.
import type { Matrix, Paint, ResolvedEffects, ResolvedRunProps, Rgba, Stroke, LineEnd, ResolvedBlipEffect } from '@pptx/resolve';
import type { FontFace, GlyphRun } from '@pptx/text';
import { patternPixels } from './patterns.ts';
import type { MissingPrim, PathPrim, PathSpec, Primitive, RenderShape, RenderSlide, TextPrim } from './tree.ts';

/** The subset of CanvasRenderingContext2D (and OffscreenCanvasRenderingContext2D) the renderer uses. */
export interface Ctx {
  save(): void;
  restore(): void;
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  fill(path: Path2D, rule?: CanvasFillRule): void;
  stroke(path: Path2D): void;
  clip(path: Path2D, rule?: CanvasFillRule): void;
  fillRect(x: number, y: number, w: number, h: number): void;
  strokeRect(x: number, y: number, w: number, h: number): void;
  fillText(text: string, x: number, y: number, maxWidth?: number): void;
  drawImage(image: CanvasImageSource, sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number): void;
  createLinearGradient(x0: number, y0: number, x1: number, y1: number): CanvasGradient;
  createRadialGradient(x0: number, y0: number, r0: number, x1: number, y1: number, r1: number): CanvasGradient;
  createPattern(image: CanvasImageSource, repetition: string | null): CanvasPattern | null;
  setLineDash(segments: number[]): void;
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  lineCap: CanvasLineCap;
  lineJoin: CanvasLineJoin;
  miterLimit: number;
  globalAlpha: number;
  globalCompositeOperation: GlobalCompositeOperation;
  shadowBlur: number;
  shadowColor: string;
  shadowOffsetX: number;
  shadowOffsetY: number;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  /** CSS filter (blur for soft edges and reflections); absent in contexts that lack it. */
  filter?: string;
}

/** How animation changes a shape while it is drawn (Phase 5): the shape's own XML is never touched. */
export interface ShapeOverride {
  hidden?: boolean;
  /** Multiplies the shape's opacity, 0..1. */
  opacity?: number;
  /** Applied in slide space after the shape's own transform: motion paths, growth, spin. */
  matrix?: Matrix;
  /** Slide-space clip for wipe-like effects. */
  clip?: Path2D;
  /** Emphasis colours: replace the fill, the line colour, or the text colour. */
  fill?: Paint;
  lineColor?: Rgba;
  textColor?: Rgba;
  /** Text builds: per paragraph index, what happens to that paragraph alone. */
  paragraphs?: ReadonlyMap<number, ParagraphOverride>;
}

/** One paragraph's animation state: like a shape's, applied to that paragraph's lines only. */
export interface ParagraphOverride {
  hidden?: boolean;
  opacity?: number;
  /** Slide-space matrix applied before the shape's own override matrix: the paragraph flying in, growing... */
  matrix?: Matrix;
  /** Slide-space clip (wipes and other reveals of the paragraph). */
  clip?: Path2D;
  textColor?: Rgba;
}

export interface DrawOptions {
  /** Device pixels per EMU. */
  scale: number;
  /**
   * A transform the drawing sits under, applied after every transform the renderer sets: the pixel ratio and node
   * transform a Konva sceneFunc draws in, for instance. Without it the renderer owns the context's transform.
   */
  base?: Matrix;
  /** Animation state per shape, consulted for every shape drawn. */
  override?: (shape: RenderShape) => ShapeOverride | undefined;
  /** Decoded images by part name; undefined while not loaded (a grey box is drawn instead). */
  image: (part: string) => CanvasImageSource | undefined;
  /** Makes scratch canvases for patterns and picture effects; without one those fall back to plain colours. */
  createCanvas?: (w: number, h: number) => OffscreenCanvas;
}

const css = (c: Rgba, alpha = 1) => `rgba(${Math.round(c.r * 255)},${Math.round(c.g * 255)},${Math.round(c.b * 255)},${+(c.a * alpha).toFixed(4)})`;
const toDOM = (m: Matrix) => new DOMMatrix([m[0], m[1], m[2], m[3], m[4], m[5]]);

// ---------------------------------------------------------------------------------------------------------- paths

const pathCache = new WeakMap<PathSpec, Path2D>();

/** The Path2D of a geometry path in its shape's box coordinates. */
export function toPath2D(spec: PathSpec): Path2D {
  let p = pathCache.get(spec);
  if (p) return p;
  p = new Path2D();
  for (const c of spec.commands) {
    switch (c.op) {
      case 'M': p.moveTo(c.x, c.y); break;
      case 'L': p.lineTo(c.x, c.y); break;
      case 'Q': p.quadraticCurveTo(c.x1, c.y1, c.x, c.y); break;
      case 'C': p.bezierCurveTo(c.x1, c.y1, c.x2, c.y2, c.x, c.y); break;
      case 'Z': p.closePath(); break;
    }
  }
  pathCache.set(spec, p);
  return p;
}

/** The same path in slide coordinates, so strokes keep their nominal width inside scaled groups. */
function inSlide(spec: PathSpec, m: Matrix): Path2D {
  const p = new Path2D();
  p.addPath(toPath2D(spec), toDOM(m));
  return p;
}

// ---------------------------------------------------------------------------------------------------------- paints

function shadeColor(c: Rgba, mode: string): Rgba {
  // path fill modes give preset shapes their 3D look (can, cube, bevel): lighten or darken the fill
  const k = mode === 'darken' ? 0.6 : mode === 'darkenLess' ? 0.8 : mode === 'lighten' ? 0.4 : mode === 'lightenLess' ? 0.2 : 0;
  if (mode.startsWith('darken')) return { r: c.r * k, g: c.g * k, b: c.b * k, a: c.a };
  if (mode.startsWith('lighten')) return { r: c.r + (1 - c.r) * k, g: c.g + (1 - c.g) * k, b: c.b + (1 - c.b) * k, a: c.a };
  return c;
}

function shadePaint(p: Paint, mode: string): Paint {
  if (mode === 'norm' || mode === 'none') return p;
  if (p.kind === 'solid') return { kind: 'solid', color: shadeColor(p.color, mode) };
  if (p.kind === 'linear' || p.kind === 'path') return { ...p, stops: p.stops.map(s => ({ pos: s.pos, color: shadeColor(s.color, mode) })) };
  return p;
}

const patternCache = new Map<string, CanvasImageSource>();

function patternImage(p: Extract<Paint, { kind: 'pattern' }>, opts: DrawOptions): CanvasImageSource | undefined {
  if (!opts.createCanvas) return undefined;
  const key = `${p.preset}|${css(p.fg)}|${css(p.bg)}`;
  let img = patternCache.get(key);
  if (!img) {
    const canvas = opts.createCanvas(8, 8);
    const c = canvas.getContext('2d');
    if (!c) return undefined;
    c.putImageData(new ImageData(patternPixels(p.preset, p.fg, p.bg), 8, 8), 0, 0);
    img = canvas;
    patternCache.set(key, img);
  }
  return img;
}

/** Fill `path` (in the shape box, current transform = box to device) with a paint. */
function fillWithPaint(ctx: Ctx, path: Path2D, paint: Paint, box: { w: number; h: number }, opts: DrawOptions, slideMatrix: Matrix): void {
  const w = box.w || 1;
  const h = box.h || 1;
  switch (paint.kind) {
    case 'none':
    case 'group':
      return;
    case 'solid':
      ctx.fillStyle = css(paint.color);
      ctx.fill(path);
      return;
    case 'linear': {
      const stops = paint.stops.length ? paint.stops : [{ pos: 0, color: { r: 0, g: 0, b: 0, a: 1 } }];
      const rad = (paint.angle * Math.PI) / 180;
      const cos = Math.cos(rad);
      const sin = Math.sin(rad);
      ctx.save();
      // scaled: the angle is set in a unit square and stretched with the shape; otherwise it holds in the real box
      const sx = paint.scaled ? w : 1;
      const sy = paint.scaled ? h : 1;
      const bw = paint.scaled ? 1 : w;
      const bh = paint.scaled ? 1 : h;
      ctx.transform(sx, 0, 0, sy, 0, 0);
      const half = Math.abs((bw / 2) * cos) + Math.abs((bh / 2) * sin);
      const g = ctx.createLinearGradient(bw / 2 - cos * half, bh / 2 - sin * half, bw / 2 + cos * half, bh / 2 + sin * half);
      for (const s of stops) g.addColorStop(Math.min(1, Math.max(0, s.pos)), css(s.color));
      ctx.fillStyle = g;
      const scaled = new Path2D();
      scaled.addPath(path, new DOMMatrix([1 / sx, 0, 0, 1 / sy, 0, 0]));
      ctx.fill(scaled);
      ctx.restore();
      return;
    }
    case 'path': {
      // the focus is the centre of fillToRect; stop 0 is drawn there, the last stop at the farthest corner
      const r = paint.fillToRect;
      const fx = r.l + (1 - r.l - r.r) / 2;
      const fy = r.t + (1 - r.t - r.b) / 2;
      const radius = Math.max(Math.hypot(fx, fy), Math.hypot(1 - fx, fy), Math.hypot(fx, 1 - fy), Math.hypot(1 - fx, 1 - fy));
      ctx.save();
      ctx.transform(w, 0, 0, h, 0, 0);
      const g = ctx.createRadialGradient(fx, fy, 0, fx, fy, paint.path === 'circle' ? radius : radius * 0.75);
      for (const s of paint.stops) g.addColorStop(Math.min(1, Math.max(0, s.pos)), css(s.color));
      ctx.fillStyle = g;
      const unit = new Path2D();
      unit.addPath(path, new DOMMatrix([1 / w, 0, 0, 1 / h, 0, 0]));
      ctx.fill(unit);
      ctx.restore();
      return;
    }
    case 'pattern': {
      const img = patternImage(paint, opts);
      const pattern = img ? ctx.createPattern(img, 'repeat') : null;
      if (!pattern) {
        ctx.fillStyle = css(paint.fg);
        ctx.fill(path);
        return;
      }
      // pattern cells are a fixed size on the slide (one pixel at 96 DPI), whatever the shape's transform
      const device = new DOMMatrix([opts.scale, 0, 0, opts.scale, 0, 0]).multiply(toDOM(slideMatrix));
      pattern.setTransform(device.inverse().multiply(new DOMMatrix([9525 * opts.scale, 0, 0, 9525 * opts.scale, 0, 0])));
      ctx.fillStyle = pattern;
      ctx.fill(path);
      return;
    }
    case 'image': {
      const img = imageFor(paint, opts);
      ctx.save();
      ctx.clip(path);
      if (!img) {
        ctx.fillStyle = 'rgba(200,200,200,1)';
        ctx.fillRect(0, 0, w, h);
        ctx.restore();
        return;
      }
      const iw = (img as { width: number }).width || 1;
      const ih = (img as { height: number }).height || 1;
      const src = paint.srcRect;
      const sx = src.l * iw;
      const sy = src.t * ih;
      const sw = iw * (1 - src.l - src.r);
      const sh = ih * (1 - src.t - src.b);
      if (paint.mode.kind === 'stretch') {
        const f = paint.mode.fillRect;
        const dx = f.l * w;
        const dy = f.t * h;
        const dw = w * (1 - f.l - f.r);
        const dh = h * (1 - f.t - f.b);
        if (sw > 0 && sh > 0 && dw > 0 && dh > 0) drawCropped(ctx, img, iw, ih, sx, sy, sw, sh, dx, dy, dw, dh);
      } else {
        const t = paint.mode;
        const pattern = ctx.createPattern(img, 'repeat');
        if (pattern) {
          // tile scale is a fraction of the picture's size at 96 DPI
          const kx = 9525 * t.sx;
          const ky = 9525 * t.sy;
          pattern.setTransform(new DOMMatrix([kx, 0, 0, ky, t.tx, t.ty]));
          ctx.fillStyle = pattern;
          ctx.fillRect(0, 0, w, h);
        }
      }
      ctx.restore();
      return;
    }
  }
}

/** drawImage with a source rectangle that may reach outside the picture (negative crop extends it). */
function drawCropped(ctx: Ctx, img: CanvasImageSource, iw: number, ih: number, sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number): void {
  const x0 = Math.max(0, sx);
  const y0 = Math.max(0, sy);
  const x1 = Math.min(iw, sx + sw);
  const y1 = Math.min(ih, sy + sh);
  if (x1 <= x0 || y1 <= y0) return;
  const kx = dw / sw;
  const ky = dh / sh;
  ctx.drawImage(img, x0, y0, x1 - x0, y1 - y0, dx + (x0 - sx) * kx, dy + (y0 - sy) * ky, (x1 - x0) * kx, (y1 - y0) * ky);
}

const processed = new Map<string, CanvasImageSource>();

/** The picture to draw for an image paint, with its picture effects applied (grayscale, brightness, duotone...). */
function imageFor(paint: Extract<Paint, { kind: 'image' }>, opts: DrawOptions): CanvasImageSource | undefined {
  const part = paint.image.svgPart ?? paint.image.part;
  const raw = (part ? opts.image(part) : undefined) ?? (paint.image.part ? opts.image(paint.image.part) : undefined);
  if (!raw) return undefined;
  const effects = paint.effects.filter(e => e.kind !== 'other');
  if (!effects.length || !opts.createCanvas) return raw;
  const key = `${part}|${JSON.stringify(effects)}`;
  let out = processed.get(key);
  if (out) return out;
  const w = (raw as { width: number }).width;
  const h = (raw as { height: number }).height;
  if (!w || !h) return raw;
  const canvas = opts.createCanvas(w, h);
  const c = canvas.getContext('2d');
  if (!c) return raw;
  c.drawImage(raw, 0, 0);
  const data = c.getImageData(0, 0, w, h);
  applyBlipEffects(data.data, effects);
  c.putImageData(data, 0, 0);
  out = canvas;
  processed.set(key, out);
  return out;
}

/** Picture effects of a:blip, applied to RGBA pixels in place. */
export function applyBlipEffects(px: Uint8ClampedArray, effects: ResolvedBlipEffect[]): void {
  for (const e of effects) {
    for (let i = 0; i < px.length; i += 4) {
      const r = px[i]! / 255;
      const g = px[i + 1]! / 255;
      const b = px[i + 2]! / 255;
      const y = 0.299 * r + 0.587 * g + 0.114 * b;
      let o: [number, number, number] = [r, g, b];
      let a = px[i + 3]! / 255;
      switch (e.kind) {
        case 'grayscl': o = [y, y, y]; break;
        case 'lum': {
          // brightness shifts, contrast scales about mid grey
          const f = (v: number) => (v - 0.5) * (1 + e.contrast) + 0.5 + e.bright;
          o = [f(r), f(g), f(b)];
          break;
        }
        case 'biLevel': o = y >= e.thresh ? [1, 1, 1] : [0, 0, 0]; break;
        case 'alphaModFix': a *= e.amt; break;
        case 'duotone': {
          const [c0, c1] = [e.colors[0], e.colors[1] ?? e.colors[0]];
          if (c0 && c1) o = [c0.r + (c1.r - c0.r) * y, c0.g + (c1.g - c0.g) * y, c0.b + (c1.b - c0.b) * y];
          break;
        }
        case 'clrChange': {
          const near = Math.abs(r - e.from.r) < 0.02 && Math.abs(g - e.from.g) < 0.02 && Math.abs(b - e.from.b) < 0.02;
          if (near) { o = [e.to.r, e.to.g, e.to.b]; if (e.useA) a *= e.to.a; }
          break;
        }
        case 'clrRepl': o = [e.color.r, e.color.g, e.color.b]; break;
        default: break;
      }
      px[i] = Math.round(Math.min(1, Math.max(0, o[0])) * 255);
      px[i + 1] = Math.round(Math.min(1, Math.max(0, o[1])) * 255);
      px[i + 2] = Math.round(Math.min(1, Math.max(0, o[2])) * 255);
      px[i + 3] = Math.round(Math.min(1, Math.max(0, a)) * 255);
    }
  }
}

// ---------------------------------------------------------------------------------------------------------- strokes

const END_SIZE = { sm: 2, med: 3, lg: 5 } as const;

/** Start and end of an open path with their directions, for line ends. */
function endpoints(spec: PathSpec): { start: [number, number]; startDir: [number, number]; end: [number, number]; endDir: [number, number] } | undefined {
  const pts: [number, number][] = [];
  for (const c of spec.commands) {
    if (c.op === 'M' || c.op === 'L') pts.push([c.x, c.y]);
    else if (c.op === 'Q') pts.push([c.x1, c.y1], [c.x, c.y]);
    else if (c.op === 'C') pts.push([c.x1, c.y1], [c.x2, c.y2], [c.x, c.y]);
  }
  if (pts.length < 2) return undefined;
  const dir = (a: [number, number], b: [number, number]): [number, number] => {
    const dx = b[0] - a[0];
    const dy = b[1] - a[1];
    const l = Math.hypot(dx, dy) || 1;
    return [dx / l, dy / l];
  };
  let i = 1;
  while (i < pts.length - 1 && pts[i]![0] === pts[0]![0] && pts[i]![1] === pts[0]![1]) i++;
  let j = pts.length - 2;
  while (j > 0 && pts[j]![0] === pts.at(-1)![0] && pts[j]![1] === pts.at(-1)![1]) j--;
  return { start: pts[0]!, startDir: dir(pts[i]!, pts[0]!), end: pts.at(-1)!, endDir: dir(pts[j]!, pts.at(-1)!) };
}

/** A line end at `tip` pointing along `dir`, in slide coordinates. */
function lineEnd(ctx: Ctx, e: LineEnd, tip: [number, number], dir: [number, number], width: number, color: string): void {
  const w = END_SIZE[e.w] * width;
  const len = END_SIZE[e.len] * width;
  const [dx, dy] = dir;
  const nx = -dy;
  const ny = dx;
  const at = (back: number, side: number): [number, number] => [tip[0] - dx * back + nx * side, tip[1] - dy * back + ny * side];
  const p = new Path2D();
  const poly = (...ps: [number, number][]) => {
    p.moveTo(ps[0]![0], ps[0]![1]);
    for (const q of ps.slice(1)) p.lineTo(q[0], q[1]);
    p.closePath();
  };
  ctx.fillStyle = color;
  ctx.strokeStyle = color;
  switch (e.type) {
    case 'triangle': poly(tip, at(len, w / 2), at(len, -w / 2)); ctx.fill(p); return;
    case 'stealth': poly(tip, at(len, w / 2), at(len * 0.6, 0), at(len, -w / 2)); ctx.fill(p); return;
    case 'diamond': poly(at(-len / 2, 0), at(0, w / 2), at(len / 2, 0), at(0, -w / 2)); ctx.fill(p); return;
    case 'oval': p.ellipse(tip[0], tip[1], len / 2, w / 2, Math.atan2(dy, dx), 0, Math.PI * 2); ctx.fill(p); return;
    case 'arrow': {
      const a = at(len, w / 2);
      const b = at(len, -w / 2);
      p.moveTo(a[0], a[1]);
      p.lineTo(tip[0], tip[1]);
      p.lineTo(b[0], b[1]);
      ctx.lineWidth = width;
      ctx.setLineDash([]);
      ctx.stroke(p);
      return;
    }
  }
}

/** Pull a straight end segment back under a solid arrowhead so the line does not poke through its tip. */
function shortened(spec: PathSpec, head: LineEnd | undefined, tail: LineEnd | undefined, width: number): PathSpec {
  const solid = (e: LineEnd | undefined) => e && (e.type === 'triangle' || e.type === 'stealth');
  if (!solid(head) && !solid(tail)) return spec;
  const cmds = spec.commands.map(c => ({ ...c }));
  const pull = (i: number, prev: number, e: LineEnd) => {
    const c = cmds[i];
    const p = cmds[prev];
    if (!c || !p || (c.op !== 'L' && c.op !== 'M') || (p.op !== 'L' && p.op !== 'M')) return;
    const dx = c.x - p.x;
    const dy = c.y - p.y;
    const l = Math.hypot(dx, dy);
    const back = Math.min(l * 0.9, END_SIZE[e.len] * width * 0.8);
    if (l > 0) { c.x -= (dx / l) * back; c.y -= (dy / l) * back; }
  };
  if (solid(tail)) pull(cmds.length - 1, cmds.length - 2, tail!);
  if (solid(head)) pull(0, 1, head!);
  return { ...spec, commands: cmds };
}

function strokePaths(ctx: Ctx, prim: PathPrim, stroke: Stroke, opts: DrawOptions): void {
  const color = stroke.paint.kind === 'solid' ? css(stroke.paint.color)
    : stroke.paint.kind === 'linear' || stroke.paint.kind === 'path' ? css(stroke.paint.stops[0]?.color ?? { r: 0, g: 0, b: 0, a: 1 })
    : stroke.paint.kind === 'pattern' ? css(stroke.paint.fg) : 'rgba(0,0,0,1)';
  const width = Math.max(stroke.width, 1 / opts.scale);          // a zero-width line is a hairline, one device pixel
  ctx.save();
  setMatrix(ctx, opts, [opts.scale, 0, 0, opts.scale, 0, 0]);
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.lineCap = stroke.cap === 'rnd' ? 'round' : stroke.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = stroke.join === 'round' ? 'round' : stroke.join === 'bevel' ? 'bevel' : 'miter';
  ctx.miterLimit = stroke.miterLimit;
  ctx.setLineDash(stroke.dash ? stroke.dash.map(d => d * width) : []);
  for (const spec of prim.paths) {
    if (!spec.stroke) continue;
    const ends = stroke.head || stroke.tail ? endpoints(spec) : undefined;
    const drawn = ends ? shortened(spec, stroke.head, stroke.tail, width) : spec;
    const p = inSlide(drawn, prim.transform);
    if (stroke.compound !== 'sng') {
      // double lines: the full width, minus a gap a third as wide down the middle
      ctx.stroke(p);
      ctx.globalCompositeOperation = 'destination-out';
      ctx.lineWidth = width / 3;
      ctx.stroke(p);
      ctx.globalCompositeOperation = 'source-over';
      ctx.lineWidth = width;
    } else {
      ctx.stroke(p);
    }
    if (ends) {
      const map = (pt: [number, number]): [number, number] => [prim.transform[0] * pt[0] + prim.transform[2] * pt[1] + prim.transform[4], prim.transform[1] * pt[0] + prim.transform[3] * pt[1] + prim.transform[5]];
      const mapDir = (d: [number, number]): [number, number] => {
        const v: [number, number] = [prim.transform[0] * d[0] + prim.transform[2] * d[1], prim.transform[1] * d[0] + prim.transform[3] * d[1]];
        const l = Math.hypot(v[0], v[1]) || 1;
        return [v[0] / l, v[1] / l];
      };
      if (stroke.tail) lineEnd(ctx, stroke.tail, map(ends.end), mapDir(ends.endDir), width, color);
      if (stroke.head) lineEnd(ctx, stroke.head, map(ends.start), mapDir(ends.startDir), width, color);
    }
  }
  ctx.restore();
}

// ---------------------------------------------------------------------------------------------------------- shapes

function compose(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
  ];
}

/** Set the context's transform to `m` (slide or shape space to device pixels), under the drawing's base transform. */
function setMatrix(ctx: Ctx, opts: DrawOptions, m: Matrix): void {
  const t = opts.base ? compose(opts.base, m) : m;
  ctx.setTransform(t[0], t[1], t[2], t[3], t[4], t[5]);
}

function setDevice(ctx: Ctx, m: Matrix, opts: DrawOptions): void {
  const k = opts.scale;
  setMatrix(ctx, opts, [m[0] * k, m[1] * k, m[2] * k, m[3] * k, m[4] * k, m[5] * k]);
}

/** Canvas shadows are in device pixels and ignore the transform: convert PowerPoint's EMU distances. */
function applyShadow(ctx: Ctx, e: ResolvedEffects, scale: number): boolean {
  const s = e.outerShadow;
  if (s) {
    const rad = (s.dir * Math.PI) / 180;
    ctx.shadowColor = css(s.color);
    ctx.shadowBlur = s.blurRad * scale;              // blurRad and shadowBlur are not the same quantity; to be fitted
    ctx.shadowOffsetX = Math.cos(rad) * s.dist * scale;
    ctx.shadowOffsetY = Math.sin(rad) * s.dist * scale;
    return true;
  }
  return false;
}

function clearShadow(ctx: Ctx): void {
  ctx.shadowColor = 'rgba(0,0,0,0)';
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

/** The full device transform of a primitive's box: base, then device scale, then the primitive's own transform. */
function deviceOf(prim: { transform: Matrix }, opts: DrawOptions): Matrix {
  const k = opts.scale, m = prim.transform;
  const scaled: Matrix = [m[0] * k, m[1] * k, m[2] * k, m[3] * k, m[4] * k, m[5] * k];
  return opts.base ? compose(opts.base, scaled) : scaled;
}

/** A shape's bounds in device pixels, with a margin, for the offscreen layers soft edges and reflections need. */
function deviceBounds(prim: PathPrim, opts: DrawOptions, pad: number): { x: number; y: number; w: number; h: number; dev: Matrix } {
  const d = deviceOf(prim, opts);
  const { w, h } = prim.box;
  const pts = [[0, 0], [w, 0], [w, h], [0, h]].map(([u, v]) => [d[0] * u! + d[2] * v! + d[4], d[1] * u! + d[3] * v! + d[5]]);
  const x0 = Math.floor(Math.min(...pts.map(p => p[0]!)) - pad), y0 = Math.floor(Math.min(...pts.map(p => p[1]!)) - pad);
  const x1 = Math.ceil(Math.max(...pts.map(p => p[0]!)) + pad), y1 = Math.ceil(Math.max(...pts.map(p => p[1]!)) + pad);
  return { x: x0, y: y0, w: Math.max(1, x1 - x0), h: Math.max(1, y1 - y0), dev: d };
}

type Layer = { canvas: CanvasImageSource; ctx: Ctx };

/** An offscreen layer with the shape's fill and line drawn as on the slide, offset to the layer's origin. */
function shapeLayer(prim: PathPrim, opts: DrawOptions, b: ReturnType<typeof deviceBounds>): Layer | undefined {
  const canvas = opts.createCanvas?.(b.w, b.h) as (CanvasImageSource & { getContext(t: '2d'): Ctx | null }) | undefined;
  const ctx = canvas?.getContext('2d');
  if (!canvas || !ctx) return undefined;
  const inner: DrawOptions = { ...opts, base: compose([1, 0, 0, 1, -b.x, -b.y], opts.base ?? [1, 0, 0, 1, 0, 0]) };
  const plain: PathPrim = { ...prim, effects: {} };
  drawPathPlain(ctx, plain, inner);
  return { canvas, ctx };
}

function drawPath(ctx: Ctx, prim: PathPrim, opts: DrawOptions): void {
  const e = prim.effects;
  const offscreen = (e.softEdge || e.reflection) && opts.createCanvas;
  if (!offscreen) {
    drawPathPlain(ctx, prim, opts);
    return;
  }
  const blurPx = Math.max(e.softEdge?.radius ?? 0, e.reflection?.blurRad ?? 0) * opts.scale;
  const b = deviceBounds(prim, opts, Math.ceil(blurPx * 2) + 2);
  const layer = shapeLayer(prim, opts, b);
  if (!layer) {
    drawPathPlain(ctx, prim, opts);
    return;
  }
  // soft edge: keep the shape only where a blurred copy of its silhouette is opaque, so its edges fade inward
  if (e.softEdge && 'filter' in layer.ctx) {
    const mask = opts.createCanvas!(b.w, b.h) as CanvasImageSource & { getContext(t: '2d'): Ctx | null };
    const m = mask.getContext('2d');
    if (m) {
      m.filter = `blur(${Math.max(0.5, (e.softEdge.radius * opts.scale) / 2)}px)`;
      m.drawImage(layer.canvas, 0, 0, b.w, b.h, 0, 0, b.w, b.h);
      layer.ctx.save();
      layer.ctx.setTransform(1, 0, 0, 1, 0, 0);
      layer.ctx.globalCompositeOperation = 'destination-in';
      layer.ctx.drawImage(mask, 0, 0, b.w, b.h, 0, 0, b.w, b.h);
      layer.ctx.restore();
    }
  }
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  // glow and shadow are cast by the layer itself; drawn from far off the canvas, only the shadow lands on it
  const cast = (color: string, blur: number, dx: number, dy: number) => {
    const far = 1e5;
    ctx.shadowColor = color;
    ctx.shadowBlur = blur;
    ctx.shadowOffsetX = dx + far;
    ctx.shadowOffsetY = dy;
    ctx.drawImage(layer.canvas, 0, 0, b.w, b.h, b.x - far, b.y, b.w, b.h);
    clearShadow(ctx);
  };
  if (e.glow) cast(css(e.glow.color), e.glow.radius * opts.scale, 0, 0);
  if (e.outerShadow) {
    const sh = e.outerShadow, rad = (sh.dir * Math.PI) / 180;
    cast(css(sh.color), sh.blurRad * opts.scale, Math.cos(rad) * sh.dist * opts.scale, Math.sin(rad) * sh.dist * opts.scale);
  }
  ctx.drawImage(layer.canvas, 0, 0, b.w, b.h, b.x, b.y, b.w, b.h);
  const r = e.reflection;
  if (r) {
    // the reflection: the layer mirrored about the shape's bottom edge, faded from stA to endA, moved by dist
    const top = b.dev[5] + Math.min(0, b.dev[3] * prim.box.h);
    const bottom = top + Math.abs(b.dev[3] * prim.box.h) + Math.abs(b.dev[2] * prim.box.w);
    const fade = opts.createCanvas!(b.w, b.h) as CanvasImageSource & { getContext(t: '2d'): Ctx | null };
    const f = fade.getContext('2d');
    if (f) {
      f.drawImage(layer.canvas, 0, 0, b.w, b.h, 0, 0, b.w, b.h);
      const height = bottom - top;
      const g = f.createLinearGradient(0, bottom - b.y - height * r.stPos, 0, bottom - b.y - height * Math.max(r.endPos, r.stPos + 0.001));
      g.addColorStop(0, `rgba(0,0,0,${r.stA})`);
      g.addColorStop(1, `rgba(0,0,0,${r.endA})`);
      f.globalCompositeOperation = 'destination-in';
      f.fillStyle = g;
      f.fillRect(0, 0, b.w, b.h);
      const dist = r.dist * opts.scale;
      ctx.setTransform(1, 0, 0, -1, 0, 2 * bottom + dist);
      if (r.blurRad && 'filter' in ctx) ctx.filter = `blur(${(r.blurRad * opts.scale) / 2}px)`;
      ctx.drawImage(fade, 0, 0, b.w, b.h, b.x, b.y, b.w, b.h);
    }
  }
  ctx.restore();
}

/** Fill, inner shadow and line of a path primitive, with its glow and outer shadow. */
function drawPathPlain(ctx: Ctx, prim: PathPrim, opts: DrawOptions): void {
  const { scale } = opts;
  ctx.save();
  setDevice(ctx, prim.transform, opts);
  // glow: the shape's own outline, blurred outward in the glow colour, under the shape
  const glow = prim.effects.glow;
  if (glow) {
    ctx.shadowColor = css(glow.color);
    ctx.shadowBlur = glow.radius * scale;
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;
    ctx.fillStyle = css(glow.color);
    for (const spec of prim.paths) if (spec.fill !== 'none') ctx.fill(toPath2D(spec));
    clearShadow(ctx);
  }
  const shadow = applyShadow(ctx, prim.effects, scale);
  for (const spec of prim.paths) {
    if (spec.fill === 'none') continue;
    fillWithPaint(ctx, toPath2D(spec), shadePaint(prim.fill, spec.fill), prim.box, opts, prim.transform);
  }
  clearShadow(ctx);
  // inner shadow: a frame around the shape, clipped to the shape, so only the frame's shadow shows inside it
  const inner = prim.effects.innerShadow;
  if (inner && prim.fill.kind !== 'none') {
    const { w, h } = prim.box;
    for (const spec of prim.paths) {
      if (spec.fill === 'none') continue;
      const shape = toPath2D(spec);
      ctx.save();
      ctx.clip(shape);
      const frame = new Path2D();
      frame.rect(-w * 2, -h * 2, w * 5, h * 5);
      frame.addPath(shape);
      const rad = (inner.dir * Math.PI) / 180;
      ctx.shadowColor = css(inner.color);
      ctx.shadowBlur = inner.blurRad * scale;
      ctx.shadowOffsetX = Math.cos(rad) * inner.dist * scale;
      ctx.shadowOffsetY = Math.sin(rad) * inner.dist * scale;
      ctx.fillStyle = css(inner.color);
      ctx.fill(frame, 'evenodd');
      ctx.restore();
    }
  }
  ctx.restore();
  if (prim.stroke) {
    ctx.save();
    // with no fill to cast it, the line casts the shadow
    if (shadow && (prim.fill.kind === 'none' || prim.paths.every(p => p.fill === 'none'))) applyShadow(ctx, prim.effects, scale);
    strokePaths(ctx, prim, prim.stroke, opts);
    ctx.restore();
  }
}

// ------------------------------------------------------------------------------------------------------------- text

const glyphPaths = new WeakMap<FontFace, Map<number, Path2D>>();

function glyphPath(face: FontFace, glyph: number): Path2D {
  let byGlyph = glyphPaths.get(face);
  if (!byGlyph) {
    byGlyph = new Map();
    glyphPaths.set(face, byGlyph);
  }
  let p = byGlyph.get(glyph);
  if (!p) {
    p = new Path2D();
    for (const op of face.outline(glyph)) {
      switch (op.op) {
        case 'M': p.moveTo(op.x, op.y); break;
        case 'L': p.lineTo(op.x, op.y); break;
        case 'Q': p.quadraticCurveTo(op.x1, op.y1, op.x, op.y); break;
        case 'C': p.bezierCurveTo(op.x1, op.y1, op.x2, op.y2, op.x, op.y); break;
        case 'Z': p.closePath(); break;
      }
    }
    byGlyph.set(glyph, p);
  }
  return p;
}

const ITALIC_SLANT = Math.tan((12 * Math.PI) / 180);

/** All glyphs of a run as one path in text box coordinates (font units scaled, y flipped, synthetic italic slanted). */
function runPath(run: GlyphRun<ResolvedRunProps>): Path2D {
  const face = run.match.face;
  const unit = run.size / face.metrics.upem;
  const slant = run.match.syntheticItalic ? ITALIC_SLANT * unit : 0;
  const path = new Path2D();
  for (let g = 0; g < run.glyphs.length; g++) {
    const x = run.x + run.xs[g]!;
    const y = run.baseline + run.ys[g]!;
    path.addPath(glyphPath(face, run.glyphs[g]!), new DOMMatrix([unit, 0, slant, -unit, x, y]));
  }
  return path;
}

function drawText(ctx: Ctx, prim: TextPrim, opts: DrawOptions): void {
  ctx.save();
  setDevice(ctx, prim.transform, opts);
  for (const line of prim.layout.lines) {
    for (const run of line.runs) {
      const props = run.props;
      const face = run.match.face;
      const unit = run.size / face.metrics.upem;
      if (props.highlight && !run.bullet) {
        ctx.fillStyle = css(props.highlight);
        ctx.fillRect(run.x, line.top, run.advance, line.height);
      }
      const path = runPath(run);
      const fill = run.bullet && props.fill.kind !== 'solid' ? { kind: 'solid' as const, color: { r: 0, g: 0, b: 0, a: 1 } } : props.fill;
      ctx.save();
      if (props.effects) applyShadow(ctx, props.effects, opts.scale);
      fillWithPaint(ctx, path, fill, { w: prim.layout.width, h: prim.layout.height }, opts, prim.transform);
      if (run.match.syntheticBold && fill.kind === 'solid') {
        ctx.strokeStyle = css(fill.color);
        ctx.lineWidth = run.size * 0.025;
        ctx.lineJoin = 'round';
        ctx.setLineDash([]);
        ctx.stroke(path);
      }
      ctx.restore();
      if (props.outline && props.outline.paint.kind === 'solid') {
        ctx.strokeStyle = css(props.outline.paint.color);
        ctx.lineWidth = props.outline.width;
        ctx.setLineDash([]);
        ctx.stroke(path);
      }
      if (run.bullet) continue;
      const decoColor = fill.kind === 'solid' ? css(fill.color) : 'rgba(0,0,0,1)';
      if (props.underline && props.underline !== 'none') {
        const thickness = Math.max(face.metrics.underlineThickness * unit, 1 / opts.scale) * (props.underline.startsWith('heavy') || props.underline === 'thick' ? 2 : 1);
        const y = run.baseline - face.metrics.underlinePosition * unit;
        ctx.fillStyle = decoColor;
        ctx.fillRect(run.x, y - thickness / 2, run.advance, thickness);
        if (props.underline === 'dbl') ctx.fillRect(run.x, y + thickness * 1.5, run.advance, thickness);
      }
      if (props.strike && props.strike !== 'noStrike') {
        const thickness = Math.max(face.metrics.strikeoutSize * unit, 1 / opts.scale);
        const y = run.baseline - face.metrics.strikeoutPosition * unit;
        ctx.fillStyle = decoColor;
        ctx.fillRect(run.x, y - thickness / 2, run.advance, thickness);
        if (props.strike === 'dblStrike') ctx.fillRect(run.x, y - thickness * 2.5, run.advance, thickness);
      }
    }
  }
  ctx.restore();
}

function drawMissing(ctx: Ctx, prim: MissingPrim, opts: DrawOptions): void {
  ctx.save();
  setDevice(ctx, prim.transform, opts);
  ctx.fillStyle = 'rgba(235,235,235,1)';
  ctx.fillRect(0, 0, prim.box.w, prim.box.h);
  ctx.strokeStyle = 'rgba(160,160,160,1)';
  ctx.lineWidth = 12700;
  ctx.setLineDash([50800, 50800]);
  ctx.strokeRect(0, 0, prim.box.w, prim.box.h);
  ctx.setLineDash([]);
  // the label in device pixels (a font size in EMU is beyond what canvases draw), turned with the box
  const d = deviceOf(prim, opts);
  const k = Math.hypot(d[0], d[1]) || 1;
  const { w, h } = prim.box;
  ctx.setTransform(d[0] / k, d[1] / k, d[2] / k, d[3] / k, d[0] * (w / 2) + d[2] * (h / 2) + d[4], d[1] * (w / 2) + d[3] * (h / 2) + d[5]);
  ctx.fillStyle = 'rgba(110,110,110,1)';
  ctx.font = `${Math.max(1, Math.min(h / 4, 14 * 12700) * k)}px sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(prim.label, 0, 0, Math.max(1, w * k * 0.9));
  ctx.restore();
}

/** Apply an override's slide-space matrix and colours to one primitive. */
function overridden(prim: Primitive, o: ShapeOverride | undefined): Primitive {
  if (!o) return prim;
  let p: Primitive = o.matrix ? { ...prim, transform: compose(o.matrix, prim.transform) } : prim;
  if (p.kind === 'path') {
    if (o.fill && p.fill.kind !== 'none') p = { ...p, fill: o.fill };
    if (o.lineColor && p.stroke) p = { ...p, stroke: { ...p.stroke, paint: { kind: 'solid', color: o.lineColor } } };
  } else if (p.kind === 'text' && (o.textColor || o.paragraphs)) {
    const color = o.textColor;
    const paras = o.paragraphs;
    const lines = p.layout.lines
      .filter(l => !paras?.get(l.paragraph)?.hidden)
      .map(l => ({ ...l, runs: l.runs.map(r => (color ? { ...r, props: { ...r.props, fill: { kind: 'solid' as const, color } } } : r)) }));
    p = { ...p, layout: { ...p.layout, lines } };
  }
  return p;
}

export function drawPrimitive(ctx: Ctx, prim: Primitive, opts: DrawOptions): void {
  if (prim.kind === 'path') drawPath(ctx, prim, opts);
  else if (prim.kind === 'text') drawText(ctx, prim, opts);
  else drawMissing(ctx, prim, opts);
}

/** Draw one shape and its children: what the editor's Konva sceneFunc calls for each shape. */
export function drawShape(ctx: Ctx, shape: RenderShape, opts: DrawOptions): void {
  const o = opts.override?.(shape);
  if (o?.hidden) return;
  const layered = o && (o.opacity !== undefined || o.clip);
  if (layered) {
    ctx.save();
    if (o.opacity !== undefined) ctx.globalAlpha *= Math.max(0, Math.min(1, o.opacity));
    if (o.clip) {
      setMatrix(ctx, opts, [opts.scale, 0, 0, opts.scale, 0, 0]);
      ctx.clip(o.clip);
    }
  }
  for (const prim of shape.prims) if (prim.kind !== 'text') drawPrimitive(ctx, overridden(prim, o), opts);
  // members of a group move with it: pass the group's matrix down, and keep each member's own animation
  const inner: DrawOptions = o?.matrix
    ? { ...opts, override: s => { const own = opts.override?.(s); return { ...own, matrix: own?.matrix ? compose(o.matrix!, own.matrix) : o.matrix! }; } }
    : opts;
  for (const c of shape.children) drawShape(ctx, c, inner);
  for (const prim of shape.prims) {
    if (prim.kind !== 'text') continue;
    const paras = o?.paragraphs;
    const own = paras ? [...paras].filter(([, v]) => !v.hidden && (v.opacity !== undefined || v.matrix || v.clip || v.textColor)) : [];
    if (!own.length) {
      drawPrimitive(ctx, overridden(prim, o), opts);
      continue;
    }
    // paragraphs with a state of their own are drawn one by one: their own motion first, then the shape's
    const whole = o ? { ...o, paragraphs: new Map<number, ParagraphOverride>() } : undefined;
    for (const [index, state] of own) {
      const lines = prim.layout.lines.filter(l => l.paragraph === index);
      if (!lines.length) continue;
      ctx.save();
      if (state.opacity !== undefined) ctx.globalAlpha *= Math.max(0, Math.min(1, state.opacity));
      if (state.clip) {
        setMatrix(ctx, opts, [opts.scale, 0, 0, opts.scale, 0, 0]);
        ctx.clip(state.clip);
      }
      let one: Primitive = { ...prim, layout: { ...prim.layout, lines }, ...(state.matrix ? { transform: compose(state.matrix, prim.transform) } : {}) };
      one = overridden(one, state.textColor ? { ...whole, textColor: state.textColor } : whole);
      drawPrimitive(ctx, one, opts);
      ctx.restore();
    }
    const mine = new Set(own.map(([i]) => i));
    const rest = { ...prim, layout: { ...prim.layout, lines: prim.layout.lines.filter(l => !mine.has(l.paragraph)) } };
    drawPrimitive(ctx, overridden(rest, o), opts);
  }
  if (layered) ctx.restore();
}

/** Draw a whole slide: background, then every shape in order. */
export function drawSlide(ctx: Ctx, slide: RenderSlide, opts: DrawOptions): void {
  const { cx, cy } = slide.size;
  const rect: PathSpec = { commands: [{ op: 'M', x: 0, y: 0 }, { op: 'L', x: cx, y: 0 }, { op: 'L', x: cx, y: cy }, { op: 'L', x: 0, y: cy }, { op: 'Z' }], fill: 'norm', stroke: false };
  drawPath(ctx, { kind: 'path', transform: [1, 0, 0, 1, 0, 0], box: { w: cx, h: cy }, paths: [rect], fill: slide.background, stroke: null, effects: {} }, opts);
  for (const s of slide.shapes) drawShape(ctx, s, opts);
}
