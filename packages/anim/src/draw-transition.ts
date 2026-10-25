// packages/anim/src/draw-transition.ts  -- one frame of a slide transition, from the old slide's picture to the new one's
//
// The ECMA kinds follow ECMA-376's descriptions (directions name where the motion goes: wipe dir="d" moves down from
// the top). PowerPoint 2010 and later kinds turn and fold the slides in 3D; here they are 2D approximations that keep
// each one's character (a cube turns by squeezing one slide as the other widens, airplane flies the old slide off
// to the top right). Unknown kinds fade.
import { filterReveal, shuffled, type Pt, type Region } from './filters.ts';
import type { TransitionKind, TransitionOptions } from './transition.ts';

/** The part of CanvasRenderingContext2D (and OffscreenCanvasRenderingContext2D) transitions use. */
export interface TransitionCtx {
  save(): void;
  restore(): void;
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  clip(path: Path2D, rule?: CanvasFillRule): void;
  drawImage(image: CanvasImageSource, dx: number, dy: number, dw: number, dh: number): void;
  fillRect(x: number, y: number, w: number, h: number): void;
  fillStyle: string | CanvasGradient | CanvasPattern;
  globalAlpha: number;
}

export interface DrawTransitionOptions {
  /** Makes Path2D objects for the clip patterns; defaults to the global Path2D. Without one, clipped kinds fade. */
  path2d?: () => Path2D;
}

type Spec = TransitionOptions & { kind: TransitionKind };
type Img = CanvasImageSource;
type M = [number, number, number, number, number, number];

/** What every kind draws with: the context, the two pictures, the frame size and a Path2D maker. */
interface Frame { ctx: TransitionCtx; from: Img; to: Img; w: number; h: number; aspect: number; make?: () => Path2D; }

const clamp01 = (v: number) => (v < 0 ? 0 : v > 1 ? 1 : v);
const smooth = (x: number) => x * x * (3 - 2 * x);
const BLACK = 'rgba(0,0,0,1)';
const WHITE = 'rgba(255,255,255,1)';
const rect = (x0: number, y0: number, x1: number, y1: number): Pt[] => [{ x: x0, y: y0 }, { x: x1, y: y0 }, { x: x1, y: y1 }, { x: x0, y: y1 }];

/** Unit vector for a direction of motion: l, r, u, d and the corners. */
function vector(dir: string | undefined, fallback: string): Pt {
  const d = dir ?? fallback;
  return { x: d.includes('l') ? -1 : d.includes('r') ? 1 : 0, y: d.includes('u') ? -1 : d.includes('d') ? 1 : 0 };
}

/** Scale about a point, then rotate about it (degrees), then move. */
function about(cx: number, cy: number, sx: number, sy: number, deg = 0, dx = 0, dy = 0): M {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  const a = c * sx, b = s * sx, cc = -s * sy, d = c * sy;
  return [a, b, cc, d, cx + dx - a * cx - cc * cy, cy + dy - b * cx - d * cy];
}

const shiftBy = (x: number, y: number): M => [1, 0, 0, 1, x, y];

function draw(f: Frame, img: Img, m?: M, alpha = 1): void {
  if (alpha <= 0) return;
  const { ctx } = f;
  ctx.save();
  if (m) ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
  if (alpha < 1) ctx.globalAlpha *= alpha;
  ctx.drawImage(img, 0, 0, f.w, f.h);
  ctx.restore();
}

function fillAll(f: Frame, color: string): void {
  f.ctx.fillStyle = color;
  f.ctx.fillRect(0, 0, f.w, f.h);
}

function regionPath(region: Region, f: Frame): Path2D {
  const path = f.make!();
  for (const poly of region) {
    poly.forEach((p, i) => (i === 0 ? path.moveTo(p.x * f.w, p.y * f.h) : path.lineTo(p.x * f.w, p.y * f.h)));
    path.closePath();
  }
  return path;
}

/** Draw a region of `img`; with a transform, the region moves with the picture (a door's half, a comb's strip). */
function clipped(f: Frame, img: Img, region: Region, m?: M): void {
  if (!region.length || !f.make) return;
  f.ctx.save();
  if (m) f.ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
  f.ctx.clip(regionPath(region, f));
  f.ctx.drawImage(img, 0, 0, f.w, f.h);
  f.ctx.restore();
}

// ------------------------------------------------------------------------------------------- reveal patterns

function circle(c: Pt, r: number, aspect: number): Pt[] {
  return Array.from({ length: 72 }, (_, i) => {
    const a = (i / 72) * Math.PI * 2;
    return { x: c.x + (Math.cos(a) * r) / aspect, y: c.y + Math.sin(a) * r };
  });
}

function rotateAbout(q: Pt, a: number, aspect: number): Pt {
  const x = (q.x - 0.5) * aspect;
  const y = q.y - 0.5;
  return { x: 0.5 + (x * Math.cos(a) - y * Math.sin(a)) / aspect, y: 0.5 + x * Math.sin(a) + y * Math.cos(a) };
}

/** Square cells, n along the longer side, appearing in a random order. */
function cells(n: number, aspect: number, p: number, seed: number): Region {
  const cols = aspect >= 1 ? n : Math.max(1, Math.round(n * aspect));
  const rows = aspect >= 1 ? Math.max(1, Math.round(n / aspect)) : n;
  const order = shuffled(cols * rows, seed);
  return order.slice(0, Math.round(clamp01(p) * cols * rows)).map(k => {
    const c = k % cols;
    const r = Math.floor(k / cols);
    return rect(c / cols, r / rows, (c + 1) / cols, (r + 1) / rows);
  });
}

/** Hexagonal cells, each appearing at a random moment. */
function hexagons(p: number, aspect: number): Region {
  const size = 0.06;
  const dx = size * Math.sqrt(3);
  const cols = Math.ceil(aspect / dx) + 2;
  const rows = Math.ceil(1 / (size * 1.5)) + 2;
  const rank = new Map(shuffled(cols * rows, 0x4e3c).map((k, i) => [k, i]));
  const out: Region = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if ((rank.get(r * cols + c) ?? 0) >= p * cols * rows) continue;
      const cx = (c - 0.5 + (r % 2) * 0.5) * dx;
      const cy = (r - 0.5) * size * 1.5;
      out.push(Array.from({ length: 6 }, (_, k) => {
        const a = (Math.PI / 3) * k + Math.PI / 6;
        return { x: (cx + Math.cos(a) * size * 1.02) / aspect, y: cy + Math.sin(a) * size * 1.02 };
      }));
    }
  }
  return out;
}

/** Diamond or square cells that appear in a sweep across the slide, with some scatter. */
function sweep(shape: 'diamond' | 'hex', v: Pt, p: number, aspect: number): Region {
  const rows = 16;
  const cols = Math.max(1, Math.round(rows * Math.max(1, aspect)));
  const jitter = shuffled(cols * rows, 0x91e7);
  const out: Region = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const x = (c + 0.5) / cols;
      const y = (r + 0.5) / rows;
      const along = v.x < 0 ? 1 - x : v.x > 0 ? x : v.y < 0 ? 1 - y : y;
      if (along * 0.6 + (jitter[r * cols + c]! / (cols * rows)) * 0.4 >= p) continue;
      const hx = 0.5 / cols;
      const hy = 0.5 / rows;
      out.push(shape === 'diamond'
        ? [{ x, y: y - hy * 1.42 }, { x: x + hx * 1.42, y }, { x, y: y + hy * 1.42 }, { x: x - hx * 1.42, y }]
        : rect(x - hx * 1.01, y - hy * 1.01, x + hx * 1.01, y + hy * 1.01));
    }
  }
  return out;
}

/** The pattern a kind reveals the new slide with at progress p, for the kinds that are reveals. */
function revealRegion(t: Spec, p: number, aspect: number): Region | undefined {
  const r = (name: string) => filterReveal(name, p, 'in', aspect).region;
  switch (t.kind) {
    case 'wipe': {
      const d = t.dir ?? 'l';
      return [d === 'l' ? rect(1 - p, 0, 1, 1) : d === 'r' ? rect(0, 0, p, 1) : d === 'u' ? rect(0, 1 - p, 1, 1) : rect(0, 0, 1, p)];
    }
    case 'split': return r(`barn(${t.dir === 'in' ? 'in' : 'out'}${t.orient === 'vert' ? 'Vertical' : 'Horizontal'})`);
    case 'blinds': return r(t.dir === 'vert' ? 'blinds(vertical)' : 'blinds(horizontal)');
    case 'checker': return r(t.dir === 'vert' ? 'checkerboard(down)' : 'checkerboard(across)');
    case 'randomBar': return r(t.dir === 'vert' ? 'randombar(vertical)' : 'randombar(horizontal)');
    case 'circle': return [circle({ x: 0.5, y: 0.5 }, p * Math.hypot(aspect / 2, 0.5) * 1.02, aspect)];
    case 'diamond': return r('diamond(out)');
    case 'plus': return r('plus(out)');
    case 'wedge': return r('wedge');
    case 'wheel': return r(`wheel(${t.spokes ?? 4})`);
    case 'wheelReverse': return r(`wheel(${t.spokes ?? 4})`)?.map(poly => poly.map(q => ({ x: 1 - q.x, y: q.y })).reverse());
    case 'zoom': return r(t.dir === 'in' ? 'box(in)' : 'box(out)');
    case 'strips': {
      const d = t.dir ?? 'lu';
      return r(`strips(${d.includes('u') ? 'up' : 'down'}${d.includes('l') ? 'Left' : 'Right'})`);
    }
    case 'dissolve': return cells(48, aspect, p, 0x5d15);
    case 'fracture': return cells(14, aspect, p, 0xf4ac);
    case 'honeycomb': return hexagons(p, aspect);
    case 'glitter': return sweep(t.pattern === 'hexagon' ? 'hex' : 'diamond', vector(t.dir, 'l'), p, aspect);
    case 'ripple': {
      const corner = t.dir && t.dir !== 'center';
      const c = corner ? { x: t.dir!.includes('r') ? 1 : 0, y: t.dir!.includes('d') ? 1 : 0 } : { x: 0.5, y: 0.5 };
      return [circle(c, p * Math.hypot(Math.max(c.x, 1 - c.x) * aspect, Math.max(c.y, 1 - c.y)) * 1.02, aspect)];
    }
    case 'vortex': return (r('wheel(6)') ?? []).map(poly => poly.map(q => rotateAbout(q, (1 - p) * Math.PI * 0.75, aspect)));
    default: return undefined;
  }
}

// ----------------------------------------------------------------------------------------------- the kinds

function fade(f: Frame, p: number, through?: string): void {
  if (!through) {
    draw(f, f.from);
    draw(f, f.to, undefined, p);
    return;
  }
  fillAll(f, through);
  if (p < 0.5) draw(f, f.from, undefined, 1 - p * 2);
  else draw(f, f.to, undefined, p * 2 - 1);
}

/** Strips of the new slide sliding in from alternate sides (comb, shred). */
function combStrips(f: Frame, vertical: boolean, n: number, p: number): void {
  draw(f, f.from);
  for (let k = 0; k < n; k++) {
    const off = (k % 2 ? 1 : -1) * (1 - p);
    const strip = vertical ? rect(k / n, 0, (k + 1) / n, 1) : rect(0, k / n, 1, (k + 1) / n);
    clipped(f, f.to, [strip], vertical ? shiftBy(0, off * f.h) : shiftBy(off * f.w, 0));
  }
}

/** Slide-turning kinds as 2D squeezes: the old slide narrows towards one side while the new one widens. */
function turning(f: Frame, t: Spec, p: number): void {
  const { w, h } = f;
  const e = smooth(p);
  const v = vector(t.dir, t.kind === 'switch' || t.kind === 'flip' ? 'r' : 'l');
  switch (t.kind) {
    case 'prism': {
      // the edge between the two faces moves across as the prism turns; seen from inside (Box) it also recedes
      const s = t.isInverted ? 1 - 0.12 * Math.sin(Math.PI * e) : 1;
      f.ctx.save();
      const m = about(w / 2, h / 2, s, s);
      f.ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
      const first = v.y === 0 ? v.x < 0 : v.y < 0;
      const [a, b] = first ? [f.from, f.to] : [f.to, f.from];
      const k = first ? 1 - e : e;
      if (v.y === 0) {
        draw(f, a, [Math.max(1e-3, k), 0, 0, 1, 0, 0]);
        draw(f, b, [Math.max(1e-3, 1 - k), 0, 0, 1, k * w, 0]);
      } else {
        draw(f, a, [1, 0, 0, Math.max(1e-3, k), 0, 0]);
        draw(f, b, [1, 0, 0, Math.max(1e-3, 1 - k), 0, k * h]);
      }
      f.ctx.restore();
      return;
    }
    case 'flip':
      draw(f, e < 0.5 ? f.from : f.to, about(w / 2, h / 2, Math.max(1e-3, Math.abs(Math.cos(Math.PI * e))), 1 - 0.1 * Math.sin(Math.PI * e)));
      return;
    case 'switch': {
      // the slides swing apart, trade places front to back half way, and come together again
      const out = Math.sin(Math.PI * e) * 0.6 * w * (v.x || 1);
      const old = about(w / 2, h / 2, 1 - 0.2 * e, 1 - 0.2 * e, 0, out);
      const next = about(w / 2, h / 2, 0.8 + 0.2 * e, 0.8 + 0.2 * e, 0, -out * 0.5);
      if (e < 0.5) {
        draw(f, f.to, next);
        draw(f, f.from, old);
      } else {
        draw(f, f.from, old);
        draw(f, f.to, next);
      }
      return;
    }
    case 'ferris': {
      const dir = v.x || -1;
      const turn = (deg: number) => about(w / 2, h * 2.2, 1, 1, deg);
      draw(f, f.from, turn(dir * 30 * e));
      draw(f, f.to, turn(dir * 30 * (e - 1)));
      return;
    }
    default: {
      // gallery and conveyor: both slides travel, smaller in the middle of the way
      const s = 1 - (t.kind === 'gallery' ? 0.25 : 0.15) * Math.sin(Math.PI * e);
      const dir = v.x || -1;
      draw(f, f.from, about(w / 2, h / 2, s, s, 0, dir * e * w));
      draw(f, f.to, about(w / 2, h / 2, s, s, 0, dir * (e - 1) * w));
    }
  }
}

/** Kinds that take the old slide away over the new one: fly, fold, crush, peel, part or blow it off. */
function away(f: Frame, t: Spec, p: number): void {
  const { w, h } = f;
  const e = smooth(p);
  const cx = w / 2;
  const cy = h / 2;
  const flip = t.invX ? -1 : 1;
  draw(f, f.to);
  switch (t.kind) {
    case 'airplane': {
      const s = 1 - 0.85 * e;
      draw(f, f.from, about(cx, cy, s, s, -25 * e * flip, flip * e * w * 0.7, -e * h * 0.8), 1 - e * e);
      return;
    }
    case 'origami': draw(f, f.from, about(0, 0, 1 - e, 1 - e, 15 * e), 1 - e * 0.5); return;
    case 'crush': draw(f, f.from, about(cx, cy, 1 - e, 1 - e, 35 * e), 1 - e * 0.3); return;
    case 'peelOff': draw(f, f.from, about(w, 0, 1, 1, -12 * e * flip, -flip * e * w * 1.1, e * h * 1.1)); return;
    case 'drape': draw(f, f.from, about(cx, h, 1 - 0.1 * e, 1 - 0.3 * e, 0, 0, e * h), 1 - e * 0.5); return;
    case 'wind': draw(f, f.from, about(cx, cy, 1 - 0.2 * e, 1 - 0.2 * e, 0, -flip * e * w), 1 - e); return;
    case 'fallOver': draw(f, f.from, about(cx, h, 1, Math.max(1e-3, 1 - e)), 1 - e * 0.4); return;
    case 'curtains':
      if (!f.make) {
        draw(f, f.from, undefined, 1 - e);
        return;
      }
      clipped(f, f.from, [rect(0, 0, 0.5, 1)], about(0, cy, 1 - 0.5 * e, 1 + 0.05 * e, 0, -e * w * 0.5));
      clipped(f, f.from, [rect(0.5, 0, 1, 1)], about(w, cy, 1 - 0.5 * e, 1 + 0.05 * e, 0, e * w * 0.5));
      return;
    default: draw(f, f.from, undefined, 1 - e);
  }
}

function drawKind(f: Frame, t: Spec, p: number): void {
  const { w, h } = f;
  const region = f.make ? revealRegion(t, p, f.aspect) : undefined;
  if (region) {
    draw(f, f.from);
    clipped(f, f.to, region, t.kind === 'ripple' ? about(w / 2, h / 2, 1.05 - 0.05 * p, 1.05 - 0.05 * p) : undefined);
    if (t.kind === 'fracture') draw(f, f.to, undefined, p * p);
    return;
  }
  switch (t.kind) {
    case 'none': draw(f, f.to); return;
    case 'cut':
      if (t.thruBlk) fillAll(f, BLACK);
      else draw(f, f.to);
      return;
    case 'fade': fade(f, p, t.thruBlk ? BLACK : undefined); return;
    case 'flash': fade(f, p, WHITE); return;
    case 'push': case 'pan': {
      const v = vector(t.dir, 'l');
      draw(f, f.from, shiftBy(v.x * p * w, v.y * p * h));
      draw(f, f.to, shiftBy(v.x * (p - 1) * w, v.y * (p - 1) * h));
      return;
    }
    case 'cover': {
      const v = vector(t.dir, 'l');
      draw(f, f.from);
      draw(f, f.to, shiftBy(v.x * (p - 1) * w, v.y * (p - 1) * h));
      return;
    }
    case 'pull': case 'reveal': {
      const v = vector(t.dir, 'l');
      const soft = t.kind === 'reveal';
      if (soft && t.thruBlk) fillAll(f, BLACK);
      draw(f, f.to, undefined, soft ? smooth(p) : 1);
      draw(f, f.from, shiftBy(v.x * p * w, v.y * p * h), soft ? 1 - p : 1);
      return;
    }
    case 'comb': if (f.make) return combStrips(f, t.dir === 'vert', 8, smooth(p)); break;
    case 'shred': if (f.make) return combStrips(f, t.pattern !== 'rectangle', 24, smooth(p)); break;
    case 'newsflash':
      draw(f, f.from);
      draw(f, f.to, about(w / 2, h / 2, p, p, 720 * (1 - p)));
      return;
    case 'prism': case 'flip': case 'switch': case 'gallery': case 'conveyor': case 'ferris': turning(f, t, p); return;
    case 'doors': case 'window': {
      // vertical: the split line is vertical, so the halves open to the left and right
      const e = smooth(p);
      draw(f, f.to, about(w / 2, h / 2, 0.85 + 0.15 * e, 0.85 + 0.15 * e), Math.min(1, 0.4 + e));
      if (!f.make) return;
      if ((t.dir ?? 'horz') === 'vert') {
        clipped(f, f.from, [rect(0, 0, 0.5, 1)], about(0, h / 2, 1 - e, 1));
        clipped(f, f.from, [rect(0.5, 0, 1, 1)], about(w, h / 2, 1 - e, 1));
      } else {
        clipped(f, f.from, [rect(0, 0, 1, 0.5)], about(w / 2, 0, 1, 1 - e));
        clipped(f, f.from, [rect(0, 0.5, 1, 1)], about(w / 2, h, 1, 1 - e));
      }
      return;
    }
    case 'warp': case 'flythrough': case 'prestige': {
      const inward = (t.dir ?? (t.kind === 'warp' ? 'out' : 'in')) === 'in';
      const e = smooth(p);
      const sFrom = t.kind === 'prestige' ? 1 - 0.2 * e : inward ? 1 + e : 1 - 0.5 * e;
      const sTo = t.kind === 'prestige' ? 1.1 - 0.1 * e : inward ? 0.5 + 0.5 * e : 1.5 - 0.5 * e;
      if (t.kind === 'prestige') fillAll(f, BLACK);
      draw(f, f.from, about(w / 2, h / 2, sFrom, sFrom), 1 - e);
      draw(f, f.to, about(w / 2, h / 2, sTo, sTo), e);
      return;
    }
    case 'airplane': case 'origami': case 'crush': case 'peelOff': case 'curtains': case 'drape': case 'wind': case 'fallOver':
      away(f, t, p);
      return;
    case 'pageCurlSingle': case 'pageCurlDouble': {
      const e = smooth(p);
      draw(f, f.to);
      if (!f.make) return;
      if (t.kind === 'pageCurlSingle') clipped(f, f.from, [rect(0, 0, 1 - e, 1)]);
      else clipped(f, f.from, [rect(0, 0, 0.5 - e / 2, 1), rect(0.5 + e / 2, 0, 1, 1)]);
      // the shadow of the curl along its edge
      f.ctx.save();
      f.ctx.globalAlpha *= 0.25 * (1 - e);
      f.ctx.fillStyle = BLACK;
      if (t.kind === 'pageCurlSingle') f.ctx.fillRect((1 - e) * w - 0.04 * w, 0, 0.04 * w, h);
      f.ctx.restore();
      return;
    }
    default: break;
  }
  fade(f, p);
}

/**
 * Draw one frame of a transition at `progress` (0: the old slide, 1: the new one) into the w × h area at the
 * context's current transform. `from` and `to` are pictures of the two slides, of any size, drawn to fill the area.
 */
export function drawTransition(
  ctx: TransitionCtx, from: Img, to: Img, transition: Spec, progress: number, w: number, h: number, options: DrawTransitionOptions = {},
): void {
  const p = clamp01(Number.isFinite(progress) ? progress : 1);
  const P = (globalThis as { Path2D?: new () => Path2D }).Path2D;
  const make = options.path2d ?? (P ? () => new P() : undefined);
  const f: Frame = { ctx, from, to, w, h, aspect: h > 0 ? w / h : 1, ...(make ? { make } : {}) };
  ctx.save();
  if (p <= 0) draw(f, from);
  else if (p >= 1) draw(f, to);
  else drawKind(f, transition, p);
  ctx.restore();
}
