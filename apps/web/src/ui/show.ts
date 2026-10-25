// apps/web/src/ui/show.ts  -- the slide show engine: animation steps, transitions, triggers, timings
import {
  clickCount, drawTransition, interactiveSequences, overrideLookup, parseTiming, sampleFrame, scheduleFor, slideTransition,
  type Schedule, type Timing, type Transition,
} from '@pptx/anim';
import { drawSlide, type RenderShape } from '@pptx/render';
import { invert } from '@pptx/resolve';
import type { Session } from '../state/store.ts';
import { isSafeUrl, linkAt } from './links.ts';

interface SlideData {
  key: string;
  timing: Timing | undefined;
  schedule: Schedule | undefined;
  steps: number;
  triggers: Set<number>;
  transition?: Transition;
}

interface Running { from: OffscreenCanvas; to: OffscreenCanvas; spec: Transition; start: number; }

/**
 * Plays a deck on a canvas as PowerPoint's slide show does. A click plays the next animation step, and after the
 * last one moves on with the next slide's transition (drawn from a picture of this slide's last frame to one of
 * the next slide's first). Effects that start with or after the previous one play by themselves; a click on a
 * trigger shape starts its sequence; "advance after" moves on by itself. Nothing in the deck is changed.
 */
export class Show {
  index: number;
  private clicks = 0;
  private t0 = 0;
  private slideStart = 0;
  private triggers: { spid: number; start: number }[] = [];
  private running: Running | undefined;
  private raf = 0;
  private readonly cache = new Map<string, SlideData>();
  private readonly session: Session;
  private readonly canvas: HTMLCanvasElement;
  private readonly parts: string[];
  private readonly exit: () => void;
  onSlide: (part: string | undefined) => void = () => undefined;

  constructor(session: Session, canvas: HTMLCanvasElement, parts: string[], start: number, exit: () => void) {
    this.session = session;
    this.canvas = canvas;
    this.parts = parts;
    this.index = start;
    this.exit = exit;
    this.enter(performance.now());
  }

  get part(): string | undefined { return this.parts[this.index]; }

  private data(part: string): SlideData {
    const scene = this.session.scenes.get(part);
    const key = this.session.scenes.key(part);
    const hit = this.cache.get(part);
    if (hit && hit.key === key) return hit;
    const view = this.session.ed.doc.slide(part);
    const el = view.timing;
    let timing: Timing | undefined, schedule: Schedule | undefined;
    try {
      timing = el ? parseTiming(el) : undefined;
      schedule = timing ? scheduleFor(timing, scene.slide) : undefined;
    } catch {
      timing = undefined;                         // a timing tree that cannot be read plays as no animation
    }
    let transition: Transition | undefined;
    try { transition = slideTransition(view); } catch { transition = undefined; }
    const data: SlideData = {
      key, timing, schedule, steps: schedule ? clickCount(schedule) : 0,
      triggers: new Set(timing ? interactiveSequences(timing).map(s => s.trigger) : []), ...(transition ? { transition } : {}),
    };
    this.cache.set(part, data);
    return data;
  }

  /** Draw a slide at a playback position into a 2D context the size of the show canvas. */
  private paint(ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, part: string, clicks: number, ms: number, now: number): void {
    const scene = this.session.scenes.get(part);
    this.session.images.request(scene.slide.images);
    const d = this.data(part);
    const w = this.canvas.width, h = this.canvas.height;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.fillStyle = '#000';
    ctx.fillRect(0, 0, w, h);
    const frame = d.timing ? sampleFrame(d.timing, scene.slide, {
      clicks, ms, triggers: this.triggers.map(t => ({ spid: t.spid, ms: now - t.start })),
    }, scene.resolved ? { colors: scene.resolved.context.colors } : {}) : undefined;
    drawSlide(ctx as CanvasRenderingContext2D, scene.slide, {
      scale: w / scene.slide.size.cx, image: p => this.session.images.get(p), createCanvas: (a, b) => new OffscreenCanvas(a, b),
      ...(frame ? { override: overrideLookup(frame, scene.slide) } : {}),
    });
  }

  private enter(now: number): void {
    this.clicks = 0;
    this.t0 = now;
    this.slideStart = now;
    this.triggers = [];
    this.onSlide(this.part);
    this.start();
  }

  /** Keep drawing while something moves. */
  start(): void {
    cancelAnimationFrame(this.raf);
    this.raf = requestAnimationFrame(t => this.tick(t));
  }

  stop(): void { cancelAnimationFrame(this.raf); }

  private tick(now: number): void {
    const ctx = this.canvas.getContext('2d')!;
    const part = this.part;
    if (!part) {
      // after the last slide, as PowerPoint ends a show
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.fillStyle = '#000';
      ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);
      ctx.fillStyle = '#ddd';
      ctx.font = `${Math.round(this.canvas.height / 40)}px sans-serif`;
      ctx.textAlign = 'center';
      ctx.fillText('End of slide show, click to exit.', this.canvas.width / 2, this.canvas.height / 12);
      return;
    }
    const r = this.running;
    if (r) {
      const p = r.spec.duration > 0 ? (now - r.start) / r.spec.duration : 1;
      if (p < 1) {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        drawTransition(ctx, r.from, r.to, r.spec, p, this.canvas.width, this.canvas.height);
        this.raf = requestAnimationFrame(t => this.tick(t));
        return;
      }
      this.running = undefined;
      this.enter(now);
      return;
    }
    const d = this.data(part);
    const ms = now - this.t0;
    this.paint(ctx, part, this.clicks, ms, now);
    const step = d.schedule?.main[this.clicks];
    const triggered = this.triggers.some(t => now - t.start < 60_000);
    const moving = (step && ms <= step.duration + 50) || triggered;
    const after = d.transition?.advanceAfter;
    if (after !== undefined && this.clicks >= d.steps && now - this.slideStart >= after && (!step || ms > step.duration)) {
      this.next(now);
      return;
    }
    if (moving || after !== undefined) this.raf = requestAnimationFrame(t => this.tick(t));
  }

  /** The next step: an animation, else the next slide (with its transition). */
  next(now = performance.now()): void {
    const part = this.part;
    if (!part) { this.exit(); return; }
    if (this.running) { this.running.start = -Infinity; this.start(); return; }     // a click finishes a transition
    const d = this.data(part);
    if (this.clicks < d.steps) {
      this.clicks++;
      this.t0 = now;
      this.start();
      return;
    }
    this.goTo(this.index + 1, now, true);
  }

  previous(now = performance.now()): void {
    if (this.clicks > 0) {
      this.clicks--;
      this.t0 = -1e9;                              // the end of that step
      this.start();
      return;
    }
    if (this.index === 0) return;
    this.index = Math.max(0, Math.min(this.parts.length, this.index) - 1);
    this.enter(now);
    const d = this.data(this.part!);
    this.clicks = d.steps;
    this.t0 = -1e9;
    this.start();
  }

  goTo(index: number, now = performance.now(), withTransition = false): void {
    const from = this.part;
    this.index = Math.max(0, Math.min(this.parts.length, index));
    const to = this.part;
    const spec = to ? this.data(to).transition : undefined;
    if (withTransition && from && to && spec && spec.kind !== 'none' && spec.duration > 0) {
      const w = this.canvas.width, h = this.canvas.height;
      const a = new OffscreenCanvas(w, h), b = new OffscreenCanvas(w, h);
      a.getContext('2d')!.drawImage(this.canvas, 0, 0);
      this.paint(b.getContext('2d')!, to, 0, 0, now);
      this.running = { from: a, to: b, spec, start: now };
      this.onSlide(to);
      this.start();
      return;
    }
    this.running = undefined;
    this.enter(now);
  }

  /** A click on the slide: a hyperlink, a trigger shape, or else the next step. */
  click(clientX: number, clientY: number): void {
    const part = this.part;
    if (!part) { this.exit(); return; }
    const r = this.canvas.getBoundingClientRect();
    const scene = this.session.scenes.get(part);
    const k = scene.slide.size.cx / r.width;
    const x = (clientX - r.left) * k, y = (clientY - r.top) * k;
    const d = this.data(part);
    const trigger = d.triggers.size ? shapeAt(scene.slide.shapes, x, y, s => d.triggers.has(s.id)) : undefined;
    if (trigger) {
      this.triggers.push({ spid: trigger.id, start: performance.now() });
      this.start();
      return;
    }
    const link = linkAt(scene.slide, x, y);
    if (link?.url) { if (isSafeUrl(link.url)) window.open(link.url, '_blank', 'noopener,noreferrer'); return; }
    if (link?.slide) { const at = this.parts.indexOf(link.slide); if (at >= 0) this.goTo(at); return; }
    const jump = /jump=(\w+)/.exec(link?.action ?? '')?.[1];
    if (jump === 'nextslide') this.goTo(this.index + 1, performance.now(), true);
    else if (jump === 'previousslide') this.goTo(this.index - 1);
    else if (jump === 'firstslide') this.goTo(0);
    else if (jump === 'lastslide') this.goTo(this.parts.length - 1);
    else if (jump === 'endshow') this.exit();
    else if (d.transition && !d.transition.advanceOnClick && this.clicks >= d.steps) return;
    else this.next();
  }
}

/** The topmost slide shape under a point that passes a test. */
function shapeAt(shapes: RenderShape[], x: number, y: number, ok: (s: RenderShape) => boolean): RenderShape | undefined {
  for (let i = shapes.length - 1; i >= 0; i--) {
    const s = shapes[i]!;
    const inner = shapeAt(s.children, x, y, ok);
    if (inner) return inner;
    if (!ok(s)) continue;
    const m = invert(s.transform);
    const u = m[0] * x + m[2] * y + m[4], v = m[1] * x + m[3] * y + m[5];
    if (u >= 0 && v >= 0 && u <= s.box.w && v <= s.box.h) return s;
  }
  return undefined;
}
