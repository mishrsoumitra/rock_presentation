// tools/render/anim-harness/main.ts  -- the browser side of the filmstrip tool: frames of animations and transitions
//
// Same pipeline as the render harness (open, resolve, build, draw), plus @pptx/anim: each frame is the slide drawn
// with the overrides the sampler gives at one playback position, and transition frames are drawTransition between
// the previous slide (all its effects played) and this one (none played yet). @pptx/anim is not linked into
// tools/render's node_modules, so it is imported by path.
import { chartRenderer } from '@pptx/charts';
import { metafileImage, PICTURE_TYPES } from '@pptx/metafile';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide, slideContext } from '@pptx/resolve';
import { BROWSER_IMAGE_TYPES, buildSlide, drawSlide, loadImages, type RenderSlide } from '@pptx/render';
import { FontRegistry } from '@pptx/text';
import {
  drawTransition, overrideLookup, parseTiming, sampleFrame, scheduleFor, slideTransition, transitionName, type PlaybackPosition, type Timing,
  type TransitionKind,
} from '../../../packages/anim/src/index.ts';

const FONT_FILES = [
  'carlito/Carlito-Regular.ttf', 'carlito/Carlito-Bold.ttf', 'carlito/Carlito-Italic.ttf', 'carlito/Carlito-BoldItalic.ttf',
  'caladea/Caladea-Regular.ttf', 'caladea/Caladea-Bold.ttf',
  'liberation/LiberationSans-Regular.ttf', 'liberation/LiberationSans-Bold.ttf', 'liberation/LiberationSans-Italic.ttf',
  'liberation/LiberationSerif-Regular.ttf', 'liberation/LiberationSerif-Bold.ttf', 'liberation/LiberationMono-Regular.ttf',
];

/** Register a bundled font as a CSS font too: text inside EMF/WMF pictures is drawn by the canvas itself. */
function addCssFont(file: string, bytes: Uint8Array): void {
  const stem = file.slice(file.lastIndexOf('/') + 1).replace(/\.ttf$/, '');
  const [name, variant = 'Regular'] = stem.split('-') as [string, string?];
  const family = name.replace(/^Liberation(Sans|Serif|Mono)$/, 'Liberation $1');
  document.fonts.add(new FontFace(family, bytes as Uint8Array<ArrayBuffer>, { weight: /Bold/.test(variant) ? '700' : '400', style: /Italic/.test(variant) ? 'italic' : 'normal' }));
}

let fontsPromise: Promise<FontRegistry> | undefined;
function fonts(): Promise<FontRegistry> {
  fontsPromise ??= (async () => {
    const reg = new FontRegistry();
    for (const f of FONT_FILES) {
      const res = await fetch(`/fonts/${f}`);
      if (!res.ok) continue;
      const bytes = new Uint8Array(await res.arrayBuffer());
      reg.addBytes(bytes);
      addCssFont(f, bytes);
    }
    return reg;
  })();
  return fontsPromise;
}

async function decode(bytes: Uint8Array, type: string): Promise<CanvasImageSource | undefined> {
  if (PICTURE_TYPES.has(type)) return metafileImage(bytes, (w, h) => new OffscreenCanvas(w, h));
  if (!BROWSER_IMAGE_TYPES.has(type)) return undefined;
  const blob = new Blob([bytes as BlobPart], { type });
  if (type === 'image/svg+xml') {
    const img = new Image();
    img.src = URL.createObjectURL(blob);
    await img.decode();
    return img;
  }
  return createImageBitmap(blob);
}

function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export interface StripResult { index: number; part: string; png?: string; error?: string; steps: number; transition?: string; }

interface Prepared { slide: RenderSlide; timing?: Timing; colors: ReturnType<typeof slideContext>['colors']; }

interface StripOptions { width: number; frames: number; slides?: number[]; all?: boolean; }

async function renderDeckStrips(base64: string, options: StripOptions): Promise<StripResult[]> {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const reg = await fonts();
  const doc = new PresentationDocument(OpcPackage.open(bytes));
  const cache = new Map<string, CanvasImageSource>();
  const parts = doc.slideParts();
  const prepared = new Map<number, Prepared>();
  const prepare = async (index: number): Promise<Prepared> => {
    let p = prepared.get(index);
    if (p) return p;
    const part = parts[index]!;
    const slide = buildSlide(resolveSlide(doc, part), { fonts: reg, chart: chartRenderer(reg) });
    await loadImages([slide], name => {
      const pt = doc.pkg.getPart(name);
      return pt ? { bytes: pt.bytes, contentType: pt.contentType } : undefined;
    }, decode, cache);
    const el = doc.slide(part).timing;
    p = { slide, colors: slideContext(doc, part).colors, ...(el ? { timing: parseTiming(el) } : {}) };
    prepared.set(index, p);
    return p;
  };
  const W = options.width;
  const frame = (p: Prepared, at: PlaybackPosition | undefined): OffscreenCanvas => {
    const scale = W / p.slide.size.cx;
    const canvas = new OffscreenCanvas(W, Math.max(1, Math.round(p.slide.size.cy * scale)));
    const ctx = canvas.getContext('2d')!;
    const frameOverrides = at ? sampleFrame(p.timing, p.slide, at, { colors: p.colors }) : undefined;
    drawSlide(ctx, p.slide, {
      scale, image: n => cache.get(n), createCanvas: (w, h) => new OffscreenCanvas(w, h),
      ...(frameOverrides ? { override: overrideLookup(frameOverrides, p.slide) } : {}),
    });
    return canvas;
  };
  const out: StripResult[] = [];
  for (const [index, part] of parts.entries()) {
    if (options.slides && !options.slides.includes(index)) continue;
    try {
      const view = doc.slide(part);
      const transition = slideTransition(view);
      const hasTransition = !!transition && transition.kind !== 'none';
      if (!options.all && !view.timing && !hasTransition) continue;
      const p = await prepare(index);
      const H = Math.max(1, Math.round(p.slide.size.cy * (W / p.slide.size.cx)));
      const rows: { label: string; frames: { canvas: OffscreenCanvas; label: string }[] }[] = [];
      const n = Math.max(2, options.frames);
      const fractions = Array.from({ length: n }, (_, i) => i / (n - 1));
      if (hasTransition) {
        const prev = index > 0 ? await prepare(index - 1) : undefined;
        const from = prev ? frame(prev, { clicks: 1e6, ms: 0 }) : (() => {
          const c = new OffscreenCanvas(W, H);
          const g = c.getContext('2d')!;
          g.fillStyle = 'black';
          g.fillRect(0, 0, W, H);
          return c;
        })();
        const to = frame(p, { clicks: 0, ms: 0 });
        rows.push({
          label: `${transitionName(transition)} (${transition.kind}${transition.dir ? ` ${transition.dir}` : ''}) ${transition.duration} ms`,
          frames: fractions.map(f => {
            const c = new OffscreenCanvas(W, H);
            drawTransition(c.getContext('2d')!, from, to, transition, f, W, H);
            return { canvas: c, label: `${Math.round(f * 100)}%` };
          }),
        });
      }
      let steps = 0;
      if (p.timing) {
        const schedule = scheduleFor(p.timing, p.slide);
        steps = schedule.main.length;
        schedule.main.forEach((act, k) => {
          const d = Number.isFinite(act.duration) ? act.duration : 3000;
          const times = d > 0 ? fractions.map(f => f * d) : [0];
          rows.push({
            label: k === 0 ? `start, ${Math.round(act.duration)} ms` : `click ${k}, ${Math.round(act.duration)} ms`,
            frames: times.map(t => ({ canvas: frame(p, { clicks: k, ms: t }), label: `${Math.round(t)} ms` })),
          });
        });
        for (const seq of schedule.interactive) {
          seq.steps.forEach((act, k) => {
            const d = Number.isFinite(act.duration) ? act.duration : 3000;
            const times = d > 0 ? fractions.map(f => f * d) : [0];
            const clicks = Array.from({ length: k }, () => ({ spid: seq.trigger, ms: 1e6 }));
            rows.push({
              label: `trigger ${seq.trigger} #${k + 1}, ${Math.round(act.duration)} ms`,
              frames: times.map(t => ({
                canvas: frame(p, { clicks: 1e6, ms: 0, triggers: [...clicks, { spid: seq.trigger, ms: t }] }),
                label: `${Math.round(t)} ms`,
              })),
            });
          });
        }
      }
      if (!rows.length) rows.push({ label: 'static', frames: [{ canvas: frame(p, undefined), label: '' }] });
      const labelH = 16;
      const cols = Math.max(...rows.map(r => r.frames.length));
      const sheet = new OffscreenCanvas(cols * (W + 6) + 6, rows.length * (H + labelH * 2 + 6) + 6);
      const g = sheet.getContext('2d')!;
      g.fillStyle = '#d8d8d8';
      g.fillRect(0, 0, sheet.width, sheet.height);
      g.font = '12px sans-serif';
      g.textBaseline = 'top';
      rows.forEach((row, r) => {
        const y = 6 + r * (H + labelH * 2 + 6);
        g.fillStyle = '#000';
        g.fillText(row.label, 6, y);
        row.frames.forEach((f, c) => {
          const x = 6 + c * (W + 6);
          g.drawImage(f.canvas, x, y + labelH);
          g.strokeStyle = '#888';
          g.strokeRect(x - 0.5, y + labelH - 0.5, W + 1, H + 1);
          g.fillStyle = '#333';
          g.fillText(f.label, x, y + labelH + H + 2);
        });
      });
      const blob = await sheet.convertToBlob({ type: 'image/png' });
      const png = toBase64(new Uint8Array(await blob.arrayBuffer()));
      out.push({ index, part, png, steps, ...(hasTransition ? { transition: transition.kind } : {}) });
    } catch (e) {
      out.push({ index, part, steps: 0, error: e instanceof Error ? `${e.message}\n${e.stack ?? ''}` : String(e) });
    }
  }
  return out;
}

/** Two test cards, A and B, so every part of a transition shows which slide it comes from. */
function card(label: string, bg: string, fg: string, w: number, h: number): OffscreenCanvas {
  const c = new OffscreenCanvas(w, h);
  const g = c.getContext('2d')!;
  g.fillStyle = bg;
  g.fillRect(0, 0, w, h);
  g.strokeStyle = fg;
  g.lineWidth = 2;
  for (let x = 0; x <= w; x += w / 8) g.strokeRect(x, 0, 0.1, h);
  for (let y = 0; y <= h; y += h / 4.5) g.strokeRect(0, y, w, 0.1);
  g.fillStyle = fg;
  g.font = `bold ${Math.round(h * 0.5)}px sans-serif`;
  g.textAlign = 'center';
  g.textBaseline = 'middle';
  g.fillText(label, w / 2, h / 2);
  g.fillRect(0, 0, w * 0.15, h * 0.15);
  return c;
}

const GALLERY: { kind: TransitionKind; dir?: string; label?: string; [k: string]: unknown }[] = [
  { kind: 'fade' }, { kind: 'fade', thruBlk: true, label: 'fade thruBlk' },
  { kind: 'cut', thruBlk: true, label: 'cut thruBlk' }, { kind: 'push', dir: 'u' }, { kind: 'wipe', dir: 'r' },
  { kind: 'split', orient: 'vert', dir: 'out' }, { kind: 'cover', dir: 'ld' }, { kind: 'pull', dir: 'l' },
  { kind: 'randomBar', dir: 'horz' }, { kind: 'blinds', dir: 'vert' }, { kind: 'checker', dir: 'horz' },
  { kind: 'circle' }, { kind: 'diamond' }, { kind: 'plus' }, { kind: 'wedge' }, { kind: 'wheel', spokes: 4 },
  { kind: 'wheelReverse', spokes: 1 }, { kind: 'zoom', dir: 'in' }, { kind: 'newsflash' }, { kind: 'dissolve' },
  { kind: 'strips', dir: 'rd' }, { kind: 'comb', dir: 'horz' }, { kind: 'vortex', dir: 'l' }, { kind: 'ripple' },
  { kind: 'glitter', dir: 'l' }, { kind: 'honeycomb' }, { kind: 'flash' }, { kind: 'shred' },
  { kind: 'switch', dir: 'r' }, { kind: 'flip', dir: 'r' }, { kind: 'gallery', dir: 'l' },
  { kind: 'prism', dir: 'l', label: 'prism (Cube)' },
  { kind: 'prism', dir: 'u', isInverted: true, label: 'prism (Box)' }, { kind: 'doors', dir: 'vert' },
  { kind: 'window', dir: 'horz' }, { kind: 'conveyor', dir: 'l' }, { kind: 'pan', dir: 'u' },
  { kind: 'ferris', dir: 'l' }, { kind: 'warp', dir: 'in' }, { kind: 'flythrough', dir: 'in' },
  { kind: 'reveal', dir: 'l' }, { kind: 'curtains' }, { kind: 'drape' }, { kind: 'wind' }, { kind: 'prestige' },
  { kind: 'fracture' }, { kind: 'crush' }, { kind: 'peelOff' }, { kind: 'pageCurlDouble' },
  { kind: 'pageCurlSingle' }, { kind: 'airplane' }, { kind: 'origami' }, { kind: 'fallOver' }, { kind: 'morph' },
  { kind: 'unknown' },
];

/** Every transition kind between two test cards, at a few moments, as one PNG. */
async function transitionGallery(options: { width: number; frames: number }): Promise<string> {
  const W = options.width;
  const H = Math.round((W * 9) / 16);
  const a = card('A', '#f4c430', '#7a3b00', W, H);
  const b = card('B', '#2d6cdf', '#ffffff', W, H);
  const n = Math.max(3, options.frames);
  const fractions = Array.from({ length: n }, (_, i) => (i + 1) / (n + 1));
  const labelH = 16;
  const sheet = new OffscreenCanvas(n * (W + 6) + 6, GALLERY.length * (H + labelH + 6) + 6);
  const g = sheet.getContext('2d')!;
  g.fillStyle = '#d8d8d8';
  g.fillRect(0, 0, sheet.width, sheet.height);
  g.font = '12px sans-serif';
  g.textBaseline = 'top';
  GALLERY.forEach((t, r) => {
    const y = 6 + r * (H + labelH + 6);
    g.fillStyle = '#000';
    g.fillText(`${t.label ?? `${t.kind}${t.dir ? ` ${t.dir}` : ''}`}  (${transitionName(t)})`, 6, y);
    fractions.forEach((f, c) => {
      const x = 6 + c * (W + 6);
      g.save();
      g.translate(x, y + labelH);
      g.beginPath();
      g.rect(0, 0, W, H);
      g.clip();
      drawTransition(g as unknown as Parameters<typeof drawTransition>[0], a, b, t, f, W, H);
      g.restore();
    });
  });
  const blob = await sheet.convertToBlob({ type: 'image/png' });
  return toBase64(new Uint8Array(await blob.arrayBuffer()));
}

(globalThis as unknown as { transitionGallery: typeof transitionGallery }).transitionGallery = transitionGallery;
(globalThis as unknown as { renderDeckStrips: typeof renderDeckStrips }).renderDeckStrips = renderDeckStrips;
(globalThis as unknown as { harnessReady: boolean }).harnessReady = true;
