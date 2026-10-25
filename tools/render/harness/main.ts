// tools/render/harness/main.ts  -- the browser side of the render harness: open, resolve, build, draw, export PNG
import { chartRenderer } from '@pptx/charts';
import { metafileImage, PICTURE_TYPES } from '@pptx/metafile';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { BROWSER_IMAGE_TYPES, buildSlide, drawSlide, loadImages } from '@pptx/render';
import { FontRegistry } from '@pptx/text';

const FONT_FILES = [
  'carlito/Carlito-Regular.ttf', 'carlito/Carlito-Bold.ttf', 'carlito/Carlito-Italic.ttf', 'carlito/Carlito-BoldItalic.ttf',
  'caladea/Caladea-Regular.ttf', 'caladea/Caladea-Bold.ttf', 'caladea/Caladea-Italic.ttf', 'caladea/Caladea-BoldItalic.ttf',
  'liberation/LiberationSans-Regular.ttf', 'liberation/LiberationSans-Bold.ttf', 'liberation/LiberationSans-Italic.ttf', 'liberation/LiberationSans-BoldItalic.ttf',
  'liberation/LiberationSerif-Regular.ttf', 'liberation/LiberationSerif-Bold.ttf', 'liberation/LiberationSerif-Italic.ttf', 'liberation/LiberationSerif-BoldItalic.ttf',
  'liberation/LiberationMono-Regular.ttf', 'liberation/LiberationMono-Bold.ttf', 'liberation/LiberationMono-Italic.ttf', 'liberation/LiberationMono-BoldItalic.ttf',
];

let fontsPromise: Promise<FontRegistry> | undefined;
function fonts(extra: string[] = []): Promise<FontRegistry> {
  fontsPromise ??= (async () => {
    const reg = new FontRegistry();
    for (const f of [...FONT_FILES, ...extra]) {
      const res = await fetch(f.startsWith('/') ? f : `/fonts/${f}`);
      if (!res.ok) continue;
      const bytes = new Uint8Array(await res.arrayBuffer());
      reg.addBytes(bytes);
      // the same font for text inside pictures (EMF, WMF), which the canvas draws itself
      const stem = f.slice(f.lastIndexOf('/') + 1).replace(/\.ttf$/, '');
      const [name, variant = 'Regular'] = stem.split('-') as [string, string?];
      const family = name.replace(/^Liberation(Sans|Serif|Mono)$/, 'Liberation $1');
      document.fonts.add(new FontFace(family, bytes, { weight: /Bold/.test(variant) ? '700' : '400', style: /Italic/.test(variant) ? 'italic' : 'normal' }));
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
    // SVG through an <img>: scripts and external loads are disabled there (plan, security section)
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

export interface SlideResult { index: number; part: string; png?: string; error?: string; ms: number; }

async function renderDeck(base64: string, options: { width: number; slides?: number[]; fonts?: string[] }): Promise<SlideResult[]> {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const reg = await fonts(options.fonts);
  const doc = new PresentationDocument(OpcPackage.open(bytes));
  const cache = new Map<string, CanvasImageSource>();
  const out: SlideResult[] = [];
  for (const [index, part] of doc.slideParts().entries()) {
    if (options.slides && !options.slides.includes(index)) continue;
    const started = performance.now();
    try {
      const slide = buildSlide(resolveSlide(doc, part), { fonts: reg, chart: chartRenderer(reg) });
      await loadImages([slide], p => {
        const pt = doc.pkg.getPart(p);
        return pt ? { bytes: pt.bytes, contentType: pt.contentType } : undefined;
      }, decode, cache);
      const scale = options.width / slide.size.cx;
      const canvas = new OffscreenCanvas(options.width, Math.max(1, Math.round(slide.size.cy * scale)));
      const ctx = canvas.getContext('2d')!;
      drawSlide(ctx, slide, { scale, image: p => cache.get(p), createCanvas: (w, h) => new OffscreenCanvas(w, h) });
      const blob = await canvas.convertToBlob({ type: 'image/png' });
      out.push({ index, part, png: toBase64(new Uint8Array(await blob.arrayBuffer())), ms: performance.now() - started });
    } catch (e) {
      out.push({ index, part, error: e instanceof Error ? `${e.message}\n${e.stack ?? ''}` : String(e), ms: performance.now() - started });
    }
  }
  return out;
}

(globalThis as unknown as { renderDeck: typeof renderDeck }).renderDeck = renderDeck;
(globalThis as unknown as { harnessReady: boolean }).harnessReady = true;
