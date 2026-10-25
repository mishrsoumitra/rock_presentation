// apps/server/render-page/main.ts  -- the page the server's Chromium draws slides in: one PNG per slide, or a print layout
//
// Modelled on tools/render/harness/main.ts. The server hands this page, the fonts and the deck bytes out from one
// made-up origin through Playwright's request interception and refuses every other request, so nothing in a deck
// can reach the network. Decks arrive already checked by the server's worker; they are opened with the same limits.
import { chartRenderer } from '@pptx/charts';
import { metafileImage, PICTURE_TYPES } from '@pptx/metafile';
import { OpcPackage, type PackageLimits } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { BROWSER_IMAGE_TYPES, buildSlide, drawSlide, loadImages } from '@pptx/render';
import { FontRegistry } from '@pptx/text';

const EMU_PER_INCH = 914_400;
/** Decoded pictures kept for the open deck; past this, those the next slide does not use are released. */
const IMAGE_CACHE_ENTRIES = 48;

/** Which deck and how to draw it. `key` is the deck's content hash: a deck already open is not fetched again. */
export interface DeckRequest { url: string; key: string; limits: Partial<PackageLimits>; width: number; maxPixels: number; }
export interface SlideImage { index: number; png?: string; error?: string; }
export interface RenderResult { slideCount: number; size: { cx: number; cy: number }; slides: SlideImage[]; }
export interface PrintResult { slideCount: number; size: { cx: number; cy: number }; errors: { index: number; error: string }[]; }
export interface RenderPageApi {
  renderReady?: boolean;
  renderSlides(req: DeckRequest & { slides: number[] }): Promise<RenderResult>;
  printDeck(req: DeckRequest): Promise<PrintResult>;
  clearPrint(): void;
}

/** Register a bundled font as a CSS font too: text inside EMF/WMF pictures is drawn by the canvas itself. */
function addCssFont(file: string, bytes: Uint8Array): void {
  const stem = file.slice(file.lastIndexOf('/') + 1).replace(/\.ttf$/, '');
  const [name, variant = 'Regular'] = stem.split('-') as [string, string?];
  const family = name.replace(/^Liberation(Sans|Serif|Mono)$/, 'Liberation $1');
  document.fonts.add(new FontFace(family, bytes as Uint8Array<ArrayBuffer>, { weight: /Bold/.test(variant) ? '700' : '400', style: /Italic/.test(variant) ? 'italic' : 'normal' }));
}

let fontsPromise: Promise<FontRegistry> | undefined;
/** Every face under assets/fonts, in the order of the server's /fonts/manifest.json (Carlito, the default, first). */
function fonts(): Promise<FontRegistry> {
  fontsPromise ??= (async () => {
    const reg = new FontRegistry();
    const res = await fetch('/fonts/manifest.json');
    const files = res.ok ? await res.json() as string[] : [];
    const data = await Promise.all(files.map(async f => {
      const r = await fetch(`/fonts/${f}`);
      return r.ok ? new Uint8Array(await r.arrayBuffer()) : undefined;
    }));
    data.forEach((d, i) => {
      if (!d) return;
      reg.addBytes(d);
      addCssFont(files[i]!, d);
    });
    return reg;
  })();
  return fontsPromise;
}

async function decode(bytes: Uint8Array, type: string): Promise<CanvasImageSource | undefined> {
  if (PICTURE_TYPES.has(type)) return metafileImage(bytes, (w, h) => new OffscreenCanvas(w, h));
  if (!BROWSER_IMAGE_TYPES.has(type)) return undefined;
  const blob = new Blob([bytes as BlobPart], { type });
  if (type === 'image/svg+xml') {
    // SVG only through an <img>: no scripts run and nothing external loads there (plan, security section)
    const img = new Image();
    const url = URL.createObjectURL(blob);
    try {
      img.src = url;
      await img.decode();
    } finally {
      URL.revokeObjectURL(url);
    }
    return img;
  }
  return createImageBitmap(blob);
}

const message = (e: unknown) => (e instanceof Error ? e.message : String(e)).split('\n')[0]!.slice(0, 300);

function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

interface OpenDeck { key: string; doc: PresentationDocument; images: Map<string, CanvasImageSource>; failed: Set<string>; }
let open: OpenDeck | undefined;

function release(images: Map<string, CanvasImageSource>, keep: Set<string> = new Set()): void {
  for (const [part, img] of images) {
    if (keep.has(part)) continue;
    if (img instanceof ImageBitmap) img.close();
    images.delete(part);
  }
}

async function openDeck(req: DeckRequest): Promise<OpenDeck> {
  if (open?.key === req.key) return open;
  if (open) release(open.images);
  open = undefined;
  const res = await fetch(req.url);
  if (!res.ok) throw new Error(`the deck could not be loaded (${res.status})`);
  const doc = new PresentationDocument(OpcPackage.open(new Uint8Array(await res.arrayBuffer()), { limits: req.limits }));
  open = { key: req.key, doc, images: new Map(), failed: new Set() };
  return open;
}

/** The slide size in EMU; a size that is not positive cannot be drawn. */
function slideSize(doc: PresentationDocument): { cx: number; cy: number } {
  const { cx, cy } = doc.presentation.slideSize;
  if (!(cx > 0 && cy > 0 && Number.isFinite(cx) && Number.isFinite(cy))) throw new Error(`the slide size ${cx} x ${cy} EMU cannot be drawn`);
  return { cx, cy };
}

/** Draw one slide `width` pixels wide. Refuses images over `maxPixels`, which an extreme aspect ratio could ask for. */
async function drawOne(deck: OpenDeck, part: string, width: number, maxPixels: number): Promise<OffscreenCanvas> {
  const reg = await fonts();
  const slide = buildSlide(resolveSlide(deck.doc, part), { fonts: reg, chart: chartRenderer(reg) });
  if (deck.images.size > IMAGE_CACHE_ENTRIES) release(deck.images, slide.images);
  await loadImages([slide], p => {
    const pt = deck.doc.pkg.getPart(p);
    return pt ? { bytes: pt.bytes, contentType: pt.contentType } : undefined;
  }, decode, deck.images, deck.failed);
  const { cx, cy } = slide.size;
  if (!(cx > 0 && cy > 0)) throw new Error(`the slide size ${cx} x ${cy} EMU cannot be drawn`);
  const scale = width / cx;
  const height = Math.max(1, Math.round(cy * scale));
  if (width * height > maxPixels) throw new Error(`a ${width} x ${height} image is over the ${maxPixels}-pixel limit`);
  const canvas = new OffscreenCanvas(width, height);
  drawSlide(canvas.getContext('2d')!, slide, { scale, image: p => deck.images.get(p), createCanvas: (w, h) => new OffscreenCanvas(w, h) });
  return canvas;
}

/** PNGs of the given slides (0-based). A slide that fails reports its error; the others are still drawn. */
async function renderSlides(req: DeckRequest & { slides: number[] }): Promise<RenderResult> {
  const deck = await openDeck(req);
  const parts = deck.doc.slideParts();
  const slides: SlideImage[] = [];
  for (const index of req.slides) {
    const part = parts[index];
    if (!part) {
      slides.push({ index, error: `there is no slide ${index + 1}` });
      continue;
    }
    try {
      const blob = await (await drawOne(deck, part, req.width, req.maxPixels)).convertToBlob({ type: 'image/png' });
      slides.push({ index, png: toBase64(new Uint8Array(await blob.arrayBuffer())) });
    } catch (e) {
      slides.push({ index, error: message(e) });
    }
  }
  return { slideCount: parts.length, size: slideSize(deck.doc), slides };
}

const printUrls: string[] = [];

/**
 * Lay every slide out as one image per printed page, the page sized to the slide (@page). The server then calls
 * page.pdf(), so the PDF is raster: one PNG per page. A slide that fails to draw leaves its page blank, so the
 * page count always equals the slide count.
 */
async function printDeck(req: DeckRequest): Promise<PrintResult> {
  clearPrint();
  const deck = await openDeck(req);
  const parts = deck.doc.slideParts();
  const size = slideSize(deck.doc);
  const w = size.cx / EMU_PER_INCH;
  const h = size.cy / EMU_PER_INCH;
  // a constructed sheet, not a <style> element: the page's CSP allows no inline styles
  const sheet = new CSSStyleSheet();
  sheet.replaceSync(`@page { size: ${w}in ${h}in; margin: 0 }
    html, body { margin: 0; padding: 0; background: #fff }
    .slide { width: ${w}in; height: ${h}in; overflow: hidden; break-after: page }
    .slide:last-child { break-after: auto }
    .slide img { display: block; width: 100%; height: 100% }`);
  document.adoptedStyleSheets = [sheet];
  const errors: { index: number; error: string }[] = [];
  for (const [index, part] of parts.entries()) {
    const page = document.createElement('div');
    page.className = 'slide';
    document.body.append(page);
    try {
      const blob = await (await drawOne(deck, part, req.width, req.maxPixels)).convertToBlob({ type: 'image/png' });
      const url = URL.createObjectURL(blob);
      printUrls.push(url);
      const img = document.createElement('img');
      img.alt = '';
      img.src = url;
      page.append(img);
      await img.decode();
    } catch (e) {
      errors.push({ index, error: message(e) });
    }
  }
  return { slideCount: parts.length, size, errors };
}

function clearPrint(): void {
  for (const url of printUrls.splice(0)) URL.revokeObjectURL(url);
  for (const el of document.querySelectorAll('.slide')) el.remove();
  document.adoptedStyleSheets = [];
}

const api = globalThis as unknown as RenderPageApi;
api.renderSlides = renderSlides;
api.printDeck = printDeck;
api.clearPrint = clearPrint;
await fonts();
api.renderReady = true;
