// apps/web/src/engine/thumbs.render.ts  -- the thumbnail worker's work: its copy of the deck, and drawing slides
import { chartRenderer } from '@pptx/charts';
import { PresentationDocument } from '@pptx/model';
import { OpcPackage } from '@pptx/opc';
import { buildSlide, drawSlide } from '@pptx/render';
import { resolveSlide } from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';
import { loadFonts, registerCssFonts } from './fonts.ts';
import { decodeImage } from './images.ts';

export type ThumbMessage =
  | { type: 'init'; bytes: Uint8Array }
  | { type: 'update'; removed: string[]; parts: { name: string; contentType: string; bytes: Uint8Array }[] }
  | { type: 'render'; id: number; part: string; width: number };

let pkg: OpcPackage | undefined;
let doc: PresentationDocument | undefined;
let fonts: FontRegistry | undefined;
const images = new Map<string, { bytes: Uint8Array; image?: CanvasImageSource }>();
let queue: Promise<void> = Promise.resolve();

const post = (msg: unknown, transfer: Transferable[] = []) => (self as unknown as Worker).postMessage(msg, transfer);

async function image(part: string): Promise<CanvasImageSource | undefined> {
  const p = pkg!.getPart(part);
  if (!p) return undefined;
  const hit = images.get(part);
  if (hit && hit.bytes === p.bytes) return hit.image;
  // SVG cannot be decoded in a worker; the renderer falls back to the PNG PowerPoint stores beside it
  const img = await decodeImage(p.bytes, p.contentType, false).catch(() => undefined);
  images.set(part, { bytes: p.bytes, ...(img ? { image: img } : {}) });
  return img;
}

async function render(id: number, part: string, width: number): Promise<void> {
  if (!pkg || !doc) return;
  fonts ??= (await Promise.all([loadFonts(), registerCssFonts()]))[0];
  try {
    const slide = buildSlide(resolveSlide(doc, part), { fonts, chart: chartRenderer(fonts) });
    const loaded = new Map<string, CanvasImageSource>();
    for (const p of slide.images) {
      const img = await image(p);
      if (img) loaded.set(p, img);
    }
    const scale = width / slide.size.cx;
    const canvas = new OffscreenCanvas(width, Math.max(1, Math.round(slide.size.cy * scale)));
    drawSlide(canvas.getContext('2d')!, slide, { scale, image: p => loaded.get(p), createCanvas: (w, h) => new OffscreenCanvas(w, h) });
    const bitmap = canvas.transferToImageBitmap();
    post({ type: 'thumb', id, bitmap }, [bitmap]);
  } catch (e) {
    post({ type: 'thumb', id, error: e instanceof Error ? e.message : String(e) });
  }
}

/** Handle one message; messages are handled in order, so an update lands before the renders that follow it. */
export function handle(e: MessageEvent<ThumbMessage>): void {
  const m = e.data;
  queue = queue.then(async () => {
    if (m.type === 'init') {
      pkg = OpcPackage.open(m.bytes);
      doc = new PresentationDocument(pkg);
      images.clear();
    } else if (m.type === 'update' && pkg && doc) {
      for (const name of m.removed) if (pkg.getPart(name)) pkg.removePart(name, { force: true });
      for (const p of m.parts) {
        if (pkg.getPart(p.name)) pkg.setPartBytes(p.name, p.bytes);
        else pkg.addPart(p.name, p.contentType, p.bytes);
        doc.invalidate(p.name);
      }
      doc = new PresentationDocument(pkg);
    } else if (m.type === 'render') {
      await render(m.id, m.part, m.width);
    }
  }).catch(() => undefined);
}
