// apps/web/src/engine/export.ts  -- slides as pictures: PNG files and a PDF, made in the browser
import { writeZip } from '@pptx/opc';
import { drawSlide } from '@pptx/render';
import type { Session } from '../state/store.ts';
import { pdfOfImages, type PdfPage } from './pdf.ts';

/** How wide an exported slide is, in pixels: full HD for a 16:9 slide. */
export const EXPORT_WIDTH = 1920;

/** Wait until the pictures a slide shows have been decoded (or given up on), so the export is not drawn without them. */
async function picturesOf(session: Session, part: string): Promise<void> {
  const images = session.scenes.get(part).slide.images;
  session.images.request(images);
  for (let waited = 0; !session.images.ready(images) && waited < 20_000; waited += 50) await new Promise(done => setTimeout(done, 50));
}

/** Draw one slide on a canvas `width` pixels wide, on white (a slide with no background would otherwise be clear). */
export async function slideCanvas(session: Session, part: string, width = EXPORT_WIDTH): Promise<OffscreenCanvas> {
  await picturesOf(session, part);
  const size = session.ed.doc.presentation.slideSize;
  const scale = width / size.cx;
  const canvas = new OffscreenCanvas(Math.round(width), Math.max(1, Math.round(size.cy * scale)));
  const c = canvas.getContext('2d')!;
  c.fillStyle = '#fff';
  c.fillRect(0, 0, canvas.width, canvas.height);
  drawSlide(c as unknown as CanvasRenderingContext2D, session.scenes.get(part).slide, {
    scale, image: p => session.images.get(p), createCanvas: (w, h) => new OffscreenCanvas(w, h),
  });
  return canvas;
}

const bytesOf = async (blob: Blob) => new Uint8Array(await blob.arrayBuffer());

/** One slide as a PNG file. */
export async function slidePng(session: Session, part: string, width = EXPORT_WIDTH): Promise<Uint8Array> {
  return bytesOf(await (await slideCanvas(session, part, width)).convertToBlob({ type: 'image/png' }));
}

/** Every slide as a PNG, in a ZIP: "Slide 1.png", "Slide 2.png"... (numbered so they sort in order). */
export async function slidesZip(session: Session, onProgress: (done: number, of: number) => void = () => undefined): Promise<Uint8Array> {
  const parts = session.ed.doc.slideParts();
  const digits = String(parts.length).length;
  const entries: { name: string; data: Uint8Array; store: boolean }[] = [];
  for (const [i, part] of parts.entries()) {
    entries.push({ name: `Slide ${String(i + 1).padStart(digits, '0')}.png`, data: await slidePng(session, part), store: true });
    onProgress(i + 1, parts.length);
  }
  return writeZip(entries);
}

/**
 * The deck as a PDF: a page per slide, each a picture of the slide at the slide's own size (a 13.33 by 7.5 inch
 * slide is a 960 by 540 point page). Hidden slides are left out, as when printing.
 */
export async function deckPdf(session: Session, onProgress: (done: number, of: number) => void = () => undefined): Promise<Uint8Array> {
  const doc = session.ed.doc;
  const parts = doc.slideParts().filter(p => !doc.slide(p).hidden);
  const size = doc.presentation.slideSize;
  const pages: PdfPage[] = [];
  for (const [i, part] of parts.entries()) {
    const canvas = await slideCanvas(session, part);
    const jpeg = await bytesOf(await canvas.convertToBlob({ type: 'image/jpeg', quality: 0.92 }));
    pages.push({ jpeg, width: canvas.width, height: canvas.height, widthPt: size.cx / 12700, heightPt: size.cy / 12700 });
    onProgress(i + 1, parts.length);
  }
  return pdfOfImages(pages, session.name.replace(/\.[^.]+$/, ''));
}
