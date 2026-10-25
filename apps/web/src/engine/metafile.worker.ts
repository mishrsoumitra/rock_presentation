// apps/web/src/engine/metafile.worker.ts  -- EMF and WMF pictures, played off the main thread
//
// The plan rasterises metafiles in a sandbox, away from the editor. Our player is memory-safe TypeScript with hard
// limits, and it runs here, in a worker the page terminates if it overruns its deadline: a hostile picture can
// cost a grey box, never a frozen editor.
import { metafileImage } from '@pptx/metafile';
import { registerCssFonts } from './fonts.ts';

self.onmessage = async (e: MessageEvent<{ id: number; bytes: Uint8Array }>) => {
  let bitmap: ImageBitmap | undefined;
  try {
    // text in pictures is drawn by the canvas: with Carlito for Calibri and so on, as the slides are
    await registerCssFonts();
    const canvas = await metafileImage(e.data.bytes, (w, h) => new OffscreenCanvas(w, h));
    bitmap = canvas?.transferToImageBitmap();
  } catch {
    bitmap = undefined;
  }
  (self as unknown as Worker).postMessage({ id: e.data.id, bitmap }, bitmap ? [bitmap] : []);
};
