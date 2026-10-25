// tools/render/metafile-harness/main.ts  -- the browser side of the metafile harness: parse, draw, export PNG
import { canvasSize, decodeImages, drawMetafile, hasDrawing, metafileImage, parseMetafile } from '@pptx/metafile';

export interface MetafileResult { png?: string; width?: number; height?: number; error?: string; info?: string; ms: number; }

export interface MetafileOptions {
  /** Which records to play in EMF files with EMF+ records; 'auto' goes through metafileImage like the renderer. */
  plus?: 'auto' | 'emf' | 'emf+';
  /** CSS colour under the picture, or 'checker' for a checkerboard that shows transparency. */
  background?: string;
}

// Metric-compatible stand-ins for the Office fonts metafile text names, as CSS fonts for canvas text.
const FONTS: [string, string, string, string][] = [
  ['Carlito', 'carlito/Carlito-Regular.ttf', 'normal', '400'], ['Carlito', 'carlito/Carlito-Bold.ttf', 'normal', '700'],
  ['Carlito', 'carlito/Carlito-Italic.ttf', 'italic', '400'], ['Carlito', 'carlito/Carlito-BoldItalic.ttf', 'italic', '700'],
  ['Caladea', 'caladea/Caladea-Regular.ttf', 'normal', '400'], ['Caladea', 'caladea/Caladea-Bold.ttf', 'normal', '700'],
  ['Caladea', 'caladea/Caladea-Italic.ttf', 'italic', '400'], ['Caladea', 'caladea/Caladea-BoldItalic.ttf', 'italic', '700'],
  ['Liberation Sans', 'liberation/LiberationSans-Regular.ttf', 'normal', '400'], ['Liberation Sans', 'liberation/LiberationSans-Bold.ttf', 'normal', '700'],
  ['Liberation Sans', 'liberation/LiberationSans-Italic.ttf', 'italic', '400'], ['Liberation Sans', 'liberation/LiberationSans-BoldItalic.ttf', 'italic', '700'],
  ['Liberation Serif', 'liberation/LiberationSerif-Regular.ttf', 'normal', '400'], ['Liberation Serif', 'liberation/LiberationSerif-Bold.ttf', 'normal', '700'],
  ['Liberation Serif', 'liberation/LiberationSerif-Italic.ttf', 'italic', '400'], ['Liberation Serif', 'liberation/LiberationSerif-BoldItalic.ttf', 'italic', '700'],
  ['Liberation Mono', 'liberation/LiberationMono-Regular.ttf', 'normal', '400'], ['Liberation Mono', 'liberation/LiberationMono-Bold.ttf', 'normal', '700'],
];

async function loadFonts(): Promise<void> {
  await Promise.all(FONTS.map(async ([family, file, style, weight]) => {
    const face = new FontFace(family, `url(/fonts/${file})`, { style, weight });
    try {
      document.fonts.add(await face.load());
    } catch {
      // a missing font file only means a different stand-in
    }
  }));
}

function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

function fromBase64(b64: string): Uint8Array {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function renderMetafile(base64: string, options: MetafileOptions = {}): Promise<MetafileResult> {
  const bytes = fromBase64(base64);
  const started = performance.now();
  try {
    let canvas: OffscreenCanvas | undefined;
    let info = '';
    const mode = options.plus ?? 'auto';
    try {
      const mf = parseMetafile(bytes);
      info = `${mf.format} plus=${mf.plus} records=${mf.records.length} skipped=${mf.skipped} natural=${mf.width.toFixed(0)}x${mf.height.toFixed(0)} drawing=${hasDrawing(mf)}`;
      if (mode !== 'auto') {
        const { width, height } = canvasSize(mf);
        canvas = new OffscreenCanvas(width, height);
        const ctx = canvas.getContext('2d')!;
        const images = await decodeImages(mf, mode);
        drawMetafile(ctx, mf, { x: 0, y: 0, w: width, h: height }, { createCanvas: (w, h) => new OffscreenCanvas(w, h), images, plus: mode, snap: true });
      }
    } catch (e) {
      info = `parse: ${e instanceof Error ? e.message : String(e)}`;
    }
    if (mode === 'auto') canvas = await metafileImage(bytes, (w, h) => new OffscreenCanvas(w, h));
    if (!canvas) return { error: 'metafileImage gave undefined', info, ms: performance.now() - started };
    const out = new OffscreenCanvas(canvas.width, canvas.height);
    const ctx = out.getContext('2d')!;
    if (options.background === 'checker') {
      const cell = Math.max(8, Math.round(Math.max(out.width, out.height) / 64));
      for (let y = 0; y < out.height; y += cell) {
        for (let x = 0; x < out.width; x += cell) {
          ctx.fillStyle = ((x + y) / cell) % 2 ? '#ddd' : '#fff';
          ctx.fillRect(x, y, cell, cell);
        }
      }
    } else {
      ctx.fillStyle = options.background ?? '#fff';
      ctx.fillRect(0, 0, out.width, out.height);
    }
    ctx.drawImage(canvas, 0, 0);
    const blob = await out.convertToBlob({ type: 'image/png' });
    return { png: toBase64(new Uint8Array(await blob.arrayBuffer())), width: out.width, height: out.height, info, ms: performance.now() - started };
  } catch (e) {
    return { error: e instanceof Error ? `${e.message}\n${e.stack ?? ''}` : String(e), ms: performance.now() - started };
  }
}

(globalThis as unknown as { renderMetafile: typeof renderMetafile }).renderMetafile = renderMetafile;
await loadFonts();
(globalThis as unknown as { harnessReady: boolean }).harnessReady = true;
