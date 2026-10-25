// apps/web/src/engine/fonts.ts  -- the fonts text is shaped with: open, metric-compatible substitutes for Office's
import { FontRegistry } from '@pptx/text';

export const FONT_FILES = [
  'carlito/Carlito-Regular.ttf', 'carlito/Carlito-Bold.ttf', 'carlito/Carlito-Italic.ttf', 'carlito/Carlito-BoldItalic.ttf',
  'caladea/Caladea-Regular.ttf', 'caladea/Caladea-Bold.ttf', 'caladea/Caladea-Italic.ttf', 'caladea/Caladea-BoldItalic.ttf',
  'liberation/LiberationSans-Regular.ttf', 'liberation/LiberationSans-Bold.ttf', 'liberation/LiberationSans-Italic.ttf', 'liberation/LiberationSans-BoldItalic.ttf',
  'liberation/LiberationSerif-Regular.ttf', 'liberation/LiberationSerif-Bold.ttf', 'liberation/LiberationSerif-Italic.ttf', 'liberation/LiberationSerif-BoldItalic.ttf',
  'liberation/LiberationMono-Regular.ttf', 'liberation/LiberationMono-Bold.ttf', 'liberation/LiberationMono-Italic.ttf', 'liberation/LiberationMono-BoldItalic.ttf',
];

/** The CSS family, weight and style of a bundled font file, from its name ("liberation/LiberationSans-BoldItalic.ttf"). */
function cssFace(file: string): { family: string; weight: string; style: string } {
  const stem = file.slice(file.lastIndexOf('/') + 1).replace(/\.ttf$/, '');
  const [name, variant = 'Regular'] = stem.split('-') as [string, string?];
  const family = name.replace(/^Liberation(Sans|Serif|Mono)$/, 'Liberation $1');
  return { family, weight: /Bold/.test(variant) ? '700' : '400', style: /Italic/.test(variant) ? 'italic' : 'normal' };
}

let css: Promise<void> | undefined;

/**
 * The same fonts as CSS fonts of this page or worker. The renderer shapes text itself, but pictures (EMF and WMF)
 * draw their text with the canvas's own text API, which only sees fonts registered here.
 */
export function registerCssFonts(base = '/fonts/'): Promise<void> {
  const set = (globalThis as { fonts?: FontFaceSet }).fonts ?? (typeof document !== 'undefined' ? document.fonts : undefined);
  if (!set || typeof FontFace === 'undefined') return Promise.resolve();
  css ??= Promise.all(FONT_FILES.map(async file => {
    const { family, weight, style } = cssFace(file);
    try {
      set.add(await new FontFace(family, `url(${base}${file})`, { weight, style }).load());
    } catch {
      // a font that does not load leaves the browser's own fallback
    }
  })).then(() => undefined);
  return css;
}

let loading: Promise<FontRegistry> | undefined;

/** The registry, loaded once (Calibri is drawn with Carlito, Cambria with Caladea, Arial with Liberation Sans...). */
export function loadFonts(base = '/fonts/'): Promise<FontRegistry> {
  loading ??= (async () => {
    const reg = new FontRegistry();
    const files = await Promise.all(FONT_FILES.map(async f => {
      const res = await fetch(base + f);
      return res.ok ? new Uint8Array(await res.arrayBuffer()) : undefined;
    }));
    for (const bytes of files) if (bytes) reg.addBytes(bytes);
    return reg;
  })();
  return loading;
}
