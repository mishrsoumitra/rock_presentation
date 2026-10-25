// apps/server/render-page/vite.config.ts  -- builds the render page into apps/server/dist/render-page
//
//   pnpm --filter @pptx/server build:render-page
//
// The output is static: index.html, one script chunk, and harfbuzz.wasm as an asset of its own. harfbuzzjs finds its
// .wasm with new URL('harfbuzz.wasm', import.meta.url), which Vite rewrites to the emitted asset. Nothing is inlined
// as a data: URL, because the page's CSP (set by the server) allows scripts, styles and fetches from its origin only.
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';

const HERE = dirname(fileURLToPath(import.meta.url));

export const RENDER_PAGE_OUT = join(HERE, '../dist/render-page');

export default defineConfig({
  root: HERE,
  base: './',
  publicDir: false,
  logLevel: 'warn',
  build: {
    outDir: RENDER_PAGE_OUT,
    emptyOutDir: true,
    target: 'es2024',            // top-level await (harfbuzzjs, the page itself); only Playwright's Chromium loads it
    assetsInlineLimit: 0,
    modulePreload: false,
    reportCompressedSize: false,
    chunkSizeWarningLimit: 4096,
  },
  // dev server only: pre-bundling would move harfbuzzjs away from its .wasm (see tools/render/render.ts)
  optimizeDeps: { exclude: ['harfbuzzjs'] },
});
