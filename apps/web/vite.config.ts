// apps/web/vite.config.ts  -- the editor app: React, the workspace packages from source, fonts from assets/
import { fileURLToPath } from 'node:url';
import react from '@vitejs/plugin-react';
import { defineConfig, type Plugin } from 'vite';

const root = fileURLToPath(new URL('../..', import.meta.url));

/**
 * The page's content security policy, in built pages only (the dev server's React refresh preamble is an inline
 * script). Decks are untrusted: no inline script, no plugins, pictures only from blobs and data URLs, SVG only
 * through <img> (where it cannot run script), and WebAssembly for the text shaper.
 */
const CSP = "default-src 'self'; img-src 'self' blob: data:; style-src 'self' 'unsafe-inline'; script-src 'self' 'wasm-unsafe-eval'; "
  + "worker-src 'self' blob:; connect-src 'self'; font-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'";
const csp = (): Plugin => ({
  name: 'csp',
  apply: 'build',
  transformIndexHtml: html => html.replace('<head>', `<head>\n<meta http-equiv="Content-Security-Policy" content="${CSP}">`),
});

export default defineConfig({
  plugins: [react(), csp()],
  // the metric-compatible open fonts the renderer substitutes for Calibri, Cambria, Arial... (assets/fonts)
  publicDir: fileURLToPath(new URL('../../assets', import.meta.url)),
  server: { fs: { allow: [root] }, headers: { 'Cross-Origin-Opener-Policy': 'same-origin' } },
  // harfbuzzjs loads its .wasm from next to its own module; pre-bundling would move the module away from it
  optimizeDeps: { exclude: ['harfbuzzjs'] },
  worker: { format: 'es' },
  build: { target: 'es2023', sourcemap: true },
});
