// tools/render/render.ts  -- render decks to PNG in headless Chromium, with the product's own renderer
//
//   node tools/render/render.ts [--width 1280] [--out out/render] [--max-slides N] <deck or folder>...
//
// Starts the harness page with Vite, opens it in Playwright's Chromium, and writes slideN.png per deck plus a
// summary (errors per slide, time per slide). This is the "our renderer" half of oracle O3.
import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium, type Page } from 'playwright';
import { createServer, type ViteDevServer } from 'vite';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = join(HERE, '../..');

export interface RenderedSlide { index: number; part: string; png?: string; error?: string; ms: number; }

export interface Harness { page: Page; close(): Promise<void>; render(bytes: Uint8Array, width: number, slides?: number[]): Promise<RenderedSlide[]>; }

/** Start the harness: a Vite server for the page, and a Chromium page on it. */
export async function startHarness(): Promise<Harness> {
  const server: ViteDevServer = await createServer({
    root: join(HERE, 'harness'),
    publicDir: join(ROOT, 'assets'),
    logLevel: 'error',
    server: { port: 0, strictPort: false, fs: { allow: [ROOT] } },
    // harfbuzzjs finds its .wasm next to its own module; pre-bundling would move the module away from it
    optimizeDeps: { exclude: ['harfbuzzjs'] },
  });
  await server.listen();
  const url = server.resolvedUrls?.local[0];
  if (!url) throw new Error('vite did not report a URL');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors: string[] = [];
  page.on('pageerror', e => errors.push(String(e)));
  await page.goto(url);
  await page.waitForFunction(() => (globalThis as unknown as { harnessReady?: boolean }).harnessReady === true, null, { timeout: 60_000 })
    .catch(() => { throw new Error(`harness did not start: ${errors.join('\n')}`); });
  return {
    page,
    async render(bytes, width, slides) {
      const b64 = Buffer.from(bytes).toString('base64');
      return page.evaluate(
        ([data, w, s]) => (globalThis as unknown as { renderDeck: (d: string, o: object) => Promise<RenderedSlide[]> }).renderDeck(data, { width: w, ...(s ? { slides: s } : {}) }),
        [b64, width, slides] as const,
      );
    },
    async close() {
      await browser.close();
      await server.close();
    },
  };
}

function decksIn(paths: string[]): string[] {
  const out: string[] = [];
  for (const p of paths) {
    if (statSync(p).isDirectory()) out.push(...decksIn(readdirSync(p).map(n => join(p, n))));
    else if (['.pptx', '.pptm', '.potx', '.ppsx'].includes(extname(p).toLowerCase())) out.push(p);
  }
  return out.sort();
}

async function main(): Promise<number> {
  const args = process.argv.slice(2);
  const opt = (name: string, dflt: string) => {
    const i = args.indexOf(name);
    if (i < 0) return dflt;
    const v = args[i + 1] ?? dflt;
    args.splice(i, 2);
    return v;
  };
  const width = Number(opt('--width', '1280'));
  const outDir = resolve(opt('--out', join(ROOT, 'out/render')));
  const maxSlides = Number(opt('--max-slides', '0'));
  const decks = decksIn(args);
  const harness = await startHarness();
  const summary: { deck: string; slides: number; errors: { index: number; error: string }[]; ms: number }[] = [];
  try {
    for (const deck of decks) {
      const name = basename(deck).replace(/\.[^.]+$/, '');
      let results: RenderedSlide[];
      try {
        const slides = maxSlides > 0 ? Array.from({ length: maxSlides }, (_, i) => i) : undefined;
        results = await harness.render(new Uint8Array(readFileSync(deck)), width, slides);
      } catch (e) {
        summary.push({ deck, slides: 0, errors: [{ index: -1, error: String(e).slice(0, 2000) }], ms: 0 });
        console.log(`FAIL  ${name}: ${String(e).split('\n')[0]}`);
        continue;
      }
      const dir = join(outDir, name);
      mkdirSync(dir, { recursive: true });
      for (const r of results) if (r.png) writeFileSync(join(dir, `slide${r.index + 1}.png`), Buffer.from(r.png, 'base64'));
      const errors = results.filter(r => r.error).map(r => ({ index: r.index, error: r.error!.slice(0, 2000) }));
      const ms = results.reduce((a, r) => a + r.ms, 0);
      summary.push({ deck, slides: results.length, errors, ms });
      console.log(`${errors.length ? 'ERR ' : 'ok  '}  ${name}: ${results.length} slides, ${Math.round(ms)} ms${errors.length ? `, ${errors.length} failed: ${errors[0]!.error.split('\n')[0]}` : ''}`);
    }
  } finally {
    await harness.close();
  }
  mkdirSync(outDir, { recursive: true });
  writeFileSync(join(outDir, 'summary.json'), JSON.stringify(summary, null, 2));
  const failed = summary.filter(s => s.errors.length).length;
  console.log(`${summary.length} decks, ${summary.reduce((a, s) => a + s.slides, 0)} slides, ${failed} decks with errors -> ${outDir}`);
  return failed ? 1 : 0;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) process.exitCode = await main();
