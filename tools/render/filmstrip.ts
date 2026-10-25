// tools/render/filmstrip.ts  -- filmstrips of slide animations and transitions, drawn in headless Chromium
//
//   node tools/render/filmstrip.ts [--width 320] [--frames 5] [--out out/anim] [--slides 1,3] [--all] [--gallery] <deck or folder>...
//
// Starts tools/render/anim-harness with Vite, opens it in Playwright's Chromium, and writes slideN.png per slide
// that has animations or a transition (every slide with --all): a row for the transition from the previous slide,
// then a row per click step (and per trigger click), frames evenly spaced through each, labelled with their time.
// --gallery also writes transitions.png: every transition kind between two test cards.
import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import { createServer, type ViteDevServer } from 'vite';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = join(HERE, '../..');

interface StripResult { index: number; part: string; png?: string; error?: string; steps: number; transition?: string; }

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
  const flag = (name: string) => {
    const i = args.indexOf(name);
    if (i >= 0) args.splice(i, 1);
    return i >= 0;
  };
  const width = Number(opt('--width', '320'));
  const frames = Number(opt('--frames', '5'));
  const outDir = resolve(opt('--out', join(ROOT, 'out/anim')));
  const slideList = opt('--slides', '');
  const slides = slideList ? slideList.split(',').map(s => Number(s) - 1) : undefined;
  const all = flag('--all');
  const gallery = flag('--gallery');
  const decks = decksIn(args);

  const server: ViteDevServer = await createServer({
    root: join(HERE, 'anim-harness'),
    publicDir: join(ROOT, 'assets'),
    logLevel: 'error',
    server: { port: 0, strictPort: false, fs: { allow: [ROOT] } },
    optimizeDeps: { exclude: ['harfbuzzjs'] },
  });
  await server.listen();
  const url = server.resolvedUrls?.local[0];
  if (!url) throw new Error('vite did not report a URL');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors: string[] = [];
  page.on('pageerror', e => errors.push(String(e)));
  let failed = 0;
  try {
    await page.goto(url);
    await page.waitForFunction(() => (globalThis as unknown as { harnessReady?: boolean }).harnessReady === true, null, { timeout: 60_000 })
      .catch(() => { throw new Error(`harness did not start: ${errors.join('\n')}`); });
    if (gallery) {
      const png = await page.evaluate(
        o => (globalThis as unknown as { transitionGallery: (o: object) => Promise<string> }).transitionGallery(o), { width, frames });
      mkdirSync(outDir, { recursive: true });
      writeFileSync(join(outDir, 'transitions.png'), Buffer.from(png, 'base64'));
      console.log(`every transition kind -> ${join(outDir, 'transitions.png')}`);
    }
    for (const deck of decks) {
      const name = basename(deck).replace(/\.[^.]+$/, '');
      const b64 = readFileSync(deck).toString('base64');
      let results: StripResult[];
      try {
        type Harness = { renderDeckStrips: (d: string, o: object) => Promise<StripResult[]> };
        results = await page.evaluate(
          ([data, o]) => (globalThis as unknown as Harness).renderDeckStrips(data, o),
          [b64, { width, frames, ...(slides ? { slides } : {}), all }] as const,
        );
      } catch (e) {
        failed++;
        console.log(`FAIL  ${name}: ${String(e).split('\n')[0]}`);
        continue;
      }
      const dir = join(outDir, name);
      mkdirSync(dir, { recursive: true });
      for (const r of results) {
        if (r.png) writeFileSync(join(dir, `slide${r.index + 1}.png`), Buffer.from(r.png, 'base64'));
        if (r.error) failed++;
        const what = `${r.steps} steps${r.transition ? `, ${r.transition}` : ''}${r.error ? `: ${r.error.split('\n')[0]}` : ''}`;
        console.log(`${r.error ? 'ERR ' : 'ok  '}  ${name} slide ${r.index + 1}: ${what}`);
      }
    }
  } finally {
    await browser.close();
    await server.close();
  }
  console.log(`${decks.length} decks -> ${outDir}`);
  return failed ? 1 : 0;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) process.exitCode = await main();
