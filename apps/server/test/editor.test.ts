// The editor's production build, served by the server, in headless Chromium. Built pages carry a content security
// policy and are served cross-origin isolated (the dev server does neither), so this is where a worker, the text
// shaper's WebAssembly, the fonts or a decoded picture would be refused if anything were.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { FastifyInstance } from 'fastify';
import { chromium, type Browser, type Page } from 'playwright';
import { PNG } from 'pngjs';
import { build } from 'vite';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { addSlide, Editor, insertPicture, prepareImage } from '../../../packages/editor-core/src/index.ts';
import { E, emf } from '../../../packages/metafile/test/build.ts';
import { blankPresentation } from '../../web/src/engine/blank.ts';
import { cleanup, ROOT, tempDir, testServer } from './helpers.ts';

const EMU = 914_400;

/** A 2x2 uncompressed RGB TIFF of one colour: a format the browser does not decode, so the picture worker does. */
function tiff([r, g, b]: [number, number, number]): Uint8Array {
  const out = new Uint8Array(8 + 12 + 2 + 9 * 12 + 4 + 6);
  const v = new DataView(out.buffer);
  out.set([0x49, 0x49, 42, 0]);
  v.setUint32(4, 20, true);
  for (let i = 0; i < 4; i++) out.set([r, g, b], 8 + i * 3);
  const entries: [number, number, number][] = [[256, 4, 2], [257, 4, 2], [258, 3, 8], [259, 3, 1], [262, 3, 2], [273, 4, 8], [277, 3, 3], [278, 4, 2], [279, 4, 12]];
  v.setUint16(20, entries.length, true);
  entries.forEach(([tag, type, value], i) => {
    const e = 22 + i * 12;
    v.setUint16(e, tag, true);
    v.setUint16(e + 2, type, true);
    v.setUint32(e + 4, 1, true);
    if (type === 3) v.setUint16(e + 8, value, true);
    else v.setUint32(e + 8, value, true);
  });
  return out;
}

/** A one-slide deck with an EMF picture (red) on the left and a TIFF picture (green) on the right. */
async function deckWithPictures(): Promise<Uint8Array> {
  const ed = new Editor(OpcPackage.open(blankPresentation()));
  if (!ed.doc.slideParts().length) {
    const master = ed.doc.masterParts()[0]!;
    ed.run(addSlide(ed, ed.doc.layoutsOf(master).at(-1)!, 0).command);
  }
  const slide = ed.doc.slideParts()[0]!;
  const red = emf([E.createBrush(1, 0, 0x0000ff), E.select(1), E.rectangle(0, 0, 99, 99)]);
  const pictures = [
    await prepareImage(ed, { bytes: red, contentType: 'image/x-emf', ext: 'emf', width: 100, height: 100 }),
    await prepareImage(ed, { bytes: tiff([0, 200, 0]), contentType: 'image/tiff', ext: 'tiff', width: 2, height: 2 }),
  ];
  pictures.forEach((img, i) => ed.run(insertPicture(ed, slide, img, { x: EMU + i * 4 * EMU, y: EMU, w: 2 * EMU, h: 2 * EMU }).command));
  return ed.save();
}

let app: FastifyInstance;
let browser: Browser;
let page: Page;
let url: string;
const errors: string[] = [];

beforeAll(async () => {
  const webDir = tempDir('pptx-web-build-');
  await build({ configFile: join(ROOT, 'apps/web/vite.config.ts'), root: join(ROOT, 'apps/web'), logLevel: 'error', build: { outDir: webDir, emptyOutDir: true } });
  app = await testServer({ webDir });
  await app.listen({ host: '127.0.0.1', port: 0 });
  const address = app.server.address();
  url = `http://127.0.0.1:${typeof address === 'object' && address ? address.port : 0}/`;
  browser = await chromium.launch();
  page = await browser.newPage({ viewport: { width: 1400, height: 860 }, acceptDownloads: true });
  page.on('pageerror', e => errors.push(`page error: ${e.message}`));
  page.on('console', m => { if (m.type() === 'error') errors.push(`console: ${m.text()}`); });
  // a refused script, worker, font or picture is reported here, in the page or in a worker
  await page.addInitScript(() => {
    addEventListener('securitypolicyviolation', e => console.error(`CSP: ${e.violatedDirective} ${e.blockedURI}`));
  });
}, 180_000);

afterAll(async () => {
  await browser?.close();
  await app?.close();
  cleanup();
});

/** The colour on the page at a point of the slide (EMU). */
async function colourAt(x: number, y: number): Promise<[number, number, number]> {
  const p = await page.evaluate(([x, y]) => {
    const st = (globalThis as unknown as { __stage: { getAbsoluteTransform(): { point(p: { x: number; y: number }): { x: number; y: number } }; container(): HTMLElement } }).__stage;
    const q = st.getAbsoluteTransform().point({ x, y });
    const r = st.container().getBoundingClientRect();
    return { x: r.left + q.x, y: r.top + q.y };
  }, [x, y] as const);
  const shot = PNG.sync.read(await page.screenshot({ clip: { x: Math.round(p.x), y: Math.round(p.y), width: 1, height: 1 } }));
  return [shot.data[0]!, shot.data[1]!, shot.data[2]!];
}

describe('the production build of the editor', () => {
  test('is served with its policy, isolated from other origins', async () => {
    const res = await page.goto(url);
    expect(res?.status()).toBe(200);
    expect(res?.headers()['content-security-policy']).toContain("script-src 'self' 'wasm-unsafe-eval'");
    expect(res?.headers()['cross-origin-embedder-policy']).toBe('require-corp');
    await expect.poll(() => page.evaluate(() => crossOriginIsolated)).toBe(true);
    await page.getByTestId('welcome-new').waitFor();
  });

  test('opens a deck and draws its EMF and TIFF pictures, on the slide and in the thumbnail', async () => {
    await page.getByTestId('welcome-open').setInputFiles({ name: 'pictures.pptx', mimeType: 'application/octet-stream', buffer: Buffer.from(await deckWithPictures()) });
    await expect.poll(() => page.getByTestId('slide-count').textContent(), { timeout: 30_000 }).toBe('Slide 1 of 1');
    const red = ([r, g, b]: number[]) => r! > 200 && g! < 80 && b! < 80;
    const green = ([r, g, b]: number[]) => g! > 150 && r! < 80 && b! < 80;
    await expect.poll(async () => red(await colourAt(2 * EMU, 2 * EMU)), { timeout: 30_000 }).toBe(true);
    await expect.poll(async () => green(await colourAt(6 * EMU, 2 * EMU)), { timeout: 30_000 }).toBe(true);
    // the thumbnail is drawn in a worker, with its own fonts, shaper and picture decoding
    const thumbHasRed = async () => {
      const png = PNG.sync.read(await page.getByTestId('thumb-1').locator('canvas').screenshot());
      for (let i = 0; i < png.data.length; i += 4) if (red([png.data[i]!, png.data[i + 1]!, png.data[i + 2]!])) return true;
      return false;
    };
    await expect.poll(thumbHasRed, { timeout: 30_000 }).toBe(true);
  });

  test('saves a deck that opens again, through the save worker', async () => {
    await page.getByRole('tab', { name: 'File' }).click();
    const [download] = await Promise.all([page.waitForEvent('download'), page.getByTestId('save').click()]);
    const saved = OpcPackage.open(new Uint8Array(readFileSync((await download.path())!)));
    expect(saved.partNames().filter(n => /\.(emf|tiff)$/.test(n)).length).toBe(2);
  });

  test('without errors, and without anything refused by its policy', () => {
    expect(errors).toEqual([]);
  });
});
