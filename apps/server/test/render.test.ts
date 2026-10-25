// Slide PNGs and PDFs from headless Chromium, through the API; and the render page's sandbox.
import { join } from 'node:path';
import type { FastifyInstance } from 'fastify';
import pngjs from 'pngjs';
import { build } from 'vite';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { allowedRequest, RENDER_ORIGIN } from '../src/render.ts';
import { cleanup, corpusFile, deck, ROOT, tempDir, testServer, upload } from './helpers.ts';

// python-pptx's test deck with three slides of positioned autoshapes (4:3, 10 x 7.5 in)
const DECK = corpusFile('python-pptx/features~steps~test_files~shp-pos-and-size.pptx');

/** Pixels that differ from the top-left one: zero for a blank image. */
function inked(png: Buffer): { width: number; height: number; inked: number } {
  const img = pngjs.PNG.sync.read(png);
  const d = img.data;
  let n = 0;
  for (let i = 0; i < d.length; i += 4) if (d[i] !== d[0] || d[i + 1] !== d[1] || d[i + 2] !== d[2]) n++;
  return { width: img.width, height: img.height, inked: n };
}

afterAll(() => cleanup());

/** Page objects and their MediaBoxes in a PDF Chromium wrote (Skia writes no object streams). */
function pdfPages(pdf: Buffer): { pages: number; boxes: string[] } {
  const text = pdf.toString('latin1');
  return { pages: text.match(/\/Type\s*\/Page(?![s\w])/g)?.length ?? 0, boxes: [...text.matchAll(/\/MediaBox\s*\[([^\]]*)\]/g)].map(m => m[1]!.trim()) };
}

describe.skipIf(!DECK)('rendering in Chromium', () => {
  let app: FastifyInstance;
  let id: string;
  let pageDir: string;
  beforeAll(async () => {
    // a fresh build of the render page, so the test draws with the packages as they are now
    pageDir = tempDir('pptx-render-page-');
    await build({ configFile: join(ROOT, 'apps/server/render-page/vite.config.ts'), logLevel: 'error', build: { outDir: pageDir, emptyOutDir: true } });
    app = await testServer({ render: { pages: 1, pageDir } });
    const res = await upload(app, DECK!, { filename: 'positions.pptx' });
    expect(res.statusCode).toBe(201);
    id = res.json().id;
  }, 120_000);
  afterAll(async () => {
    await app?.close();
  });

  test('a slide as PNG: the asked width, the slide\'s aspect, not blank; cached afterwards', async () => {
    const res = await app.inject({ url: `/api/decks/${id}/slides/1.png?width=640` });
    expect(res.statusCode, res.body).toBe(200);
    expect(res.headers['content-type']).toBe('image/png');
    const img = inked(res.rawPayload);
    expect(img).toMatchObject({ width: 640, height: 480 });
    expect(img.inked).toBeGreaterThan(640 * 480 * 0.05);    // the blue autoshape is 192 x 256 px

    const again = await app.inject({ url: `/api/decks/${id}/slides/1.png?width=640` });
    expect(Buffer.compare(again.rawPayload, res.rawPayload)).toBe(0);
    expect(again.headers.etag).toBe(res.headers.etag);
    expect((await app.inject({ url: `/api/decks/${id}/slides/1.png?width=640`, headers: { 'if-none-match': res.headers.etag as string } })).statusCode).toBe(304);

    const wide = inked((await app.inject({ url: `/api/decks/${id}/slides/3.png?width=1000` })).rawPayload);
    expect(wide).toMatchObject({ width: 1000, height: 750 });
  });

  test('slides that do not exist, and widths over the limit', async () => {
    expect((await app.inject({ url: `/api/decks/${id}/slides/4.png` })).statusCode).toBe(404);
    expect((await app.inject({ url: `/api/decks/${id}/slides/0.png` })).statusCode).toBe(400);
    expect((await app.inject({ url: `/api/decks/${id}/slides/1.png?width=5000` })).statusCode).toBe(400);
  });

  test('the PDF has one page per slide, each the slide\'s size', async () => {
    const res = await app.inject({ url: `/api/decks/${id}/export.pdf` });
    expect(res.statusCode, res.body).toBe(200);
    expect(res.headers['content-type']).toBe('application/pdf');
    expect(res.headers['content-disposition']).toContain('attachment; filename="positions.pdf"');
    expect(res.rawPayload.subarray(0, 5).toString()).toBe('%PDF-');
    const { pages, boxes } = pdfPages(res.rawPayload);
    expect(pages).toBe(3);
    expect(boxes.every(b => b === '0 0 720 540')).toBe(true);   // 10 x 7.5 in at 72 pt per inch
  });

  test('a deck without slides has no PDF', async () => {
    const empty = (await upload(app, deck(0))).json();
    expect((await app.inject({ url: `/api/decks/${empty.id}/export.pdf` })).json().error).toBe('no-slides');
  });

  test('the render page reaches nothing but its own origin', async () => {
    const renderer = app.services.renderer!;
    const before = renderer.stats().blocked;
    // another origin: the page's CSP stops it before it is even requested
    expect((await renderer.probe('https://example.com/')).ok).toBe(false);
    expect((await renderer.probe('http://127.0.0.1:9/')).ok).toBe(false);
    // its own origin with a method the interception refuses: aborted there, and counted
    expect((await renderer.probe(`${RENDER_ORIGIN}/index.html`, 'POST')).ok).toBe(false);
    expect(renderer.stats().blocked).toBe(before + 1);
    // paths outside the page's directory are not served
    expect((await renderer.probe('/..%2F..%2Fpackage.json')).status).toBe(404);
    expect((await renderer.probe('/fonts/..%2F..%2Fpackage.json')).status).toBe(404);
    expect((await renderer.probe('/index.html')).ok).toBe(true);
  });

  test('a render past its deadline answers 504, and its page is closed rather than reused', async () => {
    const slow = await testServer({ render: { pages: 1, pageDir, timeoutMs: 5 } });
    try {
      const deckId = (await upload(slow, DECK!)).json().id;
      const res = await slow.inject({ url: `/api/decks/${deckId}/slides/2.png?width=2000` });
      expect([res.statusCode, res.json().error]).toEqual([504, 'render-timeout']);
      await expect.poll(() => slow.services.renderer!.stats().pages).toBe(0);
      expect((await slow.inject({ url: '/api/health' })).statusCode).toBe(200);
    } finally {
      await slow.close();
    }
  });

  test('closing the server closes Chromium and the worker threads', async () => {
    const short = await testServer({ render: { pages: 1, pageDir } });
    const deckId = (await upload(short, DECK!)).json().id;
    expect((await short.inject({ url: `/api/decks/${deckId}/slides/1.png?width=320` })).statusCode).toBe(200);
    const { renderer, pool } = short.services;
    expect(renderer!.stats().browser).toBe(true);
    expect(pool.stats().threads).toBeGreaterThan(0);
    await short.close();
    expect(renderer!.stats().browser).toBe(false);
    expect(pool.stats().threads).toBe(0);
    await expect(pool.run('validate', { bytes: deck(1), limits: short.services.config.limits, container: short.services.config.container })).rejects.toThrow(/shutting down/);
  });
});

test('only GETs of the render origin are allowed', () => {
  expect(allowedRequest(`${RENDER_ORIGIN}/index.html`, 'GET')).toBe(true);
  expect(allowedRequest(`${RENDER_ORIGIN}/deck/abc`, 'POST')).toBe(false);
  expect(allowedRequest('https://example.com/', 'GET')).toBe(false);
  expect(allowedRequest('http://render.pptx.invalid:8080/', 'GET')).toBe(false);
  expect(allowedRequest('data:text/html,hi', 'GET')).toBe(false);
  expect(allowedRequest('not a url', 'GET')).toBe(false);
});

test('with rendering off, slide images and PDFs answer 503', async () => {
  const app = await testServer({ render: { pages: 0 } });
  try {
    const id = (await upload(app, deck(1))).json().id;
    for (const url of [`/api/decks/${id}/slides/1.png`, `/api/decks/${id}/export.pdf`]) {
      const res = await app.inject({ url });
      expect(res.statusCode).toBe(503);
      expect(res.json().error).toBe('render-unavailable');
    }
  } finally {
    await app.close();
  }
});
