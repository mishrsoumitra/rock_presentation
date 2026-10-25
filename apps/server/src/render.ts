// apps/server/src/render.ts  -- slide PNGs and PDFs, drawn by the product's own renderer in headless Chromium
//
// The render page (render-page/, built by Vite into dist/render-page) runs @pptx/opc, model, resolve, render, text,
// charts and metafile in a real browser, as the web app will. Chromium gets the page, the fonts and the deck bytes
// from one made-up origin through Playwright's request interception: files from disk, decks from memory. Every
// other request is aborted there, whatever its origin or method; host names do not resolve at all should anything
// get past the routes; service workers, downloads and WebSockets are refused. Pages are pooled and reused; a page
// that times out or crashes is closed and replaced, and a browser that dies is relaunched on the next job.
import { randomBytes } from 'node:crypto';
import { existsSync } from 'node:fs';
import { readdir, readFile } from 'node:fs/promises';
import { extname, join, relative, resolve, sep } from 'node:path';
import { chromium, type Browser, type BrowserContext, type Page, type Route } from 'playwright';
import type { PackageLimits } from '@pptx/opc';
import type { RenderSettings } from './config.ts';
import type { DeckRequest, PrintResult, RenderPageApi, RenderResult } from '../render-page/main.ts';

/** The render page's origin. The .invalid TLD never resolves (RFC 6761); requests to it are answered by serve(). */
export const RENDER_ORIGIN = 'http://render.pptx.invalid';
const EMU_PER_INCH = 914_400;

const PAGE_CSP = [
  "default-src 'none'", "script-src 'self' 'wasm-unsafe-eval'", "connect-src 'self'", "img-src 'self' blob:", "style-src 'self'",
  "base-uri 'none'", "form-action 'none'", "frame-ancestors 'none'",
].join('; ');

const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.wasm': 'application/wasm', '.json': 'application/json', '.ttf': 'font/ttf', '.otf': 'font/otf', '.ttc': 'font/collection',
};

// Carlito first: it is the registry's default family and Calibri's substitute, as in tools/render/harness/main.ts
const FONT_ORDER = ['carlito/', 'caladea/', 'liberation/'];

export type RenderFailure = 'render-failed' | 'timeout' | 'unavailable' | 'no-slides';

export class RenderError extends Error {
  readonly code: RenderFailure;
  constructor(code: RenderFailure, message: string) {
    super(message);
    this.name = 'RenderError';
    this.code = code;
  }
}

/** A deck to draw: `key` is its content hash (the page keeps the last deck open by it), `load` reads its bytes. */
export interface DeckSource { key: string; load: () => Promise<Uint8Array>; }

/** Only GETs of the render origin get an answer; everything else Chromium asks for is aborted. */
export function allowedRequest(url: string, method: string): boolean {
  if (method !== 'GET') return false;
  try {
    return new URL(url).origin === RENDER_ORIGIN;
  } catch {
    return false;
  }
}

/** The file a URL path names under `root`, or undefined when the path is malformed or would leave it. */
function under(root: string, urlPath: string): string | undefined {
  let path: string;
  try {
    path = decodeURIComponent(urlPath);
  } catch {
    return undefined;
  }
  if (path.includes('\0')) return undefined;
  const base = resolve(root);
  const file = resolve(base, `.${path}`);
  return file.startsWith(base + sep) ? file : undefined;
}

/** Every font file under the fonts directory, as paths relative to it, Carlito first. */
async function fontManifest(dir: string): Promise<string[]> {
  const out: string[] = [];
  const walk = async (d: string): Promise<void> => {
    for (const e of await readdir(d, { withFileTypes: true })) {
      if (e.isDirectory()) await walk(join(d, e.name));
      else if (/\.(ttf|otf|ttc)$/i.test(e.name)) out.push(relative(dir, join(d, e.name)).split(sep).join('/'));
    }
  };
  if (existsSync(dir)) await walk(dir);
  const rank = (f: string) => {
    const i = FONT_ORDER.findIndex(p => f.toLowerCase().startsWith(p));
    return i < 0 ? FONT_ORDER.length : i;
  };
  return out.sort((a, b) => rank(a) - rank(b) || a.localeCompare(b));
}

const firstLine = (s: string) => s.replace(/^page\.evaluate: (Error: )?/, '').split('\n')[0]!.slice(0, 300);

export interface RenderStats { browser: boolean; pages: number; idle: number; blocked: number; }

export class Renderer {
  private readonly settings: RenderSettings;
  private readonly limits: PackageLimits;
  private browser?: Browser;
  private context?: BrowserContext;
  private starting?: Promise<BrowserContext>;
  private readonly idle: Page[] = [];
  private pages = 0;
  private closed = false;
  private fonts: string[] = [];
  private readonly files = new Map<string, Buffer>();                      // scripts, wasm and fonts, read once
  private readonly decks = new Map<string, () => Promise<Uint8Array>>();   // token -> bytes, while a job runs
  private blocked = 0;                                                     // requests the interception refused

  constructor(settings: RenderSettings, limits: PackageLimits) {
    this.settings = settings;
    this.limits = limits;
  }

  /** True when the render page has been built (pnpm --filter @pptx/server build:render-page). */
  get built(): boolean {
    return existsSync(join(this.settings.pageDir, 'index.html'));
  }

  stats(): RenderStats {
    return { browser: this.browser?.isConnected() ?? false, pages: this.pages, idle: this.idle.length, blocked: this.blocked };
  }

  /** One slide (0-based) as PNG, `width` pixels wide. */
  async renderSlide(deck: DeckSource, index: number, width: number): Promise<Uint8Array> {
    return this.withPage(this.settings.timeoutMs, async page => {
      const job = this.job(deck, width);
      try {
        const request = { ...job.request, slides: [index] };
        const r: RenderResult = await this.evaluate(page, q => (globalThis as unknown as RenderPageApi).renderSlides(q), request);
        const s = r.slides[0];
        if (!s?.png) throw new RenderError('render-failed', s?.error ?? 'no image was drawn');
        return Buffer.from(s.png, 'base64');
      } finally {
        job.done();
      }
    });
  }

  /**
   * Every slide as one page of a PDF, the page sized to the slide. Raster, not vector: each page is a PNG of the
   * slide `width` pixels wide, laid out by the render page and printed with page.pdf(). A slide that fails to draw
   * gives a blank page, reported in `errors`, so the page count is always the slide count.
   */
  async exportPdf(deck: DeckSource, width: number): Promise<{ pdf: Uint8Array; slideCount: number; errors: PrintResult['errors'] }> {
    return this.withPage(this.settings.pdfTimeoutMs, async page => {
      const job = this.job(deck, width);
      try {
        const r: PrintResult = await this.evaluate(page, q => (globalThis as unknown as RenderPageApi).printDeck(q), job.request);
        if (r.slideCount === 0) throw new RenderError('no-slides', 'The presentation has no slides.');
        const pdf = await page.pdf({
          width: `${r.size.cx / EMU_PER_INCH}in`, height: `${r.size.cy / EMU_PER_INCH}in`,
          margin: { top: '0', right: '0', bottom: '0', left: '0' }, printBackground: true, preferCSSPageSize: true,
        });
        return { pdf, slideCount: r.slideCount, errors: r.errors };
      } finally {
        job.done();
        await page.evaluate(() => (globalThis as unknown as RenderPageApi).clearPrint()).catch(() => undefined);
      }
    });
  }

  /**
   * Have the render page fetch a URL and say whether it got through. For tests and diagnostics: the page's CSP stops
   * other origins first; a request the CSP allows but the interception does not (a POST) shows up in `blocked`.
   */
  async probe(url: string, method = 'GET'): Promise<{ ok: boolean; status?: number; error?: string }> {
    return this.withPage(this.settings.timeoutMs, page => page.evaluate(async ([u, m]) => {
      try {
        const r = await fetch(u, { method: m });
        return { ok: r.ok, status: r.status };
      } catch (e) {
        return { ok: false, error: String(e) };
      }
    }, [url, method] as const));
  }

  async close(): Promise<void> {
    this.closed = true;
    await this.starting?.catch(() => undefined);
    const browser = this.browser;
    this.browser = undefined;
    this.context = undefined;
    this.idle.length = 0;
    await browser?.close().catch(() => undefined);
  }

  private job(deck: DeckSource, width: number): { request: DeckRequest; done: () => void } {
    const token = randomBytes(16).toString('hex');
    this.decks.set(token, deck.load);
    return {
      request: { url: `${RENDER_ORIGIN}/deck/${token}`, key: deck.key, limits: this.limits, width, maxPixels: this.settings.maxPixels },
      done: () => this.decks.delete(token),
    };
  }

  /** page.evaluate, with the page's own exceptions (a deck it cannot draw) as RenderErrors. */
  private async evaluate<A, R>(page: Page, fn: (arg: A) => Promise<R>, arg: A): Promise<R> {
    try {
      // A is plain JSON here, which Playwright passes through unchanged; its Unboxed<A> cannot see that generically
      return await page.evaluate(fn as (arg: unknown) => Promise<R>, arg as unknown);
    } catch (e) {
      if (page.isClosed() || e instanceof RenderError) throw e;
      throw new RenderError('render-failed', firstLine(e instanceof Error ? e.message : String(e)));
    }
  }

  /**
   * Run a job on a pooled page under a deadline. A page that overruns, crashes or fails in a way that is not the
   * deck's fault is closed rather than reused.
   */
  private async withPage<T>(timeoutMs: number, fn: (page: Page) => Promise<T>): Promise<T> {
    if (this.closed) throw new RenderError('unavailable', 'The renderer is shut down.');
    const page = this.idle.pop() ?? await this.newPage();
    let healthy = true;
    let timer: NodeJS.Timeout | undefined;
    let onCrash = () => {};
    const failure = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new RenderError('timeout', `Rendering took longer than ${timeoutMs} ms.`)), timeoutMs);
      onCrash = () => reject(new RenderError('render-failed', 'The renderer crashed.'));
      page.on('crash', onCrash);
    });
    try {
      return await Promise.race([fn(page), failure]);
    } catch (e) {
      if (!(e instanceof RenderError) || e.code === 'timeout' || page.isClosed()) healthy = false;
      throw e;
    } finally {
      clearTimeout(timer);
      page.off('crash', onCrash);
      if (healthy && !this.closed && !page.isClosed() && this.idle.length < this.settings.pages) this.idle.push(page);
      else void this.discard(page);
    }
  }

  private async newPage(): Promise<Page> {
    const context = await this.ensureContext();
    const page = await context.newPage();
    this.pages++;
    page.once('close', () => this.pages--);
    page.on('dialog', d => void d.dismiss());
    try {
      await page.goto(`${RENDER_ORIGIN}/index.html`, { timeout: 30_000 });
      await page.waitForFunction(() => (globalThis as unknown as RenderPageApi).renderReady === true, null, { timeout: 60_000 });
    } catch (e) {
      void this.discard(page);
      throw new RenderError('unavailable', `The render page did not start: ${firstLine(e instanceof Error ? e.message : String(e))}`);
    }
    return page;
  }

  /** Close a page; a page that will not close within 5 s takes the browser down with it, to be relaunched. */
  private async discard(page: Page): Promise<void> {
    const closed = page.close({ runBeforeUnload: false }).then(() => true, () => true);
    const stuck = new Promise<boolean>(r => setTimeout(() => r(false), 5000).unref());
    if (!(await Promise.race([closed, stuck]))) {
      const browser = this.browser;
      this.browser = undefined;
      this.context = undefined;
      this.idle.length = 0;
      await browser?.close().catch(() => undefined);
    }
  }

  private ensureContext(): Promise<BrowserContext> {
    if (this.context && this.browser?.isConnected()) return Promise.resolve(this.context);
    this.starting ??= this.launch().finally(() => { this.starting = undefined; });
    return this.starting;
  }

  private async launch(): Promise<BrowserContext> {
    if (!this.built) throw new RenderError('unavailable', 'The render page is not built: run "pnpm --filter @pptx/server build:render-page".');
    this.fonts = await fontManifest(this.settings.fontsDir);
    let browser: Browser;
    try {
      browser = await chromium.launch({
        headless: true,
        chromiumSandbox: this.settings.sandbox,
        args: [
          '--host-resolver-rules=MAP * ~NOTFOUND',                    // no host name resolves
          `--js-flags=--max-old-space-size=${this.settings.heapMb}`,
          '--disable-features=MediaRouter,OptimizationHints,Translate,AutofillServerCommunication,InterestFeedContentSuggestions',
          '--no-pings',
        ],
      });
    } catch (e) {
      throw new RenderError('unavailable', `Chromium did not start: ${firstLine(e instanceof Error ? e.message : String(e))}`);
    }
    try {
      browser.on('disconnected', () => {
        if (this.browser !== browser) return;
        this.browser = undefined;
        this.context = undefined;
        this.idle.length = 0;
      });
      const context = await browser.newContext({
        serviceWorkers: 'block', acceptDownloads: false, viewport: { width: 1280, height: 800 }, deviceScaleFactor: 1,
        locale: 'en-US', timezoneId: 'UTC',
      });
      await context.route(() => true, route => this.serve(route));
      await context.routeWebSocket(() => true, ws => {
        this.blocked++;
        void ws.close();
      });
      if (this.closed) throw new RenderError('unavailable', 'The renderer is shut down.');
      this.browser = browser;
      this.context = context;
      return context;
    } catch (e) {
      await browser.close().catch(() => undefined);
      throw e;
    }
  }

  /** Answer one request from Chromium: the page and its assets, the fonts, the deck of a running job. Nothing else. */
  private async serve(route: Route): Promise<void> {
    const request = route.request();
    if (!allowedRequest(request.url(), request.method())) {
      this.blocked++;
      await route.abort('blockedbyclient').catch(() => undefined);
      return;
    }
    const path = new URL(request.url()).pathname;
    try {
      if (path.startsWith('/deck/')) {
        const load = this.decks.get(path.slice('/deck/'.length));
        if (!load) return await route.fulfill({ status: 404 });
        const bytes = await load();
        const body = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        return await route.fulfill({ status: 200, body, headers: { 'content-type': 'application/octet-stream', 'cache-control': 'no-store' } });
      }
      if (path === '/fonts/manifest.json') {
        return await route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(this.fonts) });
      }
      const file = path.startsWith('/fonts/') ? under(this.settings.fontsDir, path.slice('/fonts'.length)) : under(this.settings.pageDir, path);
      const type = file && TYPES[extname(file).toLowerCase()];
      if (!file || !type) return await route.fulfill({ status: 404 });
      const html = type.startsWith('text/html');
      let body = this.files.get(file);
      if (!body) {
        body = await readFile(file).catch(() => undefined);
        if (!body) return await route.fulfill({ status: 404 });
        if (!html) this.files.set(file, body);    // the page itself is read each time, so a rebuild takes effect
      }
      const headers: Record<string, string> = { 'content-type': type, 'x-content-type-options': 'nosniff' };
      if (html) headers['content-security-policy'] = PAGE_CSP;
      await route.fulfill({ status: 200, body, headers });
    } catch {
      await route.abort('failed').catch(() => undefined);
    }
  }
}
