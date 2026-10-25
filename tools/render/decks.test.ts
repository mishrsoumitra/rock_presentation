// The editor on your own decks, end to end in Chromium: every text box of every slide can be entered where it is
// double-clicked, shows a caret you can see, and takes typing; the slide panel keeps the current slide in view.
//
//   corepack pnpm test:decks            (decks from corpus/local/, or DECKS_DIR)
//
// The decks are copies you put in corpus/local/, which git ignores: they are never committed. Without any, the
// test is skipped (as in CI).
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium, type Browser, type Page } from 'playwright';
import { PNG } from 'pngjs';
import { createServer, type ViteDevServer } from 'vite';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { PresentationDocument } from '@pptx/model';
import { OpcPackage } from '@pptx/opc';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const DIR = process.env.DECKS_DIR ?? join(ROOT, 'corpus/local');
const decks = existsSync(DIR) ? readdirSync(DIR).filter(f => /\.(pptx|pptm|potx|ppsx)$/i.test(f) && !f.startsWith('~$')).sort() : [];

/**
 * A text box to edit: its shape id, the groups to enter first, and where its first text is on screen. `under` names
 * an opaque shape drawn over that text: the text cannot be seen there, so it is not clicked (nor could it be in
 * PowerPoint).
 */
interface Target { id: number; name: string; groups: number[]; x: number; y: number; lineHeight: number; text: string; under?: string }

let server: ViteDevServer;
let browser: Browser;
let page: Page;
let url: string;
const errors: string[] = [];

const state = <T,>(fn: string): Promise<T> => page.evaluate(`(() => { const s = globalThis.__editor.getState(); ${fn} })()`) as Promise<T>;

/** Text of an element's a:t descendants, and the cNvPr id of a shape element: shared by the page-side helpers. */
const HELPERS = `
  const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
  const idOf = el => { const nv = el.children.find(c => c.kind === 'element' && c.local.startsWith('nv')); const c = nv && nv.children.find(x => x.kind === 'element' && x.local === 'cNvPr'); return Number(c && c.attrs.find(a => a.local === 'id').value); };
`;

/** The current slide's text boxes (in groups too), with the middle of each one's first piece of text on screen. */
function targets(): Promise<Target[]> {
  return state<Target[]>(`${HELPERS}
    const st = globalThis.__stage;
    const screen = (x, y) => { const q = st.getAbsoluteTransform().point({ x, y }); const r = st.container().getBoundingClientRect(); return { x: r.left + q.x, y: r.top + q.y }; };
    const scale = st.getAbsoluteTransform().getMatrix()[0];
    // every shape in drawing order, with the groups it is in
    const flat = [];
    const walk = (list, groups) => {
      for (const rs of list) {
        if (rs.layer !== 'slide' || !rs.source) continue;
        if (rs.source.kind === 'grpSp') walk(rs.children, [...groups, rs.id]);
        else flat.push({ rs, groups });
      }
    };
    walk(s.session.scenes.get(s.slide).slide.shapes, []);
    const inside = (rs, x, y) => {
      const m = rs.transform, det = m[0] * m[3] - m[1] * m[2] || 1;
      const u = (m[3] * (x - m[4]) - m[2] * (y - m[5])) / det, v = (m[0] * (y - m[5]) - m[1] * (x - m[4])) / det;
      return u >= 0 && v >= 0 && u <= rs.box.w && v <= rs.box.h;
    };
    const f = rs => rs.source.fill;
    const opaque = rs => rs.source.kind === 'pic' || rs.source.kind === 'graphicFrame'
      || (f(rs) && f(rs).kind !== 'none' && !(f(rs).kind === 'solid' && f(rs).color.a < 0.5));
    const out = [];
    flat.forEach(({ rs, groups }, i) => {
      const t = rs.prims.find(p => p.kind === 'text');
      if (!rs.source.view.txBody || !t) return;
      const line = t.layout.lines.find(l => l.runs.some(r => !r.bullet && r.end > r.start));
      if (!line) return;
      const run = line.runs.find(r => !r.bullet && r.end > r.start);
      // a little way into the first word, so the caret lands inside the text
      const lx = run.x + Math.min(run.advance / 2, 6 / scale), ly = line.top + line.height / 2, m = t.transform;
      const sx = m[0] * lx + m[2] * ly + m[4], sy = m[1] * lx + m[3] * ly + m[5];
      const cover = flat.slice(i + 1).find(o => opaque(o.rs) && inside(o.rs, sx, sy));
      const p = screen(sx, sy);
      out.push({ id: rs.id, name: rs.name, groups, x: p.x, y: p.y, lineHeight: line.height * scale, text: textOf(rs.source.view.el), ...(cover ? { under: cover.rs.name } : {}) });
    });
    return out;`);
}

/**
 * How far the caret's colour is from the slide's around it, from two screenshots of the same small region: with the
 * caret showing and with it gone (it blinks). The pixels that changed are the caret's; the most common colour with
 * it gone is the background. 0 when no caret was drawn.
 */
function caretContrast(on: Buffer, off: Buffer): number {
  const a = PNG.sync.read(on), b = PNG.sync.read(off);
  const median = (values: number[]) => values.sort((x, y) => x - y)[values.length >> 1] ?? 0;
  const background = [0, 1, 2].map(k => median(Array.from({ length: b.data.length / 4 }, (_, i) => b.data[i * 4 + k]!)));
  const caret: number[][] = [[], [], []];
  for (let i = 0; i < a.data.length; i += 4) {
    if (Math.max(...[0, 1, 2].map(k => Math.abs(a.data[i + k]! - b.data[i + k]!))) < 40) continue;
    for (const k of [0, 1, 2]) caret[k]!.push(a.data[i + k]!);
  }
  if (!caret[0]!.length) return 0;
  return Math.max(...[0, 1, 2].map(k => Math.abs(median(caret[k]!) - background[k]!)));
}

beforeAll(async () => {
  if (!decks.length) return;
  server = await createServer({ configFile: join(ROOT, 'apps/web/vite.config.ts'), root: join(ROOT, 'apps/web'), logLevel: 'error', server: { port: 0 } });
  await server.listen();
  url = server.resolvedUrls!.local[0]!;
  browser = await chromium.launch();
  page = await browser.newPage({ viewport: { width: 1400, height: 860 }, acceptDownloads: true });
  page.on('pageerror', e => errors.push(String(e)));
}, 120_000);

afterAll(async () => {
  await browser?.close();
  await server?.close();
});

describe.skipIf(decks.length === 0)('your decks in the editor', () => {
  test.each(decks)('%s', async name => {
    const bytes = readFileSync(join(DIR, name));
    const original = new PresentationDocument(OpcPackage.open(new Uint8Array(bytes)));
    await page.goto(url);
    await page.getByTestId('welcome-open').setInputFiles({ name, mimeType: 'application/octet-stream', buffer: bytes });
    await page.waitForFunction(() => (globalThis as unknown as { __editor: { getState(): { session: unknown } } }).__editor.getState().session);
    const parts = await state<string[]>('return s.session.ed.doc.slideParts();');
    expect(parts.length).toBe(original.slideParts().length);

    // the slide panel: Down through every slide keeps the current one in view, and End and Home jump
    await page.getByTestId('thumb-1').click();
    const inView = (n: number) => page.evaluate(n => {
      const panel = document.querySelector('aside.slides')!.getBoundingClientRect();
      const t = document.querySelector(`[data-testid="thumb-${n}"]`)!.getBoundingClientRect();
      return t.top >= panel.top - 1 && t.bottom <= panel.bottom + 1;
    }, n);
    for (let i = 1; i < parts.length; i++) {
      await page.keyboard.press('ArrowDown');
      expect(await state<string>('return s.slide;')).toBe(parts[i]);
      await expect.poll(() => inView(i + 1), { message: `slide ${i + 1} in view` }).toBe(true);
    }
    await page.keyboard.press('Home');
    await expect.poll(() => inView(1)).toBe(true);
    await page.keyboard.press('End');
    await expect.poll(() => inView(parts.length)).toBe(true);

    // every text box: double-click enters it, the caret shows, typing and deleting work
    const problems: string[] = [];
    const hidden: string[] = [];
    let boxes = 0;
    for (let i = 0; i < parts.length; i++) {
      await page.evaluate(p => (globalThis as unknown as { __editor: { getState(): { setSlide(p: string): void } } }).__editor.getState().setSlide(p), parts[i]!);
      await page.waitForTimeout(100);
      const list = await targets();
      for (const t of list) if (t.under) hidden.push(`slide ${i + 1} "${t.name}" (under "${t.under}")`);
      for (const [n, t] of list.filter(x => !x.under).entries()) {
        boxes++;
        const where = `slide ${i + 1} "${t.name}"`;
        await page.keyboard.press('Escape');
        await page.keyboard.press('Escape');
        for (let g = 0; g < t.groups.length; g++) await page.mouse.dblclick(t.x, t.y);
        await page.mouse.dblclick(t.x, t.y);
        const editing = await state<number | null>(`${HELPERS} return s.text ? idOf(s.text.el) : null;`);
        if (editing !== t.id) {
          problems.push(`${where}: double-click ${editing === null ? 'did not start editing' : `edited shape ${editing}`}`);
          continue;
        }
        if (!(await page.evaluate(() => !!(globalThis as unknown as { __stage: { findOne(s: string): unknown } }).__stage.findOne('.editing-frame')))) {
          problems.push(`${where}: no editing frame`);
        }
        if (n === 0) {
          // the caret blinks: one screenshot while it shows, one after it has gone; it must stand out from the slide
          const h = Math.max(8, Math.round(t.lineHeight));
          const clip = { x: Math.round(t.x - 10), y: Math.round(t.y - h / 2), width: 20, height: h };
          const on = await page.screenshot({ clip });
          await page.waitForTimeout(600);
          const off = await page.screenshot({ clip });
          const contrast = caretContrast(on, off);
          if (contrast < 80) problems.push(`${where}: the caret cannot be seen against the slide (contrast ${contrast} of 255)`);
        }
        await page.keyboard.type('Z');
        const typed = await state<string>(`${HELPERS} return s.text ? textOf(s.text.el) : '';`);
        if (typed.length !== t.text.length + 1 || !typed.includes('Z')) problems.push(`${where}: typing gave "${typed.slice(0, 60)}"`);
        await page.keyboard.press('Backspace');
        const back = await state<string>(`${HELPERS} return s.text ? textOf(s.text.el) : '';`);
        if (back !== t.text) problems.push(`${where}: Backspace gave "${back.slice(0, 60)}"`);
      }
    }
    await page.keyboard.press('Escape');
    if (hidden.length) console.log(`${name}: ${hidden.length} text boxes are hidden under other shapes, so not clicked:\n  ${hidden.join('\n  ')}`);
    expect(problems).toEqual([]);
    expect(boxes).toBeGreaterThan(0);

    // the deck saves, opens again, and says what it said before
    await page.getByRole('tab', { name: 'File' }).click();
    const [download] = await Promise.all([page.waitForEvent('download'), page.getByTestId('save').click()]);
    const saved = new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync((await download.path())!))));
    const words = (doc: PresentationDocument) => doc.slideParts().map(p => {
      const out: string[] = [];
      const walk = (e: ReturnType<typeof doc.root>) => {
        if (e.local === 't') out.push(e.children.map(c => (c.kind === 'text' ? c.value : '')).join(''));
        for (const c of e.children) if (c.kind === 'element') walk(c);
      };
      walk(doc.root(p));
      return out.join('');
    });
    expect(words(saved)).toEqual(words(original));
    expect(errors).toEqual([]);
  }, 20 * 60_000);
});
