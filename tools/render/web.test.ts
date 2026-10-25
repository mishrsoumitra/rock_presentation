// The editor app end to end, in headless Chromium: create, draw, drag, type, format, undo, save, encrypt,
// reopen, present, and edit a real deck. Saved files are checked in Node with the phase 1 oracle.
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium, type Browser, type Page } from 'playwright';
import { PNG } from 'pngjs';
import { createServer, type ViteDevServer } from 'vite';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { checkOrder, PresentationDocument, ShapeView } from '@pptx/model';
import { normalizeName, OpcPackage, resolveTarget } from '@pptx/opc';
import { duplicateSlide, Editor, insertTextBox } from '../../packages/editor-core/src/index.ts';
import { listEffects, slideTransition } from '../../packages/anim/src/index.ts';
import { blankPresentation } from '../../apps/web/src/engine/blank.ts';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
let server: ViteDevServer;
let browser: Browser;
let page: Page;
const errors: string[] = [];

beforeAll(async () => {
  server = await createServer({ configFile: join(ROOT, 'apps/web/vite.config.ts'), root: join(ROOT, 'apps/web'), logLevel: 'error', server: { port: 0 } });
  await server.listen();
  browser = await chromium.launch();
  page = await browser.newPage({ viewport: { width: 1400, height: 860 }, acceptDownloads: true });
  page.on('pageerror', e => errors.push(String(e)));
  await page.goto(server.resolvedUrls!.local[0]!);
}, 120_000);

afterAll(async () => {
  await browser?.close();
  await server?.close();
});

/** State read inside the page through the store the app exposes. */
const state = <T,>(fn: string): Promise<T> => page.evaluate(`(() => { const s = globalThis.__editor.getState(); ${fn} })()`) as Promise<T>;
const slideCount = () => state<number>('return s.session ? s.session.ed.doc.slideParts().length : 0;');
/** Text of the selected (or edited) shape, straight from its XML. */
const selectedText = () => state<string>(`
  const el = s.text?.el ?? s.selection[0]; if (!el) return '';
  const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); };
  walk(el); return out.join('');`);
/** Where a point of the slide (EMU) is on the page. */
const onScreen = (x: number, y: number) => page.evaluate(([x, y]) => {
  const st = (globalThis as unknown as { __stage: { getAbsoluteTransform(): { point(p: { x: number; y: number }): { x: number; y: number } }; container(): HTMLElement } }).__stage;
  const p = st.getAbsoluteTransform().point({ x, y });
  const r = st.container().getBoundingClientRect();
  return { x: r.left + p.x, y: r.top + p.y };
}, [x, y] as const);
const selectedXfrm = () => state<{ x: number; y: number; cx: number; cy: number }>(`
  const el = s.selection[0]; const spPr = el.children.find(c => c.local === 'spPr'); const x = spPr.children.find(c => c.local === 'xfrm');
  const off = x.children.find(c => c.local === 'off'), ext = x.children.find(c => c.local === 'ext');
  const a = (e, n) => Number(e.attrs.find(t => t.local === n).value);
  return { x: a(off, 'x'), y: a(off, 'y'), cx: a(ext, 'cx'), cy: a(ext, 'cy') };`);

async function download(): Promise<Uint8Array> {
  await page.getByRole('tab', { name: 'File' }).click();
  const [d] = await Promise.all([page.waitForEvent('download'), page.getByTestId('save').click()]);
  return new Uint8Array(readFileSync((await d.path())!));
}

/** The phase 1 oracle on a saved file: it opens, every slide is in schema order, every reference resolves. */
function oracle(bytes: Uint8Array): PresentationDocument {
  const pkg = OpcPackage.open(bytes);
  const doc = new PresentationDocument(pkg);
  for (const part of doc.slideParts()) {
    const root = doc.root(part);
    expect(checkOrder(root), part).toEqual([]);
    const ids = new Set(pkg.relationships(part).map(r => r.id));
    const walk = (e: typeof root) => {
      for (const a of e.attrs) if (a.ns.endsWith('/relationships') && a.ns.includes('officeDocument')) expect(ids.has(a.value), `${part} ${a.local}=${a.value}`).toBe(true);
      for (const c of e.children) if (c.kind === 'element') walk(c);
    };
    walk(root);
    for (const r of pkg.relationships(part)) if (!r.external) expect(pkg.getPart(resolveTarget(normalizeName(part), r.target)), r.target).toBeDefined();
  }
  return doc;
}

describe('the editor app', () => {
  test('a new presentation: draw a rectangle, drag it, undo, type, format, add slides, save', async () => {
    await page.getByTestId('welcome-new').click();
    await page.getByTestId('canvas').waitFor();
    await page.waitForFunction(() => !!(globalThis as unknown as { __stage?: unknown }).__stage);
    expect(await slideCount()).toBe(1);

    // pick a rectangle, then draw it by dragging on the slide, as in PowerPoint
    await page.getByTestId('shapes').click();
    await page.locator('[data-testid="shape-gallery"] button[data-prst="rect"]').click();
    const from = await onScreen(2000000, 2000000), to = await onScreen(4000000, 3500000);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(to.x, to.y, { steps: 5 });
    await page.mouse.up();
    expect(await state<number>('return s.selection.length;')).toBe(1);
    const before = await selectedXfrm();
    const z0 = await page.evaluate(() => (globalThis as unknown as { __stage: { scaleX(): number } }).__stage.scaleX());
    expect(Math.abs(before.x - 2000000)).toBeLessThan(2 / z0);
    expect(Math.abs(before.cx - 2000000)).toBeLessThan(2 / z0);
    expect(Math.abs(before.cy - 1500000)).toBeLessThan(2 / z0);
    // the thumbnail (drawn by the worker from its own copy of the deck) shows the new blue rectangle
    await expect.poll(() => page.evaluate(`(() => {
      const c = document.querySelector('[data-testid="thumb-1"] canvas');
      const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
      let blue = 0;
      for (let i = 0; i < d.length; i += 4) if (d[i] < 120 && d[i + 2] > 150 && d[i + 3] > 0) blue++;
      return blue;
    })()`), { timeout: 15_000 }).toBeGreaterThan(50);

    // drag by (120, 60) screen pixels, holding Alt so smart guides do not snap it (as in PowerPoint)
    const c = await onScreen(before.x + before.cx / 2, before.y + before.cy / 2);
    await page.mouse.move(c.x, c.y);
    await page.keyboard.down('Alt');
    await page.mouse.down();
    await page.mouse.move(c.x + 60, c.y + 30, { steps: 4 });
    await page.mouse.move(c.x + 120, c.y + 60, { steps: 4 });
    await page.mouse.up();
    await page.keyboard.up('Alt');
    const zoom = await page.evaluate(() => (globalThis as unknown as { __stage: { scaleX(): number } }).__stage.scaleX());
    const moved = await selectedXfrm();
    expect(Math.abs(moved.x - before.x - 120 / zoom)).toBeLessThan(3 / zoom);
    expect(Math.abs(moved.y - before.y - 60 / zoom)).toBeLessThan(3 / zoom);
    await page.keyboard.press('Control+z');
    expect(await selectedXfrm()).toEqual(before);
    await page.keyboard.press('Control+y');
    expect(await selectedXfrm()).toEqual(moved);

    // double-click, type, leave
    const c2 = await onScreen(moved.x + moved.cx / 2, moved.y + moved.cy / 2);
    await page.mouse.dblclick(c2.x, c2.y);
    await page.getByTestId('text-input').waitFor({ state: 'attached' });
    await expect.poll(() => page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('text-input');
    await page.keyboard.type('Hello world');
    expect(await selectedText()).toBe('Hello world');
    await page.keyboard.press('Escape');
    expect(await state<boolean>('return s.text === null && s.selection.length === 1;')).toBe(true);
    await page.getByTestId('bold').click();
    expect(await state<string>('return JSON.stringify(s.selection[0].children.find(c => c.local === "txBody").children.find(c => c.local === "p").children.find(c => c.local === "r").children[0].attrs.map(a => a.local + "=" + a.value));')).toContain('b=1');
    // typing undoes as one step
    await page.keyboard.press('Control+z');
    expect(await selectedText()).toBe('Hello world');
    await page.keyboard.press('Control+z');
    expect(await selectedText()).toBe('');
    await page.keyboard.press('Control+y');
    await page.keyboard.press('Control+y');

    await page.getByTestId('new-slide').click();
    expect(await slideCount()).toBe(2);
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.getByTestId('text-box').click();
    await page.keyboard.type('Second slide');
    await page.keyboard.press('Escape');
    expect(await selectedText()).toBe('Second slide');

    const saved = await download();
    const doc = oracle(saved);
    expect(doc.slideParts()).toHaveLength(2);
    const first = doc.slide(doc.slideParts()[0]!).shapes.find(s => s.geometry?.kind === 'preset' && s.geometry.prst === 'rect' && !s.placeholder)!;
    expect(first).toBeDefined();
    expect(JSON.stringify(first.txBody ? new ShapeView(first.el).txBody!.children.length : 0)).not.toBe('0');
    expect(errors).toEqual([]);
  }, 120_000);

  test('pictures: insert one, crop it to 1:1 (one undo step), and the frame keeps the picture undistorted', async () => {
    // a 200x100 PNG, half red, half blue
    const img = new PNG({ width: 200, height: 100 });
    for (let i = 0; i < 200 * 100; i++) img.data.set(i % 200 < 100 ? [220, 40, 40, 255] : [40, 40, 220, 255], i * 4);
    const png = PNG.sync.write(img);
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.locator('input[type=file][accept^="image/png"]').first().setInputFiles({ name: 'wide.png', mimeType: 'image/png', buffer: png });
    await expect.poll(() => state<string>('return s.selection[0]?.local ?? "";')).toBe('pic');
    const before = await selectedXfrm();
    expect(Math.round(before.cx / before.cy)).toBe(2);
    await page.getByRole('tab', { name: 'Home' }).click();
    await page.getByRole('button', { name: 'Crop' }).click();
    await page.getByRole('menuitem', { name: 'Aspect Ratio Square 1:1' }).click();
    const after = await selectedXfrm();
    expect(after.cx).toBe(before.cx);
    expect(Math.abs(after.cy - after.cx)).toBeLessThanOrEqual(1);
    expect(await state<string>('const f = s.selection[0].children.find(c => c.local === "blipFill"); const r = f.children.find(c => c.local === "srcRect"); return r ? r.attrs.map(a => a.local + "=" + a.value).join(" ") : "";'))
      .toBe('l=25000 r=25000');
    await page.keyboard.press('Control+z');
    expect(await selectedXfrm()).toEqual(before);
    expect(errors).toEqual([]);
  }, 60_000);

  test('animations: Fade a shape from the gallery, add a transition, play it, and save both', async () => {
    // a new rectangle away from the others, on slide 1
    await page.getByTestId('thumb-1').click();
    await page.getByRole('tab', { name: 'Home' }).click();
    await page.getByTestId('shapes').click();
    await page.locator('[data-testid="shape-gallery"] button[data-prst="rect"]').click();
    const a = await onScreen(7000000, 4500000), b = await onScreen(10000000, 6000000);
    await page.mouse.move(a.x, a.y);
    await page.mouse.down();
    await page.mouse.move(b.x, b.y, { steps: 4 });
    await page.mouse.up();
    await page.getByRole('tab', { name: 'Animations' }).click();
    await page.getByTestId('animation-gallery').click();
    await page.locator('[data-testid="effect-gallery"] button[data-effect="entr:10"]').click();
    await page.getByTestId('animation-pane-toggle').click();
    await expect.poll(() => page.getByTestId('animation-pane').locator('li').count()).toBe(1);
    expect(await page.getByTestId('animation-pane').locator('li .what').textContent()).toMatch(/^Fade: Rectangle/);
    await page.getByRole('tab', { name: 'Transitions' }).click();
    await page.getByTestId('transition-kind').click();
    await page.locator('button[data-kind="push"]').click();
    expect(await state<string>('return JSON.stringify(s.session.ed.slide(s.slide).transition() ? "yes" : "no");')).toBe('"yes"');

    // the slide show: the rectangle waits for the click, then fades in
    await page.getByRole('tab', { name: 'Slide Show' }).click();
    await page.getByTestId('present').click();
    await page.getByTestId('presenter').waitFor();
    const pixel = () => page.evaluate(() => {
      const c = document.querySelector('[data-testid="presenter"] canvas') as HTMLCanvasElement;
      const d = c.getContext('2d')!.getImageData(Math.round(c.width * (8500000 / 12192000)), Math.round(c.height * (5250000 / 6858000)), 1, 1).data;
      return [d[0], d[1], d[2]];
    });
    // the push transition into the first slide is not played (it is the first); wait for the first frames
    await page.waitForTimeout(400);
    expect(await pixel()).toEqual([255, 255, 255]);
    await page.mouse.click(20, 20);
    await page.waitForTimeout(900);
    const [r, g, bl] = await pixel();
    expect(bl! - r!).toBeGreaterThan(80);                       // accent 1 blue, #4472C4
    expect(g).toBeLessThan(160);
    await page.keyboard.press('Escape');
    await page.getByTestId('canvas').waitFor();

    const doc = oracle(await download());
    const view = doc.slide(doc.slideParts()[0]!);
    expect(listEffects(view).map(e => e.name)).toEqual(['Fade']);
    expect(slideTransition(view)?.kind).toBe('push');
    expect(errors).toEqual([]);
  }, 120_000);

  test('encrypt with a password, save, and open again with it', async () => {
    await page.getByRole('tab', { name: 'File' }).click();
    await page.getByRole('button', { name: /Encrypt with password/ }).click();
    await page.getByLabel('Password', { exact: true }).fill('s3cret!');
    await page.getByLabel('Reenter password').fill('s3cret!');
    await page.getByRole('button', { name: 'OK' }).click();
    const saved = await download();
    expect([...saved.subarray(0, 4)]).toEqual([0xd0, 0xcf, 0x11, 0xe0]);
    await page.getByRole('tab', { name: 'File' }).click();
    await page.locator('input[type=file][accept^=".pptx"]').first().setInputFiles({ name: 'secret.pptx', mimeType: 'application/octet-stream', buffer: Buffer.from(saved) });
    await page.getByTestId('password').fill('wrong');
    await page.keyboard.press('Enter');
    // a wrong password shows only after the key derivation (100,000 rounds of SHA-512), slow on a CI runner
    await expect.poll(() => page.getByText('The password is not correct').count(), { timeout: 20_000 }).toBe(1);
    await page.getByTestId('password').fill('s3cret!');
    await page.keyboard.press('Enter');
    await expect.poll(slideCount, { timeout: 20_000 }).toBe(2);
    expect(errors).toEqual([]);
  }, 120_000);

  test('the slide show opens full window and Escape ends it', async () => {
    await page.getByRole('tab', { name: 'Slide Show' }).click();
    await page.getByTestId('present').click();
    await page.getByTestId('presenter').waitFor();
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Escape');
    await page.getByTestId('canvas').waitFor();
    expect(errors).toEqual([]);
  }, 60_000);

  const realDeck = join(ROOT, 'corpus/external/python-pptx/features~steps~test_files~shp-shapes.pptx');
  test.skipIf(!existsSync(realDeck))('a real deck: nudge its first shape, save, and the file still passes the oracle', async () => {
    const file = realDeck;
    await page.getByRole('tab', { name: 'File' }).click();
    await page.locator('input[type=file][accept^=".pptx"]').first().setInputFiles(file);
    await expect.poll(() => state<string>('return s.session?.name;'), { timeout: 20_000 }).toBe('features~steps~test_files~shp-shapes.pptx');
    await page.waitForFunction(() => !!(globalThis as unknown as { __stage?: unknown }).__stage);
    await state<void>('const sc = s.session.scenes.get(s.slide); const el = sc.slide.shapes.find(x => x.layer === "slide" && x.source && x.source.view.xfrm)?.source.view.el; s.select(el ? [el] : []);');
    const before = await selectedXfrm();
    await page.locator('body').click({ position: { x: 5, y: 840 } });
    await state<void>('const sc = s.session.scenes.get(s.slide); const el = sc.slide.shapes.find(x => x.layer === "slide" && x.source && x.source.view.xfrm)?.source.view.el; s.select(el ? [el] : []);');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    expect((await selectedXfrm()).x).toBe(before.x + 2 * 76200);
    const doc = oracle(await download());
    expect(doc.slideParts().length).toBeGreaterThan(0);
    expect(errors).toEqual([]);
  }, 120_000);

  test('performance (plan 4.6): a 300-slide, 140 MB deck opens in under 3 s; redrawing a slide fits a 60 fps frame', async () => {
    const ed = new Editor(OpcPackage.open(blankPresentation()));
    const first = ed.doc.slideParts()[0]!;
    ed.run(insertTextBox(ed, first, { x: 500000, y: 500000, w: 6000000, h: 1000000 }, 'A slide with text').command);
    // the plan's numbers are for a laptop: on a shared CI runner the test opens a smaller deck and times nothing
    const ci = !!process.env.CI;
    const big = new Uint8Array((ci ? 20 : 140) * 1024 * 1024);
    for (let i = 0; i < big.length; i += 65536) crypto.getRandomValues(big.subarray(i, Math.min(big.length, i + 65536)));
    ed.pkg.ensureDefaultContentType('bin', 'application/octet-stream');
    ed.pkg.addPart('/ppt/media/payload1.bin', 'application/octet-stream', big);
    ed.pkg.addRelationship(first, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/media', '/ppt/media/payload1.bin');
    for (let i = 0; i < 299; i++) ed.run(duplicateSlide(ed, first).command);
    const dir = mkdtempSync(join(tmpdir(), 'pptx-perf-'));
    const file = join(dir, 'big.pptx');
    writeFileSync(file, ed.save());
    // a fresh page, so the decks the earlier tests opened are not in memory too
    await page.close();
    page = await browser.newPage({ viewport: { width: 1400, height: 860 }, acceptDownloads: true });
    page.on('pageerror', e => errors.push(String(e)));
    await page.goto(server.resolvedUrls!.local[0]!);
    await page.getByTestId('welcome-new').click();
    await page.getByTestId('canvas').waitFor();
    await page.getByRole('tab', { name: 'File' }).click();
    await page.locator('input[type=file][accept^=".pptx"]').first().setInputFiles(file);
    rmSync(dir, { recursive: true, force: true });
    await page.waitForFunction(() => (globalThis as unknown as { __editor: { getState(): { session?: { name: string } } } }).__editor.getState().session?.name === 'big.pptx', null, { timeout: ci ? 150_000 : 60_000 })
      .catch(async e => { throw new Error(`${e}; app says: ${await state<string>('return JSON.stringify(s.message);')}`); });
    const firstFrame = await page.evaluate(() => new Promise<number>(resolve => requestAnimationFrame(() => requestAnimationFrame(() =>
      resolve(performance.now() - performance.getEntriesByName('open-start').at(-1)!.startTime)))));
    expect(await slideCount()).toBe(300);
    if (!ci) expect(firstFrame).toBeLessThan(3000);
    const redraw = await page.evaluate(() => {
      const layer = (globalThis as unknown as { __stage: { getLayers(): { draw(): void }[] } }).__stage.getLayers()[0]!;
      const times: number[] = [];
      for (let i = 0; i < 21; i++) { const a = performance.now(); layer.draw(); times.push(performance.now() - a); }
      return times.sort((x, y) => x - y)[10]!;
    });
    if (!ci) expect(redraw).toBeLessThan(16);
    expect(errors).toEqual([]);
  }, 300_000);

  test('placeholders take clicks where they are: the title and body of a new slide, at Fit and at 200% zoom', async () => {
    // The shapes' Konva nodes do not scale their strokes with the stage (strokeScaleEnabled is off), so Konva measures
    // hitStrokeWidth in screen pixels. Dividing it by the zoom once made every node's hit band cover the whole canvas,
    // and the node drawn last took every click: the title of a new slide could not be clicked into.
    type Hit = { id: string; x: number; y: number; w: number; h: number } | null;
    type StageLike = {
      getAbsoluteTransform(): { point(p: { x: number; y: number }): { x: number; y: number } };
      container(): HTMLElement;
      scaleX(): number;
      getIntersection(p: { x: number; y: number }): { id(): string; x(): number; y(): number; width(): number; height(): number } | null;
    };
    /** The slide's own shapes: placeholder type, the centre and size of the box on the slide (EMU), and the text. */
    const shapesHere = () => state<{ type: string; name: string; x: number; y: number; w: number; h: number; text: string }[]>(`
      const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
      return s.session.scenes.get(s.slide).slide.shapes.filter(x => x.layer === 'slide' && x.source).map(rs => {
        const m = rs.transform, u = rs.box.w / 2, v = rs.box.h / 2;
        return { type: rs.source.placeholder ? rs.source.placeholder.type : '', name: rs.name, x: m[0] * u + m[2] * v + m[4], y: m[1] * u + m[3] * v + m[5], w: rs.box.w, h: rs.box.h, text: textOf(rs.source.view.el) };
      });`);
    /** The node Konva's hit canvas finds at a point of the slide (EMU): its centre and size in slide units. */
    const hitAt = (x: number, y: number) => page.evaluate(([x, y]) => {
      const st = (globalThis as unknown as { __stage: StageLike }).__stage;
      const n = st.getIntersection(st.getAbsoluteTransform().point({ x, y }));
      return (n ? { id: n.id(), x: n.x(), y: n.y(), w: n.width(), h: n.height() } : null) as Hit;
    }, [x, y] as const);
    /** Click a point of the slide (EMU), first scrolling the canvas so it is in view: at 200% the slide outgrows the window. */
    const clickSlide = async (x: number, y: number) => {
      const at = await page.evaluate(([x, y]) => {
        const st = (globalThis as unknown as { __stage: StageLike }).__stage;
        const host = document.querySelector('[data-testid="canvas"]') as HTMLElement;
        const p = st.getAbsoluteTransform().point({ x, y });
        host.scrollLeft = p.x - host.clientWidth / 2;
        host.scrollTop = p.y - host.clientHeight / 2;
        const r = st.container().getBoundingClientRect();
        return { x: r.left + p.x, y: r.top + p.y };
      }, [x, y] as const);
      await page.mouse.click(at.x, at.y);
    };
    const drawn = () => page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));

    for (const zoomButton of ['Fit', '200%'] as const) {
      await page.goto(server.resolvedUrls!.local[0]!);
      await page.getByTestId('welcome-new').click();
      await page.getByTestId('canvas').waitFor();
      await page.waitForFunction(() => !!(globalThis as unknown as { __stage?: unknown }).__stage);
      await page.getByTestId('new-slide').click();
      expect(await slideCount()).toBe(2);
      await page.locator('footer.status .zoom').getByRole('button', { name: zoomButton, exact: true }).click();
      if (zoomButton === '200%') {
        await page.waitForFunction(z => Math.abs((globalThis as unknown as { __stage: StageLike }).__stage.scaleX() - z) < 1e-9, (2 * 96) / 914400);
      }
      await drawn();

      // the new slide (Title and Content): an empty title and an empty body, both without fill
      const shapes = await shapesHere();
      const title = shapes.find(sh => sh.type === 'title'), body = shapes.find(sh => sh.type !== 'title');
      expect(shapes.map(sh => sh.text), zoomButton).toEqual(['', '']);
      expect(title, zoomButton).toBeDefined();
      expect(body, zoomButton).toBeDefined();

      // the hit canvas: the centre of each shape finds that shape's own node, not whichever node was drawn last
      const hits: Hit[] = [];
      for (const sh of shapes) {
        const n = await hitAt(sh.x, sh.y);
        hits.push(n);
        expect(n, `${zoomButton}: nothing under the centre of ${sh.name}`).not.toBeNull();
        expect(Math.abs(n!.x - sh.x), `${zoomButton}: the centre of ${sh.name} hits a node centred elsewhere`).toBeLessThan(1);
        expect(Math.abs(n!.y - sh.y), `${zoomButton}: the centre of ${sh.name} hits a node centred elsewhere`).toBeLessThan(1);
        expect(Math.abs(n!.w - sh.w), `${zoomButton}: the centre of ${sh.name} hits a node of another size`).toBeLessThan(1);
        expect(Math.abs(n!.h - sh.h), `${zoomButton}: the centre of ${sh.name} hits a node of another size`).toBeLessThan(1);
      }
      expect(new Set(hits.map(n => n?.id)).size, `${zoomButton}: every shape has its own node`).toBe(shapes.length);

      // click the centre of the title and type: the words go into the title, and the body stays empty
      await clickSlide(title!.x, title!.y);
      expect(await state<number>('return s.selection.length;'), zoomButton).toBe(1);
      await page.keyboard.type('Title test');
      let after = await shapesHere();
      expect(after.find(sh => sh.type === 'title')!.text, zoomButton).toBe('Title test');
      expect(after.find(sh => sh.type !== 'title')!.text, zoomButton).toBe('');
      await page.keyboard.press('Escape');

      // click the centre of the body and type: this goes into the body, and the title keeps its words
      await clickSlide(body!.x, body!.y);
      await page.keyboard.type('Body test');
      after = await shapesHere();
      expect(after.find(sh => sh.type === 'title')!.text, zoomButton).toBe('Title test');
      expect(after.find(sh => sh.type !== 'title')!.text, zoomButton).toBe('Body test');
      await page.keyboard.press('Escape');
    }
    expect(errors).toEqual([]);
  }, 120_000);

  test('the slide panel keeps the current slide in view: Down through forty slides, then Home and End', async () => {
    // Pressing Down past the panel's bottom once left the current tile out of sight: the panel did not scroll with
    // the keyboard. Each thumbnail now calls scrollIntoView when its slide becomes current.
    const ed = new Editor(OpcPackage.open(blankPresentation()));
    const first = ed.doc.slideParts()[0]!;
    for (let i = 0; i < 39; i++) ed.run(duplicateSlide(ed, first).command);
    const dir = mkdtempSync(join(tmpdir(), 'pptx-panel-'));
    const file = join(dir, 'forty.pptx');
    writeFileSync(file, ed.save());
    await page.goto(server.resolvedUrls!.local[0]!);
    await page.getByTestId('welcome-new').click();
    await page.getByTestId('canvas').waitFor();
    await page.getByRole('tab', { name: 'File' }).click();
    await page.locator('input[type=file][accept^=".pptx"]').first().setInputFiles(file);
    rmSync(dir, { recursive: true, force: true });
    await page.waitForFunction(() => (globalThis as unknown as { __editor: { getState(): { session?: { name: string } } } }).__editor.getState().session?.name === 'forty.pptx');
    expect(await slideCount()).toBe(40);
    /** Whether the current slide's tile lies wholly inside the panel's box (what scrolls out of it is clipped). */
    const currentInView = () => page.evaluate(() => {
      const panel = document.querySelector('aside.slides')!.getBoundingClientRect();
      const t = document.querySelector('.thumb.current')!.getBoundingClientRect();
      return t.top >= panel.top - 1 && t.bottom <= panel.bottom + 1 && t.left >= panel.left - 1 && t.right <= panel.right + 1;
    });
    const index = () => state<number>('return s.session.ed.doc.slideParts().indexOf(s.slide);');
    await page.getByTestId('thumb-1').click();                // the click also gives the panel the keyboard focus
    for (let i = 1; i < 40; i++) {
      await page.keyboard.press('ArrowDown');
      expect(await index()).toBe(i);
      await expect.poll(currentInView, { message: `slide ${i + 1} in view after Down` }).toBe(true);
    }
    await page.keyboard.press('Home');
    expect(await index()).toBe(0);
    await expect.poll(currentInView, { message: 'slide 1 in view after Home' }).toBe(true);
    await page.keyboard.press('End');
    expect(await index()).toBe(39);
    await expect.poll(currentInView, { message: 'slide 40 in view after End' }).toBe(true);
    expect(errors).toEqual([]);
  }, 120_000);
});

describe('text editing with the mouse, as in Google Slides', () => {
  type Pt = { x: number; y: number };
  type Pos = { p: number; offset: number };
  type Box = { name: string; x: number; y: number; w: number; h: number; text: string };

  /** Page-side helpers: the current slide's shapes, points of the slide on the page, names and text of elements. */
  const HELP = `
    const st = globalThis.__stage;
    const screen = (x, y) => { const q = st.getAbsoluteTransform().point({ x, y }); const r = st.container().getBoundingClientRect(); return { x: r.left + q.x, y: r.top + q.y }; };
    const nameOf = el => { const nv = el.children.find(c => c.kind === 'element' && c.local.startsWith('nv')); const c = nv && nv.children.find(x => x.kind === 'element' && x.local === 'cNvPr'); const a = c && c.attrs.find(a => a.local === 'name'); return a ? a.value : ''; };
    const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
    const own = s.session.scenes.get(s.slide).slide.shapes.filter(x => x.layer === 'slide' && x.source);
    const FOOT = ['dt', 'ftr', 'sldNum'];
    const find = which => which === 'textbox' ? own.find(rs => !rs.source.placeholder)
      : which === 'body' ? own.find(rs => rs.source.placeholder && !['title', 'ctrTitle', ...FOOT].includes(rs.source.placeholder.type))
      : own.find(rs => rs.source.placeholder && rs.source.placeholder.type === which);
    const corner = (rs, u, v) => { const m = rs.transform; return screen(m[0] * u + m[2] * v + m[4], m[1] * u + m[3] * v + m[5]); };
  `;
  /** Where the caret before `offset` of a box's paragraph (the first unless given) is on the page, mid-line. */
  const caretAt = (which: string, offset: number, paragraph = 0) => state<Pt>(`${HELP}
    const rs = find(${JSON.stringify(which)}); const t = rs.prims.filter(p => p.kind === 'text').pop();
    const line = t.layout.lines.find(l => l.paragraph === ${paragraph} && l.carets.some(c => c.offset === ${offset}));
    const c = line.carets.find(c => c.offset === ${offset}); const m = t.transform; const y = line.top + line.height / 2;
    return screen(m[0] * c.x + m[2] * y + m[4], m[1] * c.x + m[3] * y + m[5]);`);
  /** A box on the page: its name, its top left corner and size, and its text. */
  const boxOf = (which: string) => state<Box>(`${HELP}
    const rs = find(${JSON.stringify(which)}); const a = corner(rs, 0, 0), b = corner(rs, rs.box.w, rs.box.h);
    return { name: rs.name, x: a.x, y: a.y, w: b.x - a.x, h: b.y - a.y, text: textOf(rs.source.view.el) };`);
  /** What is being edited: the box's name, the caret and the anchor; null when nothing is. */
  const editing = () => state<{ box: string; caret: Pos; anchor: Pos } | null>(`${HELP}
    return s.text ? { box: nameOf(s.text.el), caret: s.text.caret, anchor: s.text.anchor } : null;`);
  const selected = () => state<number>('return s.selection.length;');
  const zoomNow = () => page.evaluate(() => (globalThis as unknown as { __stage: { scaleX(): number } }).__stage.scaleX());
  const drag = async (a: Pt, b: Pt) => {
    await page.mouse.move(a.x, a.y);
    await page.mouse.down();
    await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 3 });
    await page.mouse.move(b.x, b.y, { steps: 3 });
    await page.mouse.up();
  };
  const drawn = () => page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
  /** Pixels of the selection highlight (its blue over the white slide) inside a rectangle of the page. */
  const highlight = async (r: { x: number; y: number; w: number; h: number }) => {
    await drawn();
    const png = PNG.sync.read(await page.screenshot({ clip: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.w), height: Math.round(r.h) } }));
    let n = 0;
    for (let i = 0; i < png.data.length; i += 4) {
      if (Math.abs(png.data[i]! - 202) < 14 && Math.abs(png.data[i + 1]! - 221) < 14 && png.data[i + 2]! > 240) n++;
    }
    return n;
  };
  let empty: Pt;

  beforeAll(async () => {
    await page.goto(server.resolvedUrls!.local[0]!);
    await page.getByTestId('welcome-new').click();
    await page.getByTestId('canvas').waitFor();
    await page.waitForFunction(() => !!(globalThis as unknown as { __stage?: unknown }).__stage);
    empty = await onScreen(300000, 300000);           // the top left corner of the slide: no shape of either layout is there
  }, 60_000);

  test('one click on an empty placeholder starts editing it: a new slide, a click on its title, and the title says Hello', async () => {
    await page.getByTestId('new-slide').click();
    expect(await slideCount()).toBe(2);
    await drawn();
    const title = await boxOf('title');
    expect(title.text).toBe('');
    await page.mouse.click(title.x + title.w / 2, title.y + title.h / 2);
    expect(await editing()).toEqual({ box: title.name, caret: { p: 0, offset: 0 }, anchor: { p: 0, offset: 0 } });
    await page.keyboard.type('Hello');
    expect((await boxOf('title')).text).toBe('Hello');
    expect((await boxOf('body')).text).toBe('');
    await page.keyboard.press('Escape');
    expect(await editing()).toBeNull();
    expect(await selected()).toBe(1);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a text box: one click puts the caret where you click, the mouse selects, Shift+click extends, a click never moves the box', async () => {
    await page.getByTestId('thumb-1').click();
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.getByTestId('text-box').click();
    await page.keyboard.type('Learning It for Good');
    await page.keyboard.press('Escape');
    await page.getByRole('tab', { name: 'Home' }).click();
    await page.mouse.click(empty.x, empty.y);
    expect(await selected()).toBe(0);
    await drawn();
    expect((await boxOf('textbox')).text).toBe('Learning It for Good');

    // one click before "It": the caret at offset 9 and editing, with no selection
    const at9 = await caretAt('textbox', 9);
    await page.mouse.click(at9.x + 1, at9.y);
    expect(await editing()).toMatchObject({ caret: { p: 0, offset: 9 }, anchor: { p: 0, offset: 9 } });
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('text-input');
    // a click elsewhere while editing moves the caret
    const at15 = await caretAt('textbox', 15);
    await page.mouse.click(at15.x + 1, at15.y);
    expect(await editing()).toMatchObject({ caret: { p: 0, offset: 15 }, anchor: { p: 0, offset: 15 } });

    // dragging from the start of "Learning" to its end selects it (from a box that was not selected), and X replaces it
    await page.mouse.click(empty.x, empty.y);
    const at0 = await caretAt('textbox', 0), at8 = await caretAt('textbox', 8);
    await drag({ x: at0.x + 1, y: at0.y }, at8);
    expect(await editing()).toMatchObject({ anchor: { p: 0, offset: 0 }, caret: { p: 0, offset: 8 } });
    await page.keyboard.type('X');
    expect((await boxOf('textbox')).text).toBe('X It for Good');
    await page.keyboard.press('Control+z');
    expect((await boxOf('textbox')).text).toBe('Learning It for Good');

    // Shift+click extends the selection from the anchor
    await page.mouse.click(at9.x + 1, at9.y);
    const at12 = await caretAt('textbox', 12);
    await page.keyboard.down('Shift');
    await page.mouse.click(at12.x + 1, at12.y);
    await page.keyboard.up('Shift');
    expect(await editing()).toMatchObject({ anchor: { p: 0, offset: 9 }, caret: { p: 0, offset: 12 } });

    // one click on the text does not move the box, whether the box was selected or not
    await page.keyboard.press('Escape');
    expect(await selected()).toBe(1);
    const before = await selectedXfrm();
    const at3 = await caretAt('textbox', 3);
    await page.mouse.click(at3.x + 1, at3.y);                 // selected: a click on its text starts editing
    expect(await editing()).toMatchObject({ caret: { p: 0, offset: 3 } });
    await page.keyboard.press('Escape');
    await page.mouse.click(empty.x, empty.y);
    await page.mouse.click(at3.x + 1, at3.y);                 // not selected: the same
    expect(await editing()).toMatchObject({ caret: { p: 0, offset: 3 } });
    await page.keyboard.press('Escape');
    expect(await selectedXfrm()).toEqual(before);
    // but dragging its frame does (the top edge of the selected box, away from the resize handle at its middle);
    // Alt keeps the smart guides from snapping
    const box = await boxOf('textbox');
    const grab = { x: box.x + box.w / 4, y: box.y };
    await page.keyboard.down('Alt');
    await drag(grab, { x: grab.x, y: grab.y + 50 });
    await page.keyboard.up('Alt');
    const zoom = await zoomNow();
    const moved = await selectedXfrm();
    expect(Math.abs(moved.y - before.y - 50 / zoom)).toBeLessThan(3 / zoom);
    expect(Math.abs(moved.x - before.x)).toBeLessThan(3 / zoom);
    expect(await editing()).toBeNull();
    expect(errors).toEqual([]);
  }, 90_000);

  test('typing 50 characters in a row: the text is right; locally, the time per keystroke is reported', async () => {
    await page.mouse.click(empty.x, empty.y);
    const end = await caretAt('textbox', 20);
    await page.mouse.click(end.x - 1, end.y);
    expect(await editing()).toMatchObject({ caret: { p: 0, offset: 20 }, anchor: { p: 0, offset: 20 } });
    await page.evaluate(() => performance.clearMeasures('keystroke'));
    const typed = 'abcdefghij'.repeat(5);
    await page.keyboard.type(typed);
    expect((await boxOf('textbox')).text).toBe(`Learning It for Good${typed}`);
    const times = await page.evaluate(() => performance.getEntriesByName('keystroke').map(e => e.duration).sort((a, b) => a - b));
    expect(times).toHaveLength(50);
    if (!process.env.CI) {
      // the app's own measure: the keystroke's command, autofit, scene rebuild, React render and commit (dev build)
      console.log(`keystroke: median ${times[25]!.toFixed(1)} ms, p90 ${times[45]!.toFixed(1)} ms, max ${times[49]!.toFixed(1)} ms (dev build, headless Chromium)`);
    }
    await page.keyboard.press('Control+z');                   // the fifty went in as one step
    expect((await boxOf('textbox')).text).toBe('Learning It for Good');
    await page.keyboard.press('Escape');
    expect(errors).toEqual([]);
  }, 60_000);
});
