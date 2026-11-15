// Features the editor shares with Google Slides and PowerPoint, each driven in headless Chromium the way a person
// would use it, and each checked in the deck's XML and in the saved file. Every test starts from a new presentation.
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium, type Browser, type Page } from 'playwright';
import { PNG } from 'pngjs';
import { createServer, type ViteDevServer } from 'vite';
import { afterAll, beforeAll, beforeEach, describe, expect, test } from 'vitest';
import { OpcPackage, readZip } from '@pptx/opc';
import { cellText, Editor, notesText, slideComments, tableOf, tableSize } from '../../packages/editor-core/src/index.ts';

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
  // leaving a deck with unsaved changes asks first: the tests always leave
  page.on('dialog', d => void d.accept().catch(() => undefined));
}, 120_000);

afterAll(async () => {
  await browser?.close();
  await server?.close();
});

/** A new presentation in a freshly loaded app: one title slide, nothing selected, an empty undo history. */
beforeEach(async () => {
  errors.length = 0;
  await page.goto(server.resolvedUrls!.local[0]!);
  // nothing kept by autosave from the test before: the app starts as it does the first time
  await page.evaluate(() => new Promise(done => { const q = indexedDB.deleteDatabase('presentations'); q.onsuccess = q.onerror = q.onblocked = () => done(null); }));
  await page.goto(server.resolvedUrls!.local[0]!);
  await page.getByTestId('welcome-new').click();
  await page.getByTestId('canvas').waitFor();
  await page.waitForFunction(() => !!(globalThis as unknown as { __stage?: unknown }).__stage);
}, 60_000);

/** State read inside the page through the store the app exposes. */
const state = <T,>(fn: string): Promise<T> => page.evaluate(`(() => { const s = globalThis.__editor.getState(); ${fn} })()`) as Promise<T>;

/** Save through the File tab and open the downloaded file in Node. */
async function saved(): Promise<Editor> {
  await page.getByRole('tab', { name: 'File' }).click();
  const [d] = await Promise.all([page.waitForEvent('download'), page.getByTestId('save').click()]);
  await page.getByRole('tab', { name: 'Home' }).click();
  return new Editor(OpcPackage.open(new Uint8Array(readFileSync((await d.path())!))));
}

/** The XML of the shape being edited or selected, written compactly (no namespace declarations), for toContain checks. */
const shapeXml = () => state<string>(`
  const el = s.text?.el ?? s.selection[0]; if (!el) return '';
  const ser = e => '<' + (e.prefix ? e.prefix + ':' : '') + e.local + e.attrs.filter(a => a.prefix !== 'xmlns').map(a => ' ' + a.local + '="' + a.value + '"').join('') + '>'
    + e.children.map(c => (c.kind === 'element' ? ser(c) : c.value ?? '')).join('') + '</' + e.local + '>';
  return ser(el);`);

/** Insert a text box, type into it, and stay in it with the Home tab showing. */
async function textBox(text: string): Promise<void> {
  await page.getByRole('tab', { name: 'Insert' }).click();
  await page.getByTestId('text-box').click();
  for (const [i, line] of text.split('\n').entries()) {
    if (i) await page.keyboard.press('Enter');
    await page.keyboard.type(line);
  }
  await page.getByRole('tab', { name: 'Home' }).click();
}

/** Open a ribbon menu and choose an item of it by its label. */
async function choose(menu: string, item: string): Promise<void> {
  await page.getByTestId(menu).click();
  await page.locator('.menu-panel').getByRole('menuitemradio', { name: item, exact: true }).click();
}
const pressed = (testid: string) => page.getByTestId(testid).getAttribute('aria-pressed');
/** The ticked item of a ribbon menu (opened to look, then closed). */
async function ticked(menu: string): Promise<string[]> {
  await page.getByTestId(menu).click();
  const labels = await page.locator('.menu-panel [aria-checked="true"]').allInnerTexts();
  await page.getByTestId(menu).click();
  return labels.map(l => l.replace('✓', '').trim());
}

/** Where a point of the slide (EMU) is on the page. */
const slidePoint = (x: number, y: number) => page.evaluate(([x, y]) => {
  const st = (globalThis as unknown as { __stage: { getAbsoluteTransform(): { point(p: { x: number; y: number }): { x: number; y: number } }; container(): HTMLElement } }).__stage;
  const p = st.getAbsoluteTransform().point({ x, y });
  const r = st.container().getBoundingClientRect();
  return { x: r.left + p.x, y: r.top + p.y };
}, [x, y] as const);
/** The slide as it is now is on screen: two frames have been drawn since the last change (a person is never faster). */
const frames = () => page.evaluate(() => new Promise<void>(done => requestAnimationFrame(() => requestAnimationFrame(() => done()))));
/** Draw a preset shape on the slide by dragging over a box given in EMU; it ends up selected. */
async function drawShape(prst: string, x: number, y: number, w: number, h: number): Promise<void> {
  await frames();
  await page.getByRole('tab', { name: 'Home' }).click();
  await page.getByTestId('shapes').click();
  await page.locator(`[data-testid="shape-gallery"] button[data-prst="${prst}"]`).first().click();
  const a = await slidePoint(x, y), b = await slidePoint(x + w, y + h);
  await page.mouse.move(a.x, a.y);
  await page.mouse.down();
  await page.mouse.move(b.x, b.y, { steps: 4 });
  await page.mouse.up();
  await frames();
}
/** Names of the selected shapes, and of the shapes the slide draws (hidden ones are not drawn). */
const selectedNames = () => state<string[]>(`
  const nameOf = el => { const nv = el.children.find(c => c.kind === 'element' && c.local.startsWith('nv')); const c = nv && nv.children.find(x => x.kind === 'element' && x.local === 'cNvPr'); const a = c && c.attrs.find(a => a.local === 'name'); return a ? a.value : ''; };
  return s.selection.map(nameOf);`);
const drawnNames = () => state<string[]>(`return s.session.scenes.get(s.slide).slide.shapes.filter(x => x.layer === 'slide' && x.source).map(x => x.source.view.name);`);

describe('text formatting from the ribbon', () => {
  test('superscript and subscript toggle and replace each other; the buttons show which is on', async () => {
    await textBox('Hello world');
    await page.keyboard.press('Control+a');
    await page.getByTestId('superscript').click();
    expect(await shapeXml()).toContain('baseline="30000"');
    expect([await pressed('superscript'), await pressed('subscript')]).toEqual(['true', 'false']);
    await page.getByTestId('subscript').click();
    expect(await shapeXml()).toContain('baseline="-25000"');
    expect([await pressed('superscript'), await pressed('subscript')]).toEqual(['false', 'true']);
    await page.getByTestId('subscript').click();
    expect(await shapeXml()).toContain('baseline="0"');
    expect([await pressed('superscript'), await pressed('subscript')]).toEqual(['false', 'false']);
    await page.keyboard.press('Control+z');
    expect(await shapeXml()).toContain('baseline="-25000"');
    expect(errors).toEqual([]);
  }, 60_000);

  test('superscript with only a caret applies to what is typed next', async () => {
    await textBox('x');
    await page.getByTestId('superscript').click();
    expect(await pressed('superscript')).toBe('true');
    await page.keyboard.type('2');
    const xml = await shapeXml();
    expect(xml).toMatch(/<a:r><a:rPr[^>]*><\/rPr><a:t>x<\/t><\/r><a:r><a:rPr[^>]*baseline="30000"[^>]*><\/rPr><a:t>2<\/t><\/r>/);
    expect(errors).toEqual([]);
  }, 60_000);

  test('highlight colour, character spacing, and Clear All Formatting, which keeps the text and typing going', async () => {
    await textBox('Hello world');
    await page.keyboard.press('Control+a');
    await page.getByTestId('highlight').click();
    await page.locator('.menu-panel .swatch[title="Yellow"]').click();
    expect(await shapeXml()).toContain('<a:highlight><a:srgbClr val="FFFF00"></srgbClr></highlight>');
    await choose('char-spacing', 'Loose');
    expect(await shapeXml()).toContain('spc="300"');
    expect(await ticked('char-spacing')).toEqual(['Loose']);
    await page.getByTestId('bold').click();
    expect(await shapeXml()).toContain('b="1"');

    await page.getByTestId('clear-formatting').click();
    const cleared = await shapeXml();
    expect(cleared).not.toContain('highlight');
    expect(cleared).not.toContain('spc=');
    expect(cleared).not.toContain('b="1"');
    expect(cleared).toContain('<a:t>Hello world</t>');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Clear Formatting');
    expect(await ticked('char-spacing')).toEqual(['Normal']);
    // still editing: typing replaces the selection, as before
    await page.keyboard.type('Bye');
    expect(await shapeXml()).toContain('<a:t>Bye</t>');
    expect(errors).toEqual([]);
  }, 60_000);

  test('line spacing and space before and after apply to the paragraph of the caret only', async () => {
    await textBox('First\nSecond');
    await choose('line-spacing', '1.5');
    await choose('line-spacing', 'Before: 6 pt');
    await choose('line-spacing', 'After: 12 pt');
    const xml = await shapeXml();
    expect(xml.match(/<a:lnSpc>/g)?.length).toBe(1);
    expect(xml).toContain('<a:pPr><a:lnSpc><a:spcPct val="150000"></spcPct></lnSpc><a:spcBef><a:spcPts val="600"></spcPts></spcBef><a:spcAft><a:spcPts val="1200"></spcPts></spcAft></pPr>');
    expect(xml.indexOf('<a:lnSpc>')).toBeGreaterThan(xml.indexOf('First'));          // the second paragraph, where the caret is
    expect(await ticked('line-spacing')).toEqual(['1.5', 'Before: 6 pt', 'After: 12 pt']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('align text, margins and columns apply to a selected shape as well as to the box being edited', async () => {
    await textBox('Some text');
    await choose('align-text', 'Middle');
    await choose('columns', 'Two Columns');
    expect(await shapeXml()).toMatch(/<a:bodyPr[^>]*anchor="ctr"[^>]*>/);
    expect(await shapeXml()).toMatch(/<a:bodyPr[^>]*numCol="2"[^>]*>/);
    expect(await ticked('align-text')).toContain('Middle');
    expect(await ticked('columns')).toEqual(['Two Columns']);
    // leave the text: the box stays selected, and the menus still act on it
    await page.keyboard.press('Escape');
    expect(await state<boolean>('return !!s.text;')).toBe(false);
    expect(await state<number>('return s.selection.length;')).toBe(1);
    await choose('align-text', 'Margins: Wide');
    await choose('columns', 'One Column');
    const xml = await shapeXml();
    expect(xml).toMatch(/<a:bodyPr[^>]*lIns="182880"[^>]*tIns="91440"[^>]*>|<a:bodyPr[^>]*tIns="91440"[^>]*lIns="182880"[^>]*>/);
    expect(xml).not.toContain('numCol');
    expect(await ticked('align-text')).toEqual(['Middle', 'Margins: Wide']);

    const ed = await saved();
    const slide = ed.doc.slideParts()[0]!;
    const box = ed.doc.slide(slide).shapes.find(sh => sh.name.startsWith('TextBox'))!;
    expect(box.txBody!.children.find(c => c.kind === 'element' && c.local === 'bodyPr')).toMatchObject({ local: 'bodyPr' });
    expect(errors).toEqual([]);
  }, 60_000);
});

/**
 * A clipboard event on the element that has the keyboard, as the browser sends one for Ctrl/Cmd+C, X or V, with a
 * clipboard of its own (the test never touches the computer's real clipboard). Returns what the clipboard holds
 * afterwards and whether the app handled the event.
 */
const clipboardEvent = (type: 'copy' | 'cut' | 'paste', data: Record<string, string> = {}) => page.evaluate(([type, data]) => {
  const dt = new DataTransfer();
  for (const [k, v] of Object.entries(data)) dt.setData(k, v);
  const e = new ClipboardEvent(type, { clipboardData: dt, bubbles: true, cancelable: true });
  (document.activeElement ?? document.body).dispatchEvent(e);
  return { plain: dt.getData('text/plain'), html: dt.getData('text/html'), handled: e.defaultPrevented };
}, [type, data] as const);

describe('the clipboard in text', () => {
  /** The runs of the box being edited, a paragraph per line: "[b]" and "[i]" before a run that has them. */
  const runs = () => state<string[]>(`
    const body = s.text.el.children.find(c => c.local === 'txBody');
    const flag = (r, n) => { const rPr = r.children.find(c => c.local === 'rPr'); return !!rPr && rPr.attrs.some(a => a.local === n && a.value !== '0' && a.value !== 'none'); };
    return body.children.filter(c => c.local === 'p').map(p => p.children.filter(c => c.local === 'r').map(r =>
      (flag(r, 'b') ? '[b]' : '') + (flag(r, 'i') ? '[i]' : '') + (flag(r, 'u') ? '[u]' : '') + r.children.find(c => c.local === 't').children.map(c => c.value ?? '').join('')).join('|'));`);
  /** Select characters from..to of the first paragraph with the keyboard. */
  async function select(from: number, to: number): Promise<void> {
    await page.keyboard.press('Control+Home');
    for (let i = 0; i < from; i++) await page.keyboard.press('ArrowRight');
    for (let i = from; i < to; i++) await page.keyboard.press('Shift+ArrowRight');
  }

  test('copy puts the selected text on the clipboard as plain text and as HTML that keeps bold; cut takes it out of the box', async () => {
    await textBox('Hello big world');
    await select(6, 9);
    await page.getByTestId('bold').click();
    await page.keyboard.press('Control+a');
    expect(await clipboardEvent('copy')).toEqual({ plain: 'Hello big world', html: '<p>Hello <b>big</b> world</p>', handled: true });
    expect(await runs()).toEqual(['Hello |[b]big| world']);             // copying changes nothing

    await select(6, 10);
    expect(await clipboardEvent('cut')).toEqual({ plain: 'big ', html: '<p><b>big</b> </p>', handled: true });
    expect(await runs()).toEqual(['Hello |world']);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Delete');
    await page.keyboard.press('Control+z');
    expect(await runs()).toEqual(['Hello |[b]big| world']);

    // with only a caret there is nothing to copy: the event is left alone
    await page.keyboard.press('Control+End');
    expect(await clipboardEvent('copy')).toEqual({ plain: '', html: '', handled: false });
    expect(await clipboardEvent('cut')).toEqual({ plain: '', html: '', handled: false });
    expect(errors).toEqual([]);
  }, 60_000);

  test('paste: plain text is typed where the caret is; HTML from another app keeps its bold, italic, underline and paragraphs', async () => {
    await textBox('Start.');
    expect((await clipboardEvent('paste', { 'text/plain': ' Plain text' })).handled).toBe(true);
    expect(await runs()).toEqual(['Start. Plain text']);

    await page.keyboard.press('Control+z');
    const html = '<meta charset="utf-8"><p>one <b>two</b> <span style="font-style:italic">three</span></p><p><u>four</u><script>alert(1)</script></p>';
    await clipboardEvent('paste', { 'text/plain': 'one two three\nfour', 'text/html': html });
    expect(await runs()).toEqual(['Start.|one |[b]two| |[i]three', '[u]four']);
    // the paste is one undo step
    await page.keyboard.press('Control+z');
    expect(await runs()).toEqual(['Start.']);
    await page.keyboard.press('Control+y');
    expect(await runs()).toEqual(['Start.|one |[b]two| |[i]three', '[u]four']);

    // pasting over a selection replaces it
    await page.keyboard.press('Control+a');
    await clipboardEvent('paste', { 'text/plain': 'All new' });
    expect(await runs()).toEqual(['All new']);
    // HTML with no formatting in it is pasted as its plain text (so it takes the look of where it lands)
    await clipboardEvent('paste', { 'text/plain': '!', 'text/html': '<span>!</span>' });
    expect(await runs()).toEqual(['All new!']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('what is copied in one box pastes into another with its formatting', async () => {
    await textBox('Hello big world');
    await select(6, 9);
    await page.keyboard.press('Control+b');
    await page.keyboard.press('Control+a');
    const copied = await clipboardEvent('copy');
    expect(copied.html).toContain('<b>big</b>');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    await textBox('');
    await clipboardEvent('paste', { 'text/plain': copied.plain, 'text/html': copied.html });
    expect(await runs()).toEqual(['Hello |[b]big| world']);
    const ed = await saved();
    expect(new TextDecoder().decode(ed.pkg.getPart(ed.doc.slideParts()[0]!)!.bytes).match(/<a:t>big<\/a:t>/g)?.length).toBe(2);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('the clipboard on the slide', () => {
  /** Names of the shapes drawn on the slide that are not placeholders, and the text of each. */
  const drawn = () => state<string[]>(`
    const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
    return s.session.ed.doc.slide(s.slide).shapes.filter(sh => !sh.placeholder).map(sh => sh.name.replace(/ \\d+$/, '') + (textOf(sh.el) ? ':' + textOf(sh.el) : ''));`);
  /** A rectangle under the placeholders with "Hi" in it, selected (not being typed in). */
  async function rectWithText(): Promise<void> {
    await drawShape('rect', 1000000, 5500000, 1500000, 800000);
    await page.keyboard.type('Hi');
    await page.keyboard.press('Escape');
    await frames();
  }
  /** A paste of a picture file, as a screenshot tool or an image editor puts one on the clipboard. */
  const pastePicture = () => page.evaluate(`(async () => {
    const canvas = new OffscreenCanvas(40, 30);
    const c = canvas.getContext('2d');
    c.fillStyle = '#c00';
    c.fillRect(0, 0, 40, 30);
    const dt = new DataTransfer();
    dt.items.add(new File([await canvas.convertToBlob({ type: 'image/png' })], 'shot.png', { type: 'image/png' }));
    const e = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
    document.body.dispatchEvent(e);
    return e.defaultPrevented;
  })()`) as Promise<boolean>;

  test('copy puts the shapes on the clipboard with their text; pasting that clipboard brings them back whole, here and on another slide', async () => {
    await rectWithText();
    const copied = await clipboardEvent('copy');
    expect(copied.handled).toBe(true);
    expect(copied.plain).toBe('Hi');
    expect(copied.html).toMatch(/^<div data-pptx-clip="[a-z0-9-]+"><p>Hi<\/p><\/div>$/);

    await clipboardEvent('paste', { 'text/plain': copied.plain, 'text/html': copied.html });
    await expect.poll(drawn).toEqual(['Rectangle:Hi', 'Rectangle:Hi']);
    // the copy is a shape with the fill and place of the first (set a little off it), not a text box
    const xs = await state<number[]>(`return s.session.ed.doc.slide(s.slide).shapes.filter(sh => !sh.placeholder).map(sh => sh.xfrm.x);`);
    expect(xs[1]! - xs[0]!).toBe(152400);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Paste');
    expect(await state<number>('return s.selection.length;')).toBe(1);

    await page.getByTestId('new-slide').click();
    await clipboardEvent('paste', { 'text/plain': copied.plain, 'text/html': copied.html });
    await expect.poll(drawn).toEqual(['Rectangle:Hi']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('cut takes the shapes off the slide and paste puts them back; with nothing selected copy and cut are left alone', async () => {
    expect((await clipboardEvent('copy')).handled).toBe(false);
    expect((await clipboardEvent('cut')).handled).toBe(false);
    await rectWithText();
    const cut = await clipboardEvent('cut');
    expect(cut.handled).toBe(true);
    expect(await drawn()).toEqual([]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toMatch(/Delete/);
    await clipboardEvent('paste', { 'text/plain': cut.plain, 'text/html': cut.html });
    await expect.poll(drawn).toEqual(['Rectangle:Hi']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('text from another app becomes a text box, with its bold and its paragraphs; a picture becomes a picture', async () => {
    await clipboardEvent('paste', { 'text/plain': 'From a note' });
    await expect.poll(drawn).toEqual(['TextBox:From a note']);
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');

    await clipboardEvent('paste', { 'text/plain': 'Title\nbody', 'text/html': '<h1>Title</h1><p>body <i>slanted</i></p>' });
    await expect.poll(async () => (await drawn()).length).toBe(2);
    expect(await shapeXml()).toMatch(/<a:rPr[^>]*b="1"[^>]*><\/rPr><a:t>Title<\/t>/);
    expect(await shapeXml()).toMatch(/<a:rPr[^>]*i="1"[^>]*><\/rPr><a:t>slanted<\/t>/);
    expect((await shapeXml()).match(/<a:p>/g)?.length).toBe(2);
    // one undo takes the pasted text away, the next the box
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Control+z');
    await page.keyboard.press('Control+z');
    expect(await drawn()).toEqual(['TextBox:From a note']);

    expect(await pastePicture()).toBe(true);
    await expect.poll(drawn, { timeout: 10_000 }).toEqual(['TextBox:From a note', 'Picture']);
    expect(await state<string>('return s.selection[0].local;')).toBe('pic');
    // empty text pastes nothing
    await clipboardEvent('paste', { 'text/plain': '   ' });
    await clipboardEvent('paste', { 'text/plain': '' });
    expect((await drawn()).length).toBe(2);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a clipboard from an earlier copy, or from another window, is pasted as what it says, not as the shapes in memory', async () => {
    await rectWithText();
    await clipboardEvent('copy');
    await page.keyboard.press('Escape');
    await clipboardEvent('paste', { 'text/plain': 'Other', 'text/html': '<div data-pptx-clip="someone-else"><p>Other</p></div>' });
    await expect.poll(drawn).toEqual(['Rectangle:Hi', 'TextBox:Other']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('the notes box and other text fields keep their own copy and paste', async () => {
    await rectWithText();
    await page.getByTestId('notes').click();
    await page.keyboard.type('note');
    expect((await clipboardEvent('copy')).handled).toBe(false);
    expect((await clipboardEvent('paste', { 'text/plain': 'x' })).handled).toBe(false);
    expect(await drawn()).toEqual(['Rectangle:Hi']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('the Copy and Paste buttons and the right-click menu use the same clipboard', async () => {
    await rectWithText();
    await page.getByRole('button', { name: 'Copy', exact: true }).click();
    expect(await state<boolean>('return !!s.clipboard;')).toBe(true);
    await page.getByRole('button', { name: 'Paste', exact: true }).click();
    await expect.poll(drawn).toEqual(['Rectangle:Hi', 'Rectangle:Hi']);
    await page.getByRole('button', { name: 'Cut', exact: true }).click();
    expect(await drawn()).toEqual(['Rectangle:Hi']);
    await page.getByRole('button', { name: 'Paste', exact: true }).click();
    await expect.poll(async () => (await drawn()).length).toBe(2);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a picture of the selected shapes can be made for other apps: a PNG cropped to them, with their colours', async () => {
    await drawShape('rect', 1000000, 5500000, 1500000, 800000);
    // (the app's own module, loaded in the page as the app loads it)
    const png = await page.evaluate(`(async () => {
      const m = await import('/src/ui/clipboard.ts');
      const blob = await m.shapesPng(globalThis.__editor.getState().selection);
      if (!blob) return null;
      const bmp = await createImageBitmap(blob);
      const c = new OffscreenCanvas(bmp.width, bmp.height);
      const ctx = c.getContext('2d');
      ctx.drawImage(bmp, 0, 0);
      const mid = ctx.getImageData(bmp.width >> 1, bmp.height >> 1, 1, 1).data;
      const corner = ctx.getImageData(0, 0, 1, 1).data;
      return { type: blob.type, w: bmp.width, h: bmp.height, mid: [mid[0], mid[1], mid[2], mid[3]], corner: corner[3] };
    })()`) as { type: string; w: number; h: number; mid: number[]; corner: number } | null;
    expect(png).not.toBeNull();
    expect(png!.type).toBe('image/png');
    // 1.5M by 0.8M EMU at two pixels per 9525 EMU, to within the drag's pixel
    expect(Math.abs(png!.w - 315)).toBeLessThan(6);
    expect(Math.abs(png!.h - 168)).toBeLessThan(6);
    expect(png!.mid).toEqual([0x44, 0x72, 0xc4, 255]);                   // the rectangle's accent fill
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('copy, cut and paste slides', () => {
  /** What tells the slides apart: the text of each slide's text boxes. */
  const marks = () => state<string[]>(`
    const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
    return s.session.ed.doc.slideParts().map(p => s.session.ed.doc.slide(p).shapes.filter(sh => sh.name.startsWith('TextBox')).map(sh => textOf(sh.el)).join('+'));`);
  const current = () => state<number>('return s.session.ed.doc.slideParts().indexOf(s.slide) + 1;');
  /** Three slides, each with a text box saying which it is; the slide panel focused, on the first slide. */
  async function threeSlides(): Promise<void> {
    for (const [i, mark] of ['one', 'two', 'three'].entries()) {
      if (i) await page.getByTestId('new-slide').click();
      await textBox(mark);
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
    }
    await page.getByTestId('thumb-1').click();
    await page.locator('aside.slides').focus();
  }

  test('with the slide panel focused: copy and paste puts a copy after the current slide; cut and paste moves a slide', async () => {
    await threeSlides();
    const copied = await clipboardEvent('copy');
    expect(copied.handled).toBe(true);
    expect(copied.plain).toBe('Slide 1');
    expect(copied.html).toMatch(/^<div data-pptx-slide="[a-z0-9-]+"><\/div>$/);
    expect(await marks()).toEqual(['one', 'two', 'three']);

    await page.getByTestId('thumb-2').click();
    await page.locator('aside.slides').focus();
    await clipboardEvent('paste', { 'text/plain': copied.plain, 'text/html': copied.html });
    expect(await marks()).toEqual(['one', 'two', 'one', 'three']);
    expect(await current()).toBe(3);                                     // on the pasted slide
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Paste Slide');

    // cut the first slide and paste it at the end
    await page.getByTestId('thumb-1').click();
    await page.locator('aside.slides').focus();
    const cut = await clipboardEvent('cut');
    expect(cut.handled).toBe(true);
    expect(await marks()).toEqual(['two', 'one', 'three']);
    await page.getByTestId('thumb-3').click();
    await page.locator('aside.slides').focus();
    await clipboardEvent('paste', { 'text/plain': cut.plain, 'text/html': cut.html });
    expect(await marks()).toEqual(['two', 'one', 'three', 'one']);

    // a clipboard that is not that slide (copied since, elsewhere) pastes nothing here
    expect((await clipboardEvent('paste', { 'text/plain': 'x', 'text/html': '<div data-pptx-slide="other"></div>' })).handled).toBe(false);
    expect((await clipboardEvent('paste', { 'text/plain': 'just text' })).handled).toBe(false);
    expect((await marks()).length).toBe(4);

    const ed = await saved();
    expect(ed.doc.slideParts().length).toBe(4);
    expect(errors).toEqual([]);
  }, 60_000);

  test('the only slide of a deck cannot be cut; shapes on the slide are not copied while the panel has the focus', async () => {
    await textBox('alone');
    await page.keyboard.press('Escape');                                 // the box stays selected
    await page.locator('aside.slides').focus();
    expect((await clipboardEvent('cut')).handled).toBe(false);
    expect(await marks()).toEqual(['alone']);
    const copied = await clipboardEvent('copy');
    expect(copied.html).toContain('data-pptx-slide');                    // the slide, not the selected box
    expect(errors).toEqual([]);
  }, 60_000);

  test('from the right-click menu of a thumbnail: Copy Slide, then Paste Slide on another', async () => {
    await threeSlides();
    await page.getByTestId('thumb-3').click({ button: 'right' });
    await page.getByTestId('menu-copy-slide').click();
    await page.getByTestId('thumb-1').click({ button: 'right' });
    await page.getByTestId('menu-paste-slide').click();
    expect(await marks()).toEqual(['one', 'three', 'two', 'three']);
    await page.getByTestId('thumb-2').click({ button: 'right' });
    await page.getByTestId('menu-cut-slide').click();
    expect(await marks()).toEqual(['one', 'two', 'three']);
    await page.keyboard.press('Control+z');
    expect(await marks()).toEqual(['one', 'three', 'two', 'three']);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('hyperlinks', () => {
  /** The slide's external relationships (addresses) and the slides it links to. */
  const rels = () => state<{ urls: string[]; slides: number }>(`
    const r = s.session.ed.pkg.relationships(s.slide);
    return { urls: r.filter(x => x.external).map(x => x.target), slides: r.filter(x => !x.external && x.type.endsWith('/slide')).length };`);
  /** The runs of the box being edited: "[link]" before a run that carries a link. */
  const runs = () => state<string>(`
    const body = s.text.el.children.find(c => c.local === 'txBody');
    return body.children.filter(c => c.local === 'p').map(p => p.children.filter(c => c.local === 'r').map(r => {
      const rPr = r.children.find(c => c.local === 'rPr');
      return (rPr && rPr.children.some(c => c.local === 'hlinkClick') ? '[link]' : '') + r.children.find(c => c.local === 't').children.map(c => c.value ?? '').join('');
    }).join('|')).join('/');`);
  async function select(from: number, to: number): Promise<void> {
    await page.keyboard.press('Control+Home');
    for (let i = 0; i < from; i++) await page.keyboard.press('ArrowRight');
    for (let i = from; i < to; i++) await page.keyboard.press('Shift+ArrowRight');
  }
  const dialog = () => page.getByTestId('link-dialog');

  test('Ctrl+K links the selected words to an address typed without its scheme; the caret in the link edits or removes it', async () => {
    await textBox('Visit our site now');
    await select(6, 14);
    await page.keyboard.press('Control+k');
    await dialog().waitFor();
    expect(await page.getByTestId('link-url').inputValue()).toBe('');
    expect(await page.getByTestId('link-remove').count()).toBe(0);       // nothing to remove yet
    await page.keyboard.type('example.com/docs');
    await page.keyboard.press('Enter');
    expect(await dialog().count()).toBe(0);
    expect(await runs()).toBe('Visit |[link]our site| now');
    expect(await rels()).toEqual({ urls: ['https://example.com/docs'], slides: 0 });
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Hyperlink');
    // typing goes on in the box after the dialog closes
    await page.keyboard.press('Control+End');
    await page.keyboard.type('!');
    expect(await runs()).toBe('Visit |[link]our site| now!');

    // the caret anywhere in the link: the dialog shows it, and Apply changes all of it
    await select(9, 9);
    await page.keyboard.press('Control+k');
    expect(await page.getByTestId('link-url').inputValue()).toBe('https://example.com/docs');
    await page.getByTestId('link-url').fill('someone@example.com');
    await page.getByTestId('link-apply').click();
    expect(await rels()).toEqual({ urls: ['mailto:someone@example.com'], slides: 0 });
    expect(await runs()).toBe('Visit |[link]our site| now!');

    await page.keyboard.press('Control+k');
    await page.getByTestId('link-remove').click();
    expect(await runs()).toBe('Visit |our site| now!');
    expect(await rels()).toEqual({ urls: [], slides: 0 });
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Remove Hyperlink');
    await page.keyboard.press('Control+z');
    expect(await rels()).toEqual({ urls: ['mailto:someone@example.com'], slides: 0 });
    expect(errors).toEqual([]);
  }, 60_000);

  test('an address that is not a web or mail address is refused and nothing is linked; Cancel and Escape change nothing', async () => {
    await textBox('Click here');
    await select(6, 10);
    await page.keyboard.press('Control+k');
    for (const bad of ['javascript:alert(1)', 'file:///etc/passwd', 'not an address']) {
      await page.getByTestId('link-url').fill(bad);
      await page.getByTestId('link-apply').click();
      expect(await page.getByTestId('link-error').innerText()).toContain('web address');
      expect(await dialog().count()).toBe(1);
    }
    await page.keyboard.press('Escape');
    expect(await dialog().count()).toBe(0);
    expect(await runs()).toBe('Click here');
    expect(await rels()).toEqual({ urls: [], slides: 0 });
    await page.keyboard.press('Control+k');
    await page.getByTestId('link-url').fill('example.com');
    await dialog().getByRole('button', { name: 'Cancel' }).click();
    expect(await rels()).toEqual({ urls: [], slides: 0 });
    expect(errors).toEqual([]);
  }, 60_000);

  test('with only a caret, the address is typed in as the linked text', async () => {
    await textBox('See ');
    await page.keyboard.press('Control+k');
    await page.getByTestId('link-url').fill('https://example.com');
    await page.getByTestId('link-apply').click();
    expect(await runs()).toBe('See |[link]https://example.com');
    await page.keyboard.type(' today');
    expect((await runs()).startsWith('See |[link]https://example.com')).toBe(true);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a shape links to another slide, from Insert > Link and from the right-click menu; a click on it in the show would go there', async () => {
    await page.getByTestId('new-slide').click();
    await state('s.setSlide(s.session.ed.doc.slideParts()[0]);');
    await drawShape('rect', 1000000, 5500000, 1500000, 800000);
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.getByTestId('insert-link').click();
    await page.getByTestId('link-kind-slide').check();
    expect(await page.getByTestId('link-slide').locator('option').allInnerTexts()).toEqual(['Slide 1', 'Slide 2']);
    await page.getByTestId('link-slide').selectOption({ label: 'Slide 2' });
    await page.getByTestId('link-apply').click();
    expect(await rels()).toEqual({ urls: [], slides: 1 });
    expect(await state<boolean>(`return s.selection[0].children[0].children[0].children.some(c => c.local === 'hlinkClick');`)).toBe(true);

    // what the slide show follows at the middle of the shape
    await frames();
    const hit = await page.evaluate(`(async () => {
      const m = await import('/src/ui/links.ts');
      const s = globalThis.__editor.getState();
      const link = m.linkAt(s.session.scenes.get(s.slide).slide, 1700000, 5900000);
      return link ? s.session.ed.doc.slideParts().indexOf(link.slide) + 1 : 0;
    })()`);
    expect(hit).toBe(2);

    // the menu opens the dialog on the link that is there; Remove takes it off the shape
    await page.getByRole('tab', { name: 'Home' }).click();
    const p = await slidePoint(1700000, 5900000);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByTestId('menu-link').click();
    expect(await page.getByTestId('link-kind-slide').isChecked()).toBe(true);
    expect(await page.getByTestId('link-slide').inputValue()).toContain('slide2');
    await page.getByTestId('link-remove').click();
    expect(await rels()).toEqual({ urls: [], slides: 0 });

    // with nothing selected there is nothing to link: the button is off and Ctrl+K does nothing
    await page.keyboard.press('Escape');
    await page.getByRole('tab', { name: 'Insert' }).click();
    expect(await page.getByTestId('insert-link').isDisabled()).toBe(true);
    await page.keyboard.press('Control+k');
    expect(await dialog().count()).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('links are in the saved file: the address as an external relationship the linked run points at', async () => {
    await textBox('Docs');
    await page.keyboard.press('Control+a');
    await page.keyboard.press('Control+k');
    await page.getByTestId('link-url').fill('example.com');
    await page.getByTestId('link-apply').click();
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    const ed = await saved();
    const part = ed.doc.slideParts()[0]!;
    const rel = ed.pkg.relationships(part).find(r => r.external)!;
    expect(rel.target).toBe('https://example.com');
    expect(new TextDecoder().decode(ed.pkg.getPart(part)!.bytes)).toContain(`<a:hlinkClick r:id="${rel.id}"/>`);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('fills, arrows and shadows', () => {
  /** Open a ribbon menu (by its test id) and choose a plain item by its label. */
  const pick = async (menu: string, item: string) => {
    await page.getByTestId(menu).click();
    await page.locator('.menu-panel').getByRole('menuitem', { name: item, exact: true }).click();
  };
  /** The selected shape's properties element, compactly. */
  const spPr = () => state<string>(`
    const el = s.selection[0]; const pr = el.children.find(c => c.local === 'spPr');
    const ser = e => '<' + e.local + e.attrs.filter(a => a.prefix !== 'xmlns').map(a => ' ' + a.local + '="' + a.value + '"').join('') + '>' + e.children.map(c => (c.kind === 'element' ? ser(c) : '')).join('') + '</' + e.local + '>';
    return ser(pr);`);
  /** The colour drawn at a point of the slide (EMU), as [r, g, b]. */
  const pixel = async (x: number, y: number) => {
    await frames();
    const p = await slidePoint(x, y);
    return page.evaluate(([px, py]) => {
      const c = document.querySelector('[data-testid="canvas"] canvas') as HTMLCanvasElement;
      const r = c.getBoundingClientRect();
      const d = c.getContext('2d')!.getImageData(Math.round((px - r.left) * (c.width / r.width)), Math.round((py - r.top) * (c.height / r.height)), 1, 1).data;
      return [d[0]!, d[1]!, d[2]!];
    }, [p.x, p.y] as const);
  };
  const RECT = { x: 1000000, y: 5400000, w: 2000000, h: 1000000 };

  test('a gradient in the shape\'s colour, light at the top and the colour at the bottom; then half transparent', async () => {
    await drawShape('rect', RECT.x, RECT.y, RECT.w, RECT.h);
    await pick('shape-fill', 'Gradient: Down');
    const xml = await spPr();
    expect(xml).toContain('<gradFill rotWithShape="1"><gsLst><gs pos="0"><schemeClr val="accent1"><lumMod val="40000"></lumMod><lumOff val="60000"></lumOff></schemeClr></gs><gs pos="100000"><schemeClr val="accent1"></schemeClr></gs></gsLst><lin ang="5400000" scaled="0"></lin></gradFill>');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Shape Fill');
    // drawn: lighter near the top edge than near the bottom, and the bottom close to accent 1 (#4472C4)
    await page.keyboard.press('Escape');
    const top = await pixel(RECT.x + RECT.w / 2, RECT.y + 60000), bottom = await pixel(RECT.x + RECT.w / 2, RECT.y + RECT.h - 60000);
    expect(top[0]! + top[1]! + top[2]!).toBeGreaterThan(bottom[0]! + bottom[1]! + bottom[2]! + 120);
    expect(Math.abs(bottom[0]! - 0x44) + Math.abs(bottom[1]! - 0x72) + Math.abs(bottom[2]! - 0xc4)).toBeLessThan(40);

    const mid = await slidePoint(RECT.x + RECT.w / 2, RECT.y + RECT.h / 2);
    await page.mouse.click(mid.x, mid.y);
    await pick('shape-fill', 'Transparency: 50%');
    expect((await spPr()).match(/<alpha val="50000">/g)?.length).toBe(2);
    await pick('shape-fill', 'Gradient: Right');
    expect(await spPr()).toContain('<lin ang="0" scaled="0">');
    expect(errors).toEqual([]);
  }, 60_000);

  test('a solid colour made transparent lets the slide show through; 0% makes it solid again', async () => {
    await drawShape('rect', RECT.x, RECT.y, RECT.w, RECT.h);
    await page.getByTestId('shape-fill').click();
    await page.locator('.menu-panel .swatch[title="Red"]').click();
    await pick('shape-fill', 'Transparency: 75%');
    expect(await spPr()).toContain('<solidFill><srgbClr val="FF0000"><alpha val="25000"></alpha></srgbClr></solidFill>');
    await page.keyboard.press('Escape');
    const faint = await pixel(RECT.x + RECT.w / 2, RECT.y + RECT.h / 2);
    expect(faint[0]).toBe(255);
    expect(faint[1]!).toBeGreaterThan(170);                              // mostly the white slide
    const mid = await slidePoint(RECT.x + RECT.w / 2, RECT.y + RECT.h / 2);
    await page.mouse.click(mid.x, mid.y);
    await pick('shape-fill', 'Transparency: 0%');
    expect(await spPr()).toContain('<solidFill><srgbClr val="FF0000"></srgbClr></solidFill>');
    expect(errors).toEqual([]);
  }, 60_000);

  test('a picture from a file fills the shape and is in the saved file', async () => {
    const png = new PNG({ width: 8, height: 8 });
    for (let i = 0; i < 64; i++) png.data.set([0, 160, 0, 255], i * 4);
    const file = join(mkdtempSync(join(tmpdir(), 'fill-')), 'green.png');
    writeFileSync(file, PNG.sync.write(png));

    await drawShape('ellipse', RECT.x, RECT.y, RECT.w, RECT.h);
    await page.getByTestId('shape-fill').click();
    const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.locator('.menu-panel').getByRole('menuitem', { name: 'Picture…' }).click()]);
    await chooser.setFiles(file);
    await expect.poll(spPr, { timeout: 10_000 }).toMatch(/<blipFill dpi="0" rotWithShape="1"><blip embed="rId\d+"><\/blip><srcRect><\/srcRect><stretch><fillRect><\/fillRect><\/stretch><\/blipFill>/);
    // drawn: the middle of the ellipse is the picture's green
    await page.keyboard.press('Escape');
    await expect.poll(() => pixel(RECT.x + RECT.w / 2, RECT.y + RECT.h / 2), { timeout: 10_000 }).toEqual([0, 160, 0]);
    const ed = await saved();
    expect(ed.pkg.partNames().filter(n => n.startsWith('/ppt/media/')).length).toBe(1);
    expect(errors).toEqual([]);
  }, 60_000);

  test('arrows on a line, and a shadow on a shape, from their menus', async () => {
    await drawShape('line', RECT.x, RECT.y, RECT.w, 0);
    await pick('shape-outline', 'Arrow at End');
    expect(await spPr()).toContain('<tailEnd type="triangle">');
    expect(await spPr()).not.toContain('headEnd');
    await pick('shape-outline', 'Arrows at Both Ends');
    expect(await spPr()).toMatch(/<headEnd type="triangle"><\/headEnd><tailEnd type="triangle"><\/tailEnd>/);
    await pick('shape-outline', 'Arrow at Start');
    expect(await spPr()).toContain('<headEnd type="triangle">');
    expect(await spPr()).not.toContain('tailEnd');
    await pick('shape-outline', 'Arrows: None');
    expect(await spPr()).not.toMatch(/headEnd|tailEnd/);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Shape Outline');

    await drawShape('rect', 4000000, RECT.y, 1500000, 800000);
    await pick('shape-shadow', 'Shadow: Bottom Right');
    expect(await spPr()).toContain('<effectLst><outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">');
    // drawn: just outside the bottom right corner is no longer plain white
    await page.keyboard.press('Escape');
    const beside = await pixel(4000000 + 1500000 + 25000, RECT.y + 800000 + 25000);
    expect(beside[0]! + beside[1]! + beside[2]!).toBeLessThan(760);
    const mid = await slidePoint(4700000, RECT.y + 400000);
    await page.mouse.click(mid.x, mid.y);
    await pick('shape-shadow', 'No Shadow');
    expect(await spPr()).toContain('<effectLst></effectLst>');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Shadow');
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('the format painter', () => {
  /** The fill each drawn shape has of its own ("red" for FF0000, "-" for none of its own), by shape name. */
  const fills = () => state<Record<string, string>>(`
    const out = {};
    for (const sh of s.session.ed.doc.slide(s.slide).shapes.filter(x => !x.placeholder)) {
      const pr = sh.el.children.find(c => c.local === 'spPr');
      const solid = pr && pr.children.find(c => c.local === 'solidFill');
      const clr = solid && solid.children[0];
      out[sh.name.replace(/ \\d+$/, '')] = clr ? clr.attrs.find(a => a.local === 'val').value : '-';
    }
    return out;`);
  const painter = () => page.getByTestId('format-painter');
  const click = async (x: number, y: number) => { await frames(); const p = await slidePoint(x, y); await page.mouse.click(p.x, p.y); };
  /** A red rectangle, a plain ellipse and a plain triangle in a row under the placeholders; the rectangle selected. */
  async function threeShapes(): Promise<void> {
    await drawShape('ellipse', 4000000, 5500000, 1500000, 800000);
    await drawShape('triangle', 7000000, 5500000, 1500000, 800000);
    await drawShape('rect', 1000000, 5500000, 1500000, 800000);
    await page.getByTestId('shape-fill').click();
    await page.locator('.menu-panel .swatch[title="Red"]').click();
  }
  const RECT = [1700000, 5900000] as const, OVAL = [4700000, 5900000] as const, TRIANGLE = [7750000, 6100000] as const, EMPTY = [600000, 600000] as const;

  test('one click on the painter, one click on a shape: it takes the look, is selected, and the painter goes off', async () => {
    await threeShapes();
    expect(await painter().getAttribute('aria-pressed')).toBe('false');
    await painter().click();
    expect(await painter().getAttribute('aria-pressed')).toBe('true');
    expect(await page.getByTestId('canvas').getAttribute('class')).toContain('painting');
    await click(...OVAL);
    expect(await fills()).toEqual({ Oval: 'FF0000', 'Isosceles Triangle': '-', Rectangle: 'FF0000' });
    expect(await selectedNames()).toEqual(['Oval 3']);
    expect(await painter().getAttribute('aria-pressed')).toBe('false');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Format Painter');
    // the next click is an ordinary click again
    await click(...TRIANGLE);
    expect((await fills())['Isosceles Triangle']).toBe('-');
    await page.keyboard.press('Control+z');
    expect((await fills()).Oval).toBe('-');
    expect(errors).toEqual([]);
  }, 60_000);

  test('a double click keeps the painter on for shape after shape, until Escape; a click on the empty slide also puts it down', async () => {
    await threeShapes();
    await painter().dblclick();
    expect(await painter().getAttribute('aria-pressed')).toBe('true');
    await click(...OVAL);
    await click(...TRIANGLE);
    expect(await fills()).toEqual({ Oval: 'FF0000', 'Isosceles Triangle': 'FF0000', Rectangle: 'FF0000' });
    expect(await painter().getAttribute('aria-pressed')).toBe('true');
    await page.keyboard.press('Escape');
    expect(await painter().getAttribute('aria-pressed')).toBe('false');
    expect(await selectedNames()).toEqual(['Isosceles Triangle 4']);      // Escape put the painter down, not the selection

    await painter().click();
    await click(...EMPTY);
    expect(await painter().getAttribute('aria-pressed')).toBe('false');
    // clicking the painter again while it is on turns it off
    await click(...RECT);
    await painter().click();
    await painter().click();
    expect(await painter().getAttribute('aria-pressed')).toBe('false');
    expect(errors).toEqual([]);
  }, 60_000);

  test('Ctrl+Alt+C picks a look up and Ctrl+Alt+V puts it on the selection, as often as wanted', async () => {
    await threeShapes();
    await page.keyboard.press('Control+Alt+c');
    await click(...OVAL);
    await page.keyboard.press('Control+Alt+v');
    await click(...TRIANGLE);
    await page.keyboard.press('Control+Alt+v');
    expect(await fills()).toEqual({ Oval: 'FF0000', 'Isosceles Triangle': 'FF0000', Rectangle: 'FF0000' });
    expect(await painter().getAttribute('aria-pressed')).toBe('false');
    expect(errors).toEqual([]);
  }, 60_000);

  test('picked up from selected text, only the text of the shape clicked changes', async () => {
    await threeShapes();
    await page.keyboard.type('Bold');
    await page.keyboard.press('Control+a');
    await page.getByTestId('bold').click();
    await painter().click();                                            // from the text selection: character formatting only
    await click(...OVAL);
    expect((await fills()).Oval).toBe('-');
    await page.keyboard.type('x');
    expect(await shapeXml()).toMatch(/<a:rPr[^>]*b="1"/);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('cropping a picture with handles', () => {
  /** Insert a 200 by 100 pixel picture (left half red, right half blue); it ends up selected, with the Home tab showing. */
  async function picture(): Promise<void> {
    const png = new PNG({ width: 200, height: 100 });
    for (let y = 0; y < 100; y++) for (let x = 0; x < 200; x++) png.data.set(x < 100 ? [200, 0, 0, 255] : [0, 0, 200, 255], (y * 200 + x) * 4);
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.locator('input[type=file][accept^="image/png"]').first().setInputFiles({ name: 'two.png', mimeType: 'image/png', buffer: PNG.sync.write(png) });
    await expect.poll(() => state<string>('return s.selection[0] ? s.selection[0].local : "";'), { timeout: 10_000 }).toBe('pic');
    await page.getByRole('tab', { name: 'Home' }).click();
    await frames();
  }
  /** The selected picture's frame and its crop (thousandths of a percent). */
  const pic = () => state<{ x: number; y: number; cx: number; cy: number; l: number; t: number; r: number; b: number }>(`
    const el = s.selection[0]; const find = (e, n) => e.children.find(c => c.kind === 'element' && c.local === n);
    const x = find(find(el, 'spPr'), 'xfrm'), off = find(x, 'off'), ext = find(x, 'ext'), src = find(find(el, 'blipFill'), 'srcRect');
    const num = (e, n) => { const a = e && e.attrs.find(t => t.local === n); return a ? Number(a.value) : 0; };
    return { x: num(off, 'x'), y: num(off, 'y'), cx: num(ext, 'cx'), cy: num(ext, 'cy'), l: num(src, 'l'), t: num(src, 't'), r: num(src, 'r'), b: num(src, 'b') };`);
  /** Where the crop handles are on the page, in the order the app keeps them (the fifth is the right-hand middle). */
  const handles = () => page.evaluate(() => {
    const st = (globalThis as unknown as { __stage: { find(s: string): { getClientRect(): { x: number; y: number; width: number; height: number } }[]; container(): HTMLElement } }).__stage;
    const r = st.container().getBoundingClientRect();
    return st.find('.crop-handle').map(h => { const b = h.getClientRect(); return { x: r.left + b.x + b.width / 2, y: r.top + b.y + b.height / 2 }; });
  });
  const cropMenu = async (item: string) => {
    await page.getByTestId('crop-menu').click();
    await page.locator('.menu-panel').getByRole('menuitem', { name: item, exact: true }).click();
    await frames();
  };
  const drag = async (from: { x: number; y: number }, dx: number, dy: number) => {
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(from.x + dx / 2, from.y + dy / 2, { steps: 3 });
    await page.mouse.move(from.x + dx, from.y + dy, { steps: 3 });
    await page.mouse.up();
    await frames();
  };
  const zoom = () => page.evaluate(() => (globalThis as unknown as { __stage: { scaleX(): number } }).__stage.scaleX());

  test('Crop shows eight handles; dragging one in cuts the picture there, the rest staying put; the drag is one undo step', async () => {
    await picture();
    const start = await pic();
    expect(start).toMatchObject({ cx: 1905000, cy: 952500, l: 0, t: 0, r: 0, b: 0 });      // 200 by 100 pixels at 96 dpi
    expect((await handles()).length).toBe(0);
    await cropMenu('Crop');
    const hs = await handles();
    expect(hs.length).toBe(8);

    // the right-hand middle handle, a quarter of the picture's width to the left
    const z = await zoom();
    const quarter = (start.cx / 4) * z;
    await drag(hs[4]!, -quarter, 0);
    const cut = await pic();
    expect(cut.x).toBe(start.x);
    expect(cut.y).toBe(start.y);
    expect(cut.cy).toBe(start.cy);
    expect(Math.abs(cut.cx - start.cx * 0.75)).toBeLessThan(3 / z);
    expect(Math.abs(cut.r - 25000)).toBeLessThan(400);
    expect([cut.l, cut.t, cut.b]).toEqual([0, 0, 0]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Crop');

    // the left-hand middle handle in: the frame starts further right, and the picture under it has not moved,
    // so the frame's left edge is now on the blue half
    const hs2 = await handles();
    await drag(hs2[3]!, (start.cx * 0.6) * z, 0);
    const both = await pic();
    expect(Math.abs(both.l - 60000)).toBeLessThan(400);
    expect(Math.abs(both.x - (start.x + start.cx * 0.6))).toBeLessThan(3 / z);
    await page.keyboard.press('Escape');
    expect((await handles()).length).toBe(0);
    const edge = await slidePoint(both.x + 30000, both.y + both.cy / 2);
    const px = await page.evaluate(([x, y]) => {
      const c = document.querySelector('[data-testid="canvas"] canvas') as HTMLCanvasElement;
      const r = c.getBoundingClientRect();
      const d = c.getContext('2d')!.getImageData(Math.round((x - r.left) * (c.width / r.width)), Math.round((y - r.top) * (c.height / r.height)), 1, 1).data;
      return [d[0], d[1], d[2]];
    }, [edge.x, edge.y] as const);
    expect(px).toEqual([0, 0, 200]);

    // each drag was one undo step: two undos give the whole picture back
    await page.keyboard.press('Control+z');
    expect(await pic()).toMatchObject({ x: cut.x, cx: cut.cx, l: 0 });
    await page.keyboard.press('Control+z');
    expect(await pic()).toEqual(start);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a handle cannot leave the picture; a corner crops two sides; Done Cropping, Enter and selecting something else end it', async () => {
    await picture();
    const start = await pic();
    await cropMenu('Crop');
    const z = await zoom();
    let hs = await handles();
    await drag(hs[4]!, 200, 0);                                         // outwards: there is no more picture to show
    expect(await pic()).toEqual(start);
    await drag(hs[7]!, -(start.cx / 2) * z, -(start.cy / 2) * z);       // the bottom right corner to the middle
    const corner = await pic();
    expect(Math.abs(corner.r - 50000)).toBeLessThan(600);
    expect(Math.abs(corner.b - 50000)).toBeLessThan(600);
    expect(corner.x).toBe(start.x);
    expect(corner.y).toBe(start.y);

    await cropMenu('Done Cropping');
    expect((await handles()).length).toBe(0);
    await cropMenu('Crop');
    expect((await handles()).length).toBe(8);
    await page.keyboard.press('Enter');
    expect((await handles()).length).toBe(0);
    await cropMenu('Crop');
    const empty = await slidePoint(600000, 600000);
    await page.mouse.click(empty.x, empty.y);
    expect(await state<boolean>('return !!s.cropping;')).toBe(false);

    // Reset Crop gives the whole picture back in its frame
    const mid = await slidePoint(corner.x + corner.cx / 2, corner.y + corner.cy / 2);
    await page.mouse.click(mid.x, mid.y);
    await cropMenu('Reset Crop');
    expect(await pic()).toMatchObject({ l: 0, t: 0, r: 0, b: 0 });
    hs = await handles();
    expect(hs.length).toBe(0);
    const ed = await saved();
    expect(ed.doc.slide(ed.doc.slideParts()[0]!).shapes.filter(sh => sh.kind === 'pic').length).toBe(1);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('connectors attached to shapes', () => {
  /** The frame of every drawn shape by name, and for a connector its ends and what they are attached to (shape names). */
  const scene = () => state<Record<string, { x: number; y: number; cx: number; cy: number; start?: [number, number]; end?: [number, number]; from?: string; to?: string }>>(`
    const find = (e, n) => e && e.children.find(c => c.kind === 'element' && c.local === n);
    const num = (e, n) => { const a = e && e.attrs.find(t => t.local === n); return a ? Number(a.value) : 0; };
    const shapes = s.session.ed.doc.slide(s.slide).shapes.filter(x => !x.placeholder);
    const byId = new Map(shapes.map(x => [x.id, x.name]));
    const out = {};
    for (const sh of shapes) {
      const x = sh.xfrm; const o = { x: x.x, y: x.y, cx: x.cx, cy: x.cy };
      if (sh.kind === 'cxnSp') {
        o.start = [x.flipH ? x.x + x.cx : x.x, x.flipV ? x.y + x.cy : x.y];
        o.end = [x.flipH ? x.x : x.x + x.cx, x.flipV ? x.y : x.y + x.cy];
        const nv = find(find(sh.el, 'nvCxnSpPr'), 'cNvCxnSpPr'); const st = find(nv, 'stCxn'), en = find(nv, 'endCxn');
        if (st) o.from = byId.get(num(st, 'id')) + '#' + num(st, 'idx');
        if (en) o.to = byId.get(num(en, 'id')) + '#' + num(en, 'idx');
      }
      out[sh.name.replace(/ \\d+$/, '')] = o;
    }
    return out;`);
  const A = { x: 1000000, y: 5500000, w: 1500000, h: 800000 }, B = { x: 5000000, y: 5500000, w: 1500000, h: 800000 };
  /** Two rectangles under the placeholders: "Rectangle" (A) on the left and "Rectangle: Rounded Corners" (B) on the right. */
  async function twoBoxes(): Promise<void> {
    await drawShape('rect', A.x, A.y, A.w, A.h);
    await drawShape('roundRect', B.x, B.y, B.w, B.h);
    await page.keyboard.press('Escape');
  }
  /** Drag on the slide from one point to another (EMU), a few pixels off at each end if asked. */
  async function dragSlide(from: [number, number], to: [number, number], off = 0): Promise<void> {
    await frames();
    const a = await slidePoint(...from), b = await slidePoint(...to);
    await page.mouse.move(a.x + off, a.y + off);
    await page.mouse.down();
    await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 3 });
    await page.mouse.move(b.x - off, b.y + off, { steps: 3 });
    await page.mouse.up();
    await frames();
  }
  async function drawConnector(from: [number, number], to: [number, number], off = 0): Promise<void> {
    await page.getByTestId('shapes').click();
    await page.locator('[data-testid="shape-gallery"] button[data-prst="straightConnector1"]').first().click();
    await dragSlide(from, to, off);
  }
  /** Where the end handles of the selected connector are on the page: [start, end]. */
  const endHandles = () => page.evaluate(() => {
    const st = (globalThis as unknown as { __stage: { find(s: string): { getClientRect(): { x: number; y: number; width: number; height: number } }[]; container(): HTMLElement } }).__stage;
    const r = st.container().getBoundingClientRect();
    return st.find('.end-handle').map(h => { const b = h.getClientRect(); return { x: r.left + b.x + b.width / 2, y: r.top + b.y + b.height / 2 }; });
  });
  const A_RIGHT: [number, number] = [A.x + A.w, A.y + A.h / 2], B_LEFT: [number, number] = [B.x, B.y + B.h / 2];

  test('a connector drawn from one shape to another attaches to their connection points, and follows a shape that is moved', async () => {
    await twoBoxes();
    await drawConnector(A_RIGHT, B_LEFT, 5);                             // let go a few pixels off each point: it snaps
    let sc = await scene();
    const a = sc.Rectangle!, b = sc['Rectangle: Rounded Corners']!;
    const c = sc['Straight Arrow Connector']!;
    expect(c.from).toBe('Rectangle 3#3');                               // the right side of A
    expect(c.to).toBe('Rectangle: Rounded Corners 4#1');                         // the left side of B
    expect(c.start).toEqual([a.x + a.cx, Math.round(a.y + a.cy / 2)]);
    expect(c.end).toEqual([b.x, Math.round(b.y + b.cy / 2)]);

    // drag B up and to the right: the connector's end is still on B's left side, its start still on A
    await dragSlide([b.x + b.cx / 2, b.y + b.cy / 2], [b.x + b.cx / 2 + 1500000, b.y + b.cy / 2 - 3000000]);
    sc = await scene();
    const moved = sc['Rectangle: Rounded Corners']!;
    expect(moved.y).toBeLessThan(b.y - 2000000);
    expect(sc['Straight Arrow Connector']!.end).toEqual([moved.x, Math.round(moved.y + moved.cy / 2)]);
    expect(sc['Straight Arrow Connector']!.start).toEqual([a.x + a.cx, Math.round(a.y + a.cy / 2)]);
    // moving the shape and its connector was one step
    await page.keyboard.press('Control+z');
    sc = await scene();
    expect(sc['Rectangle: Rounded Corners']).toEqual(b);
    expect(sc['Straight Arrow Connector']!.end).toEqual([b.x, Math.round(b.y + b.cy / 2)]);

    const ed = await saved();
    const xml = new TextDecoder().decode(ed.pkg.getPart(ed.doc.slideParts()[0]!)!.bytes);
    expect(xml).toMatch(/<a:stCxn id="\d+" idx="3"\/><a:endCxn id="\d+" idx="1"\/>/);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a selected connector has a handle on each end: dragged onto a connection point it attaches, dragged away it lets go', async () => {
    await twoBoxes();
    await drawConnector([3000000, 4900000], [4500000, 4900000]);         // in the open: attached to nothing
    let c = (await scene())['Straight Arrow Connector']!;
    expect([c.from, c.to]).toEqual([undefined, undefined]);
    const hs = await endHandles();
    expect(hs.length).toBe(2);

    // the end onto B's left side
    const bLeft = await slidePoint(...B_LEFT);
    await page.mouse.move(hs[1]!.x, hs[1]!.y);
    await page.mouse.down();
    await page.mouse.move(bLeft.x - 30, bLeft.y - 10, { steps: 3 });
    expect(await page.evaluate(() => (globalThis as unknown as { __stage: { find(s: string): unknown[] } }).__stage.find('.site-hint').length)).toBeGreaterThanOrEqual(4);   // B's points show
    await page.mouse.move(bLeft.x - 4, bLeft.y + 3, { steps: 3 });
    await page.mouse.up();
    await frames();
    expect(await page.evaluate(() => (globalThis as unknown as { __stage: { find(s: string): unknown[] } }).__stage.find('.site-hint').length)).toBe(0);
    let sc = await scene();
    c = sc['Straight Arrow Connector']!;
    const b = sc['Rectangle: Rounded Corners']!;
    expect(c.to).toBe('Rectangle: Rounded Corners 4#1');
    expect(c.end).toEqual([b.x, Math.round(b.y + b.cy / 2)]);
    expect(c.from).toBeUndefined();
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Move End Point');

    // the start onto A's top (it then runs from A's top to B's left); then the end away again
    const hs2 = await endHandles();
    const aTop = await slidePoint(A.x + A.w / 2, A.y);
    await page.mouse.move(hs2[0]!.x, hs2[0]!.y);
    await page.mouse.down();
    await page.mouse.move(aTop.x + 3, aTop.y - 2, { steps: 5 });
    await page.mouse.up();
    await frames();
    c = (await scene())['Straight Arrow Connector']!;
    expect(c.from).toBe('Rectangle 3#0');
    const hs3 = await endHandles();
    await page.mouse.move(hs3[1]!.x, hs3[1]!.y);
    await page.mouse.down();
    await page.mouse.move(hs3[1]!.x - 80, hs3[1]!.y - 120, { steps: 5 });
    await page.mouse.up();
    await frames();
    c = (await scene())['Straight Arrow Connector']!;
    expect([c.from, c.to]).toEqual(['Rectangle 3#0', undefined]);

    // each drag was one undo step
    await page.keyboard.press('Control+z');
    expect((await scene())['Straight Arrow Connector']!.to).toBe('Rectangle: Rounded Corners 4#1');
    await page.keyboard.press('Control+z');
    expect((await scene())['Straight Arrow Connector']!.from).toBeUndefined();
    expect(errors).toEqual([]);
  }, 60_000);

  test('a line keeps the direction it was drawn in, and a shape with no connector moves as before', async () => {
    await twoBoxes();
    await page.getByTestId('shapes').click();
    await page.locator('[data-testid="shape-gallery"] button[data-prst="line"]').first().click();
    await dragSlide([4400000, 5000000], [3000000, 4700000]);             // right to left, and upwards
    const sc = await scene();
    const line = sc['Straight Connector'] ?? sc.Line!;
    const z = await page.evaluate(() => (globalThis as unknown as { __stage: { scaleX(): number } }).__stage.scaleX());
    expect(Math.abs(line.start![0] - 4400000)).toBeLessThan(3 / z);
    expect(Math.abs(line.end![0] - 3000000)).toBeLessThan(3 / z);
    expect(line.start![1]).toBeGreaterThan(line.end![1]);
    expect([line.from, line.to]).toEqual([undefined, undefined]);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('grid view', () => {
  const marks = () => state<string[]>(`
    const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
    return s.session.ed.doc.slideParts().map(p => s.session.ed.doc.slide(p).shapes.filter(sh => sh.name.startsWith('TextBox')).map(sh => textOf(sh.el)).join('+'));`);
  const current = () => state<number>('return s.session.ed.doc.slideParts().indexOf(s.slide) + 1;');
  const grid = () => page.getByTestId('slide-grid');
  /** Five slides marked "one" to "five", shown in grid view. */
  async function fiveInGrid(): Promise<void> {
    for (const [i, mark] of ['one', 'two', 'three', 'four', 'five'].entries()) {
      if (i) await page.getByTestId('new-slide').click();
      await textBox(mark);
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
    }
    await page.getByTestId('view-grid').click();
    await grid().waitFor();
  }

  test('every slide is shown as a drawn thumbnail; the slide itself and the side panel give way; the button and Escape go back', async () => {
    await fiveInGrid();
    expect(await grid().locator('.thumb').count()).toBe(5);
    expect(await page.getByTestId('canvas').count()).toBe(0);
    expect(await page.locator('aside.slides').count()).toBe(0);
    expect(await page.getByTestId('view-grid').getAttribute('aria-pressed')).toBe('true');
    expect(await page.getByTestId('grid-thumb-5').getAttribute('class')).toContain('current');
    // the thumbnails are drawn: each is a white slide with the (small, grey-edged) marks of its text on it
    await expect.poll(() => page.evaluate(() => [...document.querySelectorAll('[data-testid="slide-grid"] canvas')].map(c => {
      const cv = c as HTMLCanvasElement;
      if (!cv.width) return 'empty';
      const d = cv.getContext('2d')!.getImageData(0, 0, cv.width, cv.height).data;
      let white = 0, marks = 0;
      for (let i = 0; i < d.length; i += 4) { if (d[i]! > 250 && d[i + 1]! > 250 && d[i + 2]! > 250 && d[i + 3]! === 255) white++; else if (d[i + 3]! > 0 && d[i]! < 235) marks++; }
      return white > d.length / 8 && marks > 3 ? 'drawn' : `white ${white} marks ${marks}`;
    })), { timeout: 15_000 }).toEqual(['drawn', 'drawn', 'drawn', 'drawn', 'drawn']);
    // wider than the side panel's
    expect(await page.getByTestId('grid-thumb-1').locator('canvas').evaluate(c => (c as HTMLElement).getBoundingClientRect().width)).toBe(240);

    await page.keyboard.press('Escape');
    expect(await grid().count()).toBe(0);
    expect(await page.getByTestId('canvas').count()).toBe(1);
    await page.getByTestId('view-grid').click();
    await page.getByTestId('view-grid').click();
    expect(await grid().count()).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a click picks a slide, the arrow keys move in all four directions, and Enter or a double click opens the slide', async () => {
    await fiveInGrid();
    await page.setViewportSize({ width: 1000, height: 860 });            // three thumbnails to a row
    await page.getByTestId('grid-thumb-1').click();
    expect(await current()).toBe(1);
    await grid().focus();
    await page.keyboard.press('ArrowRight');
    expect(await current()).toBe(2);
    await page.keyboard.press('ArrowDown');                              // a row down
    expect(await current()).toBe(5);
    await page.keyboard.press('ArrowUp');
    expect(await current()).toBe(2);
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('ArrowLeft');                              // no slide before the first
    expect(await current()).toBe(1);
    await page.keyboard.press('End');
    expect(await current()).toBe(5);
    await page.keyboard.press('Home');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Enter');
    expect(await grid().count()).toBe(0);
    expect(await current()).toBe(3);
    expect(await page.getByTestId('canvas').count()).toBe(1);

    await page.getByTestId('view-grid').click();
    await page.getByTestId('grid-thumb-4').dblclick();
    expect(await grid().count()).toBe(0);
    expect(await current()).toBe(4);
    await page.setViewportSize({ width: 1400, height: 860 });
    expect(errors).toEqual([]);
  }, 60_000);

  test('slides are reordered by dragging, deleted, duplicated, and copied and pasted in the grid', async () => {
    await fiveInGrid();
    await page.getByTestId('grid-thumb-5').dragTo(page.getByTestId('grid-thumb-2'));
    expect(await marks()).toEqual(['one', 'five', 'two', 'three', 'four']);
    await page.getByTestId('grid-thumb-1').click();
    await grid().focus();
    await page.keyboard.press('Delete');
    expect(await marks()).toEqual(['five', 'two', 'three', 'four']);
    await page.keyboard.press('Control+d');
    expect(await marks()).toEqual(['five', 'five', 'two', 'three', 'four']);
    const copied = await clipboardEvent('copy');
    expect(copied.html).toContain('data-pptx-slide');
    await page.keyboard.press('End');
    await clipboardEvent('paste', { 'text/plain': copied.plain, 'text/html': copied.html });
    expect(await marks()).toEqual(['five', 'five', 'two', 'three', 'four', 'five']);
    // the right-click menu of a thumbnail works here too
    await page.getByTestId('grid-thumb-3').click({ button: 'right' });
    await page.getByTestId('menu-delete-slide').click();
    expect(await marks()).toEqual(['five', 'five', 'three', 'four', 'five']);
    // the ribbon's New Slide adds after the current slide and the grid shows it
    await page.getByTestId('new-slide').click();
    expect(await grid().locator('.thumb').count()).toBe(6);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('rulers and guides', () => {
  const guides = () => state<{ x: number[]; y: number[] }>('return s.guides;');
  const view = async (testid: string) => { await page.getByRole('tab', { name: 'View' }).click(); await page.getByTestId(testid).click(); };
  const zoomNow = () => page.evaluate(() => (globalThis as unknown as { __stage: { scaleX(): number } }).__stage.scaleX());
  /** The selected shape's frame. */
  const frame = () => state<{ x: number; y: number; cx: number; cy: number }>(`const x = s.session.ed.view(s.slide, s.selection[0]).xfrm; return { x: x.x, y: x.y, cx: x.cx, cy: x.cy };`);
  /** Drag on the page between two page points, in small steps. */
  async function dragPage(a: { x: number; y: number }, b: { x: number; y: number }): Promise<void> {
    await page.mouse.move(a.x, a.y);
    await page.mouse.down();
    await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 4 });
    await page.mouse.move(b.x, b.y, { steps: 4 });
    await page.mouse.up();
    await frames();
  }

  test('View > Ruler shows rulers in inches from the slide\'s corner that follow the zoom; the button hides them again', async () => {
    expect(await page.getByTestId('ruler-h').count()).toBe(0);
    await view('view-ruler');
    await page.getByTestId('ruler-h').waitFor();
    await frames();
    expect(await page.getByTestId('view-ruler').getAttribute('aria-pressed')).toBe('true');
    // a 13.33 inch wide slide: the numbers 0 to 13 across, 0 to 7 down, each where that inch is on the slide
    const labels = async (id: string) => page.getByTestId(id).locator('text').evaluateAll(els => els.map(e => [e.textContent, Number(e.getAttribute(e.closest('.ruler-h') ? 'x' : 'y'))] as [string, number]));
    const across = await labels('ruler-h');
    for (const n of ['0', '1', '6', '13']) expect(across.map(l => l[0])).toContain(n);
    expect((await labels('ruler-v')).map(l => l[0])).toEqual(expect.arrayContaining(['0', '3', '7']));
    const host = (await page.getByTestId('canvas').boundingBox())!;
    const zero = await slidePoint(0, 0), six = await slidePoint(6 * 914400, 0);
    const at = (n: string) => across.filter(l => l[0] === n).map(l => l[1]).find(x => Math.abs(x - 3 - ((n === '0' ? zero.x : six.x) - host.x)) < 1.5);
    expect(at('0')).toBeDefined();
    expect(at('6')).toBeDefined();

    // zoomed to 200% the marks spread out: fewer inches fit, and an inch on the ruler is 192 pixels
    const gap = (ls: [string, number][]) => { const xs = ls.map(l => l[1]).sort((m, n) => m - n); return xs[1]! - xs[0]!; };
    const fitGap = gap(across);
    await page.locator('.status .zoom button', { hasText: '200%' }).click();
    await frames();
    const wide = await labels('ruler-h');
    expect(wide.length).toBeLessThan(across.length);
    expect(gap(wide)).toBeCloseTo(192, 0);
    expect(gap(wide)).toBeGreaterThan(fitGap * 1.5);
    await page.locator('.status .zoom button', { hasText: 'Fit' }).click();

    await page.getByTestId('view-ruler').click();
    expect(await page.getByTestId('ruler-h').count()).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('guides: added in the middle, a dragged shape snaps to one, a guide is dragged to a new place and off the slide to remove it', async () => {
    await view('add-guide-v');
    await view('add-guide-h');
    expect(await guides()).toEqual({ x: [6096000], y: [3429000] });
    await page.getByTestId('add-guide-v').click();
    expect((await guides()).x).toEqual([6096000, 6096000 + 457200]);     // not on top of the first
    expect(await page.evaluate(() => (globalThis as unknown as { __stage: { find(s: string): unknown[] } }).__stage.find('.user-guide').length)).toBe(3);

    // a rectangle dragged until its left edge is a few pixels from the first vertical guide snaps onto it
    await drawShape('rect', 1000000, 5500000, 1500000, 800000);
    const z = await zoomNow();
    const before = await frame();
    const from = await slidePoint(before.x + before.cx / 2, before.y + before.cy / 2);
    const wanted = (6096000 - before.x) * z;                             // how far right the left edge has to go
    await dragPage(from, { x: from.x + wanted + 4, y: from.y - 40 });
    expect((await frame()).x).toBe(6096000);

    // the guide itself dragged left by 200 pixels
    await page.keyboard.press('Escape');
    const g = await slidePoint(6096000, 600000);
    await dragPage(g, { x: g.x - 200, y: g.y + 5 });
    const moved = (await guides()).x[0]!;
    expect(Math.abs(moved - (6096000 - 200 / z))).toBeLessThan(3 / z);
    expect(await state<string | undefined>('return s.session.ed.log.undoLabel;')).toBe('Move');      // moving a guide is not a deck change

    // dragged off the left edge of the slide, it is gone
    const g2 = await slidePoint(moved, 600000);
    const off = await slidePoint(-300000, 600000);
    await dragPage(g2, off);
    expect((await guides()).x).toEqual([6096000 + 457200]);

    // hidden guides do not snap; Clear Guides removes them all
    await view('view-guides');
    expect(await page.evaluate(() => (globalThis as unknown as { __stage: { find(s: string): unknown[] } }).__stage.find('.user-guide').length)).toBe(0);
    await page.getByTestId('clear-guides').click();
    expect(await guides()).toEqual({ x: [], y: [] });
    expect(await page.getByTestId('clear-guides').isDisabled()).toBe(true);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a guide is pulled out of a ruler: down from the top ruler for a horizontal one, across from the left for a vertical one', async () => {
    await view('view-ruler');
    await page.getByTestId('ruler-h').waitFor();
    await frames();
    const top = (await page.getByTestId('ruler-h').boundingBox())!;
    const to = await slidePoint(3000000, 2000000);
    await dragPage({ x: top.x + 300, y: top.y + 8 }, { x: to.x, y: to.y });
    const z = await zoomNow();
    let g = await guides();
    expect(g.x).toEqual([]);
    expect(g.y.length).toBe(1);
    expect(Math.abs(g.y[0]! - 2000000)).toBeLessThan(3 / z);

    const left = (await page.getByTestId('ruler-v').boundingBox())!;
    await dragPage({ x: left.x + 8, y: left.y + 300 }, { x: to.x, y: to.y });
    g = await guides();
    expect(g.x.length).toBe(1);
    expect(Math.abs(g.x[0]! - 3000000)).toBeLessThan(3 / z);
    // pulled out but let go before reaching the slide: no guide is left
    await dragPage({ x: top.x + 300, y: top.y + 8 }, { x: top.x + 300, y: top.y + 14 });
    expect((await guides()).y.length).toBe(1);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('presenter view and the laser pointer', () => {
  /** Three slides marked in a text box and in their notes; the first slide on screen. */
  async function threeWithNotes(): Promise<void> {
    for (const [i, mark] of ['one', 'two', 'three'].entries()) {
      if (i) await page.getByTestId('new-slide').click();
      await textBox(mark);
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
      if (i !== 2) {                                                     // the last slide has no notes
        await page.getByTestId('notes').click();
        await page.keyboard.type(`Say ${mark}.`);
        await page.keyboard.press('Escape');
      }
    }
    await state('s.setSlide(s.session.ed.doc.slideParts()[0]);');
  }
  const startPresenter = async () => {
    await page.getByRole('tab', { name: 'Slide Show' }).click();
    await page.getByTestId('presenter-view').click();
    await page.getByTestId('presenter-slide').waitFor();
  };
  const text = (id: string) => page.getByTestId(id).innerText();

  test('the slide with the next slide, the notes, a count and a running timer; the keys move through the show; Escape goes back to the slide shown', async () => {
    await threeWithNotes();
    await startPresenter();
    expect(await text('presenter-count')).toBe('Slide 1 of 3');
    expect(await text('presenter-notes')).toBe('Say one.');
    expect(await page.getByTestId('presenter-next').count()).toBe(1);
    // the slide is drawn in its part of the window, not over all of it
    const slide = (await page.getByTestId('presenter-slide').locator('canvas').boundingBox())!;
    expect(slide.width).toBeGreaterThan(600);
    expect(slide.width).toBeLessThan(1000);
    expect(Math.abs(slide.width / slide.height - 16 / 9)).toBeLessThan(0.02);
    // the next slide's picture arrives
    await expect.poll(() => page.getByTestId('presenter-next').evaluate(c => (c as HTMLCanvasElement).width), { timeout: 15_000 }).toBeGreaterThan(100);

    // the timer runs, and a click starts it again
    expect(await text('presenter-timer')).toMatch(/^0:0[0-2]$/);
    await page.waitForTimeout(1300);
    expect(await text('presenter-timer')).toMatch(/^0:0[1-4]$/);
    await page.getByTestId('presenter-timer').click();
    expect(await text('presenter-timer')).toBe('0:00');
    expect(await text('presenter-clock')).toMatch(/\d{1,2}:\d{2}/);

    await page.keyboard.press('ArrowRight');
    await expect.poll(() => text('presenter-count')).toBe('Slide 2 of 3');
    expect(await text('presenter-notes')).toBe('Say two.');
    await page.getByTestId('presenter-nextbtn').click();
    await expect.poll(() => text('presenter-count')).toBe('Slide 3 of 3');
    expect(await text('presenter-notes')).toBe('No notes for this slide.');
    expect(await page.getByTestId('presenter-next-end').innerText()).toBe('End of slide show');
    await page.getByTestId('presenter-prev').click();
    await expect.poll(() => text('presenter-count')).toBe('Slide 2 of 3');

    await page.keyboard.press('Escape');
    await page.getByTestId('canvas').waitFor();
    expect(await state<number>('return s.session.ed.doc.slideParts().indexOf(s.slide) + 1;')).toBe(2);
    expect(errors).toEqual([]);
  }, 60_000);

  test('after the last slide the count says the show has ended; End Show leaves it', async () => {
    await threeWithNotes();
    await state('s.setSlide(s.session.ed.doc.slideParts()[2]);');
    await startPresenter();
    expect(await text('presenter-count')).toBe('Slide 3 of 3');
    await page.keyboard.press('ArrowRight');
    await expect.poll(() => text('presenter-count')).toBe('End of slide show');
    await page.getByTestId('presenter-end').click();
    await page.getByTestId('canvas').waitFor();
    expect(await page.getByTestId('presenter').count()).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('L turns on a laser pointer that follows the mouse, and while it is on a click does not advance the show', async () => {
    await threeWithNotes();
    await startPresenter();
    const box = (await page.getByTestId('presenter-slide').boundingBox())!;
    await page.mouse.move(box.x + 100, box.y + 100);
    expect(await page.getByTestId('laser').count()).toBe(0);
    await page.keyboard.press('l');
    expect(await page.getByTestId('presenter-laser').getAttribute('aria-pressed')).toBe('true');
    await page.mouse.move(box.x + 220, box.y + 160);
    const dot = (await page.getByTestId('laser').boundingBox())!;
    expect(Math.abs(dot.x + dot.width / 2 - (box.x + 220))).toBeLessThan(2);
    expect(Math.abs(dot.y + dot.height / 2 - (box.y + 160))).toBeLessThan(2);
    await page.mouse.click(box.x + 220, box.y + 160);
    expect(await text('presenter-count')).toBe('Slide 1 of 3');           // pointing is not clicking
    await page.keyboard.press('l');
    expect(await page.getByTestId('laser').count()).toBe(0);
    await page.mouse.click(box.x + 220, box.y + 160);
    await expect.poll(() => text('presenter-count')).toBe('Slide 2 of 3');
    await page.keyboard.press('Escape');

    // the plain slide show has the pointer too
    await page.getByTestId('canvas').waitFor();
    await page.getByRole('tab', { name: 'Slide Show' }).click();
    await page.getByTestId('present').click();
    await page.getByTestId('presenter').waitFor();
    expect(await page.getByTestId('presenter-notes').count()).toBe(0);
    await page.keyboard.press('L');
    await page.mouse.move(400, 300);
    await page.mouse.move(420, 310);
    expect(await page.getByTestId('laser').count()).toBe(1);
    await page.keyboard.press('Escape');
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('saving, autosave and recovery', () => {
  const docName = () => page.getByTestId('doc-name').innerText();
  /** Write the recovery copy now instead of after the usual wait. */
  const autosaveNow = () => page.evaluate(`import('/src/state/store.ts').then(m => m.autosaveNow())`);
  /** What autosave has kept: the deck's name and whether it holds bytes, or null. */
  const kept = () => page.evaluate(() => new Promise<{ name: string; bytes: number } | null>(done => {
    const open = indexedDB.open('presentations', 1);
    open.onupgradeneeded = () => open.result.createObjectStore('autosave');
    open.onsuccess = () => {
      const get = open.result.transaction('autosave').objectStore('autosave').get('current');
      get.onsuccess = () => { open.result.close(); done(get.result ? { name: get.result.name, bytes: get.result.bytes.length } : null); };
      get.onerror = () => { open.result.close(); done(null); };
    };
    open.onerror = () => done(null);
  }));
  const boxTexts = () => state<string[]>(`
    const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
    return s.session.ed.doc.slide(s.slide).shapes.filter(sh => sh.name.startsWith('TextBox')).map(sh => textOf(sh.el));`);
  /** A file picker of the test's own in place of the browser's: it hands out a file that remembers what is written to it. */
  const fakePicker = (name: string | null) => page.evaluate(n => {
    const w = window as unknown as { __picks: unknown[]; __written: number[][]; showSaveFilePicker: (o: unknown) => Promise<unknown> };
    w.__picks = [];
    w.__written = [];
    w.showSaveFilePicker = async (options: unknown) => {
      w.__picks.push(options);
      if (n === null) throw new DOMException('The user aborted a request.', 'AbortError');
      return { name: n, getFile: async () => new File([], n), createWritable: async () => ({ write: async (data: Uint8Array) => { w.__written.push(Array.from(data)); }, close: async () => undefined }) };
    };
  }, name);
  const picks = () => page.evaluate(() => (window as unknown as { __picks: { suggestedName: string }[] }).__picks.map(p => p.suggestedName));
  const written = () => page.evaluate(() => (window as unknown as { __written: number[][] }).__written);

  test('a deck with unsaved changes is kept by autosave, offered on the welcome screen after the tab is closed, and restored whole', async () => {
    expect(await docName()).toBe('Presentation1.pptx');
    expect(await kept()).toBeNull();
    await textBox('keep me');
    await page.keyboard.press('Escape');
    expect(await docName()).toBe('Presentation1.pptx (unsaved changes)');
    await autosaveNow();
    const record = await kept();
    expect(record?.name).toBe('Presentation1.pptx');
    expect(record!.bytes).toBeGreaterThan(1000);

    // the tab goes away without a save; the app, opened again, offers the deck
    await page.goto(server.resolvedUrls!.local[0]!);
    await page.getByTestId('recover').waitFor();
    expect(await page.getByTestId('recover').innerText()).toContain('Presentation1.pptx has changes that were not saved');
    await page.getByTestId('recover-restore').click();
    await page.getByTestId('canvas').waitFor();
    expect(await boxTexts()).toEqual(['keep me']);
    expect(await docName()).toBe('Presentation1.pptx (unsaved changes)');     // restored, and still not saved anywhere

    // a real save: nothing is unsaved, and nothing is offered next time
    const [download] = await Promise.all([page.waitForEvent('download'), (async () => { await page.getByRole('tab', { name: 'File' }).click(); await page.getByTestId('save').click(); })()]);
    expect(download.suggestedFilename()).toBe('Presentation1.pptx');
    await expect.poll(docName).toBe('Presentation1.pptx');
    await expect.poll(kept).toBeNull();
    await page.goto(server.resolvedUrls!.local[0]!);
    await page.getByTestId('welcome-new').waitFor();
    await page.waitForTimeout(150);
    expect(await page.getByTestId('recover').count()).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('autosave follows the deck by itself a few seconds after a change; undoing back to the saved state leaves nothing to keep; Discard throws the kept deck away', async () => {
    await textBox('typed once');
    await page.keyboard.press('Escape');
    await expect.poll(kept, { timeout: 12_000 }).not.toBeNull();          // with no call from the test: the timer wrote it
    // undone to where the deck started: still counted as changed (an undo is an edit), so still kept
    await page.keyboard.press('Escape');
    await page.keyboard.press('Control+z');
    await page.keyboard.press('Control+z');
    expect(await docName()).toContain('unsaved changes');

    await page.goto(server.resolvedUrls!.local[0]!);
    await page.getByTestId('recover').waitFor();
    await page.getByTestId('recover-discard').click();
    expect(await page.getByTestId('recover').count()).toBe(0);
    await expect.poll(kept).toBeNull();
    expect(errors).toEqual([]);
  }, 60_000);

  test('Save asks where the first time, then writes the same file with no questions; Save As asks again', async () => {
    await fakePicker('chosen.pptx');
    await textBox('first');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Control+s');
    await expect.poll(async () => (await written()).length).toBe(1);
    expect(await picks()).toEqual(['Presentation1.pptx']);
    await expect.poll(docName).toBe('chosen.pptx');                      // named after the file, and nothing unsaved
    await expect.poll(() => page.locator('.status .message').innerText()).toContain('Saved chosen.pptx.');
    // what was written is the deck
    const first = new Editor(OpcPackage.open(new Uint8Array((await written())[0]!)));
    expect(first.doc.slide(first.doc.slideParts()[0]!).shapes.some(sh => sh.name.startsWith('TextBox'))).toBe(true);

    await page.getByTestId('new-slide').click();
    expect(await docName()).toBe('chosen.pptx (unsaved changes)');
    await page.getByRole('tab', { name: 'File' }).click();
    await page.getByTestId('save-file').click();
    await expect.poll(async () => (await written()).length).toBe(2);
    expect((await picks()).length).toBe(1);                              // no second question
    const second = new Editor(OpcPackage.open(new Uint8Array((await written())[1]!)));
    expect(second.doc.slideParts().length).toBe(2);

    await page.getByTestId('save-as').click();
    await expect.poll(async () => (await written()).length).toBe(3);
    expect(await picks()).toEqual(['Presentation1.pptx', 'chosen.pptx']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a file dialog that is cancelled saves nothing and says nothing; where the browser cannot write files, Save downloads', async () => {
    await fakePicker(null);
    await textBox('unsaved');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Control+s');
    await expect.poll(async () => (await picks()).length).toBe(1);
    expect(await written()).toEqual([]);
    expect(await docName()).toBe('Presentation1.pptx (unsaved changes)');
    expect(await page.locator('.status .message.error').count()).toBe(0);

    await page.evaluate(() => { delete (window as unknown as { showSaveFilePicker?: unknown }).showSaveFilePicker; });
    const [download] = await Promise.all([page.waitForEvent('download'), page.keyboard.press('Control+s')]);
    expect(download.suggestedFilename()).toBe('Presentation1.pptx');
    await expect.poll(docName).toBe('Presentation1.pptx');
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('export', () => {
  /** Click an export button of the File tab and read the file the browser is given. */
  async function exported(testid: string): Promise<{ name: string; bytes: Uint8Array }> {
    await page.getByRole('tab', { name: 'File' }).click();
    const [d] = await Promise.all([page.waitForEvent('download', { timeout: 30_000 }), page.getByTestId(testid).click()]);
    await page.getByRole('tab', { name: 'Home' }).click();
    return { name: d.suggestedFilename(), bytes: new Uint8Array(readFileSync((await d.path())!)) };
  }
  /** A red rectangle on the first slide and a second, hidden-then-shown slide with a blue one. */
  async function twoSlides(): Promise<void> {
    await drawShape('rect', 1000000, 5400000, 2000000, 1000000);
    await page.getByTestId('shape-fill').click();
    await page.locator('.menu-panel .swatch[title="Red"]').click();
    await page.keyboard.press('Escape');
    await page.getByTestId('new-slide').click();
    await drawShape('rect', 8000000, 400000, 2000000, 1000000);
    await page.getByTestId('shape-fill').click();
    await page.locator('.menu-panel .swatch[title="Blue"]').click();
    await page.keyboard.press('Escape');
  }
  const at = (png: PNG, x: number, y: number) => { const i = (y * png.width + x) * 4; return [png.data[i], png.data[i + 1], png.data[i + 2], png.data[i + 3]]; };

  test('the current slide as a PNG: full HD, white where the slide is empty, and the shapes where they are', async () => {
    await twoSlides();
    await state('s.setSlide(s.session.ed.doc.slideParts()[0]);');
    const file = await exported('export-png');
    expect(file.name).toBe('Presentation1 - Slide 1.png');
    const png = PNG.sync.read(Buffer.from(file.bytes));
    expect([png.width, png.height]).toEqual([1920, 1080]);
    expect(at(png, 20, 20)).toEqual([255, 255, 255, 255]);
    // the rectangle spans 1M..3M by 5.4M..6.4M EMU of a 12.192M wide slide: its middle is at about (315, 929)
    const px = at(png, Math.round((2000000 / 12192000) * 1920), Math.round((5900000 / 12192000) * 1920));
    expect(px).toEqual([255, 0, 0, 255]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Shape Fill');      // exporting changes nothing
    expect(errors).toEqual([]);
  }, 60_000);

  test('every slide as PNGs in a ZIP, named so they sort in order', async () => {
    await twoSlides();
    const file = await exported('export-pngs');
    expect(file.name).toBe('Presentation1 - slides.zip');
    const entries = readZip(file.bytes);
    expect(entries.map(e => e.name)).toEqual(['Slide 1.png', 'Slide 2.png']);
    const second = PNG.sync.read(Buffer.from(entries[1]!.data));
    expect([second.width, second.height]).toEqual([1920, 1080]);
    expect(at(second, Math.round((9000000 / 12192000) * 1920), Math.round((900000 / 12192000) * 1920))).toEqual([0, 112, 192, 255]);   // "Blue" is #0070C0
    expect(errors).toEqual([]);
  }, 60_000);

  test('the deck as a PDF: a page per slide at the slide\'s size, each a JPEG; a hidden slide is left out', async () => {
    await twoSlides();
    let file = await exported('export-pdf');
    expect(file.name).toBe('Presentation1.pdf');
    let pdf = Buffer.from(file.bytes).toString('latin1');
    expect(pdf.startsWith('%PDF-1.4')).toBe(true);
    expect(pdf.trimEnd().endsWith('%%EOF')).toBe(true);
    expect(pdf.match(/\/Type \/Page /g)?.length).toBe(2);
    expect(pdf.match(/\/MediaBox \[0 0 960 540\]/g)?.length).toBe(2);     // 13.33 by 7.5 inches
    expect(pdf.match(/\/Filter \/DCTDecode/g)?.length).toBe(2);
    expect(pdf).toContain('/Width 1920 /Height 1080');
    expect(pdf).toContain('/Title (Presentation1)');
    // each picture really is a JPEG
    const first = /\/Length (\d+) >>\nstream\n/.exec(pdf.slice(pdf.indexOf('/Subtype /Image')))!;
    const from = pdf.indexOf('/Subtype /Image') + first.index + first[0].length;
    expect([file.bytes[from], file.bytes[from + 1]]).toEqual([0xff, 0xd8]);
    expect([file.bytes[from + Number(first[1]) - 2], file.bytes[from + Number(first[1]) - 1]]).toEqual([0xff, 0xd9]);

    await page.getByRole('tab', { name: 'Slide Show' }).click();
    await page.getByRole('button', { name: 'Hide Slide' }).click();
    file = await exported('export-pdf');
    pdf = Buffer.from(file.bytes).toString('latin1');
    expect(pdf.match(/\/Type \/Page /g)?.length).toBe(1);
    await expect.poll(() => page.locator('.status .message').innerText()).toContain('Exported Presentation1.pdf.');
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('comments', () => {
  const pane = () => page.getByTestId('comments-pane');
  /** The comments shown: "author: text" each. */
  const shown = () => pane().getByTestId('comment').evaluateAll(els => els.map(e => `${e.querySelector('strong')!.textContent}: ${e.querySelector('.comment-text')?.textContent ?? '(editing)'}`));
  const openPane = async () => { await page.getByRole('tab', { name: 'Insert' }).click(); await page.getByTestId('insert-comment').click(); await pane().waitFor(); };
  const add = async (text: string) => { await page.getByTestId('comment-new').fill(text); await page.getByTestId('comment-add').click(); };

  test('Insert > Comment opens the pane with the keyboard in the new comment; a name is asked for once, and the comment is on the slide', async () => {
    await page.evaluate(() => localStorage.removeItem('presentations.commentName'));
    await openPane();
    await expect.poll(() => page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('comment-new');
    expect(await pane().innerText()).toContain('No comments on this slide.');
    await page.keyboard.type('Is this the final number?');
    expect(await page.getByTestId('comment-add').isDisabled()).toBe(true);       // no name yet
    await page.getByTestId('comment-name').fill('Asha Rao');
    await page.getByTestId('comment-add').click();
    expect(await shown()).toEqual(['Asha Rao: Is this the final number?']);
    expect(await page.getByTestId('comment-new').inputValue()).toBe('');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('New Comment');
    // the name is kept: the next comment needs only its text, and Ctrl+Enter sends it
    await page.getByTestId('comment-new').fill('And the source?');
    await page.getByTestId('comment-new').press('Control+Enter');
    expect(await shown()).toEqual(['Asha Rao: Is this the final number?', 'Asha Rao: And the source?']);
    expect(await page.evaluate(() => localStorage.getItem('presentations.commentName'))).toBe('Asha Rao');
    // typing a comment never types on the slide or runs a slide shortcut
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.length;')).toBe(2);
    // the slide's thumbnail says how many
    expect(await page.getByTestId('thumb-comments-1').innerText()).toBe('2');
    expect(errors).toEqual([]);
  }, 60_000);

  test('comments belong to their slide; they are changed and deleted in the pane, and undo covers each', async () => {
    await page.evaluate(() => localStorage.setItem('presentations.commentName', 'Ravi'));
    await openPane();
    await add('first');
    await add('second');
    await page.getByRole('tab', { name: 'Home' }).click();
    await page.getByTestId('new-slide').click();
    expect(await shown()).toEqual([]);                                   // the new slide has none
    await add('on slide two');
    await page.getByTestId('thumb-1').click();
    expect(await shown()).toEqual(['Ravi: first', 'Ravi: second']);

    await pane().getByTestId('comment').nth(1).getByTestId('comment-edit-btn').click();
    await page.getByTestId('comment-edit').fill('second, reworded');
    await page.getByTestId('comment-save').click();
    expect(await shown()).toEqual(['Ravi: first', 'Ravi: second, reworded']);
    // Escape while editing gives the change up
    await pane().getByTestId('comment').nth(0).getByTestId('comment-edit-btn').click();
    await page.getByTestId('comment-edit').fill('thrown away');
    await page.getByTestId('comment-edit').press('Escape');
    expect(await shown()).toEqual(['Ravi: first', 'Ravi: second, reworded']);

    await pane().getByTestId('comment').nth(0).getByTestId('comment-delete').click();
    expect(await shown()).toEqual(['Ravi: second, reworded']);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Delete Comment');
    await page.evaluate(() => (globalThis as unknown as { __editor: { getState(): { undo(): void } } }).__editor.getState().undo());
    expect(await shown()).toEqual(['Ravi: first', 'Ravi: second, reworded']);
    await page.evaluate(() => (globalThis as unknown as { __editor: { getState(): { undo(): void } } }).__editor.getState().undo());
    expect(await shown()).toEqual(['Ravi: first', 'Ravi: second']);

    await pane().getByRole('button', { name: 'Close' }).click();
    expect(await pane().count()).toBe(0);
    await page.getByRole('tab', { name: 'View' }).click();
    await page.getByTestId('view-comments').click();
    expect(await pane().count()).toBe(1);
    expect(errors).toEqual([]);
  }, 60_000);

  test('comments are in the saved file, where PowerPoint reads them: who wrote each, when, and what it says', async () => {
    await page.evaluate(() => localStorage.setItem('presentations.commentName', 'Asha Rao'));
    await openPane();
    await add('Please check: 4 < 5 & "quoted"');
    const ed = await saved();
    const part = ed.doc.slideParts()[0]!;
    const comments = slideComments(ed, part);
    expect(comments.length).toBe(1);
    expect(comments[0]).toMatchObject({ author: 'Asha Rao', initials: 'AR', index: 1, text: 'Please check: 4 < 5 & "quoted"' });
    expect(Math.abs(new Date(`${comments[0]!.date}Z`).getTime() - Date.now())).toBeLessThan(120_000);
    expect(ed.pkg.getPart('/ppt/commentAuthors.xml')).toBeDefined();
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('bullet and number styles', () => {
  test('pick a bullet character and a numbering scheme; the menus tick the one in force; text in a box gets a hanging indent', async () => {
    await textBox('one\ntwo');
    await page.keyboard.press('Control+a');
    expect(await ticked('bullet-styles')).toEqual(['None']);

    await choose('bullet-styles', '■ Filled Square');
    let xml = await shapeXml();
    expect(xml.match(/<a:pPr marL="285750" indent="-285750"><a:buFont typeface="Arial"><\/buFont><a:buChar char="■"><\/buChar><\/pPr>/g)?.length).toBe(2);
    expect(await ticked('bullet-styles')).toEqual(['■ Filled Square']);
    expect(await ticked('number-styles')).toEqual([]);
    expect(await pressed('bullets')).toBe('true');

    await choose('number-styles', 'a) b) c)');
    xml = await shapeXml();
    expect(xml.match(/<a:buAutoNum type="alphaLcParenR">/g)?.length).toBe(2);
    expect(xml).not.toContain('buChar');
    expect(xml).toContain('marL="342900" indent="-342900"');            // numbers get the wider indent
    expect(await ticked('number-styles')).toEqual(['a) b) c)']);
    expect([await pressed('bullets'), await pressed('numbering')]).toEqual(['false', 'true']);

    await choose('number-styles', 'I. II. III.');
    expect(await shapeXml()).toContain('<a:buAutoNum type="romanUcPeriod">');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Numbering');

    // None takes the numbers and the indent away; the plain buttons still toggle the default styles
    await choose('number-styles', 'None');
    xml = await shapeXml();
    expect(xml).toContain('<a:buNone>');
    expect(xml).toContain('marL="0" indent="0"');
    expect(await ticked('bullet-styles')).toEqual(['None']);
    await page.getByTestId('bullets').click();
    expect(await ticked('bullet-styles')).toEqual(['• Filled Round']);
    await page.getByTestId('bullets').click();
    expect(await ticked('bullet-styles')).toEqual(['None']);

    await choose('bullet-styles', '– Dash');
    const ed = await saved();
    const box = ed.doc.slide(ed.doc.slideParts()[0]!).shapes.find(sh => sh.name.startsWith('TextBox'))!;
    expect(new TextDecoder().decode(ed.pkg.getPart(ed.doc.slideParts()[0]!)!.bytes)).toContain('<a:buChar char="–"/>');
    expect(box.txBody).toBeDefined();
    expect(errors).toEqual([]);
  }, 60_000);

  test('only the paragraph of the caret changes when nothing is selected', async () => {
    await textBox('one\ntwo');
    await choose('bullet-styles', '○ Hollow Round');
    const xml = await shapeXml();
    expect(xml.match(/<a:buChar/g)?.length).toBe(1);
    expect(xml.indexOf('<a:buChar')).toBeGreaterThan(xml.indexOf('one'));
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('speaker notes', () => {
  test('type notes under the slide, undo and redo them, keep them per slide, and find them in the saved file', async () => {
    const notes = page.getByTestId('notes');
    await expect.poll(() => notes.inputValue()).toBe('');
    await notes.click();
    await page.keyboard.type('Welcome everyone');
    await page.keyboard.press('Enter');
    await page.keyboard.type('Start with the agenda');
    expect(await notes.inputValue()).toBe('Welcome everyone\nStart with the agenda');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Notes');

    // a second slide has its own, empty notes; going back shows the first slide's again
    const first = await state<string>('return s.slide;');
    await page.keyboard.press('Escape');
    await page.getByTestId('new-slide').click();
    await expect.poll(() => notes.inputValue()).toBe('');
    await notes.click();
    await page.keyboard.type('Second slide notes');
    await page.evaluate(p => (globalThis as unknown as { __editor: { getState(): { setSlide(p: string): void } } }).__editor.getState().setSlide(p), first);
    await expect.poll(() => notes.inputValue()).toBe('Welcome everyone\nStart with the agenda');

    // undo in the notes box is the editor's undo: the typing on this slide goes, then comes back
    await page.evaluate(() => (globalThis as unknown as { __editor: { getState(): { undo(): void } } }).__editor.getState().undo());
    await expect.poll(() => notes.inputValue()).toBe('Welcome everyone\nStart with the agenda');   // that undid slide 2's typing
    await notes.click();
    await page.keyboard.type('!');
    await page.keyboard.press('Control+z');
    await expect.poll(() => notes.inputValue()).toBe('Welcome everyone\nStart with the agenda');
    await page.keyboard.press('Control+y');
    await expect.poll(() => notes.inputValue()).toBe('Welcome everyone\nStart with the agenda!');

    const ed = await saved();
    const slides = ed.doc.slideParts();
    expect(notesText(ed, slides[0]!)).toBe('Welcome everyone\nStart with the agenda!');
    expect(ed.doc.notesMaster()).toBeDefined();
    expect(errors).toEqual([]);
  }, 60_000);

  test('typing in the notes never types on the slide, and slide shortcuts stay out of the notes box', async () => {
    const notes = page.getByTestId('notes');
    await notes.click();
    await page.keyboard.type('abc');
    await page.keyboard.press('Backspace');                       // would delete a selected shape if it reached the slide
    await page.keyboard.press('Control+a');
    expect(await notes.inputValue()).toBe('ab');
    expect(await state<number>('return s.selection.length;')).toBe(0);
    expect(await state<boolean>('return !!s.text;')).toBe(false);
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.length;')).toBe(2);   // the title slide's two placeholders
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('find and replace', () => {
  /** The text of every text box (not placeholder) of each slide, in order. */
  const boxTexts = () => state<string[][]>(`
    const textOf = el => { const out = []; const walk = e => { if (e.local === 't') out.push(e.children.map(c => c.value ?? '').join('')); for (const c of e.children) if (c.kind === 'element') walk(c); }; walk(el); return out.join(''); };
    return s.session.ed.doc.slideParts().map(p => s.session.ed.doc.slide(p).shapes.filter(sh => sh.name.startsWith('TextBox')).map(sh => textOf(sh.el)));`);
  const count = () => page.getByTestId('find-count').innerText();
  const found = () => state<{ slide: number; from: number; to: number } | null>(`
    if (!s.text) return null;
    return { slide: s.session.ed.doc.slideParts().indexOf(s.slide) + 1, from: Math.min(s.text.anchor.offset, s.text.caret.offset), to: Math.max(s.text.anchor.offset, s.text.caret.offset) };`);
  const focused = () => page.evaluate(() => document.activeElement?.getAttribute('data-testid'));

  /** Two slides: "alpha beta alpha" on the first, "Alpha on two" on the second; the first slide on screen. */
  async function twoSlides(): Promise<void> {
    await textBox('alpha beta alpha');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    await page.getByTestId('new-slide').click();
    await textBox('Alpha on two');
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    await state('s.setSlide(s.session.ed.doc.slideParts()[0]);');
  }

  test('Ctrl+F finds through the deck: Enter goes on, round the end to the start; Shift+Enter goes back; the found text is selected', async () => {
    await twoSlides();
    await page.keyboard.press('Control+f');
    await expect.poll(focused).toBe('find-input');
    await page.keyboard.type('alpha');
    expect(await count()).toBe('3 found');
    expect(await found()).toBeNull();                                   // typing the query moves nothing

    await page.keyboard.press('Enter');
    expect(await count()).toBe('1 of 3');
    expect(await found()).toEqual({ slide: 1, from: 0, to: 5 });
    expect(await focused()).toBe('find-input');                         // the panel keeps the keyboard
    await page.keyboard.press('Enter');
    expect([await count(), await found()]).toEqual(['2 of 3', { slide: 1, from: 11, to: 16 }]);
    await page.keyboard.press('Enter');
    expect([await count(), await found()]).toEqual(['3 of 3', { slide: 2, from: 0, to: 5 }]);
    await page.keyboard.press('Enter');
    expect([await count(), await found()]).toEqual(['1 of 3', { slide: 1, from: 0, to: 5 }]);
    await page.keyboard.press('Shift+Enter');
    expect([await count(), await found()]).toEqual(['3 of 3', { slide: 2, from: 0, to: 5 }]);
    await page.getByTestId('find-prev').click();
    expect(await count()).toBe('2 of 3');
    await page.getByTestId('find-next').click();
    expect(await count()).toBe('3 of 3');

    // options narrow the search; nothing found says so and disables the buttons
    await page.getByTestId('find-case').check();
    expect(await count()).toBe('2 found');
    await page.getByTestId('find-case').uncheck();
    await page.getByTestId('find-input').fill('alph');
    expect(await count()).toBe('3 found');
    await page.getByTestId('find-words').check();
    expect(await count()).toBe('No results');
    expect(await page.getByTestId('find-next').isDisabled()).toBe(true);
    expect(await page.getByTestId('replace-all').isDisabled()).toBe(true);

    // Escape closes the panel; the ribbon's button opens and closes it too
    await page.getByTestId('find-input').press('Escape');
    expect(await page.getByTestId('find-bar').count()).toBe(0);
    await page.getByTestId('find').click();
    expect(await page.getByTestId('find-bar').count()).toBe(1);
    await page.getByTestId('find').click();
    expect(await page.getByTestId('find-bar').count()).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('Replace changes the match shown and moves to the next; undo brings the word back', async () => {
    await twoSlides();
    await page.keyboard.press('Control+h');
    await expect.poll(focused).toBe('replace-input');
    await page.getByTestId('find-input').fill('alpha');
    await page.getByTestId('replace-input').fill('omega');
    await page.getByTestId('replace-one').click();                      // nothing shown yet: the first press finds
    expect([await count(), await boxTexts()]).toEqual(['1 of 3', [['alpha beta alpha'], ['Alpha on two']]]);
    await page.getByTestId('replace-one').click();
    expect(await boxTexts()).toEqual([['omega beta alpha'], ['Alpha on two']]);
    expect([await count(), await found()]).toEqual(['1 of 2', { slide: 1, from: 11, to: 16 }]);
    await page.getByTestId('replace-one').click();
    expect(await boxTexts()).toEqual([['omega beta omega'], ['Alpha on two']]);
    expect([await count(), await found()]).toEqual(['1 of 1', { slide: 2, from: 0, to: 5 }]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Replace');

    await page.getByTestId('replace-input').press('Control+z');
    expect(await boxTexts()).toEqual([['omega beta alpha'], ['Alpha on two']]);
    expect(await count()).toMatch(/2/);
    expect(errors).toEqual([]);
  }, 60_000);

  test('Replace All changes every slide in one undo step, says how many, and the saved file has the new text', async () => {
    await twoSlides();
    await page.getByTestId('find').click();
    await page.getByTestId('find-input').fill('ALPHA');
    await page.getByTestId('replace-input').fill('gamma');
    await page.getByTestId('replace-all').click();
    expect(await boxTexts()).toEqual([['gamma beta gamma'], ['gamma on two']]);
    await expect.poll(() => page.locator('.status .message').innerText()).toContain('Replaced 3 occurrences.');
    expect(await count()).toBe('No results');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Replace All');

    await page.getByTestId('find-input').press('Control+z');
    expect(await boxTexts()).toEqual([['alpha beta alpha'], ['Alpha on two']]);
    expect(await count()).toBe('3 found');
    await page.getByTestId('find-input').press('Control+Shift+z');
    expect(await boxTexts()).toEqual([['gamma beta gamma'], ['gamma on two']]);

    await page.getByTestId('find-close').click();
    const ed = await saved();
    const texts = ed.doc.slideParts().map(part => ed.doc.slide(part).shapes.filter(sh => sh.name.startsWith('TextBox')).map(sh => sh.txBody ? sh.txBody.children.length : 0));
    expect(texts.map(t => t.length)).toEqual([1, 1]);
    expect(new TextDecoder().decode(ed.pkg.getPart(ed.doc.slideParts()[0]!)!.bytes)).toContain('gamma beta gamma');
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('the right-click menu', () => {
  /** Where a point of the slide (EMU) is on the page. */
  const onScreen = (x: number, y: number) => page.evaluate(([x, y]) => {
    const st = (globalThis as unknown as { __stage: { getAbsoluteTransform(): { point(p: { x: number; y: number }): { x: number; y: number } }; container(): HTMLElement } }).__stage;
    const p = st.getAbsoluteTransform().point({ x, y });
    const r = st.container().getBoundingClientRect();
    return { x: r.left + p.x, y: r.top + p.y };
  }, [x, y] as const);
  /** Draw a rectangle on the slide by dragging, at a box given in EMU; it ends up selected. */
  async function rect(x: number, y: number, w: number, h: number): Promise<void> {
    await drawnOnScreen();
    await page.getByTestId('shapes').click();
    await page.locator('[data-testid="shape-gallery"] button[data-prst="rect"]').click();
    const a = await onScreen(x, y), b = await onScreen(x + w, y + h);
    await page.mouse.move(a.x, a.y);
    await page.mouse.down();
    await page.mouse.move(b.x, b.y, { steps: 4 });
    await page.mouse.up();
  }
  /** The slide as it is now is on screen: two frames have been drawn since the last change (a person is never faster). */
  const drawnOnScreen = () => page.evaluate(() => new Promise<void>(done => requestAnimationFrame(() => requestAnimationFrame(() => done()))));
  const rightClick = async (x: number, y: number) => { await drawnOnScreen(); const p = await onScreen(x, y); await page.mouse.click(p.x, p.y, { button: 'right' }); };
  const menu = () => page.getByTestId('context-menu');
  const labels = async () => (await menu().getByRole('menuitem').allInnerTexts()).map(t => t.split('\n')[0]!);
  const disabled = async () => (await menu().locator('button:disabled').allInnerTexts()).map(t => t.split('\n')[0]!);
  /** Names of the slide's own drawn shapes (not its placeholders), in z-order, and how many are selected. */
  const drawn = () => state<string[]>(`return s.session.ed.doc.slide(s.slide).shapes.filter(sh => !sh.placeholder).map(sh => sh.name);`);
  const selected = () => state<number>('return s.selection.length;');
  const EMPTY = [600000, 600000] as const;            // a spot of the title slide with nothing on it

  test('on a shape: clipboard, duplicate, delete, order and grouping; on the empty slide: paste, select all and slide commands', async () => {
    await rect(1000000, 5500000, 1500000, 800000);
    expect(await drawn()).toEqual(['Rectangle 3']);

    await rightClick(1700000, 5900000);
    expect(await menu().getAttribute('data-kind')).toBe('shape');
    expect(await labels()).toEqual(['Cut', 'Copy', 'Paste', 'Duplicate', 'Delete', 'Edit Text', 'Bring to Front', 'Bring Forward', 'Send Backward', 'Send to Back', 'Group', 'Ungroup', 'Link…', 'Alt Text…', 'Selection Pane…']);
    expect(await disabled()).toEqual(['Paste', 'Group', 'Ungroup']);     // nothing copied yet; one shape, not a group
    await page.getByTestId('menu-duplicate').click();
    expect(await menu().count()).toBe(0);                                // choosing closes the menu
    expect((await drawn()).length).toBe(2);

    // the empty slide: nothing is selected there, and the menu is the slide's
    await rightClick(...EMPTY);
    expect(await selected()).toBe(0);
    expect(await menu().getAttribute('data-kind')).toBe('slide');
    expect(await labels()).toEqual(['Paste', 'Select All', 'New Comment', 'New Slide', 'Duplicate Slide', 'Delete Slide', 'Hide Slide']);
    expect(await disabled()).toEqual(['Paste', 'Delete Slide']);         // the only slide cannot be deleted
    await page.getByTestId('menu-select-all').click();
    expect(await selected()).toBe(4);                                    // two placeholders and the two rectangles

    // a right click on one of several selected shapes keeps them all selected: Group is for all of them
    await page.keyboard.press('Escape');
    await rightClick(1700000, 5900000);
    await menu().waitFor();
    await page.keyboard.press('Escape');
    expect(await menu().count()).toBe(0);
    expect(await selected()).toBe(1);                                    // Escape closed the menu, not the selection
    await page.keyboard.press('Control+a');
    await rightClick(1700000, 5900000);
    expect(await selected()).toBe(4);
    expect(await disabled()).not.toContain('Group');
    await page.getByTestId('menu-group').click();
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.length;')).toBe(1);
    await rightClick(1700000, 5900000);
    expect(await disabled()).toEqual(['Paste', 'Group']);                 // one group selected: Ungroup is on, Group is not
    await page.getByTestId('menu-ungroup').click();
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.length;')).toBe(4);
    expect(errors).toEqual([]);
  }, 60_000);

  test('copy on one slide and paste from the menu on another; cut and delete take the shape away', async () => {
    await rect(1000000, 5500000, 1500000, 800000);
    await rightClick(1700000, 5900000);
    await page.getByTestId('menu-copy').click();
    await page.getByTestId('new-slide').click();
    await rightClick(...EMPTY);
    expect(await disabled()).not.toContain('Paste');
    await page.getByTestId('menu-paste').click();
    await expect.poll(drawn).toEqual(['Rectangle 3']);
    expect(await selected()).toBe(1);

    await rightClick(1700000, 5900000);
    await page.getByTestId('menu-cut').click();
    expect(await drawn()).toEqual([]);
    await rightClick(...EMPTY);
    await page.getByTestId('menu-paste').click();
    await expect.poll(() => drawn().then(d => d.length)).toBe(1);
    const at = await state<{ x: number; y: number }>(`
      const sp = s.selection[0].children.find(c => c.local === 'spPr'), x = sp.children.find(c => c.local === 'xfrm'), off = x.children.find(c => c.local === 'off');
      return { x: Number(off.attrs.find(a => a.local === 'x').value), y: Number(off.attrs.find(a => a.local === 'y').value) };`);
    await rightClick(at.x + 300000, at.y + 300000);
    await page.getByTestId('menu-delete').click();
    expect(await drawn()).toEqual([]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toMatch(/Delete/);
    expect(errors).toEqual([]);
  }, 60_000);

  test('the keyboard works in the menu, a press elsewhere closes it, and it stays inside the window', async () => {
    await rect(1000000, 5500000, 1500000, 800000);
    await rightClick(1700000, 5900000);
    // focus starts on the first command; Down skips the disabled Paste to Duplicate; Enter runs it
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('menu-cut');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('menu-duplicate');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');                                 // round the top to the last command
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('menu-selection-pane');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');                                 // past Alt Text and Link, and Group and Ungroup, which are off
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('menu-back');
    await page.keyboard.press('Enter');
    expect(await menu().count()).toBe(0);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Send to Back');
    // keys typed while the menu is open do nothing to the slide
    await rightClick(1700000, 5900000);
    await page.keyboard.press('Delete');
    expect(await drawn()).toEqual(['Rectangle 3']);
    // a press elsewhere closes it
    await page.mouse.click(700, 20);
    expect(await menu().count()).toBe(0);

    // opened at the bottom right corner of the slide area, the whole menu is still in the window
    const host = await page.getByTestId('canvas').boundingBox();
    await page.mouse.click(host!.x + host!.width - 3, host!.y + host!.height - 3, { button: 'right' });
    const box = await menu().boundingBox();
    expect(box!.x + box!.width).toBeLessThanOrEqual(1400);
    expect(box!.y + box!.height).toBeLessThanOrEqual(860);
    expect(errors).toEqual([]);
  }, 60_000);

  test('Edit Text puts the caret in the shape; the menu gives the keyboard back to the text afterwards', async () => {
    await rect(1000000, 5500000, 1500000, 800000);
    await rightClick(1700000, 5900000);
    await page.getByTestId('menu-edit-text').click();
    await page.keyboard.type('Hi');
    expect(await shapeXml()).toContain('<a:t>Hi</t>');
    // a right click while editing, then Escape: typing goes on in the box
    await rightClick(1700000, 5900000);
    await menu().waitFor();
    await page.keyboard.press('Escape');
    await page.keyboard.type('!');
    expect(await shapeXml()).toContain('<a:t>Hi!</t>');
    // two right clicks in a row are not a double click: no word is selected, so typing adds to the text
    await rightClick(1700000, 5900000);
    await page.keyboard.press('Escape');
    await rightClick(1700000, 5900000);
    await page.keyboard.press('Escape');
    await page.keyboard.type('?');
    expect(await shapeXml()).toContain('<a:t>Hi!?</t>');
    expect(errors).toEqual([]);
  }, 60_000);

  test('on a slide thumbnail: new, duplicate, hide, move and delete, on the slide that was clicked', async () => {
    const thumbMenu = async (n: number) => { await page.getByTestId(`thumb-${n}`).click({ button: 'right' }); await menu().waitFor(); };
    const order = () => state<string[]>('return s.session.ed.doc.slideParts().map(p => p.replace(/^.*slide(\\d+)\\.xml$/, "$1"));');
    const current = () => state<number>('return s.session.ed.doc.slideParts().indexOf(s.slide) + 1;');

    await thumbMenu(1);
    expect(await menu().getAttribute('data-kind')).toBe('thumb');
    expect(await labels()).toEqual(['Cut Slide', 'Copy Slide', 'Paste Slide', 'New Slide', 'Duplicate Slide', 'Delete Slide', 'Hide Slide', 'Move Slide Up', 'Move Slide Down']);
    expect(await disabled()).toEqual(['Cut Slide', 'Paste Slide', 'Delete Slide', 'Move Slide Up', 'Move Slide Down']);
    await page.getByTestId('menu-new-slide').click();
    await thumbMenu(2);
    await page.getByTestId('menu-duplicate-slide').click();
    expect(await order()).toEqual(['1', '2', '3']);

    // a right click on a slide that is not the current one makes it current, and the menu acts on it
    expect(await current()).toBe(3);
    await thumbMenu(1);
    expect(await current()).toBe(1);
    expect(await disabled()).toEqual(['Paste Slide', 'Move Slide Up']);
    await page.getByTestId('menu-slide-down').click();
    expect(await order()).toEqual(['2', '1', '3']);
    expect(await current()).toBe(2);
    await thumbMenu(2);
    await page.getByTestId('menu-hide-slide').click();
    expect(await state<boolean>('return s.session.ed.doc.slide(s.slide).hidden;')).toBe(true);
    await thumbMenu(2);
    expect(await labels()).toContain('Show Slide');
    await page.getByTestId('menu-delete-slide').click();
    expect(await order()).toEqual(['2', '3']);
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('the selection pane', () => {
  const rows = () => page.getByTestId('pane-row');
  const names = async () => rows().evaluateAll(els => els.map(e => (e as HTMLElement).dataset.name));
  const row = (name: string) => page.locator(`[data-testid="pane-row"][data-name="${name}"]`);
  const openPane = async () => {
    await page.getByRole('tab', { name: 'Home' }).click();
    await page.getByTestId('arrange').click();
    await page.getByRole('menuitem', { name: 'Selection Pane…' }).click();
    await page.getByTestId('selection-pane').waitFor();
  };
  /** A title slide with a rectangle and, above it, an ellipse, both under the placeholders. */
  async function twoShapes(): Promise<void> {
    await drawShape('rect', 1000000, 5500000, 1500000, 800000);
    await drawShape('ellipse', 4000000, 5500000, 1500000, 800000);
  }

  test('lists the shapes topmost first; a click selects, Shift+click adds, and the arrows change the order', async () => {
    await twoShapes();
    await openPane();
    expect(await names()).toEqual(['Oval 4', 'Rectangle 3', 'Subtitle 2', 'Title 1']);
    expect(await rows().evaluateAll(els => els.filter(e => e.classList.contains('selected')).map(e => (e as HTMLElement).dataset.name))).toEqual(['Oval 4']);

    await row('Rectangle 3').click();
    expect(await selectedNames()).toEqual(['Rectangle 3']);
    await row('Oval 4').click({ modifiers: ['Shift'] });
    expect(await selectedNames()).toEqual(['Rectangle 3', 'Oval 4']);
    await row('Oval 4').click({ modifiers: ['Shift'] });
    expect(await selectedNames()).toEqual(['Rectangle 3']);

    await page.getByTestId('pane-forward').click();
    expect(await names()).toEqual(['Rectangle 3', 'Oval 4', 'Subtitle 2', 'Title 1']);
    await page.getByTestId('pane-backward').click();
    await page.getByTestId('pane-backward').click();
    expect(await names()).toEqual(['Oval 4', 'Subtitle 2', 'Rectangle 3', 'Title 1']);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Send Backward');

    // a shape picked on the slide is the one marked in the pane
    await page.keyboard.press('Escape');
    const p = await slidePoint(4700000, 5900000);
    await page.mouse.click(p.x, p.y);
    expect(await rows().evaluateAll(els => els.filter(e => e.classList.contains('selected')).map(e => (e as HTMLElement).dataset.name))).toEqual(['Oval 4']);

    await page.getByTestId('selection-pane').getByRole('button', { name: 'Close' }).click();
    expect(await page.getByTestId('selection-pane').count()).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a double click renames: Enter keeps the new name, Escape and an empty name keep the old one', async () => {
    await twoShapes();
    await openPane();
    await row('Rectangle 3').dblclick();
    await page.getByTestId('pane-rename').fill('Logo');
    await page.keyboard.press('Enter');
    expect(await names()).toEqual(['Oval 4', 'Logo', 'Subtitle 2', 'Title 1']);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Rename');
    expect(await drawnNames()).toContain('Logo');                       // typing a name never typed into the shape

    await row('Logo').dblclick();
    await page.getByTestId('pane-rename').fill('Thrown away');
    await page.keyboard.press('Escape');
    expect(await names()).toContain('Logo');
    await row('Logo').dblclick();
    await page.getByTestId('pane-rename').fill('   ');
    await page.keyboard.press('Enter');
    expect(await names()).toContain('Logo');
    // Delete in the name box edits the name; it does not delete the shape
    await row('Logo').dblclick();
    await page.getByTestId('pane-rename').press('End');
    await page.getByTestId('pane-rename').press('Backspace');
    await page.keyboard.press('Enter');
    expect(await names()).toEqual(['Oval 4', 'Log', 'Subtitle 2', 'Title 1']);

    await page.keyboard.press('Control+z');
    expect(await names()).toContain('Logo');
    const ed = await saved();
    expect(ed.doc.slide(ed.doc.slideParts()[0]!).shapes.map(sh => sh.name)).toContain('Logo');
    expect(errors).toEqual([]);
  }, 60_000);

  test('the dot hides a shape and shows it again; hidden shapes stay in the list and can still be selected', async () => {
    await twoShapes();
    await openPane();
    await row('Rectangle 3').getByTestId('pane-eye').click();
    expect(await drawnNames()).not.toContain('Rectangle 3');
    expect(await names()).toContain('Rectangle 3');
    expect(await row('Rectangle 3').getAttribute('class')).toContain('hidden-shape');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Hide');
    // a click where it was now finds nothing there
    await frames();
    const p = await slidePoint(1700000, 5900000);
    await page.mouse.click(p.x, p.y);
    expect(await selectedNames()).toEqual([]);
    // it can be picked in the pane, and shown from there
    await row('Rectangle 3').click();
    expect(await selectedNames()).toEqual(['Rectangle 3']);
    await row('Rectangle 3').getByTestId('pane-eye').click();
    expect(await drawnNames()).toContain('Rectangle 3');

    await page.getByTestId('pane-hide-all').click();
    expect(await drawnNames()).toEqual([]);
    expect(await page.getByTestId('pane-hide-all').isDisabled()).toBe(true);
    await page.getByTestId('pane-show-all').click();
    expect((await drawnNames()).sort()).toEqual(['Oval 4', 'Rectangle 3', 'Subtitle 2', 'Title 1']);
    expect(await page.getByTestId('pane-show-all').isDisabled()).toBe(true);

    await row('Oval 4').getByTestId('pane-eye').click();
    const ed = await saved();
    const shapes = ed.doc.slide(ed.doc.slideParts()[0]!).shapes;
    expect(shapes.filter(sh => sh.hidden).map(sh => sh.name)).toEqual(['Oval 4']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('alt text: typed for the selected shape, kept when the box is left, and opened from the right-click menu', async () => {
    await twoShapes();
    const p = await slidePoint(1700000, 5900000);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByTestId('menu-alt-text').click();
    await expect.poll(() => page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('alt-text');
    await page.keyboard.type('A blue rectangle');
    await page.keyboard.press('Backspace');                             // edits the text, not the shape
    await page.keyboard.type('e');
    expect(await selectedNames()).toEqual(['Rectangle 3']);
    await page.getByTestId('selection-pane').getByText('Selection', { exact: true }).click();    // leave the box
    expect(await state<string>(`return s.selection[0].children[0].children[0].attrs.find(a => a.local === 'descr').value;`)).toBe('A blue rectangle');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Alt Text');

    // another shape shows its own (empty) alt text; coming back shows what was typed
    await row('Oval 4').click();
    expect(await page.getByTestId('alt-text').inputValue()).toBe('');
    await row('Rectangle 3').click();
    expect(await page.getByTestId('alt-text').inputValue()).toBe('A blue rectangle');
    // Escape gives up a change
    await page.getByTestId('alt-text').fill('Something else');
    await page.getByTestId('alt-text').press('Escape');
    expect(await page.getByTestId('alt-text').inputValue()).toBe('A blue rectangle');
    // with two shapes selected there is no alt text box
    await row('Oval 4').click({ modifiers: ['Shift'] });
    expect(await page.getByTestId('alt-text').count()).toBe(0);

    const ed = await saved();
    expect(ed.doc.slide(ed.doc.slideParts()[0]!).shapes.find(sh => sh.name === 'Rectangle 3')!.descr).toBe('A blue rectangle');
    expect(errors).toEqual([]);
  }, 60_000);
});

describe('tables', () => {
  /** Insert a table of a size from the Insert tab's grid; it ends up selected, with the Home tab showing. */
  async function insertTable(cols: number, rows: number): Promise<void> {
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.getByTestId('insert-table').click();
    await page.locator(`[data-testid="table-grid"] [data-size="${cols}x${rows}"]`).click();
    await page.getByRole('tab', { name: 'Home' }).click();
    await frames();
  }
  /** Rows, columns and frame of the selected table. */
  const table = () => state<{ rows: number; cols: number; x: number; y: number; cx: number; cy: number; name: string }>(`
    const el = s.selection[0];
    const find = (e, n) => e.children.find(c => c.kind === 'element' && c.local === n);
    const tbl = find(find(find(el, 'graphic'), 'graphicData'), 'tbl');
    const x = find(el, 'xfrm'), off = find(x, 'off'), ext = find(x, 'ext');
    const num = (e, n) => Number(e.attrs.find(a => a.local === n).value);
    return { rows: tbl.children.filter(c => c.local === 'tr').length, cols: find(tbl, 'tblGrid').children.filter(c => c.local === 'gridCol').length,
      x: num(off, 'x'), y: num(off, 'y'), cx: num(ext, 'cx'), cy: num(ext, 'cy'), name: find(find(el, 'nvGraphicFramePr'), 'cNvPr').attrs.find(a => a.local === 'name').value };`);

  /** The text of every cell, row by row; a cell covered by a merge reads "<" (from the left) or "^" (from above). */
  const cellTexts = () => state<string[][]>(`
    const find = (e, n) => e.children.find(c => c.kind === 'element' && c.local === n);
    const el = s.session.ed.doc.slide(s.slide).shapes.find(sh => sh.kind === 'graphicFrame').el;
    const tbl = find(find(find(el, 'graphic'), 'graphicData'), 'tbl');
    const textOf = e => { const out = []; const ps = find(e, 'txBody').children.filter(c => c.local === 'p'); for (const p of ps) { let t = ''; const walk = x => { if (x.local === 't') t += x.children.map(c => c.value ?? '').join(''); for (const c of x.children) if (c.kind === 'element') walk(c); }; walk(p); out.push(t); } return out.join('/'); };
    const flag = (tc, n) => tc.attrs.some(a => a.local === n && a.value === '1');
    return tbl.children.filter(c => c.local === 'tr').map(tr => tr.children.filter(c => c.local === 'tc').map(tc => flag(tc, 'hMerge') ? '<' : flag(tc, 'vMerge') ? '^' : textOf(tc)));`);
  /** The middle of a cell on the page (rows at their file heights, which one line of text does not grow). */
  const cellPoint = async (r: number, c: number) => {
    const at = await state<{ x: number; y: number }>(`
      const find = (e, n) => e.children.find(c => c.kind === 'element' && c.local === n);
      const el = s.session.ed.doc.slide(s.slide).shapes.find(sh => sh.kind === 'graphicFrame').el;
      const tbl = find(find(find(el, 'graphic'), 'graphicData'), 'tbl');
      const num = (e, n) => Number(e.attrs.find(a => a.local === n).value);
      const off = find(find(el, 'xfrm'), 'off');
      const ws = find(tbl, 'tblGrid').children.filter(c => c.local === 'gridCol').map(c => num(c, 'w'));
      const hs = tbl.children.filter(c => c.local === 'tr').map(t => num(t, 'h'));
      const sum = (a, n) => a.slice(0, n).reduce((x, y) => x + y, 0);
      return { x: num(off, 'x') + sum(ws, ${c}) + ws[${c}] / 2, y: num(off, 'y') + sum(hs, ${r}) + hs[${r}] / 2 };`);
    return slidePoint(at.x, at.y);
  };
  const clickCell = async (r: number, c: number, options: { modifiers?: ('Shift')[]; button?: 'right'; double?: boolean } = {}) => {
    await frames();
    const p = await cellPoint(r, c);
    for (const m of options.modifiers ?? []) await page.keyboard.down(m);
    if (options.double) await page.mouse.dblclick(p.x, p.y);
    else await page.mouse.click(p.x, p.y, options.button ? { button: options.button } : {});
    for (const m of options.modifiers ?? []) await page.keyboard.up(m);
  };
  /** The cells picked: "r,c" for one cell, "r,c:r,c" for a block from the anchor to the far corner. */
  const picked = () => state<string | null>(`
    const c = s.cells; if (!c) return null;
    const a = c.anchor.r + ',' + c.anchor.c, f = c.focus.r + ',' + c.focus.c;
    return (a === f ? a : a + ':' + f) + (c.editing ? ' editing' : '');`);
  const tableMenu = async (item: string) => {
    await page.getByTestId('table-menu').click();
    await page.locator('.menu-panel').getByRole('menuitem', { name: item, exact: true }).click();
  };
  const tableMenuDisabled = async () => {
    await page.getByTestId('table-menu').click();
    const off = await page.locator('.menu-panel button:disabled').allInnerTexts();
    await page.getByTestId('table-menu').click();
    return off;
  };

  test('typing in cells: a click picks a cell, typing fills it, Tab moves on and adds a row at the end, Escape and a click elsewhere keep the text', async () => {
    await insertTable(2, 2);
    await clickCell(0, 0);
    expect(await picked()).toBe('0,0');
    await page.keyboard.type('Name');
    expect(await picked()).toBe('0,0 editing');
    expect(await page.getByTestId('cell-editor').inputValue()).toBe('Name');
    await page.keyboard.press('Tab');
    expect(await cellTexts()).toEqual([['Name', ''], ['', '']]);
    expect(await picked()).toBe('0,1 editing');
    await page.keyboard.type('Score');
    await page.keyboard.press('Tab');
    await page.keyboard.type('Asha');
    await page.keyboard.press('Tab');
    await page.keyboard.type('91');
    await page.keyboard.press('Tab');                                   // from the last cell: a new row
    expect(await cellTexts()).toEqual([['Name', 'Score'], ['Asha', '91'], ['', '']]);
    expect(await picked()).toBe('2,0 editing');
    await page.keyboard.type('Ravi');
    await page.keyboard.press('Shift+Tab');                             // back to the cell before
    expect(await picked()).toBe('1,1 editing');
    expect(await page.getByTestId('cell-editor').inputValue()).toBe('91');
    await page.keyboard.press('Escape');
    expect(await picked()).toBe('1,1');
    expect(await page.getByTestId('cell-editor').count()).toBe(0);
    expect(await cellTexts()).toEqual([['Name', 'Score'], ['Asha', '91'], ['Ravi', '']]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Typing');

    // a double click opens the cell with its text; Enter is a new line; a click on another cell keeps what was typed
    await clickCell(2, 1, { double: true });
    expect(await picked()).toBe('2,1 editing');
    await page.keyboard.type('78');
    await page.keyboard.press('Enter');
    await page.keyboard.type('(retake)');
    await clickCell(0, 0);
    expect(await cellTexts()).toEqual([['Name', 'Score'], ['Asha', '91'], ['Ravi', '78/(retake)']]);
    expect(await picked()).toBe('0,0');

    // Enter opens the picked cell with the text it has; leaving it unchanged adds no undo step
    const undo = await state<string>('return s.session.ed.log.undoLabel;');
    await page.keyboard.press('Enter');
    expect(await page.getByTestId('cell-editor').inputValue()).toBe('Name');
    await page.keyboard.press('Escape');
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe(undo);
    await page.keyboard.press('Control+z');
    expect(await cellTexts()).toEqual([['Name', 'Score'], ['Asha', '91'], ['Ravi', '']]);

    // typing is never lost: a click on another shape, or on the empty slide, while a cell is being typed in keeps it
    await clickCell(2, 1);
    await page.keyboard.type('55');
    const away = await slidePoint(600000, 600000);
    await page.mouse.click(away.x, away.y);
    expect(await cellTexts()).toEqual([['Name', 'Score'], ['Asha', '91'], ['Ravi', '55']]);
    expect(await state<number>('return s.selection.length;')).toBe(0);
    expect(await page.getByTestId('cell-editor').count()).toBe(0);

    const ed = await saved();
    const tbl = tableOf(ed.doc.slide(ed.doc.slideParts()[0]!).shapes.find(sh => sh.kind === 'graphicFrame')!)!;
    expect([cellText(tbl, 0, 0), cellText(tbl, 1, 1), cellText(tbl, 2, 0)]).toEqual(['Name', '91', 'Ravi']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('the arrow keys move between cells, Shift picks a block, Delete empties cells; without cells picked Delete removes the table', async () => {
    await insertTable(3, 2);
    for (const [r, c, t] of [[0, 0, 'a'], [0, 1, 'b'], [0, 2, 'c'], [1, 0, 'd'], [1, 1, 'e'], [1, 2, 'f']] as const) {
      await clickCell(r, c);
      await page.keyboard.type(t);
      await page.keyboard.press('Escape');
    }
    await clickCell(0, 0);
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowDown');
    expect(await picked()).toBe('1,1');
    await page.keyboard.press('ArrowDown');                              // no row below: it stays
    expect(await picked()).toBe('1,1');
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowUp');
    expect(await picked()).toBe('1,1:0,2');
    const x0 = await state<number>(`const el = s.selection[0]; return Number(el.children.find(c => c.local === 'xfrm').children[0].attrs[0].value);`);
    await page.keyboard.press('Delete');
    expect(await cellTexts()).toEqual([['a', '', ''], ['d', '', '']]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Clear Cells');
    await page.keyboard.press('Control+z');
    expect(await cellTexts()).toEqual([['a', 'b', 'c'], ['d', 'e', 'f']]);
    // the arrows moved cells, not the table
    expect(await state<number>(`const el = s.selection[0]; return Number(el.children.find(c => c.local === 'xfrm').children[0].attrs[0].value);`)).toBe(x0);

    // Shift+click picks from the cell picked to the cell clicked
    await clickCell(0, 0);
    await clickCell(1, 1, { modifiers: ['Shift'] });
    expect(await picked()).toBe('0,0:1,1');

    // Escape lets go of the cells; the table is still selected, and Delete now removes it
    await page.keyboard.press('Escape');
    expect(await picked()).toBeNull();
    expect(await state<number>('return s.selection.length;')).toBe(1);
    await page.keyboard.press('Delete');
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.filter(sh => sh.kind === "graphicFrame").length;')).toBe(0);
    expect(errors).toEqual([]);
  }, 60_000);

  test('Rows & Columns: insert beside the picked cell, delete, merge a block and split it; the menu offers only what can be done', async () => {
    await insertTable(2, 2);
    for (const [r, c, t] of [[0, 0, 'a'], [0, 1, 'b'], [1, 0, 'c'], [1, 1, 'd']] as const) {
      await clickCell(r, c);
      await page.keyboard.type(t);
      await page.keyboard.press('Escape');
    }
    // nothing picked: the table is selected but the cell commands are off
    await page.keyboard.press('Escape');
    expect((await tableMenuDisabled()).length).toBe(8);

    await clickCell(0, 0);
    expect(await tableMenuDisabled()).toEqual(['Merge Cells', 'Split Cell']);
    await tableMenu('Insert Row Above');
    expect(await cellTexts()).toEqual([['', ''], ['a', 'b'], ['c', 'd']]);
    expect(await picked()).toBe('1,0');                                  // still on "a"
    await tableMenu('Insert Row Below');
    expect(await cellTexts()).toEqual([['', ''], ['a', 'b'], ['', ''], ['c', 'd']]);
    await tableMenu('Insert Column Left');
    expect(await cellTexts()).toEqual([['', '', ''], ['', 'a', 'b'], ['', '', ''], ['', 'c', 'd']]);
    expect(await picked()).toBe('1,1');
    await tableMenu('Insert Column Right');
    expect(await cellTexts()).toEqual([['', '', '', ''], ['', 'a', '', 'b'], ['', '', '', ''], ['', 'c', '', 'd']]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Insert Column');

    await tableMenu('Delete Row');
    expect(await cellTexts()).toEqual([['', '', '', ''], ['', '', '', ''], ['', 'c', '', 'd']]);
    await tableMenu('Delete Column');
    expect(await cellTexts()).toEqual([['', '', ''], ['', '', ''], ['', '', 'd']]);
    await page.keyboard.press('Control+z');
    await page.keyboard.press('Control+z');
    expect(await cellTexts()).toEqual([['', '', '', ''], ['', 'a', '', 'b'], ['', '', '', ''], ['', 'c', '', 'd']]);

    // merge "a" with the cell right of it and the two under them, then split it again
    await clickCell(1, 1);
    await clickCell(2, 2, { modifiers: ['Shift'] });
    expect(await tableMenuDisabled()).toEqual(['Split Cell']);
    await tableMenu('Merge Cells');
    expect(await cellTexts()).toEqual([['', '', '', ''], ['', 'a', '<', 'b'], ['', '^', '<', ''], ['', 'c', '', 'd']]);
    expect(await picked()).toBe('1,1');
    expect(await tableMenuDisabled()).toEqual(['Merge Cells']);
    // a click anywhere in the merged cell picks it, and the arrow keys step over all of it
    await clickCell(2, 2);
    await page.keyboard.press('ArrowRight');
    expect(await picked()).toBe('2,3');
    await page.keyboard.press('ArrowLeft');
    await tableMenu('Split Cell');
    expect(await cellTexts()).toEqual([['', '', '', ''], ['', 'a', '', 'b'], ['', '', '', ''], ['', 'c', '', 'd']]);

    // picking every row leaves no row to keep: Delete Row is off; the same for columns
    await clickCell(0, 0);
    await clickCell(3, 0, { modifiers: ['Shift'] });
    expect(await tableMenuDisabled()).toEqual(['Delete Row', 'Split Cell']);
    await clickCell(0, 0);
    for (let i = 0; i < 3; i++) await page.keyboard.press('Shift+ArrowRight');      // the table is now wider than the slide: by keyboard
    expect(await picked()).toBe('0,0:0,3');
    expect(await tableMenuDisabled()).toEqual(['Delete Column', 'Split Cell']);
    expect(errors).toEqual([]);
  }, 60_000);

  test('the right-click menu on a table has the row, column and merge commands, for the cell clicked or the block picked', async () => {
    await insertTable(2, 2);
    await clickCell(1, 1, { button: 'right' });
    expect(await picked()).toBe('1,1');
    const labels = (await page.getByTestId('context-menu').getByRole('menuitem').allInnerTexts()).map(t => t.split('\n')[0]);
    expect(labels).toEqual(expect.arrayContaining(['Insert Row Above', 'Insert Row Below', 'Insert Column Left', 'Insert Column Right', 'Delete Row', 'Delete Column', 'Merge Cells', 'Split Cell', 'Cut', 'Bring to Front']));
    await page.getByTestId('menu-row-below').click();
    expect((await cellTexts()).length).toBe(3);

    // a right click inside a picked block keeps the block, so Merge is for all of it
    await clickCell(0, 0);
    await clickCell(1, 1, { modifiers: ['Shift'] });
    await clickCell(1, 0, { button: 'right' });
    expect(await picked()).toBe('0,0:1,1');
    await page.getByTestId('menu-merge').click();
    expect(await cellTexts()).toEqual([['', '<'], ['^', '<'], ['', '']]);
    // outside the block, the right click picks the cell it is on
    await clickCell(2, 1, { button: 'right' });
    expect(await picked()).toBe('2,1');
    await page.keyboard.press('Escape');
    expect(errors).toEqual([]);
  }, 60_000);

  test('formatting from the ribbon applies to the cells picked, or to the whole table when none is', async () => {
    await insertTable(2, 2);
    for (const [r, c, t] of [[0, 0, 'a'], [0, 1, 'b'], [1, 0, 'c'], [1, 1, 'd']] as const) {
      await clickCell(r, c);
      await page.keyboard.type(t);
      await page.keyboard.press('Escape');
    }
    /** Which cells' text is bold, row by row. */
    const bold = () => state<boolean[][]>(`
      const find = (e, n) => e.children.find(c => c.kind === 'element' && c.local === n);
      const el = s.session.ed.doc.slide(s.slide).shapes.find(sh => sh.kind === 'graphicFrame').el;
      const tbl = find(find(find(el, 'graphic'), 'graphicData'), 'tbl');
      const isBold = tc => { let b = false; const walk = x => { if (x.local === 'rPr' && x.attrs.some(a => a.local === 'b' && a.value === '1')) b = true; for (const c of x.children) if (c.kind === 'element') walk(c); }; walk(tc); return b; };
      return tbl.children.filter(c => c.local === 'tr').map(tr => tr.children.filter(c => c.local === 'tc').map(isBold));`);
    await clickCell(1, 0);
    await page.getByTestId('bold').click();
    expect(await bold()).toEqual([[false, false], [true, false]]);
    expect(await pressed('bold')).toBe('true');                         // the button shows the picked cell
    await clickCell(1, 1);
    expect(await pressed('bold')).toBe('false');
    await clickCell(0, 0);
    await clickCell(0, 1, { modifiers: ['Shift'] });
    await page.getByTestId('center').click();
    expect(await state<number>(`
      const el = s.selection[0]; let n = 0; const walk = x => { if (x.local === 'pPr' && x.attrs.some(a => a.local === 'algn' && a.value === 'ctr')) n++; for (const c of x.children) if (c.kind === 'element') walk(c); }; walk(el); return n;`)).toBe(2);
    // no cells picked: the whole table. Its first cell is in the header row, which the table style makes bold, so
    // the first press takes bold off everywhere and the second puts it on everywhere
    await page.keyboard.press('Escape');
    expect(await pressed('bold')).toBe('true');
    await page.getByTestId('bold').click();
    expect(await bold()).toEqual([[false, false], [false, false]]);
    await page.getByTestId('bold').click();
    expect(await bold()).toEqual([[true, true], [true, true]]);
    expect(errors).toEqual([]);
  }, 60_000);

  test('find and replace reach into cells: Find picks the cell that holds the match, Replace All changes cells too', async () => {
    await insertTable(2, 2);
    for (const [r, c, t] of [[0, 0, 'red'], [0, 1, 'blue'], [1, 0, 'green'], [1, 1, 'dark red']] as const) {
      await clickCell(r, c);
      await page.keyboard.type(t);
      await page.keyboard.press('Escape');
    }
    await page.keyboard.press('Escape');
    await page.keyboard.press('Control+f');
    await page.keyboard.type('red');
    expect(await page.getByTestId('find-count').innerText()).toBe('2 found');
    await page.keyboard.press('Enter');
    expect([await page.getByTestId('find-count').innerText(), await picked()]).toEqual(['1 of 2', '0,0']);
    await page.keyboard.press('Enter');
    expect([await page.getByTestId('find-count').innerText(), await picked()]).toEqual(['2 of 2', '1,1']);
    await page.getByTestId('replace-input').fill('pink');
    await page.getByTestId('replace-one').click();
    expect(await cellTexts()).toEqual([['red', 'blue'], ['green', 'dark pink']]);
    await page.getByTestId('replace-all').click();
    expect(await cellTexts()).toEqual([['pink', 'blue'], ['green', 'dark pink']]);
    expect(errors).toEqual([]);
  }, 60_000);

  test('picked cells copy as rows of tab-separated text, and such text pastes into cells from the picked one, as with a spreadsheet', async () => {
    await insertTable(3, 3);
    for (const [r, c, t] of [[0, 0, 'a'], [0, 1, 'b'], [1, 0, 'c'], [1, 1, 'd']] as const) {
      await clickCell(r, c);
      await page.keyboard.type(t);
      await page.keyboard.press('Escape');
    }
    await clickCell(0, 0);
    await clickCell(1, 1, { modifiers: ['Shift'] });
    const copied = await clipboardEvent('copy');
    expect(copied).toEqual({ plain: 'a\tb\nc\td', html: '<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>', handled: true });

    // from a spreadsheet: two rows of three, pasted at the middle cell; what does not fit in the table is left out
    await clickCell(1, 1);
    await clipboardEvent('paste', { 'text/plain': '1\t2\t3\n4\t5\t6\n7\t8\t9\n' });
    expect(await cellTexts()).toEqual([['a', 'b', ''], ['c', '1', '2'], ['', '4', '5']]);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Paste');
    await page.keyboard.press('Control+z');                              // one step
    expect(await cellTexts()).toEqual([['a', 'b', ''], ['c', 'd', ''], ['', '', '']]);

    // cut empties the picked cells and keeps the table
    await clickCell(0, 0);
    await clickCell(0, 1, { modifiers: ['Shift'] });
    expect((await clipboardEvent('cut')).plain).toBe('a\tb');
    expect(await cellTexts()).toEqual([['', '', ''], ['c', 'd', ''], ['', '', '']]);
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.filter(sh => sh.kind === "graphicFrame").length;')).toBe(1);
    expect(errors).toEqual([]);
  }, 60_000);

  test('Insert > Table: the grid says the size under the pointer, a click inserts that table in the middle of the slide', async () => {
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.getByTestId('insert-table').click();
    expect(await page.getByTestId('table-grid-size').innerText()).toBe('Insert Table');
    await page.locator('[data-testid="table-grid"] [data-size="4x3"]').hover();
    expect(await page.getByTestId('table-grid-size').innerText()).toBe('4 x 3 Table');
    expect(await page.locator('[data-testid="table-grid"] .on').count()).toBe(12);
    await page.locator('[data-testid="table-grid"] [data-size="4x3"]').click();
    expect(await page.getByTestId('table-grid').count()).toBe(0);         // picking closes the menu

    const t = await table();
    expect(t).toMatchObject({ rows: 3, cols: 4, name: 'Table 3', cx: 8128000, cy: 3 * 370840 });
    expect(t.x).toBe((12192000 - 8128000) / 2);                          // centred on a 16:9 slide
    expect(Math.abs(t.y + t.cy / 2 - 6858000 / 2)).toBeLessThan(2);
    expect(await state<string>('return s.session.ed.log.undoLabel;')).toBe('Insert Table');

    // the table is drawn: its header row is filled with the theme's accent colour
    await frames();
    const p = await slidePoint(t.x + 200000, t.y + 100000);
    const px = await page.evaluate(([x, y]) => {
      const c = document.querySelector('[data-testid="canvas"] canvas') as HTMLCanvasElement;
      const r = c.getBoundingClientRect();
      const d = c.getContext('2d')!.getImageData(Math.round((x - r.left) * (c.width / r.width)), Math.round((y - r.top) * (c.height / r.height)), 1, 1).data;
      return [d[0], d[1], d[2]];
    }, [p.x, p.y] as const);
    expect(px).toEqual([0x44, 0x72, 0xc4]);                              // Office accent 1

    await page.keyboard.press('Control+z');
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.length;')).toBe(2);
    await page.keyboard.press('Control+y');
    expect(await state<number>('return s.session.ed.doc.slide(s.slide).shapes.length;')).toBe(3);
    expect(errors).toEqual([]);
  }, 60_000);

  test('a table moves and resizes like any shape, its columns scaling with it, and is in the saved file', async () => {
    await insertTable(2, 2);
    const t = await table();
    await page.keyboard.press('ArrowRight');
    expect((await table()).x).toBe(t.x + 76200);
    // resize through the store's own command path: the frame doubles in width
    await state(`
      const ed = s.session.ed; const v = ed.view(s.slide, s.selection[0]); const x = v.xfrm;
      const find = (e, n) => e.children.find(c => c.kind === 'element' && c.local === n);
      globalThis.__before = find(find(find(find(v.el, 'graphic'), 'graphicData'), 'tbl'), 'tblGrid').children.filter(c => c.local === 'gridCol').map(c => Number(c.attrs[0].value));`);
    const br = await slidePoint(t.x + 76200 + t.cx, t.y + t.cy / 2);
    await page.mouse.move(br.x, br.y);
    await page.mouse.down();
    await page.mouse.move(br.x + 100, br.y, { steps: 5 });
    await page.mouse.up();
    const after = await table();
    expect(after.cx).toBeGreaterThan(t.cx);
    const cols = await state<number[]>(`
      const find = (e, n) => e.children.find(c => c.kind === 'element' && c.local === n);
      return find(find(find(find(s.selection[0], 'graphic'), 'graphicData'), 'tbl'), 'tblGrid').children.filter(c => c.local === 'gridCol').map(c => Number(c.attrs[0].value));`);
    expect(cols.reduce((a, b) => a + b, 0)).toBe(after.cx);              // the columns are what was resized

    const ed = await saved();
    const frame = ed.doc.slide(ed.doc.slideParts()[0]!).shapes.find(sh => sh.kind === 'graphicFrame')!;
    expect(tableSize(tableOf(frame)!)).toEqual({ rows: 2, cols: 2 });
    expect(cellText(tableOf(frame)!, 0, 0)).toBe('');
    expect(ed.doc.tableStylesPart()).toBeDefined();
    expect(errors).toEqual([]);
  }, 60_000);
});
