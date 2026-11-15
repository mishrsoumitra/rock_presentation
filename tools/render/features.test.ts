// Features the editor shares with Google Slides and PowerPoint, each driven in headless Chromium the way a person
// would use it, and each checked in the deck's XML and in the saved file. Every test starts from a new presentation.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium, type Browser, type Page } from 'playwright';
import { createServer, type ViteDevServer } from 'vite';
import { afterAll, beforeAll, beforeEach, describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { Editor, notesText } from '../../packages/editor-core/src/index.ts';

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
}, 120_000);

afterAll(async () => {
  await browser?.close();
  await server?.close();
});

/** A new presentation in a freshly loaded app: one title slide, nothing selected, an empty undo history. */
beforeEach(async () => {
  errors.length = 0;
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
