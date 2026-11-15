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
