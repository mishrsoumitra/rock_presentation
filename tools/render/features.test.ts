// Features the editor shares with Google Slides and PowerPoint, each driven in headless Chromium the way a person
// would use it, and each checked in the deck's XML and in the saved file. Every test starts from a new presentation.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium, type Browser, type Page } from 'playwright';
import { createServer, type ViteDevServer } from 'vite';
import { afterAll, beforeAll, beforeEach, describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { cellText, Editor, notesText, tableOf, tableSize } from '../../packages/editor-core/src/index.ts';

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
    expect(await labels()).toEqual(['Cut', 'Copy', 'Paste', 'Duplicate', 'Delete', 'Edit Text', 'Bring to Front', 'Bring Forward', 'Send Backward', 'Send to Back', 'Group', 'Ungroup', 'Alt Text…', 'Selection Pane…']);
    expect(await disabled()).toEqual(['Paste', 'Group', 'Ungroup']);     // nothing copied yet; one shape, not a group
    await page.getByTestId('menu-duplicate').click();
    expect(await menu().count()).toBe(0);                                // choosing closes the menu
    expect((await drawn()).length).toBe(2);

    // the empty slide: nothing is selected there, and the menu is the slide's
    await rightClick(...EMPTY);
    expect(await selected()).toBe(0);
    expect(await menu().getAttribute('data-kind')).toBe('slide');
    expect(await labels()).toEqual(['Paste', 'Select All', 'New Slide', 'Duplicate Slide', 'Delete Slide', 'Hide Slide']);
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
    await page.keyboard.press('ArrowUp');                                 // past Alt Text, and Group and Ungroup, which are off
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
    expect(await labels()).toEqual(['New Slide', 'Duplicate Slide', 'Delete Slide', 'Hide Slide', 'Move Slide Up', 'Move Slide Down']);
    expect(await disabled()).toEqual(['Delete Slide', 'Move Slide Up', 'Move Slide Down']);
    await page.getByTestId('menu-new-slide').click();
    await thumbMenu(2);
    await page.getByTestId('menu-duplicate-slide').click();
    expect(await order()).toEqual(['1', '2', '3']);

    // a right click on a slide that is not the current one makes it current, and the menu acts on it
    expect(await current()).toBe(3);
    await thumbMenu(1);
    expect(await current()).toBe(1);
    expect(await disabled()).toEqual(['Move Slide Up']);
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
