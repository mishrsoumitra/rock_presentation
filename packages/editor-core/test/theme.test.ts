// Theme colours and fonts: another scheme changes what takes its colour or font from the theme, on every slide,
// and nothing that has a colour or font of its own.
import { OpcPackage } from '@pptx/opc';
import { checkOrder, PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  addSlide, COLOR_SCHEMES, FONT_SCHEMES, formatRuns, insertShape, insertTextBox, setFill, setThemeColors, setThemeFonts, THEME_SLOTS, themeOfSlide,
  themePalette, themeSchemes, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';
const hex = (c: { r: number; g: number; b: number }) => [c.r, c.g, c.b].map(v => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
const GREEN = COLOR_SCHEMES.find(s => s.name === 'Green')!;

/** A shape filled by the theme (accent 1) and one filled red outright, on the first slide; a themed shape on a second slide. */
function deck(): { ed: Editor; theme: string; themed: number; red: number; second: string; other: number } {
  const ed = openEditor(deckBytes());
  const a = insertShape(ed, SLIDE, 'rect', { x: 0, y: 0, w: 1000000, h: 1000000 });
  ed.run(a.command);
  const b = insertShape(ed, SLIDE, 'rect', { x: 2000000, y: 0, w: 1000000, h: 1000000 });
  ed.run(b.command);
  ed.run(setFill(ed, SLIDE, [ed.shape(SLIDE, b.id)!], { kind: 'srgb', hex: 'FF0000' }));
  const s2 = addSlide(ed, LAYOUT);
  ed.run(s2.command);
  const c = insertShape(ed, s2.part, 'ellipse', { x: 0, y: 0, w: 1000000, h: 1000000 });
  ed.run(c.command);
  ed.log.clear();
  return { ed, theme: themeOfSlide(ed, SLIDE)!, themed: a.id, red: b.id, second: s2.part, other: c.id };
}
const fillOf = (ed: Editor, part: string, id: number) => { const f = resolveSlide(ed.doc, part).shapes.find(s => s.id === id && s.layer === 'slide')!.fill; return f.kind === 'solid' ? hex(f.color) : f.kind; };

describe('theme colours', () => {
  test('every scheme offered has all twelve colours, each a six-digit hex, and a name of its own', () => {
    expect(new Set(COLOR_SCHEMES.map(s => s.name)).size).toBe(COLOR_SCHEMES.length);
    for (const s of COLOR_SCHEMES) for (const slot of THEME_SLOTS) expect(s.colors[slot], `${s.name} ${slot}`).toMatch(/^[0-9A-F]{6}$/);
  });

  test('another scheme recolours what is coloured by the theme, on every slide, and leaves colours given outright', () => {
    const { ed, theme, themed, red, second, other } = deck();
    expect(fillOf(ed, SLIDE, themed)).toBe('4472C4');
    expect(cycle(ed, () => setThemeColors(ed, theme, GREEN))).toEqual([]);
    ed.run(setThemeColors(ed, theme, GREEN));
    expect(fillOf(ed, SLIDE, themed)).toBe('549E39');
    expect(fillOf(ed, second, other)).toBe('549E39');
    expect(fillOf(ed, SLIDE, red)).toBe('FF0000');
    expect(themeSchemes(ed, theme).colors).toBe('Green');
    // the colour picker shows the new colours
    expect(themePalette(ed, SLIDE).find(c => c.slot === 'accent1')!.base.hex).toBe('#549E39');
    expect(checkOrder(ed.doc.root(theme))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Theme Colors');
    ed.undo();
    expect(fillOf(ed, SLIDE, themed)).toBe('4472C4');
  });

  test('a saved deck opens with the scheme; choosing the scheme a theme has changes nothing that is drawn', () => {
    const { ed, theme, themed } = deck();
    ed.run(setThemeColors(ed, theme, GREEN));
    const back = openEditor(ed.save());
    expect(fillOf(back, SLIDE, themed)).toBe('549E39');
    expect(new PresentationDocument(OpcPackage.open(ed.save())).slideParts().length).toBe(2);
    ed.run(setThemeColors(ed, theme, COLOR_SCHEMES[0]!));
    expect(fillOf(ed, SLIDE, themed)).toBe('4472C4');
  });
});

describe('theme fonts', () => {
  /** The fonts the resolver gives a text box's two runs: the first left to the theme, the second set to Georgia. */
  function fonts(ed: Editor, id: number): string[] {
    return resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.text!.paragraphs[0]!.runs.map(r => r.props.latin);
  }

  test('another pair changes text that takes its font from the theme, and leaves text with a font of its own', () => {
    const { ed, theme } = deck();
    const box = insertTextBox(ed, SLIDE, { x: 0, y: 3000000, w: 4000000, h: 500000 }, 'themed own');
    ed.run(box.command);
    ed.run(formatRuns(ed, SLIDE, ed.shape(SLIDE, box.id)!.txBody!, { start: { p: 0, offset: 7 }, end: { p: 0, offset: 10 } }, { latin: 'Georgia' }));
    ed.log.clear();
    const before = fonts(ed, box.id);
    expect(before[1]).toBe('Georgia');
    const pair = FONT_SCHEMES.find(f => f.name === 'Times New Roman and Arial')!;
    expect(cycle(ed, () => setThemeFonts(ed, theme, pair))).toEqual([]);
    ed.run(setThemeFonts(ed, theme, pair));
    expect(fonts(ed, box.id)).toEqual(['Arial', 'Georgia']);              // body text takes the minor font
    expect(themeSchemes(ed, theme)).toMatchObject({ fonts: 'Times New Roman and Arial', major: 'Times New Roman', minor: 'Arial' });
    expect(checkOrder(ed.doc.root(theme))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Theme Fonts');
    ed.undo();
    expect(fonts(ed, box.id)).toEqual(before);
  });

  test('every pair offered names two fonts and has a name of its own', () => {
    expect(new Set(FONT_SCHEMES.map(f => f.name)).size).toBe(FONT_SCHEMES.length);
    for (const f of FONT_SCHEMES) expect(f.major.length > 0 && f.minor.length > 0).toBe(true);
  });
});
