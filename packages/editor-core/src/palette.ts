// packages/editor-core/src/palette.ts  -- the theme colour picker: twelve slots and PowerPoint's tints and shades
import { resolveColor, rgbToHsl, slideContext, toHex } from '@pptx/resolve';
import type { Editor } from './editor.ts';
import type { ColorSpec } from './text.ts';

export interface Swatch { label: string; color: ColorSpec; hex: string; }
export interface PaletteColumn { slot: string; name: string; base: Swatch; variants: Swatch[]; }

/** The columns PowerPoint shows: the colour map's names, so a dark master's "Text 1" is the light colour. */
const SLOTS: [string, string][] = [
  ['bg1', 'Background 1'], ['tx1', 'Text 1'], ['bg2', 'Background 2'], ['tx2', 'Text 2'],
  ['accent1', 'Accent 1'], ['accent2', 'Accent 2'], ['accent3', 'Accent 3'], ['accent4', 'Accent 4'], ['accent5', 'Accent 5'], ['accent6', 'Accent 6'],
];

type Pair = [number, number | undefined, string];

/**
 * The variants under each colour, as lumMod and lumOff pairs. Which set a colour gets depends on its lightness.
 * The pairs are the ones PowerPoint writes (every pair occurs in PowerPoint-saved files of the test corpus), not
 * derived: black gets lighter tints in small steps, white darker shades, very dark and very light colours
 * wider steps, everything else 80/60/40% lighter and 25/50% darker.
 */
const VARIANTS: Record<'black' | 'dark' | 'mid' | 'light' | 'white', Pair[]> = {
  black: [[50000, 50000, 'Lighter 50%'], [65000, 35000, 'Lighter 35%'], [75000, 25000, 'Lighter 25%'], [85000, 15000, 'Lighter 15%'], [95000, 5000, 'Lighter 5%']],
  dark: [[10000, 90000, 'Lighter 90%'], [25000, 75000, 'Lighter 75%'], [50000, 50000, 'Lighter 50%'], [75000, 25000, 'Lighter 25%'], [90000, 10000, 'Lighter 10%']],
  mid: [[20000, 80000, 'Lighter 80%'], [40000, 60000, 'Lighter 60%'], [60000, 40000, 'Lighter 40%'], [75000, undefined, 'Darker 25%'], [50000, undefined, 'Darker 50%']],
  light: [[90000, undefined, 'Darker 10%'], [75000, undefined, 'Darker 25%'], [50000, undefined, 'Darker 50%'], [25000, undefined, 'Darker 75%'], [10000, undefined, 'Darker 90%']],
  white: [[95000, undefined, 'Darker 5%'], [85000, undefined, 'Darker 15%'], [75000, undefined, 'Darker 25%'], [65000, undefined, 'Darker 35%'], [50000, undefined, 'Darker 50%']],
};

export function lightnessClass(l: number): keyof typeof VARIANTS {
  const pct = Math.round(l * 100);
  if (pct <= 0) return 'black';
  if (pct >= 100) return 'white';
  if (pct < 20) return 'dark';
  if (pct > 80) return 'light';
  return 'mid';
}

/** The theme colour picker for a slide: each column a theme colour and its five variants, with their hex values. */
export function themePalette(ed: Editor, part: string): PaletteColumn[] {
  const colors = slideContext(ed.doc, part).colors;
  const hexOf = (val: string, lumMod?: number, lumOff?: number) => toHex(resolveColor({
    base: { kind: 'scheme', name: val },
    transforms: [...(lumMod !== undefined ? [{ name: 'lumMod', val: lumMod / 100000 }] : []), ...(lumOff !== undefined ? [{ name: 'lumOff', val: lumOff / 100000 }] : [])],
  }, colors));
  return SLOTS.map(([slot, name]) => {
    const base = resolveColor({ base: { kind: 'scheme', name: slot }, transforms: [] }, colors);
    const { l } = rgbToHsl(base.r, base.g, base.b);
    return {
      slot, name,
      base: { label: name, color: { kind: 'scheme', val: slot }, hex: toHex(base) },
      variants: VARIANTS[lightnessClass(l)].map(([mod, off, label]) => ({
        label: `${name}, ${label}`,
        color: { kind: 'scheme', val: slot, lumMod: mod, ...(off !== undefined ? { lumOff: off } : {}) },
        hex: hexOf(slot, mod, off),
      })),
    };
  });
}

/** PowerPoint's ten standard colours, under the theme colours. */
export const STANDARD_COLORS: Swatch[] = [
  ['C00000', 'Dark Red'], ['FF0000', 'Red'], ['FFC000', 'Orange'], ['FFFF00', 'Yellow'], ['92D050', 'Light Green'],
  ['00B050', 'Green'], ['00B0F0', 'Light Blue'], ['0070C0', 'Blue'], ['002060', 'Dark Blue'], ['7030A0', 'Purple'],
].map(([hex, label]) => ({ label: label!, color: { kind: 'srgb', hex: hex! }, hex: `#${hex}` }));
