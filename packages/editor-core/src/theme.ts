// packages/editor-core/src/theme.ts  -- the theme's colours and fonts: reading them, and choosing others
import { attr, child, path, setAttr } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { element, esc } from './xml.ts';

/** The twelve colours of a theme, in the order the scheme lists them. */
export const THEME_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;
export type ThemeSlot = (typeof THEME_SLOTS)[number];
/** A colour scheme: a name and a six-digit hex colour for each slot. */
export interface ColorScheme { name: string; colors: Record<ThemeSlot, string>; }
/** A font scheme: a name, the font of headings (major) and the font of body text (minor). */
export interface FontScheme { name: string; major: string; minor: string; }

const scheme = (name: string, hex: string): ColorScheme => {
  const v = hex.split(' ');
  return { name, colors: Object.fromEntries(THEME_SLOTS.map((s, i) => [s, v[i]!])) as Record<ThemeSlot, string> };
};

/** Colour schemes to choose from. "Office" is the scheme of the Office theme; the others are this editor's own. */
export const COLOR_SCHEMES: ColorScheme[] = [
  scheme('Office', '000000 FFFFFF 44546A E7E6E6 4472C4 ED7D31 A5A5A5 FFC000 5B9BD5 70AD47 0563C1 954F72'),
  scheme('Grayscale', '000000 FFFFFF 000000 F8F8F8 DDDDDD B2B2B2 969696 808080 5F5F5F 4D4D4D 5F5F5F 919191'),
  scheme('Blue', '000000 FFFFFF 17406D DBEFF9 0F6FC6 009DD9 0BD0D9 10CF9B 7CCA62 A5C249 F49100 85DFD0'),
  scheme('Green', '000000 FFFFFF 455F51 E3DED1 549E39 8AB833 C0CF3A 029676 4AB5C4 0989B1 6B9F25 BA6906'),
  scheme('Orange', '000000 FFFFFF 4E3B30 FBEEC9 F0A22E A5644E B58B80 C3986D A19574 C17529 AD1F1F FFC42F'),
  scheme('Red', '000000 FFFFFF 323232 E5C243 A5300F D55816 E19825 B19C7D 7F5F52 B27D49 6B9F25 B26B02'),
  scheme('Violet', '000000 FFFFFF 373545 DCD8DC AD84C6 8784C7 5D739A 6997AF 84ACB6 6F8183 69A020 8C8C8C'),
];

/** Font pairs to choose from: heading font, body font. */
export const FONT_SCHEMES: FontScheme[] = [
  { name: 'Office', major: 'Calibri Light', minor: 'Calibri' },
  { name: 'Aptos', major: 'Aptos Display', minor: 'Aptos' },
  { name: 'Arial', major: 'Arial', minor: 'Arial' },
  { name: 'Cambria and Calibri', major: 'Cambria', minor: 'Calibri' },
  { name: 'Georgia and Verdana', major: 'Georgia', minor: 'Verdana' },
  { name: 'Times New Roman and Arial', major: 'Times New Roman', minor: 'Arial' },
  { name: 'Courier New', major: 'Courier New', minor: 'Courier New' },
];

/** The theme a slide uses: its master's. */
export function themeOfSlide(ed: Editor, slide: string): string | undefined {
  const layout = ed.doc.layoutOf(slide);
  const master = layout ? ed.doc.masterOf(layout) : ed.doc.masterParts()[0];
  return master ? ed.doc.themeOf(master) : undefined;
}

const elementsOf = (ed: Editor, theme: string) => child(ed.doc.root(theme), 'a:themeElements');

/** The name of a theme's colour scheme and of its font scheme, and its heading and body fonts. */
export function themeSchemes(ed: Editor, theme: string): { colors: string; fonts: string; major: string; minor: string } {
  const el = elementsOf(ed, theme);
  const fonts = child(el, 'a:fontScheme');
  return {
    colors: attr(child(el, 'a:clrScheme'), 'name') ?? '', fonts: attr(fonts, 'name') ?? '',
    major: attr(path(fonts, 'a:majorFont', 'a:latin'), 'typeface') ?? '', minor: attr(path(fonts, 'a:minorFont', 'a:latin'), 'typeface') ?? '',
  };
}

/**
 * Give a theme another colour scheme. Everything on every slide of the theme's master that is coloured by a theme
 * colour (shapes, text, tables, backgrounds, and their tints and shades) follows; colours given outright do not.
 */
export function setThemeColors(ed: Editor, theme: string, to: ColorScheme): Command {
  const root = ed.doc.root(theme);
  return ed.edit({
    label: 'Theme Colors', parts: [theme], scope: () => [root],
    apply: () => {
      const clr = child(elementsOf(ed, theme), 'a:clrScheme');
      if (!clr) throw new Error('this theme has no colour scheme');
      setAttr(clr, 'name', to.name);
      for (const slot of THEME_SLOTS) {
        const el = child(clr, `a:${slot}`);
        if (!el) continue;
        const color = element(el, `<a:srgbClr val="${esc(to.colors[slot].replace(/^#/, '').toUpperCase())}"/>`, ed.strict);
        color.parent = el;
        el.children = [color];
        el.selfClosing = false;
      }
    },
  });
}

/**
 * Give a theme another pair of fonts. Text that takes its font from the theme (headings from the major font, body
 * text from the minor) follows; text with a font of its own does not. Fonts for other scripts are left as they are.
 */
export function setThemeFonts(ed: Editor, theme: string, to: FontScheme): Command {
  const root = ed.doc.root(theme);
  return ed.edit({
    label: 'Theme Fonts', parts: [theme], scope: () => [root],
    apply: () => {
      const fonts = child(elementsOf(ed, theme), 'a:fontScheme');
      if (!fonts) throw new Error('this theme has no font scheme');
      setAttr(fonts, 'name', to.name);
      const set = (collection: XmlElement | undefined, face: string) => {
        const latin = child(collection, 'a:latin');
        if (!latin) return;
        setAttr(latin, 'typeface', face);
        // a font's own description (panose) belongs to the font that was there
        setAttr(latin, 'panose', undefined);
      };
      set(child(fonts, 'a:majorFont'), to.major);
      set(child(fonts, 'a:minorFont'), to.minor);
    },
  });
}
