// packages/model/src/theme.ts  -- a:theme (ECMA-376 Part 1, 20.1.6)
import type { XmlElement } from '@pptx/xml';
import { attr, child, children, is, path } from './dom.ts';
import { colorIn, type ColorDef } from './drawingml.ts';

/** The fonts of one theme font collection: one typeface per script class, plus per-script overrides keyed by
 *  ISO 15924 script code ("Jpan", "Hans", "Arab", ...). Empty typeface means "use the application default". */
export interface FontCollection { latin: string; ea: string; cs: string; scripts: Record<string, string>; }

export interface ThemeDef {
  name: string;
  colorSchemeName: string;
  colors: Map<string, ColorDef>;            // dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink
  major: FontCollection;
  minor: FontCollection;
  fillStyles: XmlElement[];                 // a:fillStyleLst: fillRef idx 1..999
  lineStyles: XmlElement[];                 // a:lnStyleLst: lnRef idx 1..
  effectStyles: XmlElement[];               // a:effectStyleLst/a:effectStyle
  bgFillStyles: XmlElement[];               // a:bgFillStyleLst: fillRef idx 1001..
}

export const SCHEME_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;

function fontCollection(el: XmlElement | undefined): FontCollection {
  const scripts: Record<string, string> = {};
  for (const f of children(el, 'a:font')) {
    const script = attr(f, 'script');
    if (script) scripts[script] = attr(f, 'typeface') ?? '';
  }
  return {
    latin: attr(child(el, 'a:latin'), 'typeface') ?? '',
    ea: attr(child(el, 'a:ea'), 'typeface') ?? '',
    cs: attr(child(el, 'a:cs'), 'typeface') ?? '',
    scripts,
  };
}

const fillChildren = (el: XmlElement | undefined) =>
  children(el).filter(c => ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'].some(n => is(c, n)));

/** a:theme, or a:themeOverride (which has the same theme elements without a name). */
export function parseTheme(root: XmlElement): ThemeDef {
  const elements = is(root, 'a:themeOverride') ? root : child(root, 'a:themeElements');
  const clrScheme = child(elements, 'a:clrScheme');
  const colors = new Map<string, ColorDef>();
  for (const slot of SCHEME_SLOTS) {
    const c = colorIn(child(clrScheme, `a:${slot}`));
    if (c) colors.set(slot, c);
  }
  const fontScheme = child(elements, 'a:fontScheme');
  const fmt = child(elements, 'a:fmtScheme');
  return {
    name: attr(root, 'name') ?? '',
    colorSchemeName: attr(clrScheme, 'name') ?? '',
    colors,
    major: fontCollection(child(fontScheme, 'a:majorFont')),
    minor: fontCollection(child(fontScheme, 'a:minorFont')),
    fillStyles: fillChildren(child(fmt, 'a:fillStyleLst')),
    lineStyles: children(child(fmt, 'a:lnStyleLst'), 'a:ln'),
    effectStyles: children(child(fmt, 'a:effectStyleLst'), 'a:effectStyle'),
    bgFillStyles: fillChildren(child(fmt, 'a:bgFillStyleLst')),
  };
}

/** The theme's name as shown in PowerPoint's Design tab. */
export const themeName = (root: XmlElement) => attr(root, 'name') ?? attr(path(root, 'a:themeElements', 'a:clrScheme'), 'name') ?? '';
