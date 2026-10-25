// packages/resolve/src/color.ts  -- DrawingML colour resolution (ECMA-376 Part 1, 20.1.2.3)
//
// Which colour space each transform works in is not spelled out by the spec. This follows LibreOffice's
// implementation, which was tuned against PowerPoint over many years: lumMod/lumOff/sat*/hue* in HSL computed from
// sRGB; tint, shade and the red/green/blue transforms in linear-light RGB; alpha on its own channel. The plan's
// colour oracle (a grid of transforms exported by PowerPoint) is how to check and fit this; until it runs, the
// results match LibreOffice's.
import type { ColorBase, ColorDef, ColorTransform } from '@pptx/model';

/** sRGB, every channel 0..1 (alpha 1 is opaque). */
export interface Rgba { r: number; g: number; b: number; a: number; }

export interface ColorContext {
  /** Theme colour slots: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink. */
  theme: ReadonlyMap<string, ColorDef>;
  /** The colour map: names used in the file (bg1, tx1, bg2, tx2, accentN, hlink, folHlink) to theme slots. */
  clrMap: Readonly<Record<string, string>>;
  /** The colour that replaces phClr: the one on the style reference being applied. */
  phClr?: Rgba;
}

const clamp01 = (v: number) => (v < 0 ? 0 : v > 1 ? 1 : v);
const toLinear = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c: number) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);

interface Hsl { h: number; s: number; l: number; }

export function rgbToHsl(r: number, g: number, b: number): Hsl {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return { h: h * 60, s, l };
}

export function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  const hh = (((h % 360) + 360) % 360) / 360;
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return [f(hh + 1 / 3), f(hh), f(hh - 1 / 3)];
}

/** ST_SystemColorVal fallbacks, used only when a:sysClr has no lastClr (Windows defaults). */
const SYSTEM_COLORS: Record<string, number> = {
  windowText: 0x000000, window: 0xffffff, menu: 0xf0f0f0, menuText: 0x000000, btnFace: 0xf0f0f0, btnText: 0x000000,
  btnShadow: 0xa0a0a0, btnHighlight: 0xffffff, highlight: 0x0078d7, highlightText: 0xffffff, grayText: 0x6d6d6d,
  captionText: 0x000000, activeCaption: 0x99b4d1, inactiveCaption: 0xbfcddb, infoText: 0x000000, infoBk: 0xffffe1,
  appWorkspace: 0xababab, background: 0x000000, scrollBar: 0xc8c8c8, windowFrame: 0x646464, '3dDkShadow': 0x696969, '3dLight': 0xe3e3e3,
  activeBorder: 0xb4b4b4, inactiveBorder: 0xf4f7fc, inactiveCaptionText: 0x000000, hotLight: 0x0066cc,
  gradientActiveCaption: 0xb9d1ea, gradientInactiveCaption: 0xd7e4f2, menuHighlight: 0x3399ff, menuBar: 0xf0f0f0,
};

/** ST_PresetColorVal: the CSS named colours, with Office's own abbreviated aliases (dkBlue, ltGray, medPurple, ...). */
const CSS_COLORS: Record<string, number> = {
  aliceBlue: 0xf0f8ff, antiqueWhite: 0xfaebd7, aqua: 0x00ffff, aquamarine: 0x7fffd4, azure: 0xf0ffff, beige: 0xf5f5dc, bisque: 0xffe4c4,
  black: 0x000000, blanchedAlmond: 0xffebcd, blue: 0x0000ff, blueViolet: 0x8a2be2, brown: 0xa52a2a, burlyWood: 0xdeb887, cadetBlue: 0x5f9ea0,
  chartreuse: 0x7fff00, chocolate: 0xd2691e, coral: 0xff7f50, cornflowerBlue: 0x6495ed, cornsilk: 0xfff8dc, crimson: 0xdc143c, cyan: 0x00ffff,
  darkBlue: 0x00008b, darkCyan: 0x008b8b, darkGoldenrod: 0xb8860b, darkGray: 0xa9a9a9, darkGrey: 0xa9a9a9, darkGreen: 0x006400,
  darkKhaki: 0xbdb76b, darkMagenta: 0x8b008b, darkOliveGreen: 0x556b2f, darkOrange: 0xff8c00, darkOrchid: 0x9932cc, darkRed: 0x8b0000,
  darkSalmon: 0xe9967a, darkSeaGreen: 0x8fbc8f, darkSlateBlue: 0x483d8b, darkSlateGray: 0x2f4f4f, darkSlateGrey: 0x2f4f4f,
  darkTurquoise: 0x00ced1, darkViolet: 0x9400d3, deepPink: 0xff1493, deepSkyBlue: 0x00bfff, dimGray: 0x696969, dimGrey: 0x696969,
  dodgerBlue: 0x1e90ff, firebrick: 0xb22222, floralWhite: 0xfffaf0, forestGreen: 0x228b22, fuchsia: 0xff00ff, gainsboro: 0xdcdcdc,
  ghostWhite: 0xf8f8ff, gold: 0xffd700, goldenrod: 0xdaa520, gray: 0x808080, grey: 0x808080, green: 0x008000, greenYellow: 0xadff2f,
  honeydew: 0xf0fff0, hotPink: 0xff69b4, indianRed: 0xcd5c5c, indigo: 0x4b0082, ivory: 0xfffff0, khaki: 0xf0e68c, lavender: 0xe6e6fa,
  lavenderBlush: 0xfff0f5, lawnGreen: 0x7cfc00, lemonChiffon: 0xfffacd, lightBlue: 0xadd8e6, lightCoral: 0xf08080, lightCyan: 0xe0ffff,
  lightGoldenrodYellow: 0xfafad2, lightGray: 0xd3d3d3, lightGrey: 0xd3d3d3, lightGreen: 0x90ee90, lightPink: 0xffb6c1, lightSalmon: 0xffa07a,
  lightSeaGreen: 0x20b2aa, lightSkyBlue: 0x87cefa, lightSlateGray: 0x778899, lightSlateGrey: 0x778899, lightSteelBlue: 0xb0c4de,
  lightYellow: 0xffffe0, lime: 0x00ff00, limeGreen: 0x32cd32, linen: 0xfaf0e6, magenta: 0xff00ff, maroon: 0x800000,
  mediumAquamarine: 0x66cdaa, mediumBlue: 0x0000cd, mediumOrchid: 0xba55d3, mediumPurple: 0x9370db, mediumSeaGreen: 0x3cb371,
  mediumSlateBlue: 0x7b68ee, mediumSpringGreen: 0x00fa9a, mediumTurquoise: 0x48d1cc, mediumVioletRed: 0xc71585, midnightBlue: 0x191970,
  mintCream: 0xf5fffa, mistyRose: 0xffe4e1, moccasin: 0xffe4b5, navajoWhite: 0xffdead, navy: 0x000080, oldLace: 0xfdf5e6, olive: 0x808000,
  oliveDrab: 0x6b8e23, orange: 0xffa500, orangeRed: 0xff4500, orchid: 0xda70d6, paleGoldenrod: 0xeee8aa, paleGreen: 0x98fb98,
  paleTurquoise: 0xafeeee, paleVioletRed: 0xdb7093, papayaWhip: 0xffefd5, peachPuff: 0xffdab9, peru: 0xcd853f, pink: 0xffc0cb, plum: 0xdda0dd,
  powderBlue: 0xb0e0e6, purple: 0x800080, red: 0xff0000, rosyBrown: 0xbc8f8f, royalBlue: 0x4169e1, saddleBrown: 0x8b4513, salmon: 0xfa8072,
  sandyBrown: 0xf4a460, seaGreen: 0x2e8b57, seaShell: 0xfff5ee, sienna: 0xa0522d, silver: 0xc0c0c0, skyBlue: 0x87ceeb, slateBlue: 0x6a5acd,
  slateGray: 0x708090, slateGrey: 0x708090, snow: 0xfffafa, springGreen: 0x00ff7f, steelBlue: 0x4682b4, tan: 0xd2b48c, teal: 0x008080,
  thistle: 0xd8bfd8, tomato: 0xff6347, turquoise: 0x40e0d0, violet: 0xee82ee, wheat: 0xf5deb3, white: 0xffffff, whiteSmoke: 0xf5f5f5,
  yellow: 0xffff00, yellowGreen: 0x9acd32,
};
const PRESET_COLORS: Record<string, number> = { ...CSS_COLORS };
for (const [name, v] of Object.entries(CSS_COLORS)) {
  const alias = name.replace(/^dark/, 'dk').replace(/^light/, 'lt').replace(/^medium/, 'med');
  if (alias !== name) PRESET_COLORS[alias] = v;
}

const fromHex = (n: number, a = 1): Rgba => ({ r: ((n >> 16) & 255) / 255, g: ((n >> 8) & 255) / 255, b: (n & 255) / 255, a });
const BLACK: Rgba = { r: 0, g: 0, b: 0, a: 1 };

function baseColor(base: ColorBase, ctx: ColorContext, depth: number): Rgba {
  switch (base.kind) {
    case 'srgb': return { r: base.r / 255, g: base.g / 255, b: base.b / 255, a: 1 };
    case 'scrgb': return { r: toSrgb(clamp01(base.r)), g: toSrgb(clamp01(base.g)), b: toSrgb(clamp01(base.b)), a: 1 };
    case 'hsl': {
      const [r, g, b] = hslToRgb(base.h, clamp01(base.s), clamp01(base.l));
      return { r, g, b, a: 1 };
    }
    case 'preset': return fromHex(PRESET_COLORS[base.name] ?? 0);
    case 'sys': return base.last ? { r: base.last.r / 255, g: base.last.g / 255, b: base.last.b / 255, a: 1 } : fromHex(SYSTEM_COLORS[base.name] ?? 0);
    case 'scheme': {
      if (base.name === 'phClr') return ctx.phClr ?? BLACK;
      // bg1/tx1/... go through the colour map; theme slot names (dk1, accent1) may also be used directly
      const slot = ctx.clrMap[base.name] ?? base.name;
      const def = ctx.theme.get(slot);
      if (!def || depth > 4) return BLACK;
      return resolveColorAt(def, ctx, depth + 1);
    }
  }
}

function applyTransform(c: Rgba, t: ColorTransform): Rgba {
  const v = t.val;
  switch (t.name) {
    case 'alpha': return { ...c, a: clamp01(v) };
    case 'alphaMod': return { ...c, a: clamp01(c.a * v) };
    case 'alphaOff': return { ...c, a: clamp01(c.a + v) };
    case 'lum': case 'lumMod': case 'lumOff': case 'sat': case 'satMod': case 'satOff': case 'hue': case 'hueMod': case 'hueOff': case 'comp': {
      const h = rgbToHsl(c.r, c.g, c.b);
      if (t.name === 'lum') h.l = clamp01(v);
      else if (t.name === 'lumMod') h.l = clamp01(h.l * v);
      else if (t.name === 'lumOff') h.l = clamp01(h.l + v);
      else if (t.name === 'sat') h.s = clamp01(v);
      else if (t.name === 'satMod') h.s = clamp01(h.s * v);
      else if (t.name === 'satOff') h.s = clamp01(h.s + v);
      else if (t.name === 'hue') h.h = v;
      else if (t.name === 'hueMod') h.h = h.h * v;
      else if (t.name === 'hueOff') h.h = h.h + v;
      else h.h = h.h + 180;
      const [r, g, b] = hslToRgb(h.h, h.s, h.l);
      return { r, g, b, a: c.a };
    }
    case 'tint': {
      // "A 10% tint is 10% of the input color combined with 90% white", in linear light
      const f = (x: number) => toSrgb(1 - (1 - toLinear(x)) * clamp01(v));
      return { r: f(c.r), g: f(c.g), b: f(c.b), a: c.a };
    }
    case 'shade': {
      const f = (x: number) => toSrgb(toLinear(x) * clamp01(v));
      return { r: f(c.r), g: f(c.g), b: f(c.b), a: c.a };
    }
    case 'red': case 'redMod': case 'redOff': case 'green': case 'greenMod': case 'greenOff': case 'blue': case 'blueMod': case 'blueOff': {
      const channel = t.name.startsWith('red') ? 'r' : t.name.startsWith('green') ? 'g' : 'b';
      const lin = toLinear(c[channel]);
      const next = t.name.endsWith('Mod') ? lin * v : t.name.endsWith('Off') ? lin + v : v;
      return { ...c, [channel]: toSrgb(clamp01(next)) };
    }
    case 'gray': {
      const y = 0.3 * c.r + 0.59 * c.g + 0.11 * c.b;
      return { r: y, g: y, b: y, a: c.a };
    }
    case 'inv': return { r: 1 - c.r, g: 1 - c.g, b: 1 - c.b, a: c.a };
    case 'gamma': return { r: toSrgb(c.r), g: toSrgb(c.g), b: toSrgb(c.b), a: c.a };
    case 'invGamma': return { r: toLinear(c.r), g: toLinear(c.g), b: toLinear(c.b), a: c.a };
    default: return c;     // unknown transforms are ignored, as an Office consumer does
  }
}

function resolveColorAt(def: ColorDef, ctx: ColorContext, depth: number): Rgba {
  let c = baseColor(def.base, ctx, depth);
  for (const t of def.transforms) c = applyTransform(c, t);
  return { r: clamp01(c.r), g: clamp01(c.g), b: clamp01(c.b), a: clamp01(c.a) };
}

export function resolveColor(def: ColorDef, ctx: ColorContext): Rgba {
  return resolveColorAt(def, ctx, 0);
}

/** '#RRGGBB' or '#RRGGBBAA' for debugging and tests. */
export function toHex(c: Rgba, withAlpha = false): string {
  const h = (v: number) => Math.round(clamp01(v) * 255).toString(16).padStart(2, '0').toUpperCase();
  return `#${h(c.r)}${h(c.g)}${h(c.b)}${withAlpha ? h(c.a) : ''}`;
}

export const presetColorNames = () => Object.keys(PRESET_COLORS);
