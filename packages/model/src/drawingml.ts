// packages/model/src/drawingml.ts  -- DrawingML colours, fills, lines and effects as plain data
// Every parser returns only what the XML states; `undefined` means "not stated here, inherit it".
import type { XmlElement } from '@pptx/xml';
import { attr, attrNs, child, children, is, lname } from './dom.ts';
import { ANGLE_UNIT, bool, coord, int, percent } from './units.ts';

// ---------------------------------------------------------------------------------------------------------- colours

export type ColorBase =
  | { kind: 'srgb'; r: number; g: number; b: number }                 // 0..255
  | { kind: 'scheme'; name: string }                                   // tx1, bg2, accent1, phClr, ...
  | { kind: 'sys'; name: string; last?: { r: number; g: number; b: number } }
  | { kind: 'preset'; name: string }
  | { kind: 'hsl'; h: number; s: number; l: number }                   // degrees, fractions
  | { kind: 'scrgb'; r: number; g: number; b: number };                // linear-light fractions

/** One colour transform, applied in document order. `val` is a fraction for percentages, degrees for hue and
 *  hueOff, and 0 for the transforms that take no value (comp, inv, gray, gamma, invGamma). */
export interface ColorTransform { name: string; val: number; }
export interface ColorDef { base: ColorBase; transforms: ColorTransform[]; }

const COLOR_ELEMENTS = ['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr', 'a:hslClr', 'a:scrgbClr'];
const ANGLE_TRANSFORMS = new Set(['hue', 'hueOff']);

function hex(v: string | undefined): { r: number; g: number; b: number } | undefined {
  if (!v || !/^[0-9A-Fa-f]{6}$/.test(v)) return undefined;
  const n = Number.parseInt(v, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

/** A colour element itself: a:srgbClr, a:schemeClr, a:sysClr, a:prstClr, a:hslClr or a:scrgbClr. */
export function parseColor(el: XmlElement | undefined): ColorDef | undefined {
  if (!el) return undefined;
  let base: ColorBase | undefined;
  if (is(el, 'a:srgbClr')) {
    const c = hex(attr(el, 'val'));
    base = c ? { kind: 'srgb', ...c } : undefined;
  } else if (is(el, 'a:schemeClr')) base = { kind: 'scheme', name: attr(el, 'val') ?? 'tx1' };
  else if (is(el, 'a:sysClr')) {
    const last = hex(attr(el, 'lastClr'));
    base = last ? { kind: 'sys', name: attr(el, 'val') ?? 'windowText', last } : { kind: 'sys', name: attr(el, 'val') ?? 'windowText' };
  } else if (is(el, 'a:prstClr')) base = { kind: 'preset', name: attr(el, 'val') ?? 'black' };
  else if (is(el, 'a:hslClr')) {
    base = { kind: 'hsl', h: int(attr(el, 'hue'), 0) / ANGLE_UNIT, s: percent(attr(el, 'sat'), 0), l: percent(attr(el, 'lum'), 0) };
  } else if (is(el, 'a:scrgbClr')) {
    base = { kind: 'scrgb', r: percent(attr(el, 'r'), 0), g: percent(attr(el, 'g'), 0), b: percent(attr(el, 'b'), 0) };
  }
  if (!base) return undefined;
  const transforms: ColorTransform[] = [];
  for (const t of children(el)) {
    if (!is(t, `a:${t.local}`)) continue;            // extension elements are not transforms
    const v = attr(t, 'val');
    transforms.push({ name: t.local, val: ANGLE_TRANSFORMS.has(t.local) ? int(v, 0) / ANGLE_UNIT : percent(v, 0) });
  }
  return { base, transforms };
}

/** The colour inside a wrapper such as a:solidFill, a:fgClr or a:glow (EG_ColorChoice). */
export function colorIn(parent: XmlElement | undefined): ColorDef | undefined {
  if (!parent) return undefined;
  for (const c of parent.children) if (c.kind === 'element' && COLOR_ELEMENTS.some(n => is(c, n))) return parseColor(c);
  return undefined;
}

// ------------------------------------------------------------------------------------------------------------ fills

/** Edges inset from each side, as fractions of the box (a:srcRect, a:fillRect, a:fillToRect, a:tileRect). */
export interface RelRect { l: number; t: number; r: number; b: number; }
export interface GradientStop { pos: number; color: ColorDef; }

export type BlipEffect =
  | { kind: 'alphaModFix'; amt: number }
  | { kind: 'grayscl' }
  | { kind: 'lum'; bright: number; contrast: number }
  | { kind: 'biLevel'; thresh: number }
  | { kind: 'duotone'; colors: ColorDef[] }
  | { kind: 'clrChange'; from?: ColorDef; to?: ColorDef; useA: boolean }
  | { kind: 'clrRepl'; color?: ColorDef }
  | { kind: 'alphaInv' } | { kind: 'alphaFloor' } | { kind: 'alphaCeiling' }
  | { kind: 'alphaBiLevel'; thresh: number }
  | { kind: 'tint'; hue: number; amt: number }
  | { kind: 'hsl'; hue: number; sat: number; lum: number }
  | { kind: 'other'; name: string };

export interface BlipDef {
  embed?: string;           // r:embed: relationship to the image part
  link?: string;            // r:link: external image
  svgEmbed?: string;        // asvg:svgBlip in the blip's extension list: the SVG, with the r:embed image as its PNG fallback
  effects: BlipEffect[];
}

export type FillDef =
  | { kind: 'none' }
  | { kind: 'solid'; color?: ColorDef }
  | {
      kind: 'gradient'; stops: GradientStop[];
      shade: { kind: 'linear'; angle: number; scaled: boolean } | { kind: 'path'; path: 'circle' | 'rect' | 'shape'; fillToRect: RelRect };
      tileRect: RelRect; flip: 'none' | 'x' | 'y' | 'xy'; rotWithShape: boolean;
    }
  | { kind: 'pattern'; preset: string; fg?: ColorDef; bg?: ColorDef }
  | {
      kind: 'blip'; blip?: BlipDef; srcRect: RelRect; rotWithShape: boolean; dpi?: number;
      mode: { kind: 'stretch'; fillRect: RelRect } | { kind: 'tile'; tx: number; ty: number; sx: number; sy: number; flip: string; algn: string } | null;
    }
  | { kind: 'group' };

export const FILL_ELEMENTS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

export function relRect(el: XmlElement | undefined): RelRect {
  return { l: percent(attr(el, 'l'), 0), t: percent(attr(el, 't'), 0), r: percent(attr(el, 'r'), 0), b: percent(attr(el, 'b'), 0) };
}

function parseBlipEffect(e: XmlElement): BlipEffect {
  switch (e.local) {
    case 'alphaModFix': return { kind: 'alphaModFix', amt: percent(attr(e, 'amt'), 1) };
    case 'grayscl': return { kind: 'grayscl' };
    case 'lum': return { kind: 'lum', bright: percent(attr(e, 'bright'), 0), contrast: percent(attr(e, 'contrast'), 0) };
    case 'biLevel': return { kind: 'biLevel', thresh: percent(attr(e, 'thresh'), 0.5) };
    case 'duotone': return { kind: 'duotone', colors: children(e).map(c => parseColor(c)).filter((c): c is ColorDef => !!c) };
    case 'clrChange': {
      const from = colorIn(child(e, 'a:clrFrom'));
      const to = colorIn(child(e, 'a:clrTo'));
      return { kind: 'clrChange', ...(from ? { from } : {}), ...(to ? { to } : {}), useA: bool(attr(e, 'useA'), true) };
    }
    case 'clrRepl': {
      const color = colorIn(e);
      return color ? { kind: 'clrRepl', color } : { kind: 'clrRepl' };
    }
    case 'alphaInv': return { kind: 'alphaInv' };
    case 'alphaFloor': return { kind: 'alphaFloor' };
    case 'alphaCeiling': return { kind: 'alphaCeiling' };
    case 'alphaBiLevel': return { kind: 'alphaBiLevel', thresh: percent(attr(e, 'thresh'), 0.5) };
    case 'tint': return { kind: 'tint', hue: int(attr(e, 'hue'), 0) / ANGLE_UNIT, amt: percent(attr(e, 'amt'), 0) };
    case 'hsl': return { kind: 'hsl', hue: int(attr(e, 'hue'), 0) / ANGLE_UNIT, sat: percent(attr(e, 'sat'), 0), lum: percent(attr(e, 'lum'), 0) };
    default: return { kind: 'other', name: e.local };
  }
}

export function parseBlip(blip: XmlElement | undefined): BlipDef | undefined {
  if (!blip) return undefined;
  const out: BlipDef = { effects: [] };
  const embed = attrNs(blip, 'r', 'embed');
  const link = attrNs(blip, 'r', 'link');
  if (embed) out.embed = embed;
  if (link) out.link = link;
  for (const e of children(blip)) {
    if (is(e, 'a:extLst')) {
      for (const ext of children(e, 'a:ext')) {
        const svg = child(ext, 'asvg:svgBlip');
        const svgEmbed = attrNs(svg, 'r', 'embed');
        if (svgEmbed) out.svgEmbed = svgEmbed;
      }
    } else if (is(e, `a:${e.local}`)) {
      out.effects.push(parseBlipEffect(e));
    }
  }
  return out;
}

/** One fill element: a:noFill, a:solidFill, a:gradFill, a:blipFill, a:pattFill or a:grpFill. */
export function parseFill(el: XmlElement | undefined): FillDef | undefined {
  if (!el) return undefined;
  if (is(el, 'a:noFill')) return { kind: 'none' };
  if (is(el, 'a:grpFill')) return { kind: 'group' };
  if (is(el, 'a:solidFill')) {
    const color = colorIn(el);
    return color ? { kind: 'solid', color } : { kind: 'solid' };
  }
  if (is(el, 'a:gradFill')) {
    const stops = children(child(el, 'a:gsLst'), 'a:gs')
      .map(gs => ({ pos: percent(attr(gs, 'pos'), 0), color: colorIn(gs) }))
      .filter((s): s is GradientStop => s.color !== undefined)
      .sort((x, y) => x.pos - y.pos);
    const lin = child(el, 'a:lin');
    const pathEl = child(el, 'a:path');
    const shade = pathEl
      ? { kind: 'path' as const, path: (attr(pathEl, 'path') ?? 'circle') as 'circle' | 'rect' | 'shape', fillToRect: relRect(child(pathEl, 'a:fillToRect')) }
      : { kind: 'linear' as const, angle: int(attr(lin, 'ang'), 0) / ANGLE_UNIT, scaled: bool(attr(lin, 'scaled'), false) };
    return {
      kind: 'gradient', stops, shade, tileRect: relRect(child(el, 'a:tileRect')),
      flip: (attr(el, 'flip') ?? 'none') as 'none' | 'x' | 'y' | 'xy', rotWithShape: bool(attr(el, 'rotWithShape'), true),
    };
  }
  if (is(el, 'a:pattFill')) {
    const fg = colorIn(child(el, 'a:fgClr'));
    const bg = colorIn(child(el, 'a:bgClr'));
    return { kind: 'pattern', preset: attr(el, 'prst') ?? 'pct5', ...(fg ? { fg } : {}), ...(bg ? { bg } : {}) };
  }
  if (is(el, 'a:blipFill') || is(el, 'p:blipFill')) {
    const tile = child(el, 'a:tile');
    const stretch = child(el, 'a:stretch');
    const mode = tile
      ? {
          kind: 'tile' as const, tx: coord(attr(tile, 'tx'), 0), ty: coord(attr(tile, 'ty'), 0),
          sx: percent(attr(tile, 'sx'), 1), sy: percent(attr(tile, 'sy'), 1),
          flip: attr(tile, 'flip') ?? 'none', algn: attr(tile, 'algn') ?? 'tl',
        }
      : stretch ? { kind: 'stretch' as const, fillRect: relRect(child(stretch, 'a:fillRect')) } : null;
    const blip = parseBlip(child(el, 'a:blip'));
    const dpi = int(attr(el, 'dpi'));
    return {
      kind: 'blip', ...(blip ? { blip } : {}), srcRect: relRect(child(el, 'a:srcRect')), mode,
      rotWithShape: bool(attr(el, 'rotWithShape'), true), ...(dpi !== undefined ? { dpi } : {}),
    };
  }
  return undefined;
}

/** The first fill element among the children of `parent` (EG_FillProperties). */
export function fillIn(parent: XmlElement | undefined): FillDef | undefined {
  if (!parent) return undefined;
  for (const c of parent.children) if (c.kind === 'element' && FILL_ELEMENTS.some(n => is(c, n))) return parseFill(c);
  return undefined;
}

// ------------------------------------------------------------------------------------------------------------ lines

export interface LineEndDef { type: string; w: 'sm' | 'med' | 'lg'; len: 'sm' | 'med' | 'lg'; }
export interface LineDef {
  w?: number;                                  // EMU
  cap?: 'flat' | 'rnd' | 'sq';
  cmpd?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
  algn?: 'ctr' | 'in';
  fill?: FillDef;
  dash?: { kind: 'preset'; name: string } | { kind: 'custom'; segments: { d: number; sp: number }[] };   // custom: fractions of line width
  join?: { kind: 'round' } | { kind: 'bevel' } | { kind: 'miter'; lim?: number };
  headEnd?: LineEndDef;
  tailEnd?: LineEndDef;
}

function lineEnd(el: XmlElement | undefined): LineEndDef | undefined {
  if (!el) return undefined;
  return {
    type: attr(el, 'type') ?? 'none',
    w: (attr(el, 'w') ?? 'med') as LineEndDef['w'],
    len: (attr(el, 'len') ?? 'med') as LineEndDef['len'],
  };
}

/** a:ln, or any element of type CT_LineProperties (a:lnL, a:uLn, ...). */
export function parseLine(ln: XmlElement | undefined): LineDef | undefined {
  if (!ln) return undefined;
  const out: LineDef = {};
  const w = coord(attr(ln, 'w'));
  if (w !== undefined) out.w = w;
  const cap = attr(ln, 'cap');
  if (cap) out.cap = cap as LineDef['cap'];
  const cmpd = attr(ln, 'cmpd');
  if (cmpd) out.cmpd = cmpd as LineDef['cmpd'];
  const algn = attr(ln, 'algn');
  if (algn) out.algn = algn as LineDef['algn'];
  const fill = fillIn(ln);
  if (fill) out.fill = fill;
  const prst = child(ln, 'a:prstDash');
  const cust = child(ln, 'a:custDash');
  if (prst) out.dash = { kind: 'preset', name: attr(prst, 'val') ?? 'solid' };
  else if (cust) out.dash = { kind: 'custom', segments: children(cust, 'a:ds').map(ds => ({ d: percent(attr(ds, 'd'), 1), sp: percent(attr(ds, 'sp'), 1) })) };
  if (child(ln, 'a:round')) out.join = { kind: 'round' };
  else if (child(ln, 'a:bevel')) out.join = { kind: 'bevel' };
  else {
    const miter = child(ln, 'a:miter');
    if (miter) {
      const lim = percent(attr(miter, 'lim'));
      out.join = lim === undefined ? { kind: 'miter' } : { kind: 'miter', lim };
    }
  }
  const head = lineEnd(child(ln, 'a:headEnd'));
  if (head) out.headEnd = head;
  const tail = lineEnd(child(ln, 'a:tailEnd'));
  if (tail) out.tailEnd = tail;
  return out;
}

// ---------------------------------------------------------------------------------------------------------- effects

export interface ShadowDef {
  blurRad: number; dist: number; dir: number;                  // EMU, EMU, degrees
  sx: number; sy: number; kx: number; ky: number;              // scale fractions, skew degrees
  algn: string; rotWithShape: boolean; color?: ColorDef;
}
export interface ReflectionDef {
  blurRad: number; stA: number; stPos: number; endA: number; endPos: number; dist: number; dir: number;
  fadeDir: number; sx: number; sy: number; kx: number; ky: number; algn: string; rotWithShape: boolean;
}
export interface EffectsDef {
  blur?: { rad: number; grow: boolean };
  glow?: { rad: number; color?: ColorDef };
  innerShdw?: ShadowDef;
  outerShdw?: ShadowDef;
  prstShdw?: { prst: string; dist: number; dir: number; color?: ColorDef };
  reflection?: ReflectionDef;
  softEdge?: { rad: number };
  unsupported?: string[];                                       // e.g. effectDag, fillOverlay
}

function shadow(e: XmlElement): ShadowDef {
  const color = colorIn(e);
  return {
    blurRad: coord(attr(e, 'blurRad'), 0), dist: coord(attr(e, 'dist'), 0), dir: int(attr(e, 'dir'), 0) / ANGLE_UNIT,
    sx: percent(attr(e, 'sx'), 1), sy: percent(attr(e, 'sy'), 1),
    kx: int(attr(e, 'kx'), 0) / ANGLE_UNIT, ky: int(attr(e, 'ky'), 0) / ANGLE_UNIT,
    algn: attr(e, 'algn') ?? 'b', rotWithShape: bool(attr(e, 'rotWithShape'), true),
    ...(color ? { color } : {}),
  };
}

/** a:effectLst (an empty one means "no effects", which differs from no element at all), or a:effectDag. */
export function parseEffects(el: XmlElement | undefined): EffectsDef | undefined {
  if (!el) return undefined;
  if (is(el, 'a:effectDag')) return { unsupported: ['effectDag'] };
  const out: EffectsDef = {};
  for (const e of children(el)) {
    switch (lname(e)) {
      case 'a:blur': out.blur = { rad: coord(attr(e, 'rad'), 0), grow: bool(attr(e, 'grow'), true) }; break;
      case 'a:glow': {
        const color = colorIn(e);
        out.glow = { rad: coord(attr(e, 'rad'), 0), ...(color ? { color } : {}) };
        break;
      }
      case 'a:innerShdw': out.innerShdw = shadow(e); break;
      case 'a:outerShdw': out.outerShdw = shadow(e); break;
      case 'a:prstShdw': {
        const color = colorIn(e);
        out.prstShdw = { prst: attr(e, 'prst') ?? 'shdw1', dist: coord(attr(e, 'dist'), 0), dir: int(attr(e, 'dir'), 0) / ANGLE_UNIT, ...(color ? { color } : {}) };
        break;
      }
      case 'a:reflection':
        out.reflection = {
          blurRad: coord(attr(e, 'blurRad'), 0), stA: percent(attr(e, 'stA'), 1), stPos: percent(attr(e, 'stPos'), 0),
          endA: percent(attr(e, 'endA'), 0), endPos: percent(attr(e, 'endPos'), 1), dist: coord(attr(e, 'dist'), 0),
          dir: int(attr(e, 'dir'), 0) / ANGLE_UNIT, fadeDir: int(attr(e, 'fadeDir'), 5400000) / ANGLE_UNIT,
          sx: percent(attr(e, 'sx'), 1), sy: percent(attr(e, 'sy'), 1), kx: int(attr(e, 'kx'), 0) / ANGLE_UNIT,
          ky: int(attr(e, 'ky'), 0) / ANGLE_UNIT, algn: attr(e, 'algn') ?? 'b', rotWithShape: bool(attr(e, 'rotWithShape'), true),
        };
        break;
      case 'a:softEdge': out.softEdge = { rad: coord(attr(e, 'rad'), 0) }; break;
      default: (out.unsupported ??= []).push(e.local);
    }
  }
  return out;
}

/** The effect element among the children of `parent` (EG_EffectProperties): a:effectLst or a:effectDag. */
export function effectsIn(parent: XmlElement | undefined): EffectsDef | undefined {
  return parseEffects(child(parent, 'a:effectLst') ?? child(parent, 'a:effectDag'));
}

// ------------------------------------------------------------------------------------------------- style references

/** p:style: references into the theme's style matrix, each with the colour that replaces phClr. */
export interface StyleRefs {
  ln?: { idx: number; color?: ColorDef };
  fill?: { idx: number; color?: ColorDef };
  effect?: { idx: number; color?: ColorDef };
  font?: { idx: 'major' | 'minor' | 'none'; color?: ColorDef };
}

export function parseStyle(style: XmlElement | undefined): StyleRefs | undefined {
  if (!style) return undefined;
  const ref = (name: string) => {
    const e = child(style, name);
    if (!e) return undefined;
    const color = colorIn(e);
    return { idx: int(attr(e, 'idx'), 0), ...(color ? { color } : {}) };
  };
  const out: StyleRefs = {};
  const ln = ref('a:lnRef');
  const fill = ref('a:fillRef');
  const effect = ref('a:effectRef');
  if (ln) out.ln = ln;
  if (fill) out.fill = fill;
  if (effect) out.effect = effect;
  const f = child(style, 'a:fontRef');
  if (f) {
    const color = colorIn(f);
    out.font = { idx: (attr(f, 'idx') ?? 'minor') as 'major' | 'minor' | 'none', ...(color ? { color } : {}) };
  }
  return out;
}
