// packages/resolve/src/paint.ts  -- fills, lines and effects with every colour resolved
import type { PresentationDocument, BlipEffect, ColorDef, EffectsDef, FillDef, LineDef, LineEndDef, RelRect, ShadowDef, ReflectionDef, StyleRefs, ThemeDef } from '@pptx/model';
import { effectsIn, parseFill, parseLine } from '@pptx/model';
import { resolveColor, type ColorContext, type Rgba } from './color.ts';

/** An image part to draw: the raster part, plus the SVG part when the picture has one (render the SVG). */
export interface ImageRef { part?: string; svgPart?: string; externalUrl?: string; }

export interface Stop { pos: number; color: Rgba; }
export type ResolvedBlipEffect =
  | { kind: 'alphaModFix'; amt: number }
  | { kind: 'grayscl' }
  | { kind: 'lum'; bright: number; contrast: number }
  | { kind: 'biLevel'; thresh: number }
  | { kind: 'duotone'; colors: Rgba[] }
  | { kind: 'clrChange'; from: Rgba; to: Rgba; useA: boolean }
  | { kind: 'clrRepl'; color: Rgba }
  | { kind: 'other'; name: string };

export type Paint =
  | { kind: 'none' }
  | { kind: 'solid'; color: Rgba }
  | { kind: 'linear'; angle: number; scaled: boolean; rotWithShape: boolean; stops: Stop[]; tileRect: RelRect; flip: string }
  | { kind: 'path'; path: 'circle' | 'rect' | 'shape'; fillToRect: RelRect; rotWithShape: boolean; stops: Stop[]; tileRect: RelRect }
  | { kind: 'pattern'; preset: string; fg: Rgba; bg: Rgba }
  | {
      kind: 'image'; image: ImageRef; srcRect: RelRect; rotWithShape: boolean; effects: ResolvedBlipEffect[];
      mode: { kind: 'stretch'; fillRect: RelRect } | { kind: 'tile'; tx: number; ty: number; sx: number; sy: number; flip: string; algn: string };
    }
  /** a:grpFill before it is replaced by the enclosing group's fill; resolved shapes never carry it. */
  | { kind: 'group' };

export interface LineEnd { type: string; w: 'sm' | 'med' | 'lg'; len: 'sm' | 'med' | 'lg'; }
export interface Stroke {
  width: number;                                  // EMU
  paint: Paint;
  cap: 'flat' | 'rnd' | 'sq';
  join: 'round' | 'bevel' | 'miter';
  miterLimit: number;
  dash: number[] | null;                          // alternating dash and gap lengths in multiples of the width
  compound: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
  align: 'ctr' | 'in';
  head?: LineEnd;
  tail?: LineEnd;
}

export interface ResolvedShadow {
  color: Rgba; blurRad: number; dist: number; dir: number; sx: number; sy: number; kx: number; ky: number; algn: string; rotWithShape: boolean;
}
export interface ResolvedEffects {
  outerShadow?: ResolvedShadow;
  innerShadow?: ResolvedShadow;
  glow?: { radius: number; color: Rgba };
  softEdge?: { radius: number };
  reflection?: ReflectionDef;
  blur?: { radius: number; grow: boolean };
  unsupported?: string[];
}

/** Where relationship ids in a fill resolve: the part that holds the XML (a slide, a layout, a theme). */
export interface PaintScope { doc: PresentationDocument; part: string; colors: ColorContext; }

const BLACK: Rgba = { r: 0, g: 0, b: 0, a: 1 };
const WHITE: Rgba = { r: 1, g: 1, b: 1, a: 1 };
const color = (def: ColorDef | undefined, ctx: ColorContext, dflt: Rgba = BLACK) => (def ? resolveColor(def, ctx) : dflt);

function blipEffect(e: BlipEffect, ctx: ColorContext): ResolvedBlipEffect {
  switch (e.kind) {
    case 'duotone': return { kind: 'duotone', colors: e.colors.map(c => resolveColor(c, ctx)) };
    case 'clrChange': return { kind: 'clrChange', from: color(e.from, ctx), to: color(e.to, ctx), useA: e.useA };
    case 'clrRepl': return { kind: 'clrRepl', color: color(e.color, ctx) };
    case 'alphaModFix': case 'grayscl': case 'lum': case 'biLevel': return e;
    default: return { kind: 'other', name: e.kind === 'other' ? e.name : e.kind };
  }
}

/** A fill definition with its colours and image relationships resolved. */
export function resolveFill(fill: FillDef | undefined, scope: PaintScope): Paint {
  if (!fill) return { kind: 'none' };
  const ctx = scope.colors;
  switch (fill.kind) {
    case 'none': return { kind: 'none' };
    case 'group': return { kind: 'group' };
    case 'solid': return { kind: 'solid', color: color(fill.color, ctx) };
    case 'pattern': return { kind: 'pattern', preset: fill.preset, fg: color(fill.fg, ctx), bg: color(fill.bg, ctx, WHITE) };
    case 'gradient': {
      const stops = fill.stops.map(s => ({ pos: s.pos, color: resolveColor(s.color, ctx) }));
      if (fill.shade.kind === 'linear') {
        return { kind: 'linear', angle: fill.shade.angle, scaled: fill.shade.scaled, rotWithShape: fill.rotWithShape, stops, tileRect: fill.tileRect, flip: fill.flip };
      }
      return { kind: 'path', path: fill.shade.path, fillToRect: fill.shade.fillToRect, rotWithShape: fill.rotWithShape, stops, tileRect: fill.tileRect };
    }
    case 'blip': {
      const image: ImageRef = {};
      const blip = fill.blip;
      if (blip?.embed) {
        const part = scope.doc.target(scope.part, blip.embed);
        if (part) image.part = part;
      }
      if (blip?.svgEmbed) {
        const svg = scope.doc.target(scope.part, blip.svgEmbed);
        if (svg) image.svgPart = svg;
      }
      if (blip?.link) {
        const rel = scope.doc.pkg.relationships(scope.part).find(r => r.id === blip.link);
        if (rel?.external) image.externalUrl = rel.target;
      }
      return {
        kind: 'image', image, srcRect: fill.srcRect, rotWithShape: fill.rotWithShape,
        effects: (blip?.effects ?? []).map(e => blipEffect(e, ctx)),
        mode: fill.mode ?? { kind: 'stretch', fillRect: { l: 0, t: 0, r: 0, b: 0 } },
      };
    }
  }
}

/** Dash patterns of ST_PresetLineDashVal, in multiples of the line width (ECMA-376 Part 1, 20.1.10.48). */
export const PRESET_DASHES: Record<string, number[] | null> = {
  solid: null,
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** PowerPoint's default line width when neither the shape nor its style gives one: 0.75 pt. */
export const DEFAULT_LINE_WIDTH = 9525;

/** Merge a shape's own a:ln over the line from its style reference: attributes and children it states win. */
export function mergeLine(over: LineDef | undefined, under: LineDef | undefined): LineDef | undefined {
  if (!over) return under;
  if (!under) return over;
  return { ...under, ...over };
}

/** A line definition resolved to a stroke; null when it draws nothing. */
export function resolveLine(line: LineDef | undefined, scope: PaintScope): Stroke | null {
  if (!line) return null;
  const paint = resolveFill(line.fill, scope);
  if (paint.kind === 'none' || !line.fill) return null;
  const join = line.join ?? { kind: 'round' };
  const lineEnd = (e: LineEndDef | undefined): LineEnd | undefined => (e && e.type !== 'none' ? { type: e.type, w: e.w, len: e.len } : undefined);
  const head = lineEnd(line.headEnd);
  const tail = lineEnd(line.tailEnd);
  return {
    width: line.w ?? DEFAULT_LINE_WIDTH, paint,
    cap: line.cap ?? 'flat', join: join.kind, miterLimit: join.kind === 'miter' ? join.lim ?? 8 : 8,
    dash: !line.dash ? null : line.dash.kind === 'preset' ? PRESET_DASHES[line.dash.name] ?? null : line.dash.segments.flatMap(s => [s.d, s.sp]),
    compound: line.cmpd ?? 'sng', align: line.algn ?? 'ctr',
    ...(head ? { head } : {}), ...(tail ? { tail } : {}),
  };
}

function shadowOf(s: ShadowDef, ctx: ColorContext): ResolvedShadow {
  return {
    color: color(s.color, ctx), blurRad: s.blurRad, dist: s.dist, dir: s.dir, sx: s.sx, sy: s.sy, kx: s.kx, ky: s.ky, algn: s.algn, rotWithShape: s.rotWithShape,
  };
}

/** Preset shadows (a:prstShdw shdw1..shdw20) drawn as an outer shadow with the given distance and direction. */
export function resolveEffects(effects: EffectsDef | undefined, scope: PaintScope): ResolvedEffects {
  const out: ResolvedEffects = {};
  if (!effects) return out;
  const ctx = scope.colors;
  if (effects.outerShdw) out.outerShadow = shadowOf(effects.outerShdw, ctx);
  else if (effects.prstShdw) {
    const p = effects.prstShdw;
    out.outerShadow = { color: color(p.color, ctx), blurRad: 0, dist: p.dist, dir: p.dir, sx: 1, sy: 1, kx: 0, ky: 0, algn: 'b', rotWithShape: true };
  }
  if (effects.innerShdw) out.innerShadow = shadowOf(effects.innerShdw, ctx);
  if (effects.glow) out.glow = { radius: effects.glow.rad, color: color(effects.glow.color, ctx) };
  if (effects.softEdge) out.softEdge = { radius: effects.softEdge.rad };
  if (effects.reflection) out.reflection = effects.reflection;
  if (effects.blur) out.blur = { radius: effects.blur.rad, grow: effects.blur.grow };
  if (effects.unsupported?.length) out.unsupported = effects.unsupported;
  return out;
}

// ------------------------------------------------------------------------------------------------- style matrix

/** The colour context a style reference applies in: its own colour replaces phClr. */
export function withPlaceholderColor(scope: PaintScope, ref: { color?: ColorDef } | undefined): PaintScope {
  if (!ref?.color) return scope;
  return { ...scope, colors: { ...scope.colors, phClr: resolveColor(ref.color, scope.colors) } };
}

/** The fill a fillRef (or bgRef) index selects: 0 none, 1..999 the theme's fill styles, 1001.. its background fills. */
export function themeFill(theme: ThemeDef, idx: number): FillDef | undefined {
  if (idx <= 0) return { kind: 'none' };
  const el = idx >= 1001 ? theme.bgFillStyles[idx - 1001] : theme.fillStyles[idx - 1];
  return parseFill(el);
}

export function themeLine(theme: ThemeDef, idx: number): LineDef | undefined {
  if (idx <= 0) return undefined;
  return parseLine(theme.lineStyles[idx - 1]);
}

export function themeEffects(theme: ThemeDef, idx: number): EffectsDef | undefined {
  if (idx <= 0) return undefined;
  return effectsIn(theme.effectStyles[idx - 1]);
}

/**
 * Resolve a shape's fill, line and effects from its own spPr and its style references. The theme scope is where the
 * theme's own relationships resolve (picture fills in a theme point at images through the theme part).
 */
export function resolveShapeStyle(
  spPr: { fill?: FillDef; line?: LineDef; effects?: EffectsDef },
  refs: StyleRefs | undefined,
  theme: ThemeDef,
  scope: PaintScope,
  themeScope: PaintScope,
): { fill: Paint; line: Stroke | null; effects: ResolvedEffects } {
  let fill: Paint;
  if (spPr.fill) fill = resolveFill(spPr.fill, scope);
  else if (refs?.fill) fill = resolveFill(themeFill(theme, refs.fill.idx), withPlaceholderColor(themeScope, refs.fill));
  else fill = { kind: 'none' };

  // A line from the style reference is resolved with its phClr; the shape's own a:ln overrides it piece by piece,
  // its colours resolving in the shape's own scope. Resolving each piece in its own scope keeps both right.
  const refLine = refs?.ln ? themeLine(theme, refs.ln.idx) : undefined;
  let line: Stroke | null;
  if (refLine && spPr.line) {
    const merged = mergeLine(spPr.line, refLine)!;
    const fromShape = spPr.line.fill !== undefined;
    line = resolveLine(merged, fromShape ? scope : withPlaceholderColor(themeScope, refs!.ln));
  } else if (spPr.line) line = resolveLine(spPr.line, scope);
  else if (refLine) line = resolveLine(refLine, withPlaceholderColor(themeScope, refs!.ln));
  else line = null;

  let effects: ResolvedEffects;
  if (spPr.effects) effects = resolveEffects(spPr.effects, scope);
  else if (refs?.effect) effects = resolveEffects(themeEffects(theme, refs.effect.idx), withPlaceholderColor(themeScope, refs.effect));
  else effects = {};
  return { fill, line, effects };
}

