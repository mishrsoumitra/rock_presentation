// packages/resolve/src/text.ts  -- effective text properties through the inheritance chain (plan section 2.4)
//
// For a paragraph at outline level N, nearest first:
//   1. the paragraph's a:pPr / the run's a:rPr
//   2. level N of the shape's own a:lstStyle (then its a:defPPr)
//   3. the same in the layout placeholder it inherits from, then in the master placeholder
//   4. the shape style's fontRef (font and colour), for shapes that have a p:style
//   5. level N of the master's p:txStyles bucket (title, body or other), or the notes master's p:notesStyle
//   6. level N of p:defaultTextStyle in presentation.xml
//   7. hard defaults: 18 pt, the minor theme font, tx1, left aligned, single spacing
import type { XmlElement } from '@pptx/xml';
import {
  child, parseBodyProps, parseListStyle, parseParaProps, parseRunProps, paragraphsOf, runsOf,
  type AutofitDef, type BodyPropsDef, type ColorDef, type FontCollection, type ListStyleDef, type ParaPropsDef, type RunPropsDef,
  type Spacing, type TabStop, type ThemeDef,
} from '@pptx/model';
import { resolveColor, type Rgba } from './color.ts';
import { resolveEffects, resolveFill, resolveLine, type ImageRef, type Paint, type PaintScope, type ResolvedEffects, type Stroke } from './paint.ts';

export interface ResolvedBodyProps {
  insets: { l: number; t: number; r: number; b: number };       // EMU
  anchor: 't' | 'ctr' | 'b' | 'just' | 'dist';
  anchorCtr: boolean;
  wrap: boolean;
  vert: string;
  rot: number;                                                    // degrees
  upright: boolean;
  numCol: number;
  spcCol: number;
  rtlCol: boolean;
  autofit: AutofitDef;
  vertOverflow: string;
  horzOverflow: string;
  spcFirstLastPara: boolean;
  warp?: { prst: string; av: { name: string; fmla: string }[] };
}

export interface ResolvedRunProps {
  size: number;                                                   // points, before autofit scaling
  bold: boolean;
  italic: boolean;
  underline: string;
  strike: 'noStrike' | 'sngStrike' | 'dblStrike';
  caps: 'none' | 'small' | 'all';
  spacing: number;                                                // points
  kern: number;                                                   // points; 0 = no kerning
  baseline: number;                                               // fraction of the font size
  latin: string;                                                  // typefaces, theme references resolved; '' = none given
  ea: string;
  cs: string;
  sym?: string;
  /** The theme font collection the run's fonts came from, for per-script fallbacks when ea or cs is empty. */
  scriptFonts: Record<string, string>;
  lang: string;
  altLang?: string;
  fill: Paint;
  outline: Stroke | null;
  highlight: Rgba | null;
  underlineLine?: Stroke | null;
  underlineFill?: Paint;
  effects?: ResolvedEffects;
  hyperlink?: { url?: string; action?: string; tooltip?: string; slide?: string };
  rtl: boolean;
}

export interface ResolvedBullet {
  kind: 'char' | 'autonum' | 'picture';
  char?: string;
  scheme?: string;
  startAt?: number;
  image?: ImageRef;
  font: string | null;                                            // null: the text's own font
  size: { kind: 'pct'; val: number } | { kind: 'pts'; val: number } | null;   // null: 100% of the text
  color: Rgba | null;                                             // null: the text's own colour
}

export interface ResolvedParaProps {
  level: number;
  align: string;
  marL: number;
  marR: number;
  indent: number;
  lineSpacing: Spacing;
  spaceBefore: Spacing;
  spaceAfter: Spacing;
  defTab: number;
  tabs: TabStop[];
  rtl: boolean;
  fontAlign: string;
  eaLnBrk: boolean;
  latinLnBrk: boolean;
  hangingPunct: boolean;
  bullet: ResolvedBullet | null;
}

export interface ResolvedRun {
  kind: 'text' | 'break' | 'field';
  text: string;
  props: ResolvedRunProps;
  /** Where the run is in the source: paragraph index and child index among the paragraph's runs. */
  source: { p: number; r: number };
  fieldType?: string;
}

export interface ResolvedParagraph {
  index: number;
  props: ResolvedParaProps;
  runs: ResolvedRun[];
  /** Properties of the paragraph mark (a:endParaRPr), which give an empty paragraph its height. */
  endProps: ResolvedRunProps;
}

export interface ResolvedTextBody { body: ResolvedBodyProps; paragraphs: ResolvedParagraph[]; }

export interface TextSources {
  /** The shape's own text body. */
  txBody: XmlElement;
  /** List styles nearest first: the shape's, its layout placeholder's, its master placeholder's. */
  lstStyles: (XmlElement | undefined)[];
  /** Body properties nearest first, in the same order. */
  bodyPrs: (XmlElement | undefined)[];
  fontRef?: { idx: 'major' | 'minor' | 'none'; color?: ColorDef };
  /** The master's text style for this kind of shape (or a table style's text properties). */
  bucket?: XmlElement;
  defaultTextStyle?: XmlElement;
  theme: ThemeDef;
  scope: PaintScope;
  slideNumber: number;
  /** Extra run properties applied below the shape's own (a table style's tcTxStyle, for instance). */
  extraRunProps?: RunPropsDef;
}

const DEFAULT_BODY: ResolvedBodyProps = {
  insets: { l: 91440, t: 45720, r: 91440, b: 45720 }, anchor: 't', anchorCtr: false, wrap: true, vert: 'horz', rot: 0,
  upright: false, numCol: 1, spcCol: 0, rtlCol: false, autofit: { kind: 'none' }, vertOverflow: 'overflow', horzOverflow: 'overflow',
  spcFirstLastPara: false,
};

/** First defined value of each field, nearest first. */
function merge<T extends object>(chain: (T | undefined)[]): T {
  const out = {} as T;
  for (let i = chain.length - 1; i >= 0; i--) {
    const level = chain[i];
    if (!level) continue;
    for (const [k, v] of Object.entries(level)) if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

export function resolveBodyProps(bodyPrs: (XmlElement | undefined)[]): ResolvedBodyProps {
  const m = merge<BodyPropsDef>(bodyPrs.map(b => (b ? parseBodyProps(b) : undefined)));
  return {
    insets: { l: m.lIns ?? DEFAULT_BODY.insets.l, t: m.tIns ?? DEFAULT_BODY.insets.t, r: m.rIns ?? DEFAULT_BODY.insets.r, b: m.bIns ?? DEFAULT_BODY.insets.b },
    anchor: m.anchor ?? 't', anchorCtr: m.anchorCtr ?? false, wrap: (m.wrap ?? 'square') === 'square', vert: m.vert ?? 'horz',
    rot: m.rot ?? 0, upright: m.upright ?? false, numCol: m.numCol ?? 1, spcCol: m.spcCol ?? 0, rtlCol: m.rtlCol ?? false,
    autofit: m.autofit ?? { kind: 'none' }, vertOverflow: m.vertOverflow ?? 'overflow', horzOverflow: m.horzOverflow ?? 'overflow',
    spcFirstLastPara: m.spcFirstLastPara ?? false, ...(m.prstTxWarp ? { warp: m.prstTxWarp } : {}),
  };
}

/** '+mj-lt' and friends name theme fonts; anything else is a typeface. */
function themeFont(typeface: string | undefined, theme: ThemeDef): { face: string; collection?: FontCollection } {
  if (!typeface) return { face: '' };
  const m = /^\+(mj|mn)-(lt|ea|cs)$/.exec(typeface);
  if (!m) return { face: typeface };
  const collection = m[1] === 'mj' ? theme.major : theme.minor;
  return { face: m[2] === 'lt' ? collection.latin : m[2] === 'ea' ? collection.ea : collection.cs, collection };
}

const HARD_RUN: RunPropsDef = {
  sz: 1800, b: false, i: false, u: 'none', strike: 'noStrike', cap: 'none', spc: 0, kern: 0, baseline: 0, lang: 'en-US',
  latin: { typeface: '+mn-lt' }, ea: { typeface: '+mn-ea' }, cs: { typeface: '+mn-cs' },
  fill: { kind: 'solid', color: { base: { kind: 'scheme', name: 'tx1' }, transforms: [] } },
};

interface Levels { lists: ListStyleDef[]; bucket?: ListStyleDef; deflt?: ListStyleDef; }

function paraChain(levels: Levels, level: number, own: ParaPropsDef): (ParaPropsDef | undefined)[] {
  const out: (ParaPropsDef | undefined)[] = [own];
  for (const ls of levels.lists) out.push(ls.levels[level], ls.def);
  if (levels.bucket) out.push(levels.bucket.levels[level], levels.bucket.def);
  if (levels.deflt) out.push(levels.deflt.levels[level], levels.deflt.def);
  return out;
}

function runChain(levels: Levels, level: number, own: RunPropsDef, fontRef: RunPropsDef | undefined, extra: RunPropsDef | undefined): (RunPropsDef | undefined)[] {
  const out: (RunPropsDef | undefined)[] = [own];
  for (const ls of levels.lists) out.push(ls.levels[level]?.defRPr, ls.def?.defRPr);
  out.push(extra, fontRef);
  if (levels.bucket) out.push(levels.bucket.levels[level]?.defRPr, levels.bucket.def?.defRPr);
  if (levels.deflt) out.push(levels.deflt.levels[level]?.defRPr, levels.deflt.def?.defRPr);
  out.push(HARD_RUN);
  return out;
}

function resolveRunProps(m: RunPropsDef, src: TextSources): ResolvedRunProps {
  const theme = src.theme;
  const latin = themeFont(m.latin?.typeface, theme);
  const ea = themeFont(m.ea?.typeface, theme);
  const cs = themeFont(m.cs?.typeface, theme);
  const collection = latin.collection ?? ea.collection ?? theme.minor;
  let fill = resolveFill(m.fill, src.scope);
  let underline = m.u ?? 'none';
  let hyperlink: ResolvedRunProps['hyperlink'];
  if (m.hlinkClick) {
    hyperlink = {};
    if (m.hlinkClick.rId) {
      const rel = src.scope.doc.pkg.relationships(src.scope.part).find(r => r.id === m.hlinkClick!.rId);
      if (rel?.external) hyperlink.url = rel.target;
      else if (rel) {
        const slide = src.scope.doc.target(src.scope.part, rel.id);
        if (slide) hyperlink.slide = slide;
      }
    }
    if (m.hlinkClick.action) hyperlink.action = m.hlinkClick.action;
    if (m.hlinkClick.tooltip) hyperlink.tooltip = m.hlinkClick.tooltip;
    // PowerPoint draws hyperlinks in the theme's hyperlink colour, underlined
    if (hyperlink.url || (hyperlink.action && !hyperlink.action.startsWith('ppaction://noaction'))) {
      fill = { kind: 'solid', color: resolveColor({ base: { kind: 'scheme', name: 'hlink' }, transforms: [] }, src.scope.colors) };
      underline = 'sng';
    }
  }
  const out: ResolvedRunProps = {
    size: (m.sz ?? 1800) / 100, bold: m.b ?? false, italic: m.i ?? false, underline, strike: m.strike ?? 'noStrike',
    caps: m.cap ?? 'none', spacing: (m.spc ?? 0) / 100, kern: (m.kern ?? 0) / 100, baseline: m.baseline ?? 0,
    latin: latin.face, ea: ea.face, cs: cs.face, scriptFonts: collection.scripts, lang: m.lang ?? 'en-US',
    fill, outline: resolveLine(m.ln, src.scope), highlight: m.highlight ? resolveColor(m.highlight, src.scope.colors) : null,
    rtl: m.rtl ?? false,
  };
  if (m.sym) out.sym = themeFont(m.sym.typeface, theme).face;
  if (m.altLang) out.altLang = m.altLang;
  if (m.uLn && m.uLn !== 'followText') out.underlineLine = resolveLine(m.uLn, src.scope);
  if (m.uFill && m.uFill !== 'followText') out.underlineFill = resolveFill(m.uFill, src.scope);
  if (m.effects) out.effects = resolveEffects(m.effects, src.scope);
  if (hyperlink) out.hyperlink = hyperlink;
  return out;
}

function resolveBullet(m: ParaPropsDef, src: TextSources): ResolvedBullet | null {
  const b = m.bullet;
  if (!b || b.kind === 'none') return null;
  const font = !m.buFont || m.buFont === 'followText' ? null : themeFont(m.buFont.typeface, src.theme).face || null;
  const size = !m.buSz || m.buSz === 'followText' ? null : m.buSz;
  const color = !m.buClr || m.buClr === 'followText' ? null : resolveColor(m.buClr, src.scope.colors);
  if (b.kind === 'char') return { kind: 'char', char: b.char, font, size, color };
  if (b.kind === 'autonum') return { kind: 'autonum', scheme: b.scheme, startAt: b.startAt, font, size, color };
  const image: ImageRef = {};
  if (b.embed) {
    const part = src.scope.doc.target(src.scope.part, b.embed);
    if (part) image.part = part;
  }
  return { kind: 'picture', image, font, size, color };
}

/** Resolve every paragraph and run of a text body. List style parsing is cached per call through `cache`. */
export function resolveTextBody(src: TextSources, cache = new Map<XmlElement, ListStyleDef>()): ResolvedTextBody {
  const parse = (el: XmlElement | undefined) => {
    if (!el) return undefined;
    let ls = cache.get(el);
    if (!ls) {
      ls = parseListStyle(el);
      cache.set(el, ls);
    }
    return ls;
  };
  const levels: Levels = {
    lists: src.lstStyles.map(parse).filter((l): l is ListStyleDef => l !== undefined),
    ...(src.bucket ? { bucket: parse(src.bucket)! } : {}),
    ...(src.defaultTextStyle ? { deflt: parse(src.defaultTextStyle)! } : {}),
  };
  const fontRef: RunPropsDef | undefined = src.fontRef && src.fontRef.idx !== 'none'
    ? {
        latin: { typeface: src.fontRef.idx === 'major' ? '+mj-lt' : '+mn-lt' },
        ea: { typeface: src.fontRef.idx === 'major' ? '+mj-ea' : '+mn-ea' },
        cs: { typeface: src.fontRef.idx === 'major' ? '+mj-cs' : '+mn-cs' },
        ...(src.fontRef.color ? { fill: { kind: 'solid' as const, color: src.fontRef.color } } : {}),
      }
    : undefined;

  const paragraphs = paragraphsOf(src.txBody).map((p, index) => {
    const pPr = parseParaProps(child(p, 'a:pPr'));
    const level = pPr.lvl ?? 0;
    const m = merge<ParaPropsDef>(paraChain(levels, level, pPr));
    const props: ResolvedParaProps = {
      level, align: m.algn ?? 'l', marL: m.marL ?? 0, marR: m.marR ?? 0, indent: m.indent ?? 0,
      lineSpacing: m.lnSpc ?? { kind: 'pct', val: 1 }, spaceBefore: m.spcBef ?? { kind: 'pts', val: 0 }, spaceAfter: m.spcAft ?? { kind: 'pts', val: 0 },
      defTab: m.defTabSz ?? 914400, tabs: m.tabs ?? [], rtl: m.rtl ?? false, fontAlign: m.fontAlgn ?? 'auto',
      eaLnBrk: m.eaLnBrk ?? true, latinLnBrk: m.latinLnBrk ?? false, hangingPunct: m.hangingPunct ?? true,
      bullet: resolveBullet(m, src),
    };
    const runs: ResolvedRun[] = runsOf(p).map((r, i) => {
      const rp = resolveRunProps(merge<RunPropsDef>(runChain(levels, level, parseRunProps(r.rPr), fontRef, src.extraRunProps)), src);
      if (r.kind === 'br') return { kind: 'break', text: '\n', props: rp, source: { p: index, r: i } };
      if (r.kind === 'fld') {
        const text = r.fieldType === 'slidenum' && src.slideNumber > 0 ? String(src.slideNumber) : r.text;
        return { kind: 'field', text, props: rp, source: { p: index, r: i }, ...(r.fieldType ? { fieldType: r.fieldType } : {}) };
      }
      return { kind: 'text', text: r.text, props: rp, source: { p: index, r: i } };
    });
    const endProps = resolveRunProps(merge<RunPropsDef>(runChain(levels, level, parseRunProps(child(p, 'a:endParaRPr')), fontRef, src.extraRunProps)), src);
    return { index, props, runs, endProps };
  });
  return { body: resolveBodyProps(src.bodyPrs), paragraphs };
}
