// packages/model/src/textprops.ts  -- DrawingML text properties as plain data (ECMA-376 Part 1, 21.1.2)
// As in drawingml.ts, absent means "inherit": the resolver merges these along the inheritance chain.
import type { XmlElement } from '@pptx/xml';
import { attr, attrNs, child, children, is } from './dom.ts';
import { colorIn, effectsIn, fillIn, parseLine, type ColorDef, type EffectsDef, type FillDef, type LineDef } from './drawingml.ts';
import { ANGLE_UNIT, bool, coord, int, percent } from './units.ts';

export interface FontDef { typeface: string; panose?: string; pitchFamily?: number; charset?: number; }
export interface HyperlinkDef { rId?: string; action?: string; tooltip?: string; }

export interface RunPropsDef {
  sz?: number;                         // hundredths of a point
  b?: boolean;
  i?: boolean;
  u?: string;                          // none, sng, dbl, heavy, dotted, ... wavyDbl
  strike?: 'noStrike' | 'sngStrike' | 'dblStrike';
  kern?: number;                       // hundredths of a point: kern at and above this size
  cap?: 'none' | 'small' | 'all';
  spc?: number;                        // hundredths of a point of extra spacing
  baseline?: number;                   // fraction: 0.3 superscript, -0.25 subscript
  lang?: string;
  altLang?: string;
  ln?: LineDef;
  fill?: FillDef;
  effects?: EffectsDef;
  highlight?: ColorDef;
  uLn?: LineDef | 'followText';
  uFill?: FillDef | 'followText';
  latin?: FontDef;
  ea?: FontDef;
  cs?: FontDef;
  sym?: FontDef;
  hlinkClick?: HyperlinkDef;
  hlinkMouseOver?: HyperlinkDef;
  rtl?: boolean;
}

export type Spacing = { kind: 'pct'; val: number } | { kind: 'pts'; val: number };   // pct: fraction; pts: hundredths of a point
export type BulletDef =
  | { kind: 'none' }
  | { kind: 'char'; char: string }
  | { kind: 'autonum'; scheme: string; startAt: number }
  | { kind: 'blip'; embed?: string };
export interface TabStop { pos: number; algn: 'l' | 'ctr' | 'r' | 'dec'; }

export interface ParaPropsDef {
  marL?: number;
  marR?: number;
  lvl?: number;
  indent?: number;
  algn?: 'l' | 'ctr' | 'r' | 'just' | 'justLow' | 'dist' | 'thaiDist';
  defTabSz?: number;
  rtl?: boolean;
  eaLnBrk?: boolean;
  latinLnBrk?: boolean;
  hangingPunct?: boolean;
  fontAlgn?: 'auto' | 't' | 'ctr' | 'base' | 'b';
  lnSpc?: Spacing;
  spcBef?: Spacing;
  spcAft?: Spacing;
  buClr?: ColorDef | 'followText';
  buSz?: { kind: 'pct'; val: number } | { kind: 'pts'; val: number } | 'followText';
  buFont?: FontDef | 'followText';
  bullet?: BulletDef;
  tabs?: TabStop[];
  defRPr?: RunPropsDef;
}

export type AutofitDef = { kind: 'none' } | { kind: 'norm'; fontScale: number; lnSpcReduction: number } | { kind: 'shape' };

export interface BodyPropsDef {
  rot?: number;                        // degrees
  spcFirstLastPara?: boolean;
  vertOverflow?: 'overflow' | 'ellipsis' | 'clip';
  horzOverflow?: 'overflow' | 'clip';
  vert?: 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl';
  wrap?: 'none' | 'square';
  lIns?: number;
  tIns?: number;
  rIns?: number;
  bIns?: number;
  numCol?: number;
  spcCol?: number;
  rtlCol?: boolean;
  fromWordArt?: boolean;
  anchor?: 't' | 'ctr' | 'b' | 'just' | 'dist';
  anchorCtr?: boolean;
  upright?: boolean;
  compatLnSpc?: boolean;
  autofit?: AutofitDef;
  prstTxWarp?: { prst: string; av: { name: string; fmla: string }[] };
}

function font(el: XmlElement | undefined): FontDef | undefined {
  if (!el) return undefined;
  const f: FontDef = { typeface: attr(el, 'typeface') ?? '' };
  const panose = attr(el, 'panose');
  const pitch = int(attr(el, 'pitchFamily'));
  const charset = int(attr(el, 'charset'));
  if (panose) f.panose = panose;
  if (pitch !== undefined) f.pitchFamily = pitch;
  if (charset !== undefined) f.charset = charset;
  return f;
}

function hyperlink(el: XmlElement | undefined): HyperlinkDef | undefined {
  if (!el) return undefined;
  const out: HyperlinkDef = {};
  const rId = attrNs(el, 'r', 'id');
  const action = attr(el, 'action');
  const tooltip = attr(el, 'tooltip');
  if (rId) out.rId = rId;
  if (action) out.action = action;
  if (tooltip) out.tooltip = tooltip;
  return out;
}

/** a:rPr, a:defRPr or a:endParaRPr (CT_TextCharacterProperties). */
export function parseRunProps(rPr: XmlElement | undefined): RunPropsDef {
  const out: RunPropsDef = {};
  if (!rPr) return out;
  const sz = int(attr(rPr, 'sz'));
  if (sz !== undefined) out.sz = sz;
  const b = bool(attr(rPr, 'b'));
  if (b !== undefined) out.b = b;
  const i = bool(attr(rPr, 'i'));
  if (i !== undefined) out.i = i;
  const u = attr(rPr, 'u');
  if (u) out.u = u;
  const strike = attr(rPr, 'strike');
  if (strike) out.strike = strike as RunPropsDef['strike'];
  const kern = int(attr(rPr, 'kern'));
  if (kern !== undefined) out.kern = kern;
  const cap = attr(rPr, 'cap');
  if (cap) out.cap = cap as RunPropsDef['cap'];
  const spc = int(attr(rPr, 'spc'));
  if (spc !== undefined) out.spc = spc;
  const baseline = percent(attr(rPr, 'baseline'));
  if (baseline !== undefined) out.baseline = baseline;
  const lang = attr(rPr, 'lang');
  if (lang) out.lang = lang;
  const altLang = attr(rPr, 'altLang');
  if (altLang) out.altLang = altLang;
  const ln = parseLine(child(rPr, 'a:ln'));
  if (ln) out.ln = ln;
  const fill = fillIn(rPr);
  if (fill) out.fill = fill;
  const effects = effectsIn(rPr);
  if (effects) out.effects = effects;
  const highlight = colorIn(child(rPr, 'a:highlight'));
  if (highlight) out.highlight = highlight;
  if (child(rPr, 'a:uLnTx')) out.uLn = 'followText';
  else {
    const uLn = parseLine(child(rPr, 'a:uLn'));
    if (uLn) out.uLn = uLn;
  }
  if (child(rPr, 'a:uFillTx')) out.uFill = 'followText';
  else {
    const uFill = fillIn(child(rPr, 'a:uFill'));
    if (uFill) out.uFill = uFill;
  }
  const latin = font(child(rPr, 'a:latin'));
  const ea = font(child(rPr, 'a:ea'));
  const cs = font(child(rPr, 'a:cs'));
  const sym = font(child(rPr, 'a:sym'));
  if (latin) out.latin = latin;
  if (ea) out.ea = ea;
  if (cs) out.cs = cs;
  if (sym) out.sym = sym;
  const click = hyperlink(child(rPr, 'a:hlinkClick'));
  if (click) out.hlinkClick = click;
  const over = hyperlink(child(rPr, 'a:hlinkMouseOver'));
  if (over) out.hlinkMouseOver = over;
  const rtl = child(rPr, 'a:rtl');
  if (rtl) out.rtl = bool(attr(rtl, 'val'), true);
  return out;
}

function spacing(el: XmlElement | undefined): Spacing | undefined {
  if (!el) return undefined;
  const pct = child(el, 'a:spcPct');
  if (pct) return { kind: 'pct', val: percent(attr(pct, 'val'), 1) };
  const pts = child(el, 'a:spcPts');
  if (pts) return { kind: 'pts', val: int(attr(pts, 'val'), 0) };
  return undefined;
}

/** a:pPr, a:defPPr or a:lvlNpPr (CT_TextParagraphProperties). */
export function parseParaProps(pPr: XmlElement | undefined): ParaPropsDef {
  const out: ParaPropsDef = {};
  if (!pPr) return out;
  const marL = coord(attr(pPr, 'marL'));
  if (marL !== undefined) out.marL = marL;
  const marR = coord(attr(pPr, 'marR'));
  if (marR !== undefined) out.marR = marR;
  const lvl = int(attr(pPr, 'lvl'));
  if (lvl !== undefined) out.lvl = Math.max(0, Math.min(8, lvl));
  const indent = coord(attr(pPr, 'indent'));
  if (indent !== undefined) out.indent = indent;
  const algn = attr(pPr, 'algn');
  if (algn) out.algn = algn as ParaPropsDef['algn'];
  const defTabSz = coord(attr(pPr, 'defTabSz'));
  if (defTabSz !== undefined) out.defTabSz = defTabSz;
  for (const k of ['rtl', 'eaLnBrk', 'latinLnBrk', 'hangingPunct'] as const) {
    const v = bool(attr(pPr, k));
    if (v !== undefined) out[k] = v;
  }
  const fontAlgn = attr(pPr, 'fontAlgn');
  if (fontAlgn) out.fontAlgn = fontAlgn as ParaPropsDef['fontAlgn'];
  const lnSpc = spacing(child(pPr, 'a:lnSpc'));
  if (lnSpc) out.lnSpc = lnSpc;
  const spcBef = spacing(child(pPr, 'a:spcBef'));
  if (spcBef) out.spcBef = spcBef;
  const spcAft = spacing(child(pPr, 'a:spcAft'));
  if (spcAft) out.spcAft = spcAft;

  if (child(pPr, 'a:buClrTx')) out.buClr = 'followText';
  else {
    const c = colorIn(child(pPr, 'a:buClr'));
    if (c) out.buClr = c;
  }
  if (child(pPr, 'a:buSzTx')) out.buSz = 'followText';
  else if (child(pPr, 'a:buSzPct')) out.buSz = { kind: 'pct', val: percent(attr(child(pPr, 'a:buSzPct'), 'val'), 1) };
  else if (child(pPr, 'a:buSzPts')) out.buSz = { kind: 'pts', val: int(attr(child(pPr, 'a:buSzPts'), 'val'), 0) };
  if (child(pPr, 'a:buFontTx')) out.buFont = 'followText';
  else {
    const f = font(child(pPr, 'a:buFont'));
    if (f) out.buFont = f;
  }
  if (child(pPr, 'a:buNone')) out.bullet = { kind: 'none' };
  else if (child(pPr, 'a:buChar')) {
    // one character is drawn: SmartArt writes char="••"
    const char = attr(child(pPr, 'a:buChar'), 'char') ?? '•';
    out.bullet = { kind: 'char', char: char ? [...char][0]! : '' };
  }
  else if (child(pPr, 'a:buAutoNum')) {
    const n = child(pPr, 'a:buAutoNum');
    out.bullet = { kind: 'autonum', scheme: attr(n, 'type') ?? 'arabicPeriod', startAt: int(attr(n, 'startAt'), 1) };
  } else if (child(pPr, 'a:buBlip')) {
    const embed = attrNs(child(child(pPr, 'a:buBlip'), 'a:blip'), 'r', 'embed');
    out.bullet = embed ? { kind: 'blip', embed } : { kind: 'blip' };
  }
  const tabLst = child(pPr, 'a:tabLst');
  if (tabLst) out.tabs = children(tabLst, 'a:tab').map(t => ({ pos: coord(attr(t, 'pos'), 0), algn: (attr(t, 'algn') ?? 'l') as TabStop['algn'] }));
  const defRPr = child(pPr, 'a:defRPr');
  if (defRPr) out.defRPr = parseRunProps(defRPr);
  return out;
}

/** a:bodyPr (CT_TextBodyProperties). */
export function parseBodyProps(bodyPr: XmlElement | undefined): BodyPropsDef {
  const out: BodyPropsDef = {};
  if (!bodyPr) return out;
  const rot = int(attr(bodyPr, 'rot'));
  if (rot !== undefined) out.rot = rot / ANGLE_UNIT;
  for (const k of ['spcFirstLastPara', 'rtlCol', 'fromWordArt', 'anchorCtr', 'upright', 'compatLnSpc'] as const) {
    const v = bool(attr(bodyPr, k));
    if (v !== undefined) out[k] = v;
  }
  const strings = ['vertOverflow', 'horzOverflow', 'vert', 'wrap', 'anchor'] as const;
  for (const k of strings) {
    const v = attr(bodyPr, k);
    if (v) (out as Record<string, unknown>)[k] = v;
  }
  for (const k of ['lIns', 'tIns', 'rIns', 'bIns', 'spcCol'] as const) {
    const v = coord(attr(bodyPr, k));
    if (v !== undefined) out[k] = v;
  }
  const numCol = int(attr(bodyPr, 'numCol'));
  if (numCol !== undefined) out.numCol = Math.max(1, Math.min(16, numCol));
  if (child(bodyPr, 'a:noAutofit')) out.autofit = { kind: 'none' };
  else if (child(bodyPr, 'a:spAutoFit')) out.autofit = { kind: 'shape' };
  else {
    const norm = child(bodyPr, 'a:normAutofit');
    if (norm) out.autofit = { kind: 'norm', fontScale: percent(attr(norm, 'fontScale'), 1), lnSpcReduction: percent(attr(norm, 'lnSpcReduction'), 0) };
  }
  const warp = child(bodyPr, 'a:prstTxWarp');
  if (warp) {
    out.prstTxWarp = {
      prst: attr(warp, 'prst') ?? 'textNoShape',
      av: children(child(warp, 'a:avLst'), 'a:gd').map(g => ({ name: attr(g, 'name') ?? '', fmla: attr(g, 'fmla') ?? '' })),
    };
  }
  return out;
}

/** a:lstStyle, p:titleStyle, p:bodyStyle, p:otherStyle, p:defaultTextStyle, a:defaultTextStyle: the default paragraph
 *  properties and one entry per outline level (index 0 is level 1). */
export interface ListStyleDef { def?: ParaPropsDef; levels: (ParaPropsDef | undefined)[]; }

export function parseListStyle(el: XmlElement | undefined): ListStyleDef {
  const out: ListStyleDef = { levels: [] };
  if (!el) return out;
  const def = child(el, 'a:defPPr');
  if (def) out.def = parseParaProps(def);
  for (let n = 1; n <= 9; n++) {
    const lvl = child(el, `a:lvl${n}pPr`);
    out.levels[n - 1] = lvl ? parseParaProps(lvl) : undefined;
  }
  return out;
}

/** True for elements that hold a paragraph's runs: a:r, a:br, a:fld. Math (a14:m) and others are not text runs. */
export const isRunElement = (el: XmlElement) => is(el, 'a:r') || is(el, 'a:br') || is(el, 'a:fld');
