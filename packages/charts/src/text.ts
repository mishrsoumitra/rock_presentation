// packages/charts/src/text.ts  -- chart text: properties down the chart's txPr chain, laid out as text primitives
//
// Chart text inherits differently from slide text: a title's runs take the paragraph's own a:pPr/a:defRPr, then the
// element's c:txPr, then the chart space's c:txPr, then the chart style's defaults (the chart's font size, 120% and
// bold for the chart title, bold for axis titles, tx1). Everything resolves through @pptx/resolve's text resolver,
// with that chain passed as the run properties under the text's own.
import type { XmlElement } from '@pptx/xml';
import { child, parseRunProps, path, type RunPropsDef } from '@pptx/model';
import {
  multiply, resolveTextBody, type Matrix, type ResolvedBodyProps, type ResolvedParagraph, type ResolvedParaProps,
  type ResolvedRunProps, type ResolvedTextBody,
} from '@pptx/resolve';
import type { TextPrim } from '@pptx/render';
import { layoutText, type TextLayout } from '@pptx/text';
import { textColor, type Env } from './style.ts';
import { aElement } from './xml.ts';

/** How one kind of chart text is drawn: resolved run and paragraph properties, and the txPr's rotation. */
export interface TextStyle {
  run: ResolvedRunProps;
  para: ResolvedParaProps;
  body: ResolvedBodyProps;
  /** Degrees clockwise from the txPr's a:bodyPr rot; undefined when automatic (absent, or Office's -60000000). */
  rot?: number;
  /** a:bodyPr vert: vertical text (vert, vert270, wordArtVert...). */
  vert: string;
  size: number;
}

/** A laid-out piece of text and its size (EMU). */
export interface Block { layout: TextLayout<ResolvedRunProps>; w: number; h: number; }

/** The c:txPr elements text inherits from, nearest first (the chart space's last). */
type TxPrs = (XmlElement | undefined)[];

const defRPr = (txPr: XmlElement | undefined): RunPropsDef | undefined => {
  const el = path(txPr, 'a:p', 'a:pPr', 'a:defRPr');
  return el ? parseRunProps(el) : undefined;
};

/** First defined value of each property, nearest first. */
function merge(chain: (RunPropsDef | undefined)[]): RunPropsDef {
  const out: Record<string, unknown> = {};
  for (let i = chain.length - 1; i >= 0; i--) {
    const level = chain[i];
    if (!level) continue;
    for (const [k, v] of Object.entries(level)) if (v !== undefined) out[k] = v;
  }
  return out as RunPropsDef;
}

/** The chart's font size in points: its c:txPr's, else 10 pt (ECMA-376 21.2.3.46). */
export function chartFontSize(chartTxPr: XmlElement | undefined): number {
  const sz = defRPr(chartTxPr)?.sz;
  return sz && sz > 0 ? sz / 100 : 10;
}

/**
 * Defaults under a chain of c:txPr (the chart space's last). `scaled` marks the chart title, whose size is 120% of
 * the chart's whatever the chart space's txPr says, so its size goes above that txPr rather than below it.
 */
export interface TextDefaults { size: number; bold: boolean; scaled?: boolean; }

function extraFor(env: Env, chain: TxPrs, d: TextDefaults, own?: RunPropsDef): RunPropsDef {
  const sz = Math.round(d.size * 100);
  const fallback: RunPropsDef = { sz, b: d.bold, fill: { kind: 'solid', color: textColor(env.style) } };
  const defs = chain.map(defRPr);
  if (d.scaled && defs.length) defs.splice(defs.length - 1, 0, { sz });
  return merge([own, ...defs, fallback]);
}

function rotOf(bodyPr: XmlElement | undefined): { rot?: number; vert: string } {
  const r = bodyPr?.attrs.find(a => a.local === 'rot' && a.ns === '')?.value;
  const vert = bodyPr?.attrs.find(a => a.local === 'vert' && a.ns === '')?.value ?? 'horz';
  const n = r === undefined ? NaN : Number.parseInt(r, 10);
  // -60000000 is Office's "automatic"; anything outside +-90 degrees is not a label angle
  return Number.isFinite(n) && n !== -60000000 && Math.abs(n) <= 5400000 ? { rot: n / 60000, vert } : { vert };
}

const SAMPLE = aElement('txBody', [aElement('p', [aElement('r', [aElement('t', ['0'])])])]);

/** What text resolution needs to know about where the chart is; charts have no slide number fields to fill. */
const where = (env: Env) => ({ theme: env.theme, scope: env.scope, slideNumber: 0 });

/** Rotation and direction come from the nearest a:bodyPr that states either. */
const turnOf = (bodyPrs: (XmlElement | undefined)[]) =>
  rotOf(bodyPrs.find(b => b?.attrs.some(a => a.local === 'rot' || a.local === 'vert')));

/** The style of text the chart writes itself (tick labels, legend entries, data labels) under a chain of c:txPr. */
export function textStyle(env: Env, chain: TxPrs, d: TextDefaults): TextStyle {
  const bodyPrs = chain.map(t => child(t, 'a:bodyPr'));
  const extraRunProps = extraFor(env, chain, d);
  const r = resolveTextBody({ txBody: SAMPLE, lstStyles: [], bodyPrs, ...where(env), extraRunProps });
  const p = r.paragraphs[0]!;
  return { run: p.runs[0]!.props, para: p.props, body: r.body, ...turnOf(bodyPrs), size: p.runs[0]!.props.size };
}

/**
 * Body properties for chart text: the chart places and turns its text itself, so the txPr's insets, anchoring,
 * columns and rotation do not apply inside the laid-out block.
 */
const bare = (b: ResolvedBodyProps): ResolvedBodyProps => ({
  ...b, insets: { l: 0, t: 0, r: 0, b: 0 }, wrap: false, anchor: 't', anchorCtr: false, rot: 0, vert: 'horz', numCol: 1,
  autofit: { kind: 'none' },
});

/** A text body of plain lines in one style ('\n' separates paragraphs). */
export function plainBody(style: TextStyle, text: string, align = 'ctr'): ResolvedTextBody {
  const paragraphs: ResolvedParagraph[] = text.split('\n').map((line, index) => ({
    index, props: { ...style.para, align, bullet: null }, endProps: style.run,
    runs: line ? [{ kind: 'text' as const, text: line, props: style.run, source: { p: index, r: 0 } }] : [],
  }));
  return { body: bare(style.body), paragraphs };
}

/**
 * Rich text (c:rich) as a text body: each paragraph's own a:pPr/a:defRPr sits under its runs' properties, then the
 * chain. Paragraphs without an alignment are centred, as Office centres titles and labels. Fields (a:fld) take their
 * text from `field` when it knows the field type ([VALUE], [CELLRANGE]...).
 */
export function richBody(
  env: Env, rich: XmlElement, chain: TxPrs, d: TextDefaults, field?: (type: string) => string | undefined,
): { body: ResolvedTextBody; rot?: number; vert: string } {
  const paragraphs: ResolvedParagraph[] = [];
  const bodyPrs = [child(rich, 'a:bodyPr'), ...chain.map(t => child(t, 'a:bodyPr'))];
  const lstStyles = [child(rich, 'a:lstStyle')];
  let body: ResolvedBodyProps | undefined;
  for (const p of rich.children) {
    if (p.kind !== 'element' || p.local !== 'p') continue;
    const pDef = path(p, 'a:pPr', 'a:defRPr');
    const extraRunProps = extraFor(env, chain, d, pDef ? parseRunProps(pDef) : undefined);
    const txBody = aElement('txBody', [p]);
    const r = resolveTextBody({ txBody, lstStyles, bodyPrs, ...where(env), extraRunProps });
    body ??= r.body;
    const explicit = path(p, 'a:pPr')?.attrs.some(a => a.local === 'algn');
    for (const para of r.paragraphs) {
      const runs = para.runs.map(run => {
        if (run.kind !== 'field' || !field || !run.fieldType) return run;
        const text = field(run.fieldType);
        return text === undefined ? run : { ...run, text };
      });
      const props = { ...para.props, align: explicit ? para.props.align : 'ctr', bullet: null };
      paragraphs.push({ ...para, index: paragraphs.length, runs, props });
    }
  }
  const b = body ?? textStyle(env, chain, d).body;
  return { body: { body: bare(b), paragraphs }, ...turnOf(bodyPrs) };
}

/** Lay out a text body; `maxW` wraps it at that width (EMU). */
export function block(env: Env, body: ResolvedTextBody, maxW?: number): Block {
  const wrap = maxW !== undefined && maxW > 0;
  const input = { ...body, body: { ...bare(body.body), wrap } };
  const left = { ...input, paragraphs: input.paragraphs.map(p => ({ ...p, props: { ...p.props, align: 'l' } })) };
  const measured = layoutText(left, { w: wrap ? maxW : 1e10, h: 1e10 }, env.fonts);
  const w = Math.max(0, ...measured.lines.map(l => l.left + l.width));
  const h = measured.contentHeight;
  const layout = layoutText(input, { w: w + 1, h: Math.max(h, 1) }, env.fonts);
  return { layout, w, h };
}

/** True when a body has no visible text. */
export const isEmpty = (body: ResolvedTextBody) => !body.paragraphs.some(p => p.runs.some(r => r.text.trim().length));

const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
const rotate = (deg: number): Matrix => {
  const r = (deg * Math.PI) / 180;
  return [Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0];
};

/** The transform that puts a block's point (px, py) at (x, y), turned `deg` degrees clockwise about that point. */
export function placeAt(x: number, y: number, px: number, py: number, deg = 0): Matrix {
  const m = deg ? multiply(translate(x, y), rotate(deg)) : translate(x, y);
  return multiply(m, translate(-px, -py));
}

export const textPrim = (b: Block, transform: Matrix): TextPrim => ({ kind: 'text', transform, layout: b.layout });

/** The axis-aligned box a block covers once turned by `deg` (its extent across and along the axis). */
export function rotatedSize(b: { w: number; h: number }, deg: number): { w: number; h: number } {
  const r = (deg * Math.PI) / 180;
  const c = Math.abs(Math.cos(r));
  const s = Math.abs(Math.sin(r));
  return { w: b.w * c + b.h * s, h: b.w * s + b.h * c };
}
