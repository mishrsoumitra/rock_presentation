// Timings, slides and render trees for the animation tests, written the way PowerPoint writes them.
import { parseXml, type XmlElement } from '@pptx/xml';
import { SlideView } from '@pptx/model';
import type { PathSpec, RenderShape, RenderSlide } from '@pptx/render';
import type { Matrix, Paint, ResolvedShape } from '@pptx/resolve';
import { parseTiming, type Timing } from '../src/index.ts';

export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;
export const SLIDE_PART = '/ppt/slides/slide1.xml';

// ------------------------------------------------------------------------------------------------- timing XML

export const tgt = (spid: number, sub = '') => `<p:tgtEl><p:spTgt spid="${spid}"${sub ? `>${sub}</p:spTgt>` : '/>'}</p:tgtEl>`;
export const para = (st: number, end = st) => `<p:txEl><p:pRg st="${st}" end="${end}"/></p:txEl>`;
const ctn = (dur: number | string, extra = '', delay?: number) =>
  `<p:cTn dur="${dur}"${extra}${delay === undefined ? '/>' : `><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>`}`;

export const setVis = (spid: number, to: 'visible' | 'hidden', delay = 0, sub = '') =>
  `<p:set><p:cBhvr>${ctn(1, ' fill="hold"', delay)}${tgt(spid, sub)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:to><p:strVal val="${to}"/></p:to></p:set>`;
export const filterFx = (spid: number, filter: string, dur: number, transition = 'in', sub = '') =>
  `<p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr>${ctn(dur)}${tgt(spid, sub)}</p:cBhvr></p:animEffect>`;
export const anim = (spid: number, attrName: string, from: string, to: string, dur: number, extra = '', sub = '', additive = ' additive="base"') =>
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr${additive}>${ctn(dur, ` fill="hold"${extra}`)}${tgt(spid, sub)}` +
  `<p:attrNameLst><p:attrName>${attrName}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>` +
  `<p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
export const motion = (spid: number, path: string, dur: number, extra = '') =>
  `<p:animMotion origin="layout" path="${path}" pathEditMode="relative"><p:cBhvr>${ctn(dur, ` fill="hold"${extra}`)}${tgt(spid)}` +
  '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>';
export const scale = (spid: number, by: number, dur: number, extra = ' fill="hold"') =>
  `<p:animScale><p:cBhvr>${ctn(dur, extra)}${tgt(spid)}</p:cBhvr><p:by x="${by * 100000}" y="${by * 100000}"/></p:animScale>`;

/** An effect par: a preset container with its trigger and delay. */
export const effect = (cls: string, id: number, sub: number, nodeType: string, behaviors: string, attrs = '', delay = 0) =>
  `<p:par><p:cTn presetID="${id}" presetClass="${cls}" presetSubtype="${sub}" fill="hold" nodeType="${nodeType}"${attrs}>` +
  `<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${behaviors}</p:childTnLst></p:cTn></p:par>`;
export const fadeIn = (spid: number, nodeType = 'clickEffect', dur = 500, sub = '') =>
  effect('entr', 10, 0, nodeType, setVis(spid, 'visible', 0, sub) + filterFx(spid, 'fade', dur, 'in', sub));
export const fadeOut = (spid: number, nodeType = 'clickEffect', dur = 500) =>
  effect('exit', 10, 0, nodeType, filterFx(spid, 'fade', dur, 'out') + setVis(spid, 'hidden', dur - 1));

/** A step of a click group ("with" effects together; "after" steps start when the step before ends). */
export const step = (delay: number, ...effects: string[]) =>
  `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${effects.join('')}</p:childTnLst></p:cTn></p:par>`;
/** A click group; `auto` starts it with the main sequence instead of on a click. */
export const clickGroup = (steps: string[], auto = false) =>
  `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="indefinite"/>${auto ? '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : ''}` +
  `</p:stCondLst><p:childTnLst>${steps.join('')}</p:childTnLst></p:cTn></p:par>`;
const slideCond = (evt: string) => `<p:cond evt="${evt}" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>`;
export const mainSeq = (groups: string[]) =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groups.join('')}</p:childTnLst></p:cTn>` +
  `<p:prevCondLst>${slideCond('onPrev')}</p:prevCondLst><p:nextCondLst>${slideCond('onNext')}</p:nextCondLst></p:seq>`;
export const interactiveSeq = (trigger: number, groups: string[]) =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
  `<p:stCondLst><p:cond evt="onClick" delay="0">${tgt(trigger)}</p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>` +
  `<p:childTnLst>${groups.join('')}</p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0">${tgt(trigger)}</p:cond></p:nextCondLst></p:seq>`;
export const timingXml = (sequences: string[], bldLst = '') =>
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${sequences.join('')}` +
  `</p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst}</p:timing>`;

/** Parse a p:timing written without namespace declarations. */
export function timing(xml: string): Timing {
  return parseTiming(parseXml(xml.replace(/^<p:timing>/, `<p:timing ${NS}>`)).root);
}

/** The main sequence with one click group per entry; each entry is the effects of its first step. */
export const clicks = (...groups: string[][]) => timing(timingXml([mainSeq(groups.map(g => clickGroup([step(0, ...g)])))]));

// -------------------------------------------------------------------------------------------------- slide XML

const xfrm = (x: number, y: number, cx: number, cy: number) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;

/** A shape with a fill and optional text paragraphs. */
export const spXml = (id: number, paragraphs: string[] = [], options: { fill?: boolean; placeholder?: boolean } = {}) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr>${options.placeholder ? '<p:ph idx="1"/>' : ''}</p:nvPr></p:nvSpPr>` +
  `<p:spPr>${xfrm(1000000 * id, 1000000, 2000000, 1000000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  `${options.fill === false ? '' : '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>'}</p:spPr>` +
  (paragraphs.length ? `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs.map(t => `<a:p><a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p>`).join('')}</p:txBody>` : '') +
  '</p:sp>';
export const picXml = (id: number) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr>${xfrm(500000, 500000, 1000000, 1000000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
export const chartXml = (id: number) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId8"/></a:graphicData></a:graphic></p:graphicFrame>';

export const slideXml = (shapes: string, extra = '') =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${NS}><p:cSld><p:spTree>` +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  `${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${extra}</p:sld>`;

/** A slide view over XML, counting how often its owner is touched. */
export function slideView(xml: string): { view: SlideView; touches: () => number } {
  let n = 0;
  const view = new SlideView(parseXml(xml).root, { part: SLIDE_PART, touch: () => { n++; } });
  return { view, touches: () => n };
}

// ----------------------------------------------------------------------------------------------- render trees

export const SLIDE = { cx: 9144000, cy: 6858000 };

const rectSpec = (w: number, h: number): PathSpec => ({
  commands: [{ op: 'M', x: 0, y: 0 }, { op: 'L', x: w, y: 0 }, { op: 'L', x: w, y: h }, { op: 'L', x: 0, y: h }, { op: 'Z' }], fill: 'norm', stroke: false,
});

export interface FakeShapeOptions {
  paragraphs?: string[];
  flipV?: boolean;
  fill?: Paint;
  layer?: 'master' | 'layout' | 'slide';
  part?: string;
  modelId?: string;
  diagram?: boolean;
  children?: RenderShape[];
}

/** A render shape with a box and a transform, as the sampler sees it (no primitives to draw). */
export function shape(id: number, x: number, y: number, w: number, h: number, o: FakeShapeOptions = {}): RenderShape {
  const transform: Matrix = o.flipV ? [1, 0, 0, -1, x, y + h] : [1, 0, 0, 1, x, y];
  const el = parseXml(`<p:sp ${NS}${o.modelId ? ` modelId="${o.modelId}"` : ''}/>`).root as XmlElement;
  const props = { fill: { kind: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } } };
  const source = {
    id, name: `Shape ${id}`, kind: 'sp', layer: o.layer ?? 'slide', part: o.part ?? SLIDE_PART, view: { el }, placeholder: null,
    xfrm: { x, y, cx: w, cy: h, rot: 0, flipH: false, flipV: !!o.flipV }, transform, textTransform: transform, geometry: null,
    fill: o.fill ?? { kind: 'solid', color: { r: 0.2, g: 0.4, b: 0.8, a: 1 } }, line: null, effects: {},
    text: o.paragraphs ? { body: {}, paragraphs: o.paragraphs.map((t, index) => ({ index, props: {}, runs: [{ kind: 'text', text: t, props, source: { p: index, r: 0 } }], endProps: props })) } : null,
    children: [], ...(o.diagram ? { graphic: { kind: 'diagram' } } : {}),
  } as unknown as ResolvedShape;
  return { id, name: source.name, layer: o.layer ?? 'slide', transform, box: { w, h }, outline: [rectSpec(w, h)], prims: [], children: o.children ?? [], source };
}

export function renderSlide(shapes: RenderShape[]): RenderSlide {
  return { part: SLIDE_PART, size: SLIDE, background: { kind: 'none' }, shapes, images: new Set() };
}

// ------------------------------------------------------------------------------------------- canvas stand-ins

/** A Path2D that records what is drawn into it (Node has none). */
export class FakePath {
  readonly ops: string[] = [];
  moveTo(x: number, y: number): void { this.ops.push(`M${Math.round(x)},${Math.round(y)}`); }
  lineTo(x: number, y: number): void { this.ops.push(`L${Math.round(x)},${Math.round(y)}`); }
  closePath(): void { this.ops.push('Z'); }
  /** The points of the path, for area and bounds checks. */
  points(): [number, number][] {
    return this.ops.filter(o => o !== 'Z').map(o => o.slice(1).split(',').map(Number) as [number, number]);
  }
}
export const fakePath = () => new FakePath() as unknown as Path2D;

/** Install FakePath as the global Path2D for code that looks it up (drawTransition). */
export function installPath2D(): void {
  (globalThis as { Path2D?: unknown }).Path2D ??= FakePath;
}

export interface Recorded { op: string; image?: unknown; alpha: number; args: number[]; }

/** A 2D context that records drawImage, fillRect, clip and transform calls, with the alpha in effect. */
export function recordingContext(): { ctx: import('../src/index.ts').TransitionCtx; calls: Recorded[] } {
  const calls: Recorded[] = [];
  const stack: number[] = [];
  const ctx = {
    globalAlpha: 1,
    fillStyle: '#000' as string | CanvasGradient | CanvasPattern,
    save() { stack.push(this.globalAlpha); },
    restore() { this.globalAlpha = stack.pop() ?? 1; },
    transform(...args: number[]) { calls.push({ op: 'transform', alpha: this.globalAlpha, args }); },
    clip(path: Path2D) { calls.push({ op: 'clip', alpha: this.globalAlpha, args: [(path as unknown as FakePath).ops.length] }); },
    drawImage(image: CanvasImageSource, ...args: number[]) { calls.push({ op: 'drawImage', image, alpha: this.globalAlpha, args }); },
    fillRect(...args: number[]) { calls.push({ op: 'fillRect', alpha: this.globalAlpha, args }); },
  };
  return { ctx, calls };
}
