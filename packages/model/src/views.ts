// packages/model/src/views.ts  -- typed views over shapes, slides and the presentation part
//
// A view wraps an element of the lossless tree and reads or writes it in place. Whatever a view does not know about
// stays in the tree untouched. Writes go through insertInOrder (schema order) and then `touch`, which marks the part
// dirty so it is re-serialized on save.
import type { XmlElement } from '@pptx/xml';
import { attr, attrNs, child, children, createElement, is, lname, path, setAttr, setAttrNs } from './dom.ts';
import { effectiveChildren } from './mc.ts';
import { nsKey, type NsKey } from './ns.ts';
import { ensureChild, insertInOrder } from './order.ts';
import { bool, coord, int } from './units.ts';

/** Namespaces whose mc:Choice branches this build reads. None yet: Office's extension branches (ink, 3D models,
 *  math, chartEx, zoom) all carry a fallback picture or shape, which is what a consumer must use otherwise. */
export const UNDERSTOOD: ReadonlySet<NsKey> = new Set();

export interface Xfrm { x: number; y: number; cx: number; cy: number; rot: number; flipH: boolean; flipV: boolean; }
export interface ChildSpace { chOff: { x: number; y: number }; chExt: { cx: number; cy: number }; }
export interface Placeholder {
  type: string;                     // title, body, ctrTitle, subTitle, dt, ftr, sldNum, obj, pic, tbl, chart, dgm, media, clipArt, sldImg, hdr
  idx: number;
  hasIdx: boolean;                  // idx was written: matching by index is only meaningful then (it defaults to 0)
  orient: 'horz' | 'vert';
  sz: 'full' | 'half' | 'quarter';
  hasCustomPrompt: boolean;
}
export type ShapeKind = 'sp' | 'pic' | 'grpSp' | 'graphicFrame' | 'cxnSp' | 'contentPart';

const NV: Record<ShapeKind, string> = {
  sp: 'nvSpPr', pic: 'nvPicPr', grpSp: 'nvGrpSpPr', graphicFrame: 'nvGraphicFramePr', cxnSp: 'nvCxnSpPr', contentPart: 'nvContentPartPr',
};

/** The kind of a shape element in a slide's tree (p:) or a SmartArt drawing (dsp:), or undefined for anything else. */
export function shapeKind(el: XmlElement): ShapeKind | undefined {
  const k = nsKey(el.ns);
  if (k !== 'p' && k !== 'dsp') return undefined;
  switch (el.local) {
    case 'sp': case 'pic': case 'grpSp': case 'graphicFrame': case 'cxnSp': case 'contentPart': return el.local;
    default: return undefined;
  }
}

/** Something that owns a part: views call `touch` after writing so the part is saved from its tree. */
export interface PartOwner { readonly part: string; touch(): void; }

function readXfrm(x: XmlElement | undefined): Xfrm | null {
  const off = child(x, 'a:off');
  const ext = child(x, 'a:ext');
  if (!x || !off || !ext) return null;
  return {
    x: coord(attr(off, 'x'), 0), y: coord(attr(off, 'y'), 0), cx: coord(attr(ext, 'cx'), 0), cy: coord(attr(ext, 'cy'), 0),
    rot: int(attr(x, 'rot'), 0), flipH: bool(attr(x, 'flipH'), false), flipV: bool(attr(x, 'flipV'), false),
  };
}

export class ShapeView {
  readonly el: XmlElement;
  readonly kind: ShapeKind;
  readonly owner: PartOwner | null;
  private readonly k: string;           // 'p' or 'dsp'

  constructor(el: XmlElement, owner: PartOwner | null = null) {
    const kind = shapeKind(el);
    if (!kind) throw new Error(`${lname(el)} is not a shape`);
    this.el = el;
    this.kind = kind;
    this.owner = owner;
    this.k = nsKey(el.ns)!;
  }

  private touch(): void {
    this.owner?.touch();
  }

  get nv(): XmlElement | undefined { return child(this.el, `${this.k}:${NV[this.kind]}`); }
  get cNvPr(): XmlElement | undefined { return child(this.nv, `${this.k}:cNvPr`); }
  get nvPr(): XmlElement | undefined { return child(this.nv, `${this.k}:nvPr`); }
  get id(): number { return int(attr(this.cNvPr, 'id'), 0); }
  get name(): string { return attr(this.cNvPr, 'name') ?? ''; }
  get descr(): string | undefined { return attr(this.cNvPr, 'descr'); }
  get hidden(): boolean { return bool(attr(this.cNvPr, 'hidden'), false); }
  get hlinkClick(): XmlElement | undefined { return child(this.cNvPr, 'a:hlinkClick'); }
  /** p:sp useBgFill="1": the shape is filled with the slide background. */
  get useBgFill(): boolean { return bool(attr(this.el, 'useBgFill'), false); }

  get placeholder(): Placeholder | null {
    const ph = child(this.nvPr, 'p:ph');
    if (!ph) return null;
    return {
      type: attr(ph, 'type') ?? 'obj', idx: int(attr(ph, 'idx'), 0), hasIdx: attr(ph, 'idx') !== undefined,
      orient: (attr(ph, 'orient') ?? 'horz') as Placeholder['orient'], sz: (attr(ph, 'sz') ?? 'full') as Placeholder['sz'],
      hasCustomPrompt: bool(attr(ph, 'hasCustomPrompt'), false),
    };
  }

  /** p:spPr, or p:grpSpPr for a group; graphic frames have none. */
  get spPr(): XmlElement | undefined {
    if (this.kind === 'graphicFrame') return undefined;
    return child(this.el, `${this.k}:${this.kind === 'grpSp' ? 'grpSpPr' : 'spPr'}`);
  }

  private xfrmElement(): XmlElement | undefined {
    return this.kind === 'graphicFrame' ? child(this.el, 'p:xfrm') : child(this.spPr, 'a:xfrm');
  }

  /** The shape's own transform; null for a placeholder that takes its position from the layout or master. */
  get xfrm(): Xfrm | null { return readXfrm(this.xfrmElement()); }

  /** A group's child coordinate space: children are placed in it and mapped onto the group's own off/ext. */
  get childSpace(): ChildSpace | null {
    if (this.kind !== 'grpSp') return null;
    const x = this.xfrmElement();
    const chOff = child(x, 'a:chOff');
    const chExt = child(x, 'a:chExt');
    if (!chOff || !chExt) return null;
    return {
      chOff: { x: coord(attr(chOff, 'x'), 0), y: coord(attr(chOff, 'y'), 0) },
      chExt: { cx: coord(attr(chExt, 'cx'), 0), cy: coord(attr(chExt, 'cy'), 0) },
    };
  }

  /** Write the transform where this kind of shape keeps it: p:spPr/a:xfrm, p:grpSpPr/a:xfrm (keeping a group's
   *  child space), or a graphic frame's own p:xfrm. A placeholder without one gets one. */
  setXfrm(v: Xfrm, childSpace?: ChildSpace): void {
    let x = this.xfrmElement();
    if (!x) {
      const container = this.kind === 'graphicFrame' ? this.el : ensureChild(this.el, `${this.k}:${this.kind === 'grpSp' ? 'grpSpPr' : 'spPr'}`);
      x = createElement(container, this.kind === 'graphicFrame' ? 'p:xfrm' : 'a:xfrm');
      insertInOrder(container, x);
    }
    setAttr(x, 'rot', v.rot ? Math.round(v.rot) : undefined);
    setAttr(x, 'flipH', v.flipH ? '1' : undefined);
    setAttr(x, 'flipV', v.flipV ? '1' : undefined);
    const off = ensureChild(x, 'a:off');
    setAttr(off, 'x', Math.round(v.x));
    setAttr(off, 'y', Math.round(v.y));
    const ext = ensureChild(x, 'a:ext');
    setAttr(ext, 'cx', Math.max(0, Math.round(v.cx)));
    setAttr(ext, 'cy', Math.max(0, Math.round(v.cy)));
    if (this.kind === 'grpSp') {
      const space = childSpace ?? this.childSpace ?? { chOff: { x: v.x, y: v.y }, chExt: { cx: v.cx, cy: v.cy } };
      const chOff = ensureChild(x, 'a:chOff');
      setAttr(chOff, 'x', Math.round(space.chOff.x));
      setAttr(chOff, 'y', Math.round(space.chOff.y));
      const chExt = ensureChild(x, 'a:chExt');
      setAttr(chExt, 'cx', Math.max(0, Math.round(space.chExt.cx)));
      setAttr(chExt, 'cy', Math.max(0, Math.round(space.chExt.cy)));
    }
    this.touch();
  }

  get geometry(): { kind: 'preset'; prst: string; av: { name: string; fmla: string }[] } | { kind: 'custom'; el: XmlElement } | undefined {
    const prst = child(this.spPr, 'a:prstGeom');
    if (prst) {
      return {
        kind: 'preset', prst: attr(prst, 'prst') ?? 'rect',
        av: children(child(prst, 'a:avLst'), 'a:gd').map(g => ({ name: attr(g, 'name') ?? '', fmla: attr(g, 'fmla') ?? '' })),
      };
    }
    const cust = child(this.spPr, 'a:custGeom');
    return cust ? { kind: 'custom', el: cust } : undefined;
  }

  get style(): XmlElement | undefined { return child(this.el, `${this.k}:style`); }
  get txBody(): XmlElement | undefined { return child(this.el, `${this.k}:txBody`); }
  /** SmartArt drawing shapes place their text in their own rectangle (dsp:txXfrm). */
  get txXfrm(): Xfrm | null { return readXfrm(child(this.el, 'dsp:txXfrm')); }
  /** Whether that rectangle says how its text is turned (a rot attribute, even "0"). */
  get txXfrmHasRot(): boolean { return attr(child(this.el, 'dsp:txXfrm'), 'rot') !== undefined; }
  get blipFill(): XmlElement | undefined { return child(this.el, `${this.k}:blipFill`); }

  get graphicData(): { uri: string; el: XmlElement } | undefined {
    const gd = path(this.el, 'a:graphic', 'a:graphicData');
    return gd ? { uri: attr(gd, 'uri') ?? '', el: gd } : undefined;
  }

  /** Video or audio attached to the shape (a:videoFile, a:audioFile, or p14:media in an extension). */
  get media(): { kind: 'video' | 'audio'; link?: string; embed?: string } | undefined {
    const nvPr = this.nvPr;
    const video = child(nvPr, 'a:videoFile');
    const audio = child(nvPr, 'a:audioFile');
    let embed: string | undefined;
    for (const ext of children(child(nvPr, 'p:extLst'), 'p:ext')) {
      const m = child(ext, 'p14:media');
      if (m) embed = attrNs(m, 'r', 'embed');
    }
    const el = video ?? audio;
    if (!el && !embed) return undefined;
    const link = attrNs(el, 'r', 'link');
    return { kind: audio ? 'audio' : 'video', ...(link ? { link } : {}), ...(embed ? { embed } : {}) };
  }

  /** A group's member shapes, with mc:AlternateContent resolved. */
  get children(): ShapeView[] {
    if (this.kind !== 'grpSp') return [];
    return shapesIn(this.el, this.owner);
  }
}

/** The shapes directly under a shape tree or group element. */
export function shapesIn(tree: XmlElement | undefined, owner: PartOwner | null, understood: ReadonlySet<NsKey> = UNDERSTOOD): ShapeView[] {
  if (!tree) return [];
  return effectiveChildren(tree, understood).filter(c => shapeKind(c) !== undefined).map(c => new ShapeView(c, owner));
}

export type SlideKind = 'slide' | 'layout' | 'master' | 'notes' | 'notesMaster' | 'handoutMaster';
const SLIDE_ROOTS: Record<string, SlideKind> = {
  'p:sld': 'slide', 'p:sldLayout': 'layout', 'p:sldMaster': 'master', 'p:notes': 'notes', 'p:notesMaster': 'notesMaster', 'p:handoutMaster': 'handoutMaster',
};

/** A slide, layout, master, notes page or notes/handout master. */
export class SlideView {
  readonly root: XmlElement;
  readonly kind: SlideKind;
  readonly owner: PartOwner;

  constructor(root: XmlElement, owner: PartOwner) {
    const kind = SLIDE_ROOTS[lname(root)];
    if (!kind) throw new Error(`${lname(root)} is not a slide-like part`);
    this.root = root;
    this.kind = kind;
    this.owner = owner;
  }

  get cSld(): XmlElement | undefined { return child(this.root, 'p:cSld'); }
  get name(): string { return attr(this.cSld, 'name') ?? ''; }
  get spTree(): XmlElement | undefined { return child(this.cSld, 'p:spTree'); }
  get shapes(): ShapeView[] { return shapesIn(this.spTree, this.owner); }
  /**
   * The preview pictures of the slide's ActiveX controls (p:controls, beside the shape tree). Office draws them
   * above every shape; the controls themselves are never run.
   */
  get controlPictures(): ShapeView[] {
    const controls = child(this.cSld, 'p:controls');
    if (!controls) return [];
    return effectiveChildren(controls, UNDERSTOOD)
      .filter(c => is(c, 'p:control'))
      .flatMap(c => shapesIn(c, this.owner).filter(s => s.kind === 'pic'));
  }
  get bg(): XmlElement | undefined { return child(this.cSld, 'p:bg'); }
  /** Slides and layouts: whether shapes from the layer below are drawn. */
  get showMasterSp(): boolean { return bool(attr(this.root, 'showMasterSp'), true); }
  /** p:sld show="0" hides a slide in the slide show. */
  get hidden(): boolean { return !bool(attr(this.root, 'show'), true); }
  /** p:sldLayout type: title, obj, twoObj, blank, ... ("cust" when absent). */
  get layoutType(): string { return attr(this.root, 'type') ?? 'cust'; }

  /** A master's colour map: scheme names used in the file (bg1, tx1, ...) to theme slots (lt1, dk1, ...). */
  get clrMap(): Record<string, string> | undefined {
    const m = child(this.root, 'p:clrMap');
    return m ? Object.fromEntries(m.attrs.filter(a => a.ns === '').map(a => [a.local, a.value])) : undefined;
  }

  /** A slide's or layout's colour map override, or undefined when it follows the master. */
  get clrMapOverride(): Record<string, string> | undefined {
    const o = child(child(this.root, 'p:clrMapOvr'), 'a:overrideClrMapping');
    return o ? Object.fromEntries(o.attrs.filter(a => a.ns === '').map(a => [a.local, a.value])) : undefined;
  }

  /** A master's text styles: title placeholders, body placeholders, everything else. */
  get txStyles(): { title?: XmlElement; body?: XmlElement; other?: XmlElement } {
    const t = child(this.root, 'p:txStyles');
    const title = child(t, 'p:titleStyle');
    const body = child(t, 'p:bodyStyle');
    const other = child(t, 'p:otherStyle');
    return { ...(title ? { title } : {}), ...(body ? { body } : {}), ...(other ? { other } : {}) };
  }

  /** Notes masters keep one list style for all notes text. */
  get notesStyle(): XmlElement | undefined { return child(this.root, 'p:notesStyle'); }
  get hf(): XmlElement | undefined { return child(this.root, 'p:hf'); }

  /** The slide transition, resolving mc:AlternateContent: PowerPoint 2010 and later put p14 transitions in a Choice. */
  transition(understood: ReadonlySet<NsKey> = UNDERSTOOD): XmlElement | undefined {
    return effectiveChildren(this.root, understood).find(c => is(c, 'p:transition'));
  }
  get timing(): XmlElement | undefined { return child(this.root, 'p:timing'); }
}

export interface SlideIdEntry { id: number; rId: string; }
export interface SectionDef { name: string; id: string; slideIds: number[]; }

const SECTIONS_EXT = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';

export class PresentationView {
  readonly root: XmlElement;
  readonly owner: PartOwner;

  constructor(root: XmlElement, owner: PartOwner) {
    if (!is(root, 'p:presentation')) throw new Error(`${lname(root)} is not p:presentation`);
    this.root = root;
    this.owner = owner;
  }

  get slideSize(): { cx: number; cy: number; type?: string } {
    const s = child(this.root, 'p:sldSz');
    const type = attr(s, 'type');
    return { cx: coord(attr(s, 'cx'), 9144000), cy: coord(attr(s, 'cy'), 6858000), ...(type ? { type } : {}) };
  }
  get notesSize(): { cx: number; cy: number } {
    const s = child(this.root, 'p:notesSz');
    return { cx: coord(attr(s, 'cx'), 6858000), cy: coord(attr(s, 'cy'), 9144000) };
  }
  get firstSlideNum(): number { return int(attr(this.root, 'firstSlideNum'), 1); }
  get defaultTextStyle(): XmlElement | undefined { return child(this.root, 'p:defaultTextStyle'); }

  get slideIds(): SlideIdEntry[] {
    return children(child(this.root, 'p:sldIdLst'), 'p:sldId').map(e => ({ id: int(attr(e, 'id'), 0), rId: attrNs(e, 'r', 'id') ?? '' }));
  }
  get masterIds(): SlideIdEntry[] {
    return children(child(this.root, 'p:sldMasterIdLst'), 'p:sldMasterId').map(e => ({ id: int(attr(e, 'id'), 0), rId: attrNs(e, 'r', 'id') ?? '' }));
  }
  get notesMasterRId(): string | undefined { return attrNs(path(this.root, 'p:notesMasterIdLst', 'p:notesMasterId'), 'r', 'id'); }
  get embeddedFonts(): { typeface: string; regular?: string; bold?: string; italic?: string; boldItalic?: string }[] {
    return children(child(this.root, 'p:embeddedFontLst'), 'p:embeddedFont').map(f => {
      const out: { typeface: string; regular?: string; bold?: string; italic?: string; boldItalic?: string } = { typeface: attr(child(f, 'p:font'), 'typeface') ?? '' };
      for (const style of ['regular', 'bold', 'italic', 'boldItalic'] as const) {
        const id = attrNs(child(f, `p:${style}`), 'r', 'id');
        if (id) out[style] = id;
      }
      return out;
    });
  }

  /** Sections (PowerPoint 2010 extension p14:sectionLst). */
  get sections(): SectionDef[] {
    for (const ext of children(child(this.root, 'p:extLst'), 'p:ext')) {
      if (attr(ext, 'uri') !== SECTIONS_EXT) continue;
      return children(child(ext, 'p14:sectionLst'), 'p14:section').map(s => ({
        name: attr(s, 'name') ?? '', id: attr(s, 'id') ?? '',
        slideIds: children(child(s, 'p14:sldIdLst'), 'p14:sldId').map(e => int(attr(e, 'id'), 0)),
      }));
    }
    return [];
  }

  /** Add a slide id entry for relationship `rId`, at `index` (default: the end). Returns the new slide id:
   *  one more than the largest in use, at least 256, as PowerPoint numbers them. */
  addSlideId(rId: string, index?: number): number {
    const lst = ensureChild(this.root, 'p:sldIdLst');
    const used = children(lst, 'p:sldId').map(e => int(attr(e, 'id'), 0));
    const id = Math.max(255, ...used) + 1;
    const el = createElement(lst, 'p:sldId', { id });
    setAttrNs(el, 'r', 'id', rId);
    const entries = children(lst, 'p:sldId');
    if (index === undefined || index >= entries.length) lst.children.push(el);
    else lst.children.splice(lst.children.indexOf(entries[index]!), 0, el);
    lst.selfClosing = false;
    this.owner.touch();
    return id;
  }

  removeSlideId(id: number): void {
    const lst = child(this.root, 'p:sldIdLst');
    if (!lst) return;
    lst.children = lst.children.filter(c => !(c.kind === 'element' && is(c, 'p:sldId') && int(attr(c, 'id'), 0) === id));
    this.owner.touch();
  }

  /** Move a slide id entry to position `to` among the entries. */
  moveSlideId(id: number, to: number): void {
    const lst = child(this.root, 'p:sldIdLst');
    if (!lst) return;
    const entries = children(lst, 'p:sldId');
    const el = entries.find(e => int(attr(e, 'id'), 0) === id);
    if (!el) return;
    const rest = entries.filter(e => e !== el);
    rest.splice(Math.max(0, Math.min(to, rest.length)), 0, el);
    const others = lst.children.filter(c => !(c.kind === 'element' && is(c, 'p:sldId')));
    lst.children = [...rest, ...others];
    this.owner.touch();
  }

  setSlideSize(cx: number, cy: number, type?: string): void {
    const s = ensureChild(this.root, 'p:sldSz');
    setAttr(s, 'cx', Math.round(cx));
    setAttr(s, 'cy', Math.round(cy));
    setAttr(s, 'type', type);
    this.owner.touch();
  }
}

