// packages/resolve/src/context.ts  -- the layers, theme and colour map a slide is resolved in
import type { PresentationDocument, SlideKind, SlideView, ThemeDef } from '@pptx/model';
import type { ColorContext } from './color.ts';
import type { PaintScope } from './paint.ts';

/** PowerPoint's colour map when a master has none. */
export const DEFAULT_CLR_MAP: Readonly<Record<string, string>> = {
  bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
  accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink',
};

const EMPTY_THEME: ThemeDef = {
  name: '', colorSchemeName: '', colors: new Map(), major: { latin: '', ea: '', cs: '', scripts: {} },
  minor: { latin: '', ea: '', cs: '', scripts: {} }, fillStyles: [], lineStyles: [], effectStyles: [], bgFillStyles: [],
};

export interface Layer { part: string; view: SlideView; }

export interface SlideContext {
  doc: PresentationDocument;
  /** The part being resolved: a slide, or a layout or master shown on its own. */
  part: string;
  kind: SlideKind;
  slide?: Layer;
  layout?: Layer;
  master?: Layer;
  themePart?: string;
  theme: ThemeDef;
  clrMap: Readonly<Record<string, string>>;
  colors: ColorContext;
  /** 1-based slide number for slide number fields, counting from the presentation's firstSlideNum. */
  slideNumber: number;
}

/** Build the context for a slide, layout, master or notes page. */
export function slideContext(doc: PresentationDocument, part: string): SlideContext {
  const view = doc.slide(part);
  let slide: Layer | undefined;
  let layout: Layer | undefined;
  let master: Layer | undefined;
  if (view.kind === 'slide') {
    slide = { part, view };
    const l = doc.layoutOf(part);
    if (l) layout = { part: l, view: doc.slide(l) };
    const m = l ? doc.masterOf(l) : undefined;
    if (m) master = { part: m, view: doc.slide(m) };
  } else if (view.kind === 'layout') {
    layout = { part, view };
    const m = doc.masterOf(part);
    if (m) master = { part: m, view: doc.slide(m) };
  } else if (view.kind === 'master' || view.kind === 'notesMaster' || view.kind === 'handoutMaster') {
    master = { part, view };
  } else if (view.kind === 'notes') {
    slide = { part, view };
    const nm = doc.notesMaster();
    if (nm) master = { part: nm, view: doc.slide(nm) };
  }
  const themePart = master ? doc.themeOf(master.part) : undefined;
  const theme = themePart ? doc.theme(themePart) : EMPTY_THEME;
  // A slide's (or layout's) own override wins; otherwise the master's map applies (p:clrMapOvr/a:masterClrMapping).
  const own = view.kind === 'slide' || view.kind === 'layout' || view.kind === 'notes' ? view.clrMapOverride : undefined;
  const clrMap = own ?? master?.view.clrMap ?? DEFAULT_CLR_MAP;
  let slideNumber = 1;
  if (view.kind === 'slide') {
    const index = doc.slideParts().indexOf(part);
    slideNumber = doc.presentation.firstSlideNum + Math.max(0, index);
  }
  return {
    doc, part, kind: view.kind, ...(slide ? { slide } : {}), ...(layout ? { layout } : {}), ...(master ? { master } : {}),
    ...(themePart ? { themePart } : {}), theme, clrMap, colors: { theme: theme.colors, clrMap }, slideNumber,
  };
}

/** The paint scope for XML that lives in `part` (relationship ids resolve against that part's relationships). */
export const scopeFor = (ctx: SlideContext, part: string): PaintScope => ({ doc: ctx.doc, part, colors: ctx.colors });
/** The paint scope for the theme's style matrix. */
export const themeScope = (ctx: SlideContext): PaintScope => scopeFor(ctx, ctx.themePart ?? ctx.part);
