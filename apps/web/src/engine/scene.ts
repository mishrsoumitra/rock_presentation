// apps/web/src/engine/scene.ts  -- render trees of slides, rebuilt only when a part they depend on changed
import { chartRenderer } from '@pptx/charts';
import type { Editor } from '@pptx/editor-core';
import { buildSlide, buildSlideShape, type RenderShape, type RenderSlide } from '@pptx/render';
import { resolveSlide, resolveSlideShape, type ResolvedShape, type ResolvedSlide } from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';
import type { XmlElement } from '@pptx/xml';

export interface Scene { slide: RenderSlide; resolved: ResolvedSlide | undefined; error?: string; }

/** Shapes the editor said the next changes to a part touch, and the part version those changes end at. */
interface Expected { part: string; els: Set<XmlElement>; from: number; to: number; }

/**
 * The render tree of each slide, cached under a key made of the versions of every part it is drawn from: the
 * slide, its layout, master and theme, and the presentation (slide size). A command bumps only what it touched.
 *
 * Typing rebuilds less: the editor says which shape its command touches (`expectChange`), and when the slide alone
 * changed, only that shape is resolved and built again; the other shapes keep their trees (and their identity).
 */
export class SceneCache {
  private readonly cache = new Map<string, { key: string; scene: Scene }>();
  private readonly ed: Editor;
  private readonly fonts: FontRegistry;
  private readonly chart: ReturnType<typeof chartRenderer>;
  private expected: Expected | null = null;
  /** Each slide's layout, master and theme, as of the slide's and the presentation's versions (what can change them). */
  private readonly chains = new Map<string, { at: string; parts: (string | undefined)[] }>();

  constructor(ed: Editor, fonts: FontRegistry) {
    this.ed = ed;
    this.fonts = fonts;
    this.chart = chartRenderer(fonts);
  }

  key(part: string): string {
    const d = this.ed.doc;
    // the chain is read from relationships, which are parsed on every lookup: every thumbnail asks on every change
    const at = `${this.ed.versionOf(part)}:${this.ed.versionOf(d.mainPart)}`;
    let chain = this.chains.get(part);
    if (!chain || chain.at !== at) {
      const layout = d.pkg.getPart(part) ? d.layoutOf(part) : undefined;
      const master = layout ? d.masterOf(layout) : d.masterOf(part);
      const theme = master ? d.themeOf(master) : undefined;
      chain = { at, parts: [part, layout, master, theme, d.mainPart] };
      this.chains.set(part, chain);
    }
    return chain.parts.map(p => (p ? this.ed.versionOf(p) : '-')).join(':');
  }

  /**
   * The next command on `part` changes only `el` (a shape of the slide's own tree, or something inside one): the
   * slide's other shapes are kept when the tree is next asked for. Several calls before that add up. A change the
   * editor did not announce, or one that touched the layout, master or theme, rebuilds the whole slide as before.
   */
  expectChange(part: string, el: XmlElement): void {
    const v = this.ed.versionOf(part);
    const x = this.expected;
    if (x && x.part === part && x.to === v) {
      x.els.add(el);
      x.to = v + 1;
    } else this.expected = { part, els: new Set([el]), from: v, to: v + 1 };
  }

  /** The announced change did not happen (its command failed): the next change is not it. */
  forgetExpected(part: string): void {
    if (this.expected?.part === part) this.expected = null;
  }

  get(part: string): Scene {
    const key = this.key(part);
    const hit = this.cache.get(part);
    if (hit && hit.key === key) return hit.scene;
    const scene = (hit && this.patch(part, hit.key, key, hit.scene)) ?? this.build(part);
    this.cache.set(part, { key, scene });
    return scene;
  }

  /** The slide's tree with only the announced shapes rebuilt, or undefined when the whole slide must be. */
  private patch(part: string, oldKey: string, newKey: string, prev: Scene): Scene | undefined {
    const x = this.expected;
    if (!x || x.part !== part) return undefined;
    this.expected = null;
    // every change since the cached tree was announced, and nothing but the slide part moved
    const [oldSlide, ...oldRest] = oldKey.split(':');
    const [, ...newRest] = newKey.split(':');
    if (Number(oldSlide) !== x.from || this.ed.versionOf(part) !== x.to || oldRest.join(':') !== newRest.join(':')) return undefined;
    if (!prev.resolved || prev.error) return undefined;
    const own = this.ed.doc.slide(part);
    const tops = [...own.shapes, ...own.controlPictures].filter(s => !s.hidden);
    const touched = new Set<XmlElement>();
    for (const el of x.els) {
      const top = tops.find(s => contains(s.el, el));
      if (top) touched.add(top.el);
    }
    const prevResolved = new Map<XmlElement, ResolvedShape>();
    for (const s of prev.resolved.shapes) if (s.layer === 'slide') prevResolved.set(s.view.el, s);
    const prevBuilt = new Map<XmlElement, RenderShape>();
    for (const s of prev.slide.shapes) if (s.layer === 'slide') prevBuilt.set(s.source.view.el, s);
    const images = new Set(prev.slide.images);
    const resolvedShapes = prev.resolved.shapes.filter(s => s.layer !== 'slide');
    const shapes = prev.slide.shapes.filter(s => s.layer !== 'slide');
    for (const v of tops) {
      let r = touched.has(v.el) ? undefined : prevResolved.get(v.el);
      let b = r ? prevBuilt.get(v.el) : undefined;
      if (!r || !b) {
        r = resolveSlideShape(prev.resolved, v);
        b = buildSlideShape(prev.resolved, r, { fonts: this.fonts, chart: this.chart }, images);
      }
      resolvedShapes.push(r);
      shapes.push(b);
    }
    return { slide: { ...prev.slide, shapes, images }, resolved: { ...prev.resolved, shapes: resolvedShapes } };
  }

  private build(part: string): Scene {
    this.forgetExpected(part);                        // the whole tree is fresh: what was announced is in it
    try {
      const resolved = resolveSlide(this.ed.doc, part);
      return { slide: buildSlide(resolved, { fonts: this.fonts, chart: this.chart }), resolved };
    } catch (e) {
      // a slide that cannot be read is shown empty, with the reason, rather than taking the editor down
      const size = this.ed.doc.presentation.slideSize;
      return {
        resolved: undefined, error: e instanceof Error ? e.message : String(e),
        slide: {
          part, size: { cx: size.cx, cy: size.cy }, background: { kind: 'solid', color: { r: 1, g: 1, b: 1, a: 1 } }, images: new Set(), shapes: [],
        },
      };
    }
  }

  drop(part: string): void {
    this.cache.delete(part);
    this.chains.delete(part);
  }
}

function contains(root: XmlElement, el: XmlElement): boolean {
  for (let e: XmlElement | null = el; e; e = e.parent) if (e === root) return true;
  return false;
}
