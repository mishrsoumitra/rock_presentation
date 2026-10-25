// apps/web/src/engine/scene.ts  -- render trees of slides, rebuilt only when a part they depend on changed
import { chartRenderer } from '@pptx/charts';
import type { Editor } from '@pptx/editor-core';
import { buildSlide, type RenderSlide } from '@pptx/render';
import { resolveSlide, type ResolvedSlide } from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';

export interface Scene { slide: RenderSlide; resolved: ResolvedSlide | undefined; error?: string; }

/**
 * The render tree of each slide, cached under a key made of the versions of every part it is drawn from: the
 * slide, its layout, master and theme, and the presentation (slide size). A command bumps only what it touched.
 */
export class SceneCache {
  private readonly cache = new Map<string, { key: string; scene: Scene }>();
  private readonly ed: Editor;
  private readonly fonts: FontRegistry;
  private readonly chart: ReturnType<typeof chartRenderer>;

  constructor(ed: Editor, fonts: FontRegistry) {
    this.ed = ed;
    this.fonts = fonts;
    this.chart = chartRenderer(fonts);
  }

  key(part: string): string {
    const d = this.ed.doc;
    const layout = d.pkg.getPart(part) ? d.layoutOf(part) : undefined;
    const master = layout ? d.masterOf(layout) : d.masterOf(part);
    const theme = master ? d.themeOf(master) : undefined;
    return [part, layout, master, theme, d.mainPart].map(p => (p ? this.ed.versionOf(p) : '-')).join(':');
  }

  get(part: string): Scene {
    const key = this.key(part);
    const hit = this.cache.get(part);
    if (hit && hit.key === key) return hit.scene;
    let scene: Scene;
    try {
      const resolved = resolveSlide(this.ed.doc, part);
      scene = { slide: buildSlide(resolved, { fonts: this.fonts, chart: this.chart }), resolved };
    } catch (e) {
      // a slide that cannot be read is shown empty, with the reason, rather than taking the editor down
      const size = this.ed.doc.presentation.slideSize;
      scene = {
        resolved: undefined, error: e instanceof Error ? e.message : String(e),
        slide: {
          part, size: { cx: size.cx, cy: size.cy }, background: { kind: 'solid', color: { r: 1, g: 1, b: 1, a: 1 } }, images: new Set(), shapes: [],
        },
      };
    }
    this.cache.set(part, { key, scene });
    return scene;
  }

  drop(part: string): void { this.cache.delete(part); }
}
