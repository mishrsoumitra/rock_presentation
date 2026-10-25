// apps/web/src/ui/links.ts  -- what a click on a slide in the slide show hits: hyperlinks on text and on shapes
import type { RenderShape, RenderSlide } from '@pptx/render';
import { invert, type Matrix } from '@pptx/resolve';

export interface Link { url?: string; action?: string; slide?: string; }

const apply = (m: Matrix, x: number, y: number) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]] as const;

function inShape(s: RenderShape, x: number, y: number): Link | undefined {
  for (let i = s.children.length - 1; i >= 0; i--) {
    const hit = inShape(s.children[i]!, x, y);
    if (hit) return hit;
  }
  // a link on a run of text wins over the shape's own
  for (const p of s.prims) {
    if (p.kind !== 'text') continue;
    const [u, v] = apply(invert(p.transform), x, y);
    for (const line of p.layout.lines) {
      if (v < line.top || v > line.top + line.height) continue;
      for (const run of line.runs) {
        const link = run.props.hyperlink;
        if (link && u >= run.x && u <= run.x + run.advance) return link;
      }
    }
  }
  const link = s.source?.hyperlink;
  if (!link) return undefined;
  const [u, v] = apply(invert(s.transform), x, y);
  return u >= 0 && v >= 0 && u <= s.box.w && v <= s.box.h ? link : undefined;
}

/** The link under a point of the slide (EMU), topmost shape first. */
export function linkAt(slide: RenderSlide, x: number, y: number): Link | undefined {
  for (let i = slide.shapes.length - 1; i >= 0; i--) {
    const hit = inShape(slide.shapes[i]!, x, y);
    if (hit) return hit;
  }
  return undefined;
}

/** Only web and mail links leave the presentation; anything else (file paths, javascript:) is not followed. */
export const isSafeUrl = (url: string) => /^(https?:|mailto:)/i.test(url.trim());
