// apps/web/src/ui/prompts.ts  -- "Click to add title": what an empty placeholder shows while editing
import { child, children, paragraphText, ShapeView } from '@pptx/model';
import { shapeTextLayout, type RenderShape, type TextPrim } from '@pptx/render';
import { matchLayoutPlaceholder, type ResolvedTextBody } from '@pptx/resolve';
import type { FontRegistry } from '@pptx/text';
import type { Editor } from '@pptx/editor-core';

/** PowerPoint's prompts by placeholder type; date, footer and slide number placeholders show none. */
const PROMPTS: Record<string, string> = {
  title: 'Click to add title', ctrTitle: 'Click to add title', subTitle: 'Click to add subtitle', body: 'Click to add text', obj: 'Click to add text',
  pic: 'Click to add picture', tbl: 'Click to add table', chart: 'Click to add chart', dgm: 'Click to add SmartArt graphic',
  media: 'Click to add media', clipArt: 'Click to add clip art',
};

const hasText = (t: ResolvedTextBody | null) => !!t?.paragraphs.some(p => p.runs.some(r => r.text.length > 0));

/**
 * The prompt for an empty slide placeholder, laid out in the placeholder's own text formatting (as PowerPoint
 * shows it), or undefined. A layout placeholder with a custom prompt (hasCustomPrompt) gives its own text.
 */
export function promptOf(ed: Editor, part: string, rs: RenderShape, fonts: FontRegistry): TextPrim | undefined {
  const s = rs.source;
  const ph = s?.placeholder;
  if (!s || !ph || rs.layer !== 'slide' || hasText(s.text) || !s.text) return undefined;
  let text = PROMPTS[ph.type];
  const layout = ed.doc.layoutOf(part);
  const own = layout ? matchLayoutPlaceholder(ph, ed.doc.slide(layout).shapes) : undefined;
  if (own?.placeholder?.hasCustomPrompt) {
    const custom = children(own.txBody, 'a:p').map(paragraphText).join('\n').trim();
    if (custom) text = custom;
  }
  if (!text) return undefined;
  const p0 = s.text.paragraphs[0];
  if (!p0) return undefined;
  const body: ResolvedTextBody = {
    ...s.text,
    paragraphs: [{ ...p0, runs: [{ kind: 'text', text, props: p0.endProps, source: { p: 0, r: 0 } }] }],
  };
  const t = shapeTextLayout(s, fonts, body);
  return t ? { kind: 'text', transform: t.transform, layout: t.layout } : undefined;
}

/** Whether a shape element is a placeholder (for the dashed outline empty ones get). */
export const isPlaceholder = (v: ShapeView) => !!child(v.nvPr, 'p:ph');
