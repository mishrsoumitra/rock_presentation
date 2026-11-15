// packages/editor-core/src/wordart.ts  -- WordArt: text with a fill, an outline and a shadow of its own
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { topSlot, type Rect } from './shapes.ts';
import { patchRunProps, type TextEffect } from './text.ts';
import { child, children } from '@pptx/model';
import { element, esc, insertAt, nextShapeId } from './xml.ts';

/** A WordArt style: what the gallery calls it, and how its text is painted. */
export interface WordArtStyle { id: string; name: string; effect: TextEffect; }

const scheme = (val: string, lumMod?: number, lumOff?: number) => ({ kind: 'scheme' as const, val, ...(lumMod !== undefined ? { lumMod } : {}), ...(lumOff !== undefined ? { lumOff } : {}) });

/** The gallery: looks in the theme's colours, so they change with the theme, as PowerPoint's do. */
export const WORDART_STYLES: WordArtStyle[] = [
  { id: 'fill-dark-shadow', name: 'Fill: Black, Text colour 1; Shadow', effect: { fill: scheme('tx1'), shadow: true } },
  { id: 'fill-accent1-shadow', name: 'Fill: Accent colour 1; Shadow', effect: { fill: scheme('accent1'), shadow: true } },
  { id: 'outline-accent2', name: 'Outline: Accent colour 2; No fill', effect: { fill: 'none', outline: { color: scheme('accent2'), width: 22225 } } },
  { id: 'fill-light-outline-accent1', name: 'Fill: White; Outline: Accent colour 1; Shadow', effect: { fill: scheme('bg1'), outline: { color: scheme('accent1'), width: 15875 }, shadow: true } },
  { id: 'gradient-accent4', name: 'Gradient fill: Accent colour 4', effect: { fill: { from: scheme('accent4', 60000, 40000), to: scheme('accent4', 75000), angle: 90 } } },
  { id: 'fill-accent5-outline-light', name: 'Fill: Accent colour 5; Outline: White; Shadow', effect: { fill: scheme('accent5'), outline: { color: scheme('bg1'), width: 9525 }, shadow: true } },
];

/**
 * Insert WordArt: a text box of large text in one of the gallery's looks, sized to its text and not wrapping, as
 * PowerPoint makes one. The look is ordinary character formatting, so it can be changed or cleared like any other.
 */
export function insertWordArt(ed: Editor, part: string, style: WordArtStyle, box: Rect, text = 'Your text here'): { command: Command; id: number } {
  const tree = ed.spTree(part);
  const id = nextShapeId(ed.doc.root(part));
  const r = Math.round;
  const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr><a:xfrm><a:off x="${r(box.x)}" y="${r(box.y)}"/><a:ext cx="${r(box.w)}" cy="${r(box.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
    + `<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="5400" b="1"/><a:t>${esc(text)}</a:t></a:r><a:endParaRPr lang="en-US" sz="5400" b="1"/></a:p></p:txBody></p:sp>`;
  const command = ed.edit({
    label: 'Insert WordArt', parts: [part], scope: () => [tree],
    apply: () => {
      const sp = element(tree, xml, ed.strict);
      insertAt(tree, topSlot(tree), sp);                                  // in the tree first: schema order is known by where an element is
      const p = child(child(sp, 'p:txBody'), 'a:p')!;
      for (const props of [...children(p, 'a:r').map(run => child(run, 'a:rPr')!), child(p, 'a:endParaRPr')!]) patchRunProps(ed, props, { effect: style.effect });
    },
  });
  return { command, id };
}
