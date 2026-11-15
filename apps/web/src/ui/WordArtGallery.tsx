// apps/web/src/ui/WordArtGallery.tsx  -- the WordArt looks, each shown as a letter painted the way the text will be
import type { CSSProperties } from 'react';
import { themePalette, WORDART_STYLES, type ColorSpec, type WordArtStyle } from '@pptx/editor-core';
import { useEditor } from '../state/store.ts';

/** A look as CSS, close enough to choose by: the fill (or a gradient clipped to the letter), the outline, the shadow. */
export function previewStyle(style: WordArtStyle, hexOf: (c: ColorSpec) => string): CSSProperties {
  const e = style.effect;
  const css: CSSProperties = {};
  if (e.fill === 'none') css.color = 'transparent';
  else if ('from' in e.fill) Object.assign(css, { backgroundImage: `linear-gradient(${e.fill.angle + 90}deg, ${hexOf(e.fill.from)}, ${hexOf(e.fill.to)})`, WebkitBackgroundClip: 'text', backgroundClip: 'text', color: 'transparent' });
  else css.color = hexOf(e.fill);
  if (e.outline) css.WebkitTextStroke = `${Math.max(1, Math.round(e.outline.width / 12700))}px ${hexOf(e.outline.color)}`;
  if (e.shadow) css.filter = 'drop-shadow(1.5px 1.5px 1px rgba(0, 0, 0, 0.45))';
  return css;
}

/** The gallery: a button a look, named for a screen reader and on hover. */
export function WordArtGallery({ onPick, testid = 'wordart-gallery' }: { onPick: (style: WordArtStyle) => void; testid?: string }) {
  const s = useEditor(st => st.session);
  const slide = useEditor(st => st.slide);
  if (!s || !slide) return null;
  const palette = themePalette(s.ed, slide);
  // the theme's own colour stands for its tints and shades here: a preview, not the paint
  const hexOf = (c: ColorSpec) => (c.kind === 'srgb' ? `#${c.hex.replace(/^#/, '')}` : palette.find(p => p.slot === c.val)?.base.hex ?? '#444444');
  const hash = (h: string) => (h.startsWith('#') ? h : `#${h}`);
  return (
    <div className="wordart-gallery" data-testid={testid}>
      {WORDART_STYLES.map(w => (
        <button key={w.id} type="button" className="wordart" title={w.name} aria-label={w.name} data-style={w.id} onClick={() => onPick(w)} onMouseDown={e => e.preventDefault()}>
          <span style={previewStyle(w, c => hash(hexOf(c)))}>A</span>
        </button>
      ))}
    </div>
  );
}
