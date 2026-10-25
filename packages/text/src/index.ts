export { FontFace, type FontMetrics, type PathOp } from './font.ts';
export { FontRegistry, SUBSTITUTES, type FontMatch } from './registry.ts';
export { fontClasses, fontClassOf, hbScriptOf, scriptOfLang, type FontClass } from './script.ts';
export {
  autonumberLabel, caretAt, EMU_PER_PT, hitTest, layoutText,
  type Autofit, type BodyStyle, type BulletStyle, type GlyphRun, type Line, type ParagraphInput, type ParaStyle, type RunInput,
  type RunStyle, type Spacing, type TextInput, type TextLayout,
} from './layout.ts';
