export type { MissingPrim, PathPrim, PathSpec, Primitive, RenderShape, RenderSlide, TextPrim } from './tree.ts';
export { buildSlide, shapeTextLayout, textAreaOf, type BuildOptions } from './build.ts';
export { applyBlipEffects, drawPrimitive, drawShape, drawSlide, toPath2D, type Ctx, type DrawOptions, type ParagraphOverride, type ShapeOverride } from './canvas.ts';
export { BROWSER_IMAGE_TYPES, loadImages, type Decoder } from './images.ts';
export { PATTERNS, patternPixels } from './patterns.ts';
