export { CommandLog, compound, edit, type Command, type CommandLogOptions, type EditSpec } from './history.ts';
export { Editor, type ChangeListener } from './editor.ts';
export { capture, restore, type Snapshot } from './snapshot.ts';
export {
  align, distribute, effectiveXfrm, flip, moveBy, parentGroup, parentMatrix, rotateQuarter, setTransforms, unionBounds, visualBounds,
  type Alignment, type Box, type TransformChange,
} from './transform.ts';
export { group, rename, setAltText, ungroup, zOrder, type ZOrder } from './arrange.ts';
export {
  copyShapes, deleteShapes, duplicateShapes, insertShape, insertTextBox, paragraphsXml, pasteShapes, shapeName, type Rect, type ShapeClip,
} from './shapes.ts';
export {
  bodyPlainText, deleteText, formatBody, formatParagraphs, formatRuns, layoutToModel, modelToLayout, paragraphLength, paragraphs,
  patchRunProps, replaceText, textIn, typeText,
  type BodyPatch, type BulletSpec, type ColorSpec, type ParaPatch, type RunPatch, type SpacingSpec, type TextEdit, type TextOptions, type TextPos,
  type TextRange,
} from './text.ts';
export { cropPicture, insertPicture, naturalSize, prepareImage, prepareSvg, replacePicture, retarget, type ImageData, type SvgImage } from './pictures.ts';
export { addSlide, applyLayout, deleteSlide, duplicateSlide, moveSlide, setBackground, setSlideHidden } from './slides.ts';
export { importMasters, partText, prepareTemplate } from './templates.ts';
export { lightnessClass, STANDARD_COLORS, themePalette, type PaletteColumn, type Swatch } from './palette.ts';
export { setSlideSize, type ScaleMode } from './slide-size.ts';
export { colorXml, setFill, setLine, type LineSpec } from './format.ts';
export { addAnimation, moveAnimation, removeAnimation, setAnimationTiming, setSlideTransition, setTransitionOnAll } from './animation.ts';
export { prepareForSave } from './save.ts';
export { collectOrphans, copyPart, dropUnreferenced, findMedia, freeName, hashMedia, mediaPart, referencedIds, relRefs } from './parts.ts';
export { convertFlavour, element, fragment, idOf, nextShapeId, shapeIds, shapesOfUnit, twins, unitOf, unitsIn } from './xml.ts';
