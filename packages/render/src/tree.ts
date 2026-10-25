// packages/render/src/tree.ts  -- the render tree: what to draw, with every inherited value already applied
import type { PathCommand, PathFill } from '@pptx/geometry';
import type { Matrix, Paint, ResolvedEffects, ResolvedRunProps, ResolvedShape, Stroke } from '@pptx/resolve';
import type { TextLayout } from '@pptx/text';

/** One path of a geometry in a shape's own box (EMU), and how it is painted. */
export interface PathSpec {
  commands: PathCommand[];
  fill: PathFill;
  stroke: boolean;
}

/** A filled and stroked geometry. `transform` maps the shape's box (0..w, 0..h) to slide EMU. */
export interface PathPrim {
  kind: 'path';
  transform: Matrix;
  box: { w: number; h: number };
  paths: PathSpec[];
  fill: Paint;
  stroke: Stroke | null;
  effects: ResolvedEffects;
}

/** Laid-out text. `transform` maps the text box (0..layout.width, 0..layout.height) to slide EMU. */
export interface TextPrim {
  kind: 'text';
  transform: Matrix;
  layout: TextLayout<ResolvedRunProps>;
}

/** Content that cannot be drawn (an embedded object without a preview, an unsupported graphic): a labelled box. */
export interface MissingPrim {
  kind: 'missing';
  transform: Matrix;
  box: { w: number; h: number };
  label: string;
}

export type Primitive = PathPrim | TextPrim | MissingPrim;

export interface RenderShape {
  id: number;
  name: string;
  layer: 'master' | 'layout' | 'slide';
  transform: Matrix;
  box: { w: number; h: number };
  /** The geometry, for hit-testing and selection outlines. */
  outline: PathSpec[];
  /** Primitives in drawing order (fill and line, then text), each with its own slide-space transform. */
  prims: Primitive[];
  children: RenderShape[];
  source: ResolvedShape;
}

export interface RenderSlide {
  part: string;
  size: { cx: number; cy: number };
  background: Paint;
  shapes: RenderShape[];
  /** Image parts the slide draws: load them before drawing (see `imagesOf`). */
  images: Set<string>;
}
