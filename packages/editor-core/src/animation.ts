// packages/editor-core/src/animation.ts  -- animations and slide transitions as undoable commands
import {
  addEffect, clearTransition, moveEffect, removeEffect, setEffectTiming, setTransition,
  type AddEffectOptions, type EffectTiming, type Preset, type TransitionSpec,
} from '@pptx/anim';
import type { ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import { compound, type Command } from './history.ts';

/** Animate a shape with a preset (entrance, emphasis, exit or motion path), as the Animations gallery does. */
export function addAnimation(ed: Editor, part: string, shape: ShapeView, preset: Preset, options: AddEffectOptions = {}): Command {
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Add Animation', parts: [part], scope: () => [root],
    apply: () => { addEffect(ed.slide(part), shape.id, preset, options); },
  });
}

/** Remove the effect at an index of the animation pane. */
export function removeAnimation(ed: Editor, part: string, index: number): Command {
  const root = ed.doc.root(part);
  return ed.edit({ label: 'Remove Animation', parts: [part], scope: () => [root], apply: () => removeEffect(ed.slide(part), index) });
}

/** Move an effect within the animation pane (its click, with or after grouping is kept consistent). */
export function moveAnimation(ed: Editor, part: string, from: number, to: number): Command {
  const root = ed.doc.root(part);
  return ed.edit({ label: 'Reorder Animation', parts: [part], scope: () => [root], apply: () => moveEffect(ed.slide(part), from, to) });
}

/** An effect's start (on click, with previous, after previous), duration, delay and repeat. */
export function setAnimationTiming(ed: Editor, part: string, index: number, timing: EffectTiming): Command {
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Animation Timing', parts: [part], scope: () => [root], coalesceKey: `anim-timing:${part}:${index}`,
    apply: () => setEffectTiming(ed.slide(part), index, timing),
  });
}

/** Set a slide's transition, or remove it (null). */
export function setSlideTransition(ed: Editor, part: string, spec: TransitionSpec | null): Command {
  const root = ed.doc.root(part);
  return ed.edit({
    label: spec ? 'Transition' : 'Remove Transition', parts: [part], scope: () => [root],
    apply: () => (spec && spec.kind !== 'none' ? setTransition(ed.slide(part), spec) : clearTransition(ed.slide(part))),
  });
}

/** Apply To All: the same transition on every slide, as one undo step. */
export function setTransitionOnAll(ed: Editor, spec: TransitionSpec | null): Command {
  return compound('Apply To All', ed.doc.slideParts().map(p => setSlideTransition(ed, p, spec)));
}
