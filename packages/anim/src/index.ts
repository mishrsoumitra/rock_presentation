// packages/anim/src/index.ts  -- animations and slide transitions (plan phase 5)
export * from './timing.ts';
export { compileFormula, evaluateFormula, FormulaError, MAX_FORMULA_LENGTH, type Formula, type FormulaScope } from './formula.ts';
export { parseMotionPath, pointAt, type MotionPath } from './path.ts';
export { complement, filterReveal, supportedFilter, type Pt, type Region, type Reveal } from './filters.ts';
export {
  buildSchedule, clickCount, effectLength, waitsForClick, type Activation, type IterationCount, type Schedule, type ScheduledNode,
} from './schedule.ts';
export {
  ease, overrideLookup, sample, sampleFrame, scheduleFor, type AnimationFrame, type PlaybackPosition, type SampleOptions,
} from './sample.ts';
export {
  clearTransition, readTransition, setTransition, slideTransition, speedFor, SPEED_MS, transitionName,
  type Transition, type TransitionKind, type TransitionOptions, type TransitionSpec, type TransitionSpeed,
} from './transition.ts';
export { drawTransition, type DrawTransitionOptions, type TransitionCtx } from './draw-transition.ts';
export {
  instantiate, rescale, templateDuration, templateOf, templateXml,
  type EffectTarget, type EffectTemplate, type InstantiateOptions, type PresetSource,
} from './template.ts';
export { findPreset, presetName, presets, subtypeName, type Preset } from './presets.ts';
export {
  addEffect, clicksOf, listEffects, moveEffect, removeEffect, removeEffectsOf, setEffectTiming,
  type AddEffectOptions, type EffectInfo, type EffectTiming, type Trigger,
} from './authoring.ts';
