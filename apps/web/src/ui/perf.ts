// apps/web/src/ui/perf.ts  -- what a keystroke costs on the main thread, for the profiler and the dev console
//
// `keystrokeStart` is called where a key is handled; `keystrokeEnd` runs after the slide canvas commits the render
// that key caused, so the measure covers the store's command, autofit, the scene rebuild, React's render and its
// commit. Konva draws the layers in the animation frame after that, which the measure leaves out (a millisecond).
// In the profiler: the "keystroke" entries of the User Timing track. In development: one console line per key.

let pending = false;

export function keystrokeStart(): void {
  pending = true;
  performance.mark('keystroke-start');
}

export function keystrokeEnd(): void {
  if (!pending) return;
  pending = false;
  const m = performance.measure('keystroke', 'keystroke-start') as PerformanceMeasure | undefined;
  if (import.meta.env.DEV && m) console.debug(`keystroke ${m.duration.toFixed(1)} ms`);
}
