// apps/web/src/ui/AnimationTabs.tsx  -- the Transitions and Animations tabs, and the animation pane
import {
  listEffects, presets, slideTransition, subtypeName, transitionName, type EffectInfo, type Preset, type Trigger, type TransitionKind,
  type TransitionSpec,
} from '@pptx/anim';
import {
  addAnimation, compound, moveAnimation, removeAnimation, setAnimationTiming, setSlideTransition, setTransitionOnAll, type Command,
} from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { useEditor } from '../state/store.ts';
import { Btn, Menu, MenuItem } from './widgets.tsx';

const KINDS: TransitionKind[] = [
  'none', 'morph', 'fade', 'push', 'wipe', 'split', 'reveal', 'randomBar', 'circle', 'diamond', 'plus', 'wedge', 'wheel', 'zoom', 'cover',
  'pull', 'cut', 'dissolve', 'checker', 'blinds', 'strips', 'comb', 'newsflash', 'flash', 'vortex', 'ripple', 'honeycomb', 'glitter',
  'shred', 'switch', 'flip', 'gallery', 'doors', 'window', 'ferris', 'conveyor', 'pan', 'flythrough', 'warp', 'prism', 'curtains',
  'drape', 'wind', 'prestige', 'fracture', 'crush', 'peelOff', 'pageCurlDouble', 'pageCurlSingle', 'airplane', 'origami', 'fallOver',
];

/** Effect options per kind: label and what it writes. */
function optionsFor(kind: TransitionKind): [string, Partial<TransitionSpec>][] {
  const sides: [string, Partial<TransitionSpec>][] = [['From Left', { dir: 'l' }], ['From Right', { dir: 'r' }], ['From Top', { dir: 'u' }], ['From Bottom', { dir: 'd' }]];
  switch (kind) {
    case 'push': case 'wipe': case 'cover': case 'pull': return sides;
    case 'split': return [['Vertical Out', { orient: 'vert', dir: 'out' }], ['Vertical In', { orient: 'vert', dir: 'in' }], ['Horizontal Out', { orient: 'horz', dir: 'out' }], ['Horizontal In', { orient: 'horz', dir: 'in' }]];
    case 'zoom': return [['In', { dir: 'in' }], ['Out', { dir: 'out' }]];
    case 'randomBar': case 'blinds': case 'checker': case 'comb': return [['Vertical', { dir: 'vert' }], ['Horizontal', { dir: 'horz' }]];
    case 'wheel': return [1, 2, 3, 4, 8].map(n => [`${n} Spoke${n > 1 ? 's' : ''}`, { spokes: n }]);
    case 'fade': return [['Smoothly', { thruBlk: false }], ['Through Black', { thruBlk: true }]];
    default: return [];
  }
}

const seconds = (ms: number | undefined) => (ms === undefined || !Number.isFinite(ms) ? '' : (ms / 1000).toFixed(2));

export function TransitionsTab() {
  const s = useEditor(st => st.session);
  const part = useEditor(st => st.slide);
  useEditor(st => st.version);
  if (!s || !part) return <section className="group"><span className="hint">Open a presentation.</span></section>;
  const ed = s.ed;
  const current = slideTransition(ed.slide(part));
  const spec: TransitionSpec = current
    ? { kind: current.kind === 'unknown' ? 'fade' : current.kind, duration: current.duration, advanceOnClick: current.advanceOnClick,
      ...(current.advanceAfter !== undefined ? { advanceAfter: current.advanceAfter } : {}),
      ...(current.dir ? { dir: current.dir } : {}), ...(current.orient ? { orient: current.orient } : {}),
      ...(current.spokes ? { spokes: current.spokes } : {}), ...(current.thruBlk ? { thruBlk: true } : {}) }
    : { kind: 'none' };
  const set = (patch: Partial<TransitionSpec>, all = false) => {
    const next: TransitionSpec = { ...spec, ...patch };
    const run = useEditor.getState().run;
    run(all ? setTransitionOnAll(ed, next.kind === 'none' ? null : next) : setSlideTransition(ed, part, next.kind === 'none' ? null : next));
  };
  const options = optionsFor(spec.kind);
  return (
    <>
      <section className="group" aria-label="Transition to This Slide">
        <Menu label={`Transition: ${transitionName({ kind: spec.kind })}`} testid="transition-kind">{close => (
          <div className="kind-grid">
            {KINDS.map(k => <button key={k} type="button" className={`kind${k === spec.kind ? ' active' : ''}`} data-kind={k}
              onClick={() => { close(); set({ kind: k, ...(k === 'none' ? {} : { duration: spec.kind === 'none' ? undefined : spec.duration }), dir: undefined, orient: undefined, spokes: undefined } as Partial<TransitionSpec>); }}>
              {transitionName({ kind: k })}</button>)}
          </div>
        )}</Menu>
        {options.length > 0 && (
          <Menu label="Effect Options">{close => options.map(([label, patch]) => <MenuItem key={label} label={label} onClick={() => { close(); set(patch); }} />)}</Menu>
        )}
      </section>
      <section className="group" aria-label="Timing">
        <label className="field">Duration
          <input type="number" min="0.01" step="0.25" value={seconds(spec.kind === 'none' ? undefined : spec.duration)} disabled={spec.kind === 'none'}
            onChange={e => { const v = Number(e.target.value); if (v > 0) set({ duration: Math.round(v * 1000) }); }} aria-label="Transition duration in seconds" />
        </label>
        <label className="field"><input type="checkbox" checked={spec.advanceOnClick !== false} onChange={e => set({ advanceOnClick: e.target.checked })} />On Mouse Click</label>
        <label className="field">After
          <input type="number" min="0" step="1" value={seconds(spec.advanceAfter)} placeholder="—"
            onChange={e => set(e.target.value === '' ? { advanceAfter: undefined } : { advanceAfter: Math.round(Number(e.target.value) * 1000) } as Partial<TransitionSpec>)} aria-label="Advance after seconds" />
        </label>
        <Btn label="Apply To All" wide onClick={() => set({}, true)} testid="apply-all" />
      </section>
    </>
  );
}

const CLASSES: [Preset['presetClass'], string][] = [['entr', 'Entrance'], ['emph', 'Emphasis'], ['exit', 'Exit'], ['path', 'Motion Paths']];

/** One preset per effect (its default direction), by class, as the Animations gallery lists them. */
function gallery(): Map<string, Preset[]> {
  const out = new Map<string, Preset[]>();
  for (const [cls] of CLASSES) {
    const byId = new Map<number, Preset>();
    for (const p of presets()) {
      if (p.presetClass !== cls) continue;
      const have = byId.get(p.presetID);
      if (!have || p.presetSubtype < have.presetSubtype) byId.set(p.presetID, p);
    }
    out.set(cls, [...byId.values()].sort((a, b) => a.name.localeCompare(b.name)));
  }
  return out;
}

function selectedShape(): { part: string; view: ShapeView } | undefined {
  const st = useEditor.getState();
  if (!st.session || !st.slide || st.selection.length !== 1) return undefined;
  return { part: st.slide, view: new ShapeView(st.selection[0]!, st.session.ed.doc.owner(st.slide)) };
}

function effectsOf(part: string, spid: number): EffectInfo[] {
  const ed = useEditor.getState().session!.ed;
  return listEffects(ed.slide(part)).filter(e => e.spid === spid);
}

/** Apply a preset to the selected shape: replacing its animation (the gallery) or adding one more. */
function animate(preset: Preset, replace: boolean): void {
  const st = useEditor.getState();
  const sel = selectedShape();
  if (!st.session || !sel) { st.notify('info', 'Select one shape to animate.'); return; }
  const ed = st.session.ed;
  const old = replace ? effectsOf(sel.part, sel.view.id).map(e => e.index).sort((a, b) => b - a) : [];
  const steps: Command[] = old.map(i => removeAnimation(ed, sel.part, i));
  steps.push(addAnimation(ed, sel.part, sel.view, preset, { trigger: 'click' }));
  st.run(steps.length === 1 ? steps[0]! : compound('Animation', steps));
}

function PresetGrid({ onPick }: { onPick: (p: Preset) => void }) {
  const g = gallery();
  return (
    <div className="gallery effects" data-testid="effect-gallery">
      {CLASSES.map(([cls, title]) => (
        <div key={cls}>
          <div className="colors-title">{title}</div>
          <div className="effect-grid">
            {(g.get(cls) ?? []).map(p => <button key={`${cls}${p.presetID}`} type="button" className={`effect ${cls}`} data-effect={`${cls}:${p.presetID}`} onClick={() => onPick(p)}>{p.name}</button>)}
          </div>
        </div>
      ))}
    </div>
  );
}

export function AnimationsTab() {
  const s = useEditor(st => st.session);
  useEditor(st => st.version);
  useEditor(st => st.selection);
  const pane = useEditor(st => st.pane);
  const sel = selectedShape();
  const effects = sel ? effectsOf(sel.part, sel.view.id) : [];
  const first = effects[0];
  const run = useEditor.getState().run;
  const retime = (t: { trigger?: Trigger; duration?: number; delay?: number }) => {
    if (!s || !sel) return;
    const steps = effects.map(e => setAnimationTiming(s.ed, sel.part, e.index, t));
    if (steps.length) run(steps.length === 1 ? steps[0]! : compound('Animation Timing', steps));
  };
  const subtypes = first?.presetClass !== undefined && first.presetID !== undefined
    ? presets().filter(p => p.presetClass === first.presetClass && p.presetID === first.presetID) : [];
  return (
    <>
      <section className="group" aria-label="Animation">
        <Menu label={first ? first.name : 'Animation'} testid="animation-gallery">{close => <PresetGrid onPick={p => { close(); animate(p, true); }} />}</Menu>
        {subtypes.length > 1 && (
          <Menu label="Effect Options">{close => subtypes.map(p => (
            <MenuItem key={p.presetSubtype} label={subtypeName(p.presetClass, p.presetID, p.presetSubtype) ?? `Option ${p.presetSubtype}`} onClick={() => { close(); animate(p, true); }} />
          ))}</Menu>
        )}
        <Menu label="Add Animation">{close => <PresetGrid onPick={p => { close(); animate(p, false); }} />}</Menu>
      </section>
      <section className="group" aria-label="Timing">
        <label className="field">Start
          <select value={first?.trigger ?? 'click'} disabled={!first} onChange={e => retime({ trigger: e.target.value as Trigger })} aria-label="Start">
            <option value="click">On Click</option><option value="with">With Previous</option><option value="after">After Previous</option>
          </select>
        </label>
        <label className="field">Duration
          <input type="number" min="0.01" step="0.25" disabled={!first} value={seconds(first?.duration)} aria-label="Animation duration in seconds"
            onChange={e => { const v = Number(e.target.value); if (v > 0) retime({ duration: Math.round(v * 1000) }); }} />
        </label>
        <label className="field">Delay
          <input type="number" min="0" step="0.25" disabled={!first} value={seconds(first?.delay)} aria-label="Animation delay in seconds"
            onChange={e => { const v = Number(e.target.value); if (v >= 0) retime({ delay: Math.round(v * 1000) }); }} />
        </label>
      </section>
      <section className="group">
        <Btn icon="anim" label="Animation Pane" wide active={pane === 'animation'} onClick={() => useEditor.setState({ pane: pane === 'animation' ? 'none' : 'animation' })} testid="animation-pane-toggle" />
        <Btn icon="play" label="Preview" wide disabled={!s} onClick={() => useEditor.setState({ mode: 'present' })} />
      </section>
    </>
  );
}

const CLASS_MARK: Record<string, string> = { entr: '★', emph: '✦', exit: '✕', path: '⤳', verb: '•', mediacall: '▶' };

/** The animation pane: every effect of the slide in playing order, to select, reorder or remove. */
export function AnimationPane() {
  const s = useEditor(st => st.session);
  const part = useEditor(st => st.slide);
  useEditor(st => st.version);
  const selection = useEditor(st => st.selection);
  if (!s || !part) return null;
  const ed = s.ed;
  const effects = listEffects(ed.slide(part));
  let click = 0;
  const run = useEditor.getState().run;
  return (
    <aside className="pane" aria-label="Animation Pane" data-testid="animation-pane">
      <header><strong>Animation Pane</strong><button type="button" aria-label="Close" onClick={() => useEditor.setState({ pane: 'none' })}>×</button></header>
      {effects.length === 0 && <p className="hint">Select a shape, then pick an effect on the Animations tab.</p>}
      <ol>
        {effects.map((e, i) => {
          if (e.trigger === 'click' && e.triggerShape === undefined) click++;
          const target = e.spid !== undefined ? ed.shape(part, e.spid) : undefined;
          const selected = !!target && selection.includes(target.el);
          return (
            <li key={`${i}-${e.spid}`} className={selected ? 'selected' : ''} onClick={() => target && useEditor.getState().select([target.el])}>
              <span className="num">{e.trigger === 'click' && e.triggerShape === undefined ? click : ''}</span>
              <span className={`mark ${e.presetClass ?? ''}`}>{CLASS_MARK[e.presetClass ?? ''] ?? '•'}</span>
              <span className="what">{e.name}{target ? `: ${target.name}` : ''}{e.paragraphs ? ` ¶${e.paragraphs.st + 1}` : ''}</span>
              <span className="when">{e.trigger === 'click' ? '' : e.trigger === 'with' ? 'with' : 'after'}</span>
              <span className="tools">
                <button type="button" aria-label="Move earlier" disabled={i === 0} onClick={ev => { ev.stopPropagation(); run(moveAnimation(ed, part, i, i - 1)); }}>▲</button>
                <button type="button" aria-label="Move later" disabled={i === effects.length - 1} onClick={ev => { ev.stopPropagation(); run(moveAnimation(ed, part, i, i + 1)); }}>▼</button>
                <button type="button" aria-label="Remove" onClick={ev => { ev.stopPropagation(); run(removeAnimation(ed, part, i)); }}>✕</button>
              </span>
            </li>
          );
        })}
      </ol>
    </aside>
  );
}
