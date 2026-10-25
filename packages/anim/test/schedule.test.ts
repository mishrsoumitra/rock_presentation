// The scheduler: click steps of the main sequence, groups that start by themselves, with/after steps, delays, and
// durations stretched by repeatCount, repeatDur, autoRev, spd and iteration.
import { describe, expect, test } from 'vitest';
import { buildSchedule, clickCount, effectLength, mainSequence, parseTimeNode, type ScheduledNode } from '../src/index.ts';
import { parseXml } from '@pptx/xml';
import {
  clickGroup, clicks, effect, fadeIn, fadeOut, filterFx, interactiveSeq, mainSeq, NS, scale, setVis, step, timing, timingXml,
} from './helpers.ts';

/** Begin and end of every effect (preset container) of an activation, in activation time. */
function effects(roots: ScheduledNode[]): [number, number][] {
  const out: [number, number][] = [];
  const visit = (s: ScheduledNode, at: number) => {
    if (!Number.isFinite(s.begin)) return;
    const begin = at + s.begin;
    if (s.node.cTn.presetClass) out.push([begin, begin + s.active]);
    for (const c of s.children) visit(c, begin);
  };
  roots.forEach(r => visit(r, 0));
  return out;
}

const node = (xml: string) => parseTimeNode(parseXml(xml.replace(/^<p:(\w+)/, `<p:$1 ${NS}`)).root)!;

describe('click steps', () => {
  test('each click group waits for a click; step 0 holds what starts with the slide', () => {
    const s = buildSchedule(clicks([fadeIn(2)], [fadeIn(3)], [fadeIn(4)]));
    expect(clickCount(s)).toBe(3);
    expect(s.main[0]!.roots).toEqual([]);
    expect(s.main.slice(1).map(a => effects(a.roots))).toEqual([[[0, 500]], [[0, 500]], [[0, 500]]]);
    expect(s.main.map(a => a.duration)).toEqual([0, 500, 500, 500]);
  });

  test('a first group that starts with the main sequence (with or after previous) plays with the slide', () => {
    const t = timing(timingXml([mainSeq([clickGroup([step(0, fadeIn(2, 'afterEffect'))], true), clickGroup([step(0, fadeIn(3))])])]));
    const s = buildSchedule(t);
    expect(clickCount(s)).toBe(1);
    expect(effects(s.main[0]!.roots)).toEqual([[0, 500]]);
  });

  test('with effects start together; an after step starts at its delay, when the step before has ended', () => {
    const t = timing(timingXml([mainSeq([clickGroup([
      step(0, fadeIn(2), fadeIn(3, 'withEffect', 2000)),
      step(2000, fadeIn(4, 'afterEffect')),
    ])])]));
    expect(effects(buildSchedule(t).main[1]!.roots)).toEqual([[0, 500], [0, 2000], [2000, 2500]]);
  });

  test('an effect delay moves its start', () => {
    const t = timing(timingXml([mainSeq([clickGroup([step(0, effect('entr', 10, 0, 'clickEffect', setVis(2, 'visible') + filterFx(2, 'fade', 500), '', 750))])])]));
    expect(effects(buildSchedule(t).main[1]!.roots)).toEqual([[750, 1250]]);
  });

  test('main-sequence children without a click follow the one before, as in any seq', () => {
    const chained = '<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="300"/></p:stCondLst><p:childTnLst>' + step(0, fadeIn(3, 'afterEffect')) + '</p:childTnLst></p:cTn></p:par>';
    const t = timing(timingXml([mainSeq([clickGroup([step(0, fadeIn(2))]), chained])]));
    const s = buildSchedule(t);
    expect(clickCount(s)).toBe(1);
    expect(effects(s.main[1]!.roots)).toEqual([[0, 500], [800, 1300]]);
  });

  test('interactive sequences get their own steps, one per click on the trigger', () => {
    const t = timing(timingXml([mainSeq([clickGroup([step(0, fadeIn(2))])]), interactiveSeq(9, [
      '<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' + step(0, fadeIn(3)) + '</p:childTnLst></p:cTn></p:par>',
      clickGroup([step(0, fadeOut(3))]),
    ])]));
    const s = buildSchedule(t);
    expect(s.interactive.map(i => [i.trigger, i.steps.length])).toEqual([[9, 2]]);
    expect(effects(s.interactive[0]!.steps[1]!.roots)).toEqual([[0, 500]]);
    expect(mainSequence(t)?.cTn.children.length).toBe(1);
  });
});

describe('durations', () => {
  const one = (attrs: string, behaviours = scale(2, 1, 1000)) => {
    const s = buildSchedule(clicks([effect('emph', 6, 0, 'clickEffect', behaviours, attrs)]));
    return effects(s.main[1]!.roots)[0]!;
  };

  test('repeatCount repeats (in thousandths), repeatDur caps, autoRev doubles, spd speeds up', () => {
    expect(one('')).toEqual([0, 1000]);
    expect(one(' repeatCount="3000"')).toEqual([0, 3000]);
    expect(one(' repeatCount="2500"')).toEqual([0, 2500]);
    expect(one(' autoRev="1"')).toEqual([0, 2000]);
    expect(one(' autoRev="1" repeatCount="2000"')).toEqual([0, 4000]);
    expect(one(' repeatCount="indefinite" repeatDur="1500"')).toEqual([0, 1500]);
    expect(one(' spd="200000"')).toEqual([0, 500]);
    expect(one(' spd="-50000"')).toEqual([0, 2000]);
    expect(one(' repeatCount="indefinite"')).toEqual([0, Infinity]);
  });

  test('an end condition on the next click keeps an effect going for the rest of its step', () => {
    const fx = effect('emph', 6, 0, 'clickEffect', scale(2, 1, 1000),
      ' repeatCount="indefinite"').replace('</p:stCondLst>', '</p:stCondLst><p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>');
    expect(effects(buildSchedule(clicks([fx])).main[1]!.roots)[0]).toEqual([0, Infinity]);
  });

  test('onBegin and onEnd conditions follow other nodes, and a cycle of them does not hang', () => {
    const a = '<p:par><p:cTn id="30" dur="400"><p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="31"/></p:cond></p:stCondLst></p:cTn></p:par>';
    const b = '<p:par><p:cTn id="31" dur="400"><p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="30"/></p:cond></p:stCondLst></p:cTn></p:par>';
    const c = '<p:par><p:cTn id="32" dur="100"><p:stCondLst><p:cond delay="100"/></p:stCondLst></p:cTn></p:par>';
    const d = '<p:par><p:cTn id="33" dur="100"><p:stCondLst><p:cond evt="onEnd" delay="50"><p:tn val="32"/></p:cond></p:stCondLst></p:cTn></p:par>';
    const s = buildSchedule(timing(timingXml([mainSeq([clickGroup([step(0, a, b, c, d)])])])));
    const inner = s.main[1]!.roots[0]!.children[0]!.children;
    expect(inner.map(n => n.begin)).toEqual([Infinity, Infinity, 100, 250]);
  });

  test('an iterated effect is stretched over its letters', () => {
    const fx = effect('entr', 10, 0, 'clickEffect', setVis(2, 'visible') + filterFx(2, 'fade', 1000)).replace('</p:stCondLst>', '</p:stCondLst><p:iterate type="lt"><p:tmPct val="10000"/></p:iterate>');
    const s = buildSchedule(clicks([fx]), () => 11);
    expect(effects(s.main[1]!.roots)[0]).toEqual([0, 2000]);
  });
});

describe('effectLength', () => {
  test('a one-millisecond set is a moment: Appear lasts 0, an exit ends with its animation', () => {
    expect(effectLength(node(effect('entr', 1, 0, 'clickEffect', setVis(2, 'visible'))))).toEqual({ pass: 0, total: 0 });
    expect(effectLength(node(fadeOut(2)))).toEqual({ pass: 500, total: 500 });
    expect(effectLength(node(effect('emph', 6, 0, 'clickEffect', scale(2, 1, 800), ' repeatCount="2000"')))).toEqual({ pass: 800, total: 1600 });
  });
});
