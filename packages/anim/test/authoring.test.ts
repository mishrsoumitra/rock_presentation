// Authoring: effects added, removed, moved and retimed as XML edits that keep PowerPoint's structure, schema order,
// unique ids and the build list, and that read back as what was intended.
import { describe, expect, test } from 'vitest';
import { checkOrder, children, is, type SlideView } from '@pptx/model';
import { parseXml, serializeElement, serializeXml } from '@pptx/xml';
import {
  addEffect, clicksOf, findPreset, listEffects, moveEffect, parseTiming, removeEffect, removeEffectsOf, sample, setEffectTiming, walk,
} from '../src/index.ts';
import { chartXml, picXml, renderSlide, shape, slideView, slideXml, spXml } from './helpers.ts';

const fade = findPreset('entr', 10)!;
const fadeOut = findPreset('exit', 10)!;
const fly = findPreset('entr', 2)!;
const slide = (shapes = spXml(2, ['Hello']) + spXml(3, ['One', '', 'Two', 'Three']) + picXml(4) + chartXml(5)) => slideView(slideXml(shapes));
const timingXml = (v: SlideView) => (v.timing ? serializeElement(v.timing) : '');
const summary = (v: SlideView) => listEffects(v).map(e => [e.trigger, e.name, e.spid, e.duration, e.delay, ...(e.paragraphs ? [e.paragraphs.st] : [])]);

/** Every cTn id once, numbered 1, 2, 3, ... in document order outside sub-lists; references point at real nodes. */
function expectWellFormed(v: SlideView): void {
  expect(checkOrder(v.root)).toEqual([]);
  if (!v.timing) return;
  const ids = [...walk(parseTiming(v.timing).nodes)].map(n => n.cTn.id);
  expect(ids).toEqual(ids.map((_, i) => i + 1));
  for (const tn of serializeElement(v.timing).matchAll(/<p:tn val="(\d+)"\/>/g)) expect(ids).toContain(Number(tn[1]));
}

describe('addEffect', () => {
  test('on a slide without animations, creates p:timing the way PowerPoint writes it', () => {
    const { view, touches } = slide();
    const added = addEffect(view, 2, fade);
    expect(added.map(e => [e.index, e.trigger, e.name, e.spid, e.duration])).toEqual([[0, 'click', 'Fade', 2, 500]]);
    expect(touches()).toBeGreaterThan(0);
    expect(timingXml(view).replace(/ xmlns:\w+="[^"]*"/g, '')).toBe(
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">' +
      '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' +
      '<p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par>' +
      '<p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
      '<p:childTnLst><p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl>' +
      '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>' +
      '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="500"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:animEffect>' +
      '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>' +
      '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
      '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>' +
      '<p:bldLst><p:bldP spid="2" grpId="0" animBg="1"/></p:bldLst></p:timing>');
    expectWellFormed(view);
  });

  test('click, with and after: one group per click, after steps delayed to when the step before ends', () => {
    const { view } = slide();
    addEffect(view, 2, fly);
    addEffect(view, 3, fade, { trigger: 'with', duration: 2000 });
    addEffect(view, 4, fade, { trigger: 'after', delay: 250 });
    addEffect(view, 5, fade);
    expect(summary(view)).toEqual([['click', 'Fly In', 2, 500, 0], ['with', 'Fade', 3, 2000, 0], ['after', 'Fade', 4, 500, 250], ['click', 'Fade', 5, 500, 0]]);
    const t = parseTiming(view.timing!);
    const groups = [...walk(t.nodes)].filter(n => n.kind === 'par' && n.cTn.start[0]?.delay === 'indefinite');
    expect(groups.length).toBe(2);
    expect(groups[0]!.cTn.children.map(s => s.cTn.start[0]!.delay)).toEqual([0, 2000]);
    expect(clicksOf(view)).toBe(2);
    expectWellFormed(view);
  });

  test('builds: a text shape gets bldP and a group id per effect, a chart bldGraphic, a picture nothing', () => {
    const { view } = slide();
    addEffect(view, 2, fade);
    addEffect(view, 2, fadeOut);
    addEffect(view, 4, fade);
    addEffect(view, 5, fade);
    const t = parseTiming(view.timing!);
    expect(t.builds.map(b => [b.kind, b.spid, b.grpId])).toEqual([['paragraph', 2, 0], ['paragraph', 2, 1], ['graphic', 5, 0]]);
    expect(listEffects(view).map(e => (parseTiming(view.timing!), e.el.children.length) && /grpId="(\d+)"/.exec(serializeElement(e.el))?.[1])).toEqual(['0', '1', undefined, '0']);
    expectWellFormed(view);
  });

  test('by paragraph: one effect per paragraph with text, one build with build="p"', () => {
    const { view } = slide();
    const added = addEffect(view, 3, fly, { paragraphs: 'each' });
    expect(added.map(e => [e.trigger, e.paragraphs])).toEqual([['click', { st: 0, end: 0 }], ['click', { st: 2, end: 2 }], ['click', { st: 3, end: 3 }]]);
    // as PowerPoint writes paragraph builds of a placeholder or text box: no animBg
    expect(serializeElement(view.timing!)).toContain('<p:bldLst><p:bldP spid="3" grpId="0" build="p"/></p:bldLst>');
    expect(sample(parseTiming(view.timing!), renderSlide([shape(3, 0, 0, 1e6, 1e6, { paragraphs: ['One', '', 'Two', 'Three'] })]), { clicks: 0, ms: 0 })
      .get(3)?.paragraphs?.get(2)).toEqual({ hidden: true });
    expectWellFormed(view);
  });

  test('interactive: an effect triggered by clicking a shape goes in that shape\'s sequence', () => {
    const { view } = slide();
    addEffect(view, 2, fade);
    addEffect(view, 3, fade, { triggerShape: 4 });
    addEffect(view, 3, fadeOut, { triggerShape: 4 });
    expect(listEffects(view).map(e => [e.index, e.trigger, e.triggerShape, e.spid])).toEqual([[0, 'click', undefined, 2], [1, 'click', 4, 3], [2, 'click', 4, 3]]);
    const xml = serializeElement(view.timing!);
    expect(xml).toContain('nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cond></p:stCondLst>');
    expect(xml).toContain('<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>');
    expectWellFormed(view);
  });

  test('a motion path can be given, and the duration scales the whole effect', () => {
    const { view } = slide();
    addEffect(view, 2, findPreset('path', 42)!, { path: 'M 0 0 L 0 0.5 E', duration: 3000 });
    const xml = serializeElement(view.timing!);
    expect(xml).toContain('path="M 0 0 L 0 0.5 E"');
    expect(listEffects(view)[0]!.duration).toBe(3000);
  });
});

describe('removing, moving and retiming', () => {
  test('removing the last effect removes p:timing', () => {
    const { view } = slide();
    addEffect(view, 2, fade);
    addEffect(view, 3, fade, { trigger: 'with' });
    removeEffect(view, 0);
    expect(summary(view)).toEqual([['with', 'Fade', 3, 500, 0]]);
    expectWellFormed(view);
    removeEffect(view, 0);
    expect(view.timing).toBeUndefined();
    expect(checkOrder(view.root)).toEqual([]);
    expect(() => removeEffect(view, 0)).toThrow(RangeError);
  });

  test('moving an effect keeps its trigger and regroups; a "with" moved to the front starts with the slide', () => {
    const { view } = slide();
    addEffect(view, 2, fade);
    addEffect(view, 3, fade);
    addEffect(view, 4, fade, { trigger: 'with' });
    moveEffect(view, 2, 0);
    expect(summary(view).map(s => [s[0], s[2]])).toEqual([['with', 4], ['click', 2], ['click', 3]]);
    const first = parseTiming(view.timing!).nodes[0]!.cTn.children[0]!.cTn.children[0]!;
    expect(first.cTn.start.map(c => [c.delay, c.evt, c.tn])).toEqual([['indefinite', undefined, undefined], [0, 'onBegin', 2]]);
    expect(clicksOf(view)).toBe(2);
    moveEffect(view, 0, 2);
    expect(summary(view).map(s => s[2])).toEqual([2, 3, 4]);
    expectWellFormed(view);
  });

  test('setEffectTiming: duration scales the timeline (an exit keeps hiding one millisecond before its end), delay, trigger, repeat', () => {
    const { view } = slide();
    addEffect(view, 2, fade);
    addEffect(view, 3, fadeOut);
    setEffectTiming(view, 1, { duration: 2000, delay: 100, trigger: 'after', repeat: 2 });
    expect(summary(view)).toEqual([['click', 'Fade', 2, 500, 0], ['after', 'Fade', 3, 2000, 100]]);
    const xml = serializeElement(listEffects(view)[1]!.el);
    expect(xml).toContain('<p:cTn id="10" dur="2000"/>');
    expect(xml).toContain('<p:cond delay="1999"/>');
    expect(xml).toContain('repeatCount="2000"');
    expect(listEffects(view)[1]!.repeat).toBe(2);
    // the after step starts when the fade before it ends
    const steps = parseTiming(view.timing!).nodes[0]!.cTn.children[0]!.cTn.children[0]!.cTn.children;
    expect(steps.map(s => s.cTn.start[0]!.delay)).toEqual([0, 500]);
    expectWellFormed(view);
  });

  test('removeEffectsOf removes a deleted shape\'s effects, the sequence it triggers and its builds', () => {
    const { view } = slide();
    addEffect(view, 2, fade);
    addEffect(view, 3, fade);
    addEffect(view, 2, fadeOut, { triggerShape: 3 });
    removeEffectsOf(view, 3);
    expect(listEffects(view).map(e => [e.spid, e.triggerShape])).toEqual([[2, undefined]]);
    expect(parseTiming(view.timing!).builds.map(b => b.spid)).toEqual([2]);
    expectWellFormed(view);
    removeEffectsOf(view, 2);
    expect(view.timing).toBeUndefined();
  });
});

describe('round trip', () => {
  test('what is written serializes, parses again and lists the same effects', () => {
    const { view } = slide();
    addEffect(view, 2, fly, { trigger: 'click' });
    addEffect(view, 3, fly, { paragraphs: { st: 2, end: 3 }, trigger: 'after', delay: 300 });
    addEffect(view, 5, findPreset('emph', 6)!, { trigger: 'with' });
    const doc = { declaration: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', prolog: [], root: view.root, epilog: [] };
    const again = slideView(serializeXml(doc)).view;
    const strip = (v: SlideView) => listEffects(v).map(({ el: _el, ...rest }) => rest);
    expect(strip(again)).toEqual(strip(view));
    expect(children(again.root).filter(c => is(c, 'p:timing')).length).toBe(1);
    expect(parseXml(serializeXml(doc)).root.local).toBe('sld');
  });
});
