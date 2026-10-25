// Reading p:timing into plain data: every node kind, attribute, condition, target and build, and the XML untouched.
import { describe, expect, test } from 'vitest';
import { parseXml, serializeElement } from '@pptx/xml';
import {
  interactiveSequences, mainSequence, parseTiming, parseTmFilter, rootOf, walk, type AnimNode, type AnimClrNode, type AnimEffectNode,
  type AnimMotionNode, type AnimRotNode, type AnimScaleNode, type CmdNode, type MediaNode, type SetNode, type TimeNode,
} from '../src/index.ts';
import { clickGroup, effect, interactiveSeq, mainSeq, NS, step, tgt, timingXml } from './helpers.ts';

const BEHAVIOURS =
  '<p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>' + tgt(4, '<p:txEl><p:pRg st="1" end="2"/></p:txEl>') +
  '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>' +
  '<p:anim calcmode="lin" valueType="num" from="(0.5)" to="(#ppt_x)"><p:cBhvr additive="base"><p:cTn id="7" dur="750" accel="25000" decel="50000" spd="-100000" ' +
  'repeatCount="3000" autoRev="1" fill="remove" tmFilter="0,0; 0.5,0.2; 1,1"/>' + tgt(4, '<p:bg/>') +
  '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0" fmla="#ppt_x+$"><p:val><p:fltVal val="0"/></p:val></p:tav>' +
  '<p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav><p:tav><p:val><p:intVal val="3"/></p:val></p:tav></p:tavLst></p:anim>' +
  '<p:animClr clrSpc="hsl" dir="ccw"><p:cBhvr><p:cTn id="8" dur="2000" fill="hold"/>' + tgt(5) + '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr>' +
  '<p:by><p:hsl h="5400000" s="10000" l="-20000"/></p:by><p:to><a:schemeClr val="accent2"/></p:to></p:animClr>' +
  '<p:animEffect transition="out" filter="blinds(horizontal);blinds(vertical)" prLst="gradientSize: 0.1"><p:cBhvr><p:cTn id="9" dur="500"/>' + tgt(5) + '</p:cBhvr></p:animEffect>' +
  '<p:animMotion origin="layout" path="M 0 0 L 0.25 0 E" pathEditMode="relative" rAng="5400000" ptsTypes="AA"><p:cBhvr><p:cTn id="10" dur="2000" fill="hold"/>' + tgt(5) +
  '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:rCtr x="50000" y="-25000"/></p:animMotion>' +
  '<p:animRot by="21600000"><p:cBhvr><p:cTn id="11" dur="500" fill="hold"/>' + tgt(5) + '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>' +
  '<p:animScale zoomContents="1"><p:cBhvr><p:cTn id="12" dur="500" fill="hold"/>' + tgt(5, '<p:graphicEl><a:dgm id="{ABC}" bldStep="sp"/></p:graphicEl>') +
  '</p:cBhvr><p:from x="100000" y="100000"/><p:by x="150.000%" y="150.000%"/></p:animScale>' +
  '<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="13" dur="30093" fill="hold"/>' + tgt(6) + '</p:cBhvr></p:cmd>';

const XML = timingXml([
  mainSeq([clickGroup([step(0, effect('entr', 2, 4, 'clickEffect', BEHAVIOURS, ' grpId="1"', 250))]), clickGroup([step(0)], true)]),
  interactiveSeq(7, [clickGroup([step(0)])]),
  '<p:audio isNarration="1"><p:cMediaNode vol="80000" mute="1"><p:cTn id="20" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' +
  '<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn>' + tgt(6) + '</p:cMediaNode></p:audio>',
], '<p:bldLst><p:bldP spid="4" grpId="1" build="p" animBg="1" rev="1" advAuto="2000"/><p:bldGraphic spid="5" grpId="0"><p:bldAsOne/></p:bldGraphic>' +
  '<p:bldGraphic spid="9" grpId="0"><p:bldSub><a:bldChart bld="series" animBg="0"/></p:bldSub></p:bldGraphic><p:bldDgm spid="10" grpId="0" bld="lvlOne"/></p:bldLst>');

const root = parseXml(XML.replace('<p:timing>', `<p:timing ${NS}>`)).root;
const before = serializeElement(root);
const t = parseTiming(root);
const all = [...walk(t.nodes)];
const byId = (id: number) => all.find(n => n.cTn.id === id)! as TimeNode;

describe('parseTiming', () => {
  test('finds the time root, the main sequence and the interactive sequences', () => {
    expect(rootOf(t)?.cTn.nodeType).toBe('tmRoot');
    expect(mainSequence(t)?.cTn.nodeType).toBe('mainSeq');
    expect(mainSequence(t)?.next[0]).toMatchObject({ evt: 'onNext', target: { kind: 'slide' } });
    expect(interactiveSequences(t).map(s => s.trigger)).toEqual([7]);
  });

  test('reads conditions: delays, indefinite, events with time node references', () => {
    const groups = mainSequence(t)!.cTn.children;
    expect(groups[0]!.cTn.start).toMatchObject([{ delay: 'indefinite' }]);
    expect(groups[1]!.cTn.start).toMatchObject([{ delay: 'indefinite' }, { evt: 'onBegin', delay: 0, tn: 2 }]);
    const fx = groups[0]!.cTn.children[0]!.cTn.children[0]!;
    expect(fx.cTn).toMatchObject({ presetClass: 'entr', presetID: 2, presetSubtype: 4, nodeType: 'clickEffect', grpId: 1, fill: 'hold' });
    expect(fx.cTn.start[0]!.delay).toBe(250);
  });

  test('reads the common time node attributes in natural units', () => {
    const a = byId(7) as AnimNode;
    expect(a.cTn).toMatchObject({ dur: 750, accel: 0.25, decel: 0.5, spd: -1, repeatCount: 3, autoRev: true, fill: 'remove' });
    expect(a.cTn.tmFilter).toEqual([[0, 0], [0.5, 0.2], [1, 1]]);
  });

  test('reads every behaviour with its target and attribute names', () => {
    const set = byId(6) as SetNode;
    expect(set.kind).toBe('set');
    expect(set.to).toEqual({ kind: 'string', value: 'visible' });
    expect(set.bhvr.target).toEqual({ kind: 'shape', spid: 4, sub: { kind: 'text', range: { unit: 'paragraph', st: 1, end: 2 } } });
    const anim = byId(7) as AnimNode;
    expect(anim).toMatchObject({ kind: 'anim', calcmode: 'lin', valueType: 'num', from: '(0.5)', to: '(#ppt_x)' });
    expect(anim.bhvr).toMatchObject({ additive: 'base', attrNames: ['ppt_x'], target: { kind: 'shape', spid: 4, sub: { kind: 'bg' } } });
    expect(anim.tavs.map(v => [v.tm, v.fmla, v.val])).toEqual([
      [0, '#ppt_x+$', { kind: 'float', value: 0 }], [1, undefined, { kind: 'string', value: '#ppt_x' }], [undefined, undefined, { kind: 'int', value: 3 }],
    ]);
    const clr = byId(8) as AnimClrNode;
    expect(clr).toMatchObject({ kind: 'animClr', clrSpc: 'hsl', dir: 'ccw', by: { space: 'hsl', h: 90, s: 0.1, l: -0.2 } });
    expect(clr.to?.base).toEqual({ kind: 'scheme', name: 'accent2' });
    expect(byId(9) as AnimEffectNode).toMatchObject({ kind: 'animEffect', transition: 'out', filter: 'blinds(horizontal);blinds(vertical)', prLst: 'gradientSize: 0.1' });
    expect(byId(10) as AnimMotionNode).toMatchObject({ kind: 'animMotion', origin: 'layout', path: 'M 0 0 L 0.25 0 E', pathEditMode: 'relative', rAng: 90, rCtr: { x: 0.5, y: -0.25 } });
    expect(byId(11) as AnimRotNode).toMatchObject({ kind: 'animRot', by: 360 });
    const sc = byId(12) as AnimScaleNode;
    expect(sc).toMatchObject({ kind: 'animScale', zoomContents: true, from: { x: 1, y: 1 }, by: { x: 1.5, y: 1.5 } });
    expect(sc.bhvr.target).toMatchObject({ kind: 'shape', spid: 5, sub: { kind: 'graphic', diagram: { id: '{ABC}', bldStep: 'sp' } } });
    expect(byId(13) as CmdNode).toMatchObject({ kind: 'cmd', type: 'call', cmd: 'playFrom(0.0)' });
  });

  test('reads media nodes beside the sequences', () => {
    const audio = byId(20) as MediaNode;
    expect(audio).toMatchObject({ kind: 'audio', vol: 0.8, mute: true, isNarration: true, target: { kind: 'shape', spid: 6 } });
    expect(audio.cTn.end[0]).toMatchObject({ evt: 'onStopAudio', target: { kind: 'slide' } });
  });

  test('reads the build list', () => {
    expect(t.builds.map(b => ({ kind: b.kind, spid: b.spid, grpId: b.grpId }))).toEqual([
      { kind: 'paragraph', spid: 4, grpId: 1 }, { kind: 'graphic', spid: 5, grpId: 0 }, { kind: 'graphic', spid: 9, grpId: 0 }, { kind: 'diagram', spid: 10, grpId: 0 },
    ]);
    expect(t.builds[0]).toMatchObject({ build: 'p', animBg: true, rev: true, advAuto: 2000, bldLvl: 1 });
    expect(t.builds[1]).toMatchObject({ asOne: true });
    expect(t.builds[2]).toMatchObject({ asOne: false, sub: { kind: 'chart', bld: 'series', animBg: false } });
  });

  test('keeps a reference to each source element and never changes the XML', () => {
    for (const n of all) expect(['par', 'seq', 'excl', 'set', 'anim', 'animClr', 'animEffect', 'animMotion', 'animRot', 'animScale', 'cmd', 'audio', 'video']).toContain(n.el.local);
    expect(byId(9).el.local).toBe('animEffect');
    expect(byId(9).cTn.el.local).toBe('cTn');
    expect(serializeElement(root)).toBe(before);
  });

  test('skips what it does not know, and survives depth beyond any real file', () => {
    let deep = '<p:set><p:cBhvr><p:cTn id="1"/>' + tgt(1) + '</p:cBhvr></p:set>';
    for (let i = 0; i < 200; i++) deep = `<p:par><p:cTn><p:childTnLst>${deep}<p:unknown/></p:childTnLst></p:cTn></p:par>`;
    const parsed = parseTiming(parseXml(`<p:timing ${NS}><p:tnLst>${deep}</p:tnLst></p:timing>`, { maxDepth: 2000 }).root);
    expect([...walk(parsed.nodes)].length).toBeLessThan(200);
  });

  test('tmFilter pairs are parsed leniently and sorted', () => {
    expect(parseTmFilter('0, 0; .2, .5; .8, .5; 1, 0')).toEqual([[0, 0], [0.2, 0.5], [0.8, 0.5], [1, 0]]);
    expect(parseTmFilter('nonsense')).toBeUndefined();
  });
});
