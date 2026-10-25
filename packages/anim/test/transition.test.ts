// Slide transitions: reading every kind (PowerPoint 2010 and 2013 kinds in mc:AlternateContent included), writing
// them the way PowerPoint does, and drawing a frame of each against a recording context.
import { beforeAll, describe, expect, test } from 'vitest';
import { checkOrder, children, is } from '@pptx/model';
import { serializeElement } from '@pptx/xml';
import {
  clearTransition, drawTransition, setTransition, slideTransition, transitionName, type TransitionKind, type TransitionSpec,
} from '../src/index.ts';
import { installPath2D, recordingContext, slideView, slideXml, spXml } from './helpers.ts';

const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const withTransition = (xml: string) => slideView(slideXml(spXml(2), xml)).view;
const alternate = (choice: string, fallback = '<p:transition spd="slow"><p:fade/></p:transition>', requires = 'p14', ns = `xmlns:p14="${P14}"`) =>
  `<mc:AlternateContent xmlns:mc="${MC}" ${ns}><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;

describe('reading', () => {
  test.each([
    ['<p:transition spd="med"><p:wipe dir="r"/></p:transition>', { kind: 'wipe', dir: 'r', speed: 'med', duration: 750 }],
    ['<p:transition><p:fade thruBlk="1"/></p:transition>', { kind: 'fade', thruBlk: true, speed: 'fast', duration: 500 }],
    ['<p:transition spd="slow"><p:zoom/></p:transition>', { kind: 'zoom', dir: 'out', duration: 1000 }],
    ['<p:transition spd="slow"><p:wheel spokes="8"/></p:transition>', { kind: 'wheel', spokes: 8 }],
    ['<p:transition><p:split orient="vert"/></p:transition>', { kind: 'split', orient: 'vert', dir: 'out' }],
    ['<p:transition advClick="0" advTm="1000"><p:fade/></p:transition>', { kind: 'fade', advanceOnClick: false, advanceAfter: 1000 }],
    ['<p:transition spd="slow"/>', { kind: 'none', duration: 1000 }],
    [alternate(`<p:transition spd="med" p14:dur="700"><p:fade/></p:transition>`), { kind: 'fade', speed: 'med', duration: 700 }],
    [alternate(`<p:transition spd="slow" p14:dur="1250"><p14:vortex dir="r"/></p:transition>`), { kind: 'vortex', dir: 'r', duration: 1250 }],
    [alternate(`<p:transition spd="slow" p14:dur="2000"><p14:prism isInverted="1"/></p:transition>`), { kind: 'prism', isInverted: true }],
    [alternate(`<p:transition spd="slow" xmlns:p14="${P14}" p14:dur="1250"><p15:prstTrans prst="airplane"/></p:transition>`, undefined, 'p15', `xmlns:p15="${P15}"`), { kind: 'airplane', duration: 1250 }],
    [alternate('<p:transition spd="slow" p14:dur="1500"><p159:morph option="byObject"/></p:transition>', undefined, 'p159',
      `xmlns:p14="${P14}" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"`), { kind: 'morph', option: 'byObject', duration: 1500 }],
    // a Choice this reader does not understand falls back
    [alternate('<p:transition p14:dur="2799"><p:checker/></p:transition>', '<p:transition spd="slow"><p:checker/></p:transition>', 'p14',
      'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2007/7/12/main"'), { kind: 'checker', duration: 1000 }],
    ['<p:transition><p:fade/><p:sndAc><p:stSnd loop="1"><p:snd r:embed="rId2" name="applause.wav"/></p:stSnd></p:sndAc></p:transition>',
      { kind: 'fade', sound: { kind: 'start', embed: 'rId2', name: 'applause.wav', loop: true } }],
  ] as const)('%s', (xml, want) => {
    expect(slideTransition(withTransition(xml))).toMatchObject(want);
  });

  test('a slide without one has none', () => {
    expect(slideTransition(withTransition(''))).toBeUndefined();
  });

  test('names as the Transitions gallery shows them', () => {
    expect(transitionName({ kind: 'prism' })).toBe('Cube');
    expect(transitionName({ kind: 'prism', isInverted: true })).toBe('Box');
    expect(transitionName({ kind: 'pull' })).toBe('Uncover');
    expect(transitionName({ kind: 'warp' })).toBe('Zoom');
  });
});

describe('writing', () => {
  const written = (spec: TransitionSpec, before = '') => {
    const { view, touches } = slideView(slideXml(spXml(2), before));
    setTransition(view, spec);
    expect(touches()).toBe(1);
    expect(checkOrder(view.root)).toEqual([]);
    const holder = children(view.root).find(c => is(c, 'p:transition') || is(c, 'mc:AlternateContent'));
    return { view, xml: holder ? serializeElement(holder).replace(/ xmlns:\w+="[^"]*"/g, '') : '' };
  };

  test('an ECMA kind at its speed is a plain p:transition', () => {
    expect(written({ kind: 'wipe', dir: 'd', speed: 'med' }).xml).toBe('<p:transition spd="med"><p:wipe dir="d"/></p:transition>');
    expect(written({ kind: 'fade' }).xml).toBe('<p:transition><p:fade/></p:transition>');
    expect(written({ kind: 'push', dir: 'l' }).xml).toBe('<p:transition><p:push/></p:transition>');
  });

  test('another duration goes in p14:dur, with a Fallback for PowerPoint 2007', () => {
    const { view, xml } = written({ kind: 'fade', duration: 700 });
    expect(xml).toBe('<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="700"><p:fade/></p:transition></mc:Choice>' +
      '<mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>');
    expect(slideTransition(view)).toMatchObject({ kind: 'fade', duration: 700 });
  });

  test('PowerPoint 2010 and 2013 kinds are written with a fade for older readers', () => {
    const vortex = written({ kind: 'vortex', dir: 'r', duration: 1250 });
    expect(vortex.xml).toContain('<mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="1250"><p14:vortex dir="r"/></p:transition></mc:Choice>');
    expect(vortex.xml).toContain('<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>');
    expect(slideTransition(vortex.view)).toMatchObject({ kind: 'vortex', dir: 'r', duration: 1250 });
    const airplane = written({ kind: 'airplane', duration: 1250 });
    expect(airplane.xml).toContain('<mc:Choice Requires="p15"><p:transition spd="slow" p14:dur="1250"><p15:prstTrans prst="airplane"/></p:transition>');
    expect(slideTransition(airplane.view)).toMatchObject({ kind: 'airplane' });
    expect(slideTransition(written({ kind: 'morph', duration: 2000 }).view)).toMatchObject({ kind: 'morph', option: 'byObject', duration: 2000 });
  });

  test('replacing a transition keeps its sound and leaves one transition; clearing removes either form', () => {
    const before = '<p:transition><p:fade/><p:sndAc><p:stSnd><p:snd r:embed="rId2" name="a.wav"/></p:stSnd></p:sndAc></p:transition>';
    const { view, xml } = written({ kind: 'ripple', duration: 1400 }, before);
    expect((xml.match(/<p:snd /g) ?? []).length).toBe(2);
    expect(children(view.root).filter(c => is(c, 'p:transition') || is(c, 'mc:AlternateContent')).length).toBe(1);
    expect(slideTransition(view)).toMatchObject({ kind: 'ripple', sound: { kind: 'start', embed: 'rId2' } });
    clearTransition(view);
    expect(slideTransition(view)).toBeUndefined();
    expect(checkOrder(view.root)).toEqual([]);
  });

  test('the transition goes before p:timing', () => {
    const { view } = written({ kind: 'cut' }, '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"/></p:par></p:tnLst></p:timing>');
    expect(children(view.root).map(c => c.local)).toEqual(['cSld', 'clrMapOvr', 'transition', 'timing']);
  });
});

describe('drawing', () => {
  beforeAll(installPath2D);
  const from = { name: 'from' } as unknown as CanvasImageSource;
  const to = { name: 'to' } as unknown as CanvasImageSource;
  const KINDS: TransitionKind[] = [
    'none', 'fade', 'cut', 'push', 'wipe', 'split', 'cover', 'pull', 'randomBar', 'blinds', 'checker', 'circle', 'diamond', 'plus', 'wedge', 'wheel',
    'wheelReverse', 'zoom', 'newsflash', 'dissolve', 'strips', 'comb', 'random', 'vortex', 'ripple', 'glitter', 'honeycomb', 'flash', 'shred', 'switch',
    'flip', 'gallery', 'prism', 'doors', 'window', 'conveyor', 'pan', 'ferris', 'warp', 'flythrough', 'reveal', 'curtains', 'drape', 'wind', 'prestige',
    'fracture', 'crush', 'peelOff', 'pageCurlDouble', 'pageCurlSingle', 'airplane', 'origami', 'fallOver', 'morph', 'unknown',
  ];

  test.each(KINDS)('%s draws the old slide at 0, both at 0.5, the new one at 1', kind => {
    const frame = (p: number) => {
      const { ctx, calls } = recordingContext();
      drawTransition(ctx, from, to, { kind }, p, 1280, 720);
      return calls.filter(c => c.op === 'drawImage' || c.op === 'fillRect');
    };
    const start = frame(0);
    expect(start.map(c => c.image)).toEqual([from]);
    const end = frame(1);
    expect(end.map(c => c.image)).toEqual([to]);
    expect(end[0]!.alpha).toBe(1);
    const mid = frame(0.5);
    expect(mid.length).toBeGreaterThan(0);
    if (!['none', 'cut', 'flash'].includes(kind)) expect(mid.some(c => c.image === to)).toBe(true);
  });

  test('fade through black shows black halfway, the slides on either side', () => {
    const { ctx, calls } = recordingContext();
    drawTransition(ctx, from, to, { kind: 'fade', thruBlk: true }, 0.5, 100, 100);
    expect(calls.filter(c => c.op === 'fillRect').length).toBe(1);
    expect(calls.filter(c => c.op === 'drawImage').every(c => c.alpha === 0)).toBe(true);
  });

  test('a push moves both slides by the same amount', () => {
    const { ctx, calls } = recordingContext();
    drawTransition(ctx, from, to, { kind: 'push', dir: 'u' }, 0.25, 100, 200);
    expect(calls.filter(c => c.op === 'transform').map(c => c.args[5])).toEqual([-50, 150]);
  });

  test('a wipe clips the new slide; without Path2D it fades instead', () => {
    const { ctx, calls } = recordingContext();
    drawTransition(ctx, from, to, { kind: 'wipe', dir: 'd' }, 0.5, 100, 100);
    expect(calls.map(c => c.op)).toEqual(['drawImage', 'clip', 'drawImage']);
    const saved = (globalThis as { Path2D?: unknown }).Path2D;
    delete (globalThis as { Path2D?: unknown }).Path2D;
    try {
      const plain = recordingContext();
      drawTransition(plain.ctx, from, to, { kind: 'wipe', dir: 'd' }, 0.5, 100, 100);
      expect(plain.calls.map(c => [c.op, c.image, c.alpha])).toEqual([['drawImage', from, 1], ['drawImage', to, 0.5]]);
    } finally {
      (globalThis as { Path2D?: unknown }).Path2D = saved;
    }
  });

  test('the doors open to the left and right when vertical', () => {
    const { ctx, calls } = recordingContext();
    drawTransition(ctx, from, to, { kind: 'doors', dir: 'vert' }, 0.5, 100, 100);
    // each half is squeezed towards its own edge: x scales, y does not
    const halves = calls.filter(c => c.op === 'transform').slice(1);
    expect(halves.map(c => [c.args[0], c.args[3]])).toEqual([[0.5, 1], [0.5, 1]]);
  });
});
