// The sampler: what each shape looks like at a playback position, as renderer overrides.
import { describe, expect, test } from 'vitest';
import type { ShapeOverride } from '@pptx/render';
import { ease, overrideLookup, sample, sampleFrame, type PlaybackPosition } from '../src/index.ts';
import {
  anim, clickGroup, clicks, effect, fadeIn, fadeOut, FakePath, fakePath, filterFx, interactiveSeq, mainSeq, motion, para, renderSlide, scale,
  setVis, shape, SLIDE, step, tgt, timing, timingXml,
} from './helpers.ts';

const W = SLIDE.cx;
const H = SLIDE.cy;
// shape 2: 2 000 000 × 1 000 000 EMU at (1 000 000, 1 000 000), centre (2 000 000, 1 500 000)
const slide = renderSlide([shape(2, 1e6, 1e6, 2e6, 1e6), shape(3, 4e6, 1e6, 2e6, 1e6), shape(4, 1e6, 4e6, 2e6, 1e6),
  shape(5, 4e6, 4e6, 2e6, 2e6, { paragraphs: ['one', 'two words', 'three'] })]);
const at = (clicks: number, ms: number, triggers?: PlaybackPosition['triggers']): PlaybackPosition => ({ clicks, ms, ...(triggers ? { triggers } : {}) });
const close = (m: readonly number[] | undefined, want: number[], digits = 3) => {
  expect(m).toBeDefined();
  m!.forEach((v, i) => expect(v / (i >= 4 ? 1000 : 1)).toBeCloseTo(want[i]! / (i >= 4 ? 1000 : 1), digits));
};

describe('initial state and click steps', () => {
  const t = clicks([fadeIn(2)], [fadeOut(3)], [effect('emph', 6, 0, 'clickEffect', scale(4, 2, 1000))]);

  test('before anything plays, an entrance target is hidden; exit and emphasis targets are as on the slide', () => {
    const o = sample(t, slide, at(0, 0));
    expect(o.get(2)).toEqual({ hidden: true });
    expect(o.has(3)).toBe(false);
    expect(o.has(4)).toBe(false);
  });

  test('a click starts the next group; fades run over their duration and leave nothing behind', () => {
    expect(sample(t, slide, at(1, 0)).get(2)).toEqual({ opacity: 0 });
    expect(sample(t, slide, at(1, 250)).get(2)?.opacity).toBeCloseTo(0.5, 6);
    expect(sample(t, slide, at(1, 500)).has(2)).toBe(false);
    expect(sample(t, slide, at(2, 250)).get(3)?.opacity).toBeCloseTo(0.5, 6);
    expect(sample(t, slide, at(2, 500)).get(3)).toEqual({ hidden: true });
  });

  test('a later click finishes the earlier steps at once (seek), and after the last click all is final', () => {
    const o = sample(t, slide, at(2, 0));
    expect(o.has(2)).toBe(false);
    for (const position of [at(3, 1e9), at(9, 0)]) {
      const final = sample(t, slide, position);
      expect(final.has(2)).toBe(false);
      expect(final.get(3)).toEqual({ hidden: true });
      close(final.get(4)?.matrix, [2, 0, 0, 2, -2e6, -4.5e6]);
    }
  });
});

describe('effects', () => {
  test('Fly In from the bottom: the shape starts just below the slide and lands where it is', () => {
    const fly = effect('entr', 2, 4, 'clickEffect', setVis(2, 'visible') + anim(2, 'ppt_x', '#ppt_x', '#ppt_x', 500) + anim(2, 'ppt_y', '1+#ppt_h/2', '#ppt_y', 500));
    const t = clicks([fly]);
    const dy = (1 + 0.5e6 / H) * H - 1.5e6;                     // centre to just below the bottom edge
    close(sample(t, slide, at(1, 0)).get(2)?.matrix, [1, 0, 0, 1, 0, dy]);
    close(sample(t, slide, at(1, 250)).get(2)?.matrix, [1, 0, 0, 1, 0, dy / 2]);
    expect(sample(t, slide, at(1, 500)).has(2)).toBe(false);
  });

  test('with and after: effects in one step start together, an after step when the one before ends', () => {
    const t = timing(timingXml([mainSeq([clickGroup([step(0, fadeIn(2), fadeIn(3, 'withEffect')), step(500, fadeIn(4, 'afterEffect'))])])]));
    const mid = sample(t, slide, at(1, 250));
    expect([mid.get(2)?.opacity, mid.get(3)?.opacity, mid.get(4)]).toEqual([0.5, 0.5, { hidden: true }]);
    const later = sample(t, slide, at(1, 750));
    expect([later.has(2), later.has(3), later.get(4)?.opacity]).toEqual([false, false, 0.5]);
  });

  test('a delay holds the effect back from its trigger', () => {
    const t = clicks([effect('entr', 10, 0, 'clickEffect', setVis(2, 'visible') + filterFx(2, 'fade', 500), '', 1000)]);
    expect(sample(t, slide, at(1, 900)).get(2)).toEqual({ hidden: true });
    expect(sample(t, slide, at(1, 1250)).get(2)?.opacity).toBeCloseTo(0.5, 6);
  });

  test('fill: hold keeps the end state, remove drops it', () => {
    const hold = clicks([effect('emph', 6, 0, 'clickEffect', scale(4, 2, 1000))]);
    close(sample(hold, slide, at(1, 5000)).get(4)?.matrix, [2, 0, 0, 2, -2e6, -4.5e6]);
    const remove = clicks([effect('emph', 6, 0, 'clickEffect', scale(4, 2, 1000, ' fill="remove"'), ' fill="remove"').replace('fill="hold" nodeType', 'nodeType')]);
    close(sample(remove, slide, at(1, 500)).get(4)?.matrix, [1.5, 0, 0, 1.5, -1e6, -2.25e6]);
    expect(sample(remove, slide, at(1, 5000)).has(4)).toBe(false);
  });

  test('autoReverse plays forward then back; repeatCount repeats', () => {
    const rev = clicks([effect('emph', 6, 0, 'clickEffect', scale(4, 2, 1000, ' autoRev="1" fill="hold"'))]);
    const s = (ms: number) => sample(rev, slide, at(1, ms)).get(4)?.matrix?.[0];
    expect([s(500), s(1000), s(1500)]).toEqual([1.5, 2, 1.5]);
    expect(sample(rev, slide, at(1, 2000)).has(4)).toBe(false);          // back where it started
    const rep = clicks([effect('emph', 6, 0, 'clickEffect', scale(4, 2, 1000, ' repeatCount="2000" fill="hold"'))]);
    expect(sample(rep, slide, at(1, 1500)).get(4)?.matrix?.[0]).toBeCloseTo(1.5, 6);
    expect(sample(rep, slide, at(1, 2500)).get(4)?.matrix?.[0]).toBeCloseTo(2, 6);
  });

  test('accelerate and decelerate ease the progress, on a behaviour or on its effect', () => {
    expect(ease(0.4, 0, 0)).toBe(0.4);
    expect(ease(0.5, 0.5, 0.5)).toBeCloseTo(0.5, 12);
    expect(ease(0.25, 1, 0)).toBeCloseTo(0.0625, 12);
    expect(ease(0.75, 0, 1)).toBeCloseTo(0.9375, 12);
    const path = clicks([effect('path', 63, 0, 'clickEffect', motion(2, 'M 0 0 L 0.5 0 E', 1000), ' accel="50000" decel="50000"')]);
    expect(sample(path, slide, at(1, 250)).get(2)?.matrix?.[4]).toBeCloseTo(0.0625 * W, 0);
    expect(sample(path, slide, at(1, 500)).get(2)?.matrix?.[4]).toBeCloseTo(0.25 * W, 0);
  });

  test('a negative speed plays backwards', () => {
    const t = clicks([effect('path', 63, 0, 'clickEffect', motion(2, 'M 0 0 L 0.5 0 E', 1000, ' spd="-100000"'))]);
    expect(sample(t, slide, at(1, 0)).get(2)?.matrix?.[4]).toBeCloseTo(0.5 * W, 0);
    expect(sample(t, slide, at(1, 250)).get(2)?.matrix?.[4]).toBeCloseTo(0.375 * W, 0);
  });

  test('motion paths move by fractions of the slide, at a steady speed along curves too', () => {
    const t = clicks([effect('path', 0, 0, 'clickEffect', motion(2, 'M 0 0 L 0.25 0 L 0.25 0.25 E', 1000))]);
    close(sample(t, slide, at(1, 500)).get(2)?.matrix, [1, 0, 0, 1, 0.25 * W, 0]);
    close(sample(t, slide, at(1, 750)).get(2)?.matrix, [1, 0, 0, 1, 0.25 * W, 0.125 * H]);
  });

  test('Spin turns the shape about its centre', () => {
    const t = clicks([effect('emph', 8, 0, 'clickEffect', '<p:animRot by="21600000"><p:cBhvr><p:cTn dur="1000" fill="hold"/>' + tgt(2) +
      '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>')]);
    close(sample(t, slide, at(1, 250)).get(2)?.matrix, [0, 1, -1, 0, 2e6 + 1.5e6, 1.5e6 - 2e6]);
  });

  test('a tav formula (Spiral In) ends exactly where the shape is', () => {
    const tav = (attr: string, fmla: string) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn dur="1000" fill="hold"/>${tgt(2)}` +
      `<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0" fmla="${fmla}"><p:val><p:fltVal val="0"/></p:val></p:tav>` +
      '<p:tav tm="100000"><p:val><p:fltVal val="1"/></p:val></p:tav></p:tavLst></p:anim>';
    const t = clicks([effect('entr', 15, 0, 'clickEffect', setVis(2, 'visible') +
      tav('ppt_x', '#ppt_x+(cos(-2*pi*(1-$))*-#ppt_x-sin(-2*pi*(1-$))*(1-#ppt_y))*(1-$)') +
      tav('ppt_y', '#ppt_y+(sin(-2*pi*(1-$))*-#ppt_x+cos(-2*pi*(1-$))*(1-#ppt_y))*(1-$)'))]);
    close(sample(t, slide, at(1, 0)).get(2)?.matrix, [1, 0, 0, 1, -2e6, H - 1.5e6]);
    expect(sample(t, slide, at(1, 1000)).has(2)).toBe(false);
  });

  test('wipes and other filters clip the shape while they run', () => {
    const t = clicks([effect('entr', 22, 4, 'clickEffect', setVis(2, 'visible') + filterFx(2, 'wipe(down)', 1000))]);
    const mid = sample(t, slide, at(1, 500), { path2d: fakePath }).get(2)!;
    const ys = (mid.clip as unknown as FakePath).points().map(p => p[1]);
    // from the bottom: the lower half of the shape (its 1 000 000 EMU box, plus a margin) is showing
    expect(Math.min(...ys)).toBeCloseTo(1.5e6, -4);
    expect(Math.max(...ys)).toBeGreaterThanOrEqual(2e6);
    expect(sample(t, slide, at(1, 1000), { path2d: fakePath }).has(2)).toBe(false);
    // without a Path2D there is no clip, and nothing breaks
    expect(sample(t, slide, at(1, 500)).get(2)?.clip).toBeUndefined();
  });

  test('colour animations recolour the fill, from the shape\'s own colour', () => {
    const t = clicks([effect('emph', 1, 2, 'clickEffect', '<p:animClr clrSpc="rgb" dir="cw"><p:cBhvr><p:cTn dur="1000" fill="hold"/>' + tgt(2) +
      '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:to><a:schemeClr val="accent2"/></p:to></p:animClr>')]);
    const colors = { theme: new Map([['accent2', { base: { kind: 'srgb' as const, r: 255, g: 0, b: 0 }, transforms: [] }]]), clrMap: {} };
    const fill = sample(t, slide, at(1, 500), { colors }).get(2)?.fill;
    expect(fill?.kind).toBe('solid');
    const c = (fill as { color: { r: number; g: number; b: number } }).color;
    expect([c.r, c.g, c.b].map(v => +v.toFixed(3))).toEqual([0.6, 0.2, 0.4]);
  });
});

describe('regressions found on real decks', () => {
  test('a motion path written with exponents moves the shape (HowWeRefactor.pptx)', () => {
    const t = clicks([effect('path', 42, 0, 'clickEffect', motion(2, 'M 5E-6 -2.22222E-6 L 0.13126 0.38472 ', 1000))]);
    expect(sample(t, slide, at(1, 1000)).get(2)?.matrix?.[5]).toBeCloseTo(0.38472 * H, -1);
  });

  test('a wipe from the bottom starts at the bottom of a flipped shape too', () => {
    const flipped = renderSlide([shape(2, 1e6, 1e6, 2e6, 1e6, { flipV: true })]);
    const t = clicks([effect('entr', 22, 4, 'clickEffect', setVis(2, 'visible') + filterFx(2, 'wipe(down)', 1000))]);
    const ys = (sample(t, flipped, at(1, 500), { path2d: fakePath }).get(2)!.clip as unknown as FakePath).points().map(p => p[1]);
    expect(Math.min(...ys)).toBeCloseTo(1.5e6, -4);
  });

  test('a wipe covers what the geometry draws outside the box (a callout\'s tail)', () => {
    const callout = shape(2, 1e6, 1e6, 2e6, 1e6);
    callout.outline[0]!.commands.push({ op: 'M', x: 1e6, y: 1e6 }, { op: 'L', x: 1.2e6, y: 2e6 });
    const t = clicks([effect('entr', 22, 4, 'clickEffect', setVis(2, 'visible') + filterFx(2, 'wipe(down)', 1000))]);
    const ys = (sample(t, renderSlide([callout]), at(1, 500), { path2d: fakePath }).get(2)!.clip as unknown as FakePath).points().map(p => p[1]);
    expect(Math.max(...ys)).toBeGreaterThanOrEqual(3e6);
  });

  test('behaviours for Internet Explorer (rctx="IE") are left out: Transparency is half transparent, not invisible', () => {
    const t = clicks([effect('emph', 9, 0, 'clickEffect',
      '<p:set><p:cBhvr><p:cTn dur="indefinite"/>' + tgt(2) + '<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>' +
      '<p:to><p:strVal val="0.5"/></p:to></p:set><p:animEffect filter="image" prLst="opacity: 0.5"><p:cBhvr rctx="IE"><p:cTn dur="indefinite"/>' +
      tgt(2) + '</p:cBhvr></p:animEffect>')]);
    expect(sample(t, slide, at(1, 10_000)).get(2)).toEqual({ opacity: 0.5 });
  });

  test('a colour animation of every paragraph with text colours the shape\'s text (Color Pulse)', () => {
    const pulse = (i: number) => '<p:animClr clrSpc="rgb" dir="cw"><p:cBhvr override="childStyle"><p:cTn dur="1000" autoRev="1" fill="remove"/>' +
      tgt(5, para(i)) + '<p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr><p:to><a:srgbClr val="FFFFFF"/></p:to></p:animClr>';
    const all = clicks([effect('emph', 27, 0, 'clickEffect', pulse(0) + pulse(1) + pulse(2))]);
    expect(sample(all, slide, at(1, 1000)).get(5)?.textColor).toEqual({ r: 1, g: 1, b: 1, a: 1 });
    expect(sample(all, slide, at(1, 2000)).has(5)).toBe(false);
    // one paragraph of three cannot be coloured alone
    expect(sample(clicks([effect('emph', 27, 0, 'clickEffect', pulse(0))]), slide, at(1, 1000)).has(5)).toBe(false);
  });
});

describe('parts of shapes', () => {
  const byParagraph = clicks(...[0, 1, 2].map(i => [fadeIn(5, 'clickEffect', 500, para(i))]));

  test('a paragraph build hides the paragraphs not yet built', () => {
    expect([...sample(byParagraph, slide, at(0, 0)).get(5)!.paragraphs!]).toEqual([[0, { hidden: true }], [1, { hidden: true }], [2, { hidden: true }]]);
    const mid = sample(byParagraph, slide, at(1, 250)).get(5)!.paragraphs!;
    expect([...mid]).toEqual([[0, { opacity: 0.5 }], [1, { hidden: true }], [2, { hidden: true }]]);
    expect(sample(byParagraph, slide, at(3, 500)).has(5)).toBe(false);
  });

  test('a paragraph flying in fades in with its effect, as the override cannot move one paragraph', () => {
    const fly = effect('entr', 2, 4, 'clickEffect', setVis(5, 'visible', 0, para(1)) + anim(5, 'ppt_y', '1+#ppt_h/2', '#ppt_y', 1000, '', para(1)));
    const o = sample(clicks([fly]), slide, at(1, 250)).get(5)!;
    expect(o.matrix).toBeUndefined();
    expect(o.paragraphs?.get(1)?.opacity).toBeCloseTo(0.25, 6);
  });

  test('a paragraph with laid-out lines flies in on its own: it moves, from below, and is not faded', () => {
    // the same shape, with its text laid out: two lines, one per paragraph, each 2e5 EMU high
    const base = slide.shapes.find(s => s.id === 5)!;
    const lines = [0, 1, 2].map(i => ({
      paragraph: i, start: 0, end: 3, top: i * 2e5, height: 2e5, baseline: i * 2e5 + 1.6e5, ascent: 1.6e5, descent: 0.4e5,
      left: 0, available: 2e6, width: 1e6, carets: [], runs: [{ x: 1e5, advance: 1e6, paragraph: i }],
    }));
    const text = { kind: 'text', transform: [1, 0, 0, 1, 4e6, 4e6], layout: { lines, width: 2e6, height: 2e6 } };
    const laid = renderSlide([{ ...base, prims: [text] } as unknown as typeof base]);
    const fly = effect('entr', 2, 4, 'clickEffect', setVis(5, 'visible', 0, para(1)) + anim(5, 'ppt_y', '1+#ppt_h/2', '#ppt_y', 1000, '', para(1)));
    const mid = sample(clicks([fly]), laid, at(1, 250)).get(5)!.paragraphs!.get(1)!;
    expect(mid.opacity).toBeUndefined();
    expect(mid.matrix![5]).toBeGreaterThan(0);                    // still below where it lands
    expect(mid.matrix![4]).toBeCloseTo(0, 3);
    expect(sample(clicks([fly]), laid, at(1, 1200)).get(5)?.paragraphs?.get(1)).toBeUndefined();
  });

  test('the background of a shape built by paragraph (animBg) fades its fill and line', () => {
    const t = clicks([fadeIn(5, 'clickEffect', 500, '<p:bg/>')]);
    const before = sample(t, slide, at(0, 0)).get(5)!;
    expect(before.hidden).toBeUndefined();
    expect(before.fill).toMatchObject({ kind: 'solid', color: { a: 0 } });
    expect(sample(t, slide, at(1, 250)).get(5)!.fill).toMatchObject({ kind: 'solid', color: { a: 0.5 } });
  });

  test('a shape of an interactive sequence waits for its trigger to be clicked', () => {
    const t = timing(timingXml([mainSeq([clickGroup([step(0, fadeIn(2))])]), interactiveSeq(4, [
      '<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' + step(0, fadeIn(3)) + '</p:childTnLst></p:cTn></p:par>',
    ])]));
    expect(sample(t, slide, at(1, 1000)).get(3)).toEqual({ hidden: true });
    expect(sample(t, slide, at(0, 0, [{ spid: 4, ms: 250 }])).get(3)?.opacity).toBeCloseTo(0.5, 6);
    expect(sample(t, slide, at(0, 0, [{ spid: 4, ms: 1000 }])).has(3)).toBe(false);
  });

  test('SmartArt nodes built one by one are found by their model id', () => {
    const node = shape(1, 0, 0, 1e6, 1e6, { part: '/ppt/diagrams/drawing1.xml', modelId: '{N1}' });
    const frame = shape(7, 0, 0, 3e6, 2e6, { diagram: true, children: [node] });
    const s = renderSlide([frame]);
    const t = clicks([fadeIn(7, 'clickEffect', 500, '<p:graphicEl><a:dgm id="{N1}"/></p:graphicEl>')]);
    const f = sampleFrame(t, s, at(0, 0));
    expect(f.nodes.get(node)).toEqual({ hidden: true });
    expect(f.shapes.has(7)).toBe(false);
    const lookup = overrideLookup(f, s);
    expect(lookup(node)).toEqual({ hidden: true });
  });

  test('chart series cannot be drawn alone: the first chart build stands for the chart, the later ones are left out', () => {
    const series = (i: number) => `<p:graphicEl><a:chart seriesIdx="${i}" categoryIdx="-1" bldStep="series"/></p:graphicEl>`;
    const t = clicks([fadeIn(4, 'clickEffect', 500, series(-3))], [fadeIn(4, 'clickEffect', 500, series(0))]);
    expect(sample(t, slide, at(0, 0)).get(4)).toEqual({ hidden: true });
    expect(sample(t, slide, at(1, 1000)).has(4)).toBe(false);
    expect(sample(t, slide, at(2, 0)).has(4)).toBe(false);
  });

  test('master and layout shapes keep their own ids out of the slide\'s overrides', () => {
    const master = shape(2, 0, 0, 1, 1, { layer: 'master', part: '/ppt/slideMasters/slideMaster1.xml' });
    const s = renderSlide([master, ...slide.shapes]);
    const lookup = overrideLookup(sample(clicks([fadeIn(2)]), s, at(0, 0)), s);
    expect(lookup(master)).toBeUndefined();
    expect(lookup(s.shapes[1]!)).toEqual({ hidden: true } satisfies ShapeOverride);
  });
});
