// Every editing command on every corpus deck: each must pass the oracle and undo exactly (Phase 4 exit criterion).
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { children } from '@pptx/model';
import { corpusFiles, corpusName } from '../../../tools/corpus/list.ts';
import { expectedRefusals } from '../../opc/test/corpus-helpers.ts';
import {
  addSlide, align, applyLayout, copyShapes, cropPicture, deleteShapes, deleteSlide, deleteText, distribute, duplicateShapes, duplicateSlide, effectiveXfrm, flip,
  importMasters, addAnimation, moveAnimation, removeAnimation, setAnimationTiming, setSlideTransition, setTransitionOnAll,
  formatBody, formatParagraphs, formatRuns, group, insertPicture, insertShape, insertTextBox, moveBy, moveSlide, paragraphLength,
  pasteShapes, prepareImage, prepareSvg, rename, replacePicture, replaceText, rotateQuarter, setAltText, setBackground, setFill, setLine, setSlideHidden, setSlideSize,
  setTransforms, typeText, ungroup, unitOf, zOrder,
  type Command, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';
import { OpcPackage } from '@pptx/opc';
import { findPreset, listEffects } from '@pptx/anim';
import { deckBytes } from '../../resolve/test/deck.ts';

const MALFORMED_PARTS = new Set(['libreoffice/sd~qa~unit~data~pptx~fail~ofz35597-1.pptx']);
const decks = corpusFiles.filter(f => !expectedRefusals[corpusName(f)] && !MALFORMED_PARTS.has(corpusName(f)));

const FADE = findPreset('entr', 10)!;

/** A small template (the resolver tests' deck) whose master is imported into every corpus deck. */
const TEMPLATE = OpcPackage.open(deckBytes());
const TEMPLATE_HASHES = new Map<string, string>();

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0, 123, 64, 232, 221, 0, 0, 0, 15, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 86, 199, 45, 179, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);

const SVG = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="2" height="1"><rect width="2" height="1" fill="#c00"/></svg>');

/** The commands to try on a deck, built fresh for each run (the deck is back to its original state between runs). */
async function battery(ed: Editor): Promise<[string, () => Command][]> {
  const slides = ed.doc.slideParts();
  const slide = slides[0];
  if (!slide) return [];
  const view = ed.slide(slide);
  const top = () => ed.slide(slide).shapes;
  const first = () => top()[0]!;
  const layout = ed.doc.layoutOf(slide);
  const master = layout ? ed.doc.masterOf(layout) : undefined;
  const otherLayout = master ? ed.doc.layoutsOf(master).find(l => l !== layout) : undefined;
  const text = () => top().find(s => s.txBody && children(s.txBody, 'a:p').length > 0);
  const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 2, height: 1 });
  const svg = await prepareSvg(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 2, height: 1 }, SVG);
  const out: [string, () => Command][] = [];
  if (view.spTree) {
    out.push(['insert shape', () => insertShape(ed, slide, 'rect', { x: 10, y: 10, w: 1000, h: 1000 }).command]);
    out.push(['insert text box', () => insertTextBox(ed, slide, { x: 10, y: 10, w: 1000, h: 1000 }, 'Hello\nworld').command]);
    out.push(['insert picture', () => insertPicture(ed, slide, img).command]);
    out.push(['insert SVG picture', () => insertPicture(ed, slide, svg).command]);
    out.push(['background', () => setBackground(ed, slide, { hex: 'FF0000' })]);
  }
  if (top().length) {
    out.push(['move', () => moveBy(ed, slide, [first()], 12700, 25400)]);
    out.push(['resize', () => {
      const x = effectiveXfrm(ed, slide, first());
      return setTransforms(ed, slide, [{ shape: first(), xfrm: { ...x, cx: x.cx * 1.1, cy: x.cy * 0.9 } }], 'Resize');
    }]);
    out.push(['rotate', () => rotateQuarter(ed, slide, [first()], 1)]);
    out.push(['flip', () => flip(ed, slide, [first()], 'h')]);
    out.push(['z-order', () => zOrder(ed, slide, [first()], 'front')]);
    out.push(['duplicate', () => duplicateShapes(ed, slide, [first()]).command]);
    out.push(['delete', () => deleteShapes(ed, slide, [first()])]);
    out.push(['copy and paste', () => pasteShapes(ed, slide, copyShapes(ed, slide, [first()])).command]);
    out.push(['fill', () => setFill(ed, slide, [first()], { kind: 'scheme', val: 'accent2', lumMod: 60000, lumOff: 40000 })]);
    out.push(['outline', () => setLine(ed, slide, [first()], { color: { kind: 'srgb', hex: '00B050' }, width: 28575, dash: 'sysDash' })]);
    out.push(['rename', () => rename(ed, slide, first(), 'Renamed shape')]);
    out.push(['alt text', () => setAltText(ed, slide, first(), 'A description & "quotes"')]);
  }
  // pictures: on the first slide that has one
  const picSlide = slides.find(p => ed.slide(p).shapes.some(s => s.kind === 'pic' && s.blipFill));
  if (picSlide) {
    const pic = () => ed.slide(picSlide).shapes.find(s => s.kind === 'pic' && s.blipFill)!;
    out.push(['crop picture', () => cropPicture(ed, picSlide, pic(), { l: 10000, t: 0, r: 5000, b: 20000 })]);
    out.push(['change picture', () => replacePicture(ed, picSlide, pic(), img)]);
  }
  const shapes = top();
  if (shapes.length >= 2 && unitOf(shapes[0]!.el).parent === unitOf(shapes[1]!.el).parent) {
    out.push(['group', () => group(ed, slide, [top()[0]!, top()[1]!])]);
    out.push(['align', () => align(ed, slide, [top()[0]!, top()[1]!], 'middle')]);
  }
  if (shapes.length >= 3) out.push(['distribute', () => distribute(ed, slide, top().slice(0, 3), 'h')]);
  if (shapes.length) out.push(['send backward', () => zOrder(ed, slide, [top()[top().length - 1]!], 'backward')]);
  if (shapes.some(s => s.kind === 'grpSp' && unitOf(s.el) === s.el)) {
    out.push(['ungroup', () => ungroup(ed, slide, top().find(s => s.kind === 'grpSp' && unitOf(s.el) === s.el)!)]);
  }
  if (text()) {
    const body = () => text()!.txBody!;
    const end0 = () => ({ p: 0, offset: paragraphLength(children(body(), 'a:p')[0]!) });
    out.push(['type', () => typeText(ed, slide, body(), end0(), 'abc\nxyz').command]);
    out.push(['backspace', () => deleteText(ed, slide, body(), end0()).command]);
    out.push(['delete selection', () => {
      const ps = children(body(), 'a:p');
      return deleteText(ed, slide, body(), { start: { p: 0, offset: 0 }, end: { p: ps.length - 1, offset: paragraphLength(ps[ps.length - 1]!) } }).command;
    }]);
    out.push(['font', () => {
      const ps = children(body(), 'a:p');
      return formatRuns(ed, slide, body(), { start: { p: 0, offset: 0 }, end: { p: ps.length - 1, offset: paragraphLength(ps[ps.length - 1]!) } }, { b: true, sz: 2000, fill: { kind: 'srgb', hex: 'C00000' }, latin: 'Arial' });
    }]);
    out.push(['paragraph', () => formatParagraphs(ed, slide, body(), { start: { p: 0, offset: 0 }, end: { p: 0, offset: 0 } }, { algn: 'ctr', lvl: 1, bullet: { kind: 'char', char: '•', font: 'Arial' }, spcBef: { pts: 600 } })]);
    out.push(['text box options', () => formatBody(ed, slide, body(), { anchor: 'ctr', autofit: { kind: 'norm', fontScale: 90000 } })]);
    out.push(['replace text', () => replaceText(ed, slide, body(), 'Replaced\ntext')]);
  }
  if (top().length) out.push(['add animation', () => addAnimation(ed, slide, first(), FADE, { trigger: 'click' })]);
  const effects = listEffects(ed.slide(slide));
  if (effects.length) {
    out.push(['retime animation', () => setAnimationTiming(ed, slide, 0, { duration: 1500, delay: 100 })]);
    out.push(['remove animation', () => removeAnimation(ed, slide, 0)]);
  }
  if (effects.length > 1) out.push(['reorder animation', () => moveAnimation(ed, slide, effects.length - 1, 0)]);
  out.push(['transition', () => setSlideTransition(ed, slide, { kind: 'push', dir: 'u', duration: 800 })]);
  out.push(['transition on every slide', () => setTransitionOnAll(ed, { kind: 'fade', duration: 700 })]);
  if (layout) out.push(['new slide', () => addSlide(ed, layout, 1).command]);
  out.push(['duplicate slide', () => duplicateSlide(ed, slide).command]);
  if (slides.length > 1) out.push(['move slide', () => moveSlide(ed, slide, slides.length - 1)]);
  out.push(['delete slide', () => deleteSlide(ed, slide)]);
  out.push(['hide slide', () => setSlideHidden(ed, slide, true)]);
  if (otherLayout) out.push(['apply layout', () => applyLayout(ed, slide, otherLayout)]);
  const size = ed.doc.presentation.slideSize;
  out.push(['slide size', () => setSlideSize(ed, size.cx === 9144000 ? 12192000 : 9144000, 6858000, 'ensureFit')]);
  out.push(['import masters', () => importMasters(ed, TEMPLATE, TEMPLATE_HASHES).command]);
  return out;
}

export function runShard(shard: number, of: number): void {
  const mine = decks.filter((_, i) => i % of === shard);
  describe.skipIf(mine.length === 0)(`editing commands on corpus decks (shard ${shard + 1} of ${of})`, () => {
    test.each(mine.map(f => [corpusName(f), f]))('%s', async (_name, file) => {
      const ed = openEditor(new Uint8Array(readFileSync(file)));
      const problems: string[] = [];
      for (const [name, make] of await battery(ed)) {
        try {
          problems.push(...cycle(ed, make, { parse: false }).map(p => `${name}: ${p}`));
        } catch (e) {
          problems.push(`${name}: threw ${e instanceof Error ? e.message : String(e)}`);
        }
      }
      expect(problems).toEqual([]);
    });
  });
}
