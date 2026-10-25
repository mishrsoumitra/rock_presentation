// Every corpus deck: the timing and transition of every slide, layout and master read without an exception; every
// slide with animations schedules and samples at any position into sane overrides; rewriting the main sequence of
// a slide PowerPoint wrote changes nothing; and every transition written back reads back the same.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { beforeAll, describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { checkOrder, PresentationDocument } from '@pptx/model';
import { decodeXmlBytes, serializeElement } from '@pptx/xml';
import { resolveSlide, slideContext } from '@pptx/resolve';
import { buildSlide } from '@pptx/render';
import { FontRegistry } from '../../text/src/index.ts';
import { corpusFiles, corpusName } from '../../../tools/corpus/list.ts';
import { expectedRefusals } from '../../opc/test/corpus-helpers.ts';
import {
  clickCount, listEffects, moveEffect, parseTiming, readTransition, sampleFrame, scheduleFor, setTransition, slideTransition, type TransitionSpec,
} from '../src/index.ts';

const MALFORMED_PARTS = new Set(['libreoffice/sd~qa~unit~data~pptx~fail~ofz35597-1.pptx']);
const decks = corpusFiles.filter(f => !expectedRefusals[corpusName(f)] && !MALFORMED_PARTS.has(corpusName(f)));
const fonts = new FontRegistry();
const totals = { decksWithTiming: 0, slidesWithTiming: 0, transitions: 0, rewritten: 0 };

beforeAll(() => {
  const dir = join(import.meta.dirname, '../../../assets/fonts');
  for (const f of ['carlito/Carlito-Regular.ttf', 'liberation/LiberationSans-Regular.ttf', 'liberation/LiberationSerif-Regular.ttf']) {
    fonts.addBytes(new Uint8Array(readFileSync(join(dir, f))));
  }
});

function application(doc: PresentationDocument): string {
  const part = doc.pkg.getPart('/docProps/app.xml');
  return part ? /<Application>([^<]*)</.exec(decodeXmlBytes(part.bytes).text)?.[1] ?? '' : '';
}

describe.skipIf(decks.length === 0)('animations of every corpus deck', () => {
  test.each(decks.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    const open = () => new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync(file))));
    const doc = open();
    const slides = doc.slideParts();
    const others = new Set<string>();
    for (const m of doc.masterParts()) {
      others.add(m);
      for (const l of doc.layoutsOf(m)) others.add(l);
    }
    // layouts and masters: read only
    for (const part of others) {
      const view = doc.slide(part);
      if (view.timing) parseTiming(view.timing);
      slideTransition(view);
    }
    let withTiming = 0;
    for (const part of slides) {
      const view = doc.slide(part);
      const t = slideTransition(view);
      if (t) {
        totals.transitions++;
        expect(t.duration).toBeGreaterThanOrEqual(0);
      }
      if (!view.timing) continue;
      withTiming++;
      const timing = parseTiming(view.timing);
      const slide = buildSlide(resolveSlide(doc, part), { fonts });
      const schedule = scheduleFor(timing, slide);
      const colors = slideContext(doc, part).colors;
      for (let clicks = 0; clicks <= clickCount(schedule) + 1; clicks++) {
        for (const ms of [0, 333, 5000, Infinity]) {
          const triggers = schedule.interactive.map(i => ({ spid: i.trigger, ms }));
          const frame = sampleFrame(timing, slide, { clicks, ms, triggers }, { colors });
          for (const o of [...frame.shapes.values(), ...frame.nodes.values()]) {
            if (o.opacity !== undefined) expect(o.opacity).toBeGreaterThanOrEqual(0);
            if (o.opacity !== undefined) expect(o.opacity).toBeLessThan(1);
            if (o.matrix) expect(o.matrix.every(Number.isFinite)).toBe(true);
            // paragraphs moved, wiped or faded on their own
            for (const po of 'paragraphs' in o ? o.paragraphs?.values() ?? [] : []) {
              if (po.opacity !== undefined) expect(po.opacity).toBeGreaterThanOrEqual(0);
              if (po.opacity !== undefined) expect(po.opacity).toBeLessThan(1);
              if (po.matrix) expect(po.matrix.every(Number.isFinite)).toBe(true);
            }
          }
        }
      }
    }
    if (withTiming) {
      totals.decksWithTiming++;
      totals.slidesWithTiming += withTiming;
    }

    // rewriting what PowerPoint wrote (moving the first effect of each sequence onto itself) gives the same XML back
    if (/PowerPoint/.test(application(doc)) && !/LibreOffice/.test(application(doc))) {
      const fresh = open();
      for (const part of fresh.slideParts()) {
        const view = fresh.slide(part);
        if (!view.timing) continue;
        const before = serializeElement(view.timing);
        if (/clickPar|withGroup|afterGroup/.test(before)) continue;
        const effects = listEffects(view);
        const firsts = effects.filter((e, i) => i === 0 || e.triggerShape !== effects[i - 1]!.triggerShape).map(e => e.index);
        for (const i of firsts) moveEffect(view, i, i);
        if (firsts.length) expect(serializeElement(view.timing!), part).toBe(before);
        totals.rewritten += firsts.length;
      }
    }

    // every transition written back reads back the same, in schema order
    const again = open();
    for (const part of again.slideParts()) {
      const view = again.slide(part);
      const t = slideTransition(view);
      if (!t) continue;
      const spec: TransitionSpec = { ...t, ...(t.advanceAfter !== undefined ? { advanceAfter: t.advanceAfter } : {}) };
      setTransition(view, spec);
      const back = slideTransition(view)!;
      expect({ ...back, el: undefined, effectEl: undefined }).toEqual({ ...readTransition(t.el), el: undefined, effectEl: undefined });
      expect(checkOrder(view.root)).toEqual([]);
    }
  });

  test('covers the animated part of the corpus', () => {
    // 209 decks and 639 slides with p:timing, 366 slides with a transition, when this was written
    expect(totals.decksWithTiming).toBeGreaterThanOrEqual(200);
    expect(totals.slidesWithTiming).toBeGreaterThanOrEqual(600);
    expect(totals.transitions).toBeGreaterThanOrEqual(350);
    expect(totals.rewritten).toBeGreaterThanOrEqual(130);
  });
});
