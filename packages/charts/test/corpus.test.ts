// Every chart part in the corpus draws without an exception and with finite numbers only: a NaN in a path or a
// transform would draw nothing, or everything, in the renderer. Charts are drawn on the slides that show them,
// through buildSlide as the editor does, and every chart part is also drawn on its own.
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { buildSlide, type Primitive } from '@pptx/render';
import { corpusFiles, corpusName } from '../../../tools/corpus/list.ts';
import { expectedRefusals } from '../../opc/test/corpus-helpers.ts';
import { chartRenderer, drawChart, loadChart } from '../src/index.ts';
import { fonts } from './helpers.ts';

const CHART_PART = /^\/?ppt\/charts\/chart[^/]*\.xml$/i;

// decks that hold chart parts, found from the names in their ZIP directories without opening them
const decks = corpusFiles.filter(f => !expectedRefusals[corpusName(f)] && readFileSync(f).includes('ppt/charts/chart'));

function check(prims: Primitive[], where: string): void {
  for (const p of prims) {
    expect(p.transform.every(Number.isFinite), `${where} transform`).toBe(true);
    if (p.kind === 'path') {
      expect(Number.isFinite(p.box.w) && Number.isFinite(p.box.h), `${where} box`).toBe(true);
      for (const s of p.paths) for (const c of s.commands) for (const v of Object.values(c)) if (typeof v === 'number') expect(Number.isFinite(v), `${where} path`).toBe(true);
      if (p.stroke) expect(Number.isFinite(p.stroke.width), `${where} stroke`).toBe(true);
    } else if (p.kind === 'text') {
      for (const l of p.layout.lines) for (const r of l.runs) expect(Number.isFinite(r.x) && Number.isFinite(r.baseline) && r.xs.every(Number.isFinite), `${where} text`).toBe(true);
    } else {
      expect(Number.isFinite(p.box.w) && Number.isFinite(p.box.h), `${where} box`).toBe(true);
    }
  }
}

let drawn = 0;

describe.skipIf(decks.length === 0)('every chart in the corpus draws', () => {
  test.each(decks.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    const doc = new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync(file))));
    const reg = fonts();
    const hook = chartRenderer(reg);
    let first: ReturnType<typeof resolveSlide> | undefined;
    for (const part of doc.slideParts()) {
      const slide = resolveSlide(doc, part);
      first ??= slide;
      const chart: typeof hook = (p, box, ctx) => {
        const prims = hook(p, box, ctx);
        if (prims) {
          check(prims, `${part} ${p}`);
          drawn++;
          return prims;
        }
        // only a part that is not a chart (or cannot be read) may give undefined
        let local = '';
        try {
          local = doc.root(p).local;
        } catch {
          local = '';
        }
        expect(local, `${p} is a chart but drew nothing`).not.toBe('chartSpace');
        return prims;
      };
      expect(() => buildSlide(slide, { fonts: reg, chart })).not.toThrow();
    }
    // every chart part, also those no slide shows, drawn on its own in the first slide's colours
    if (!first) return;
    for (const p of doc.pkg.partNames().filter(n => CHART_PART.test(n))) {
      const loaded = loadChart(p, first.context, reg);
      if (!loaded) continue;
      check(drawChart(loaded.chart, loaded.env, { w: 5486400, h: 3200400 }), p);
      check(drawChart(loaded.chart, loaded.env, { w: 50000, h: 30000 }), `${p} (tiny)`);
    }
  });

  test('the charts were drawn', () => {
    expect(drawn).toBeGreaterThanOrEqual(150);
  });
});
