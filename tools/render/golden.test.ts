// Oracle O3's ratchet: slides rendered by our renderer in Chromium, compared with approved renders by pixelmatch.
// The plan compares with PowerPoint's own export (POST /oracle/export on the Windows runner); until that runs, the
// baselines are our own approved output, so this catches regressions, not fidelity. Approve baselines with
// UPDATE_GOLDEN=1 (a case without one is skipped, never approved implicitly). Failures' diffs go to out/golden-diff/.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { startHarness, type Harness } from './render.ts';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const GOLDEN = join(ROOT, 'tools/render/golden');
const DIFF = join(ROOT, 'out/golden-diff');
const WIDTH = 960;

/** The plan's per-class limits on mismatching pixels (exit criteria of phase 3). */
const LIMIT = { shapes: 0.01, text: 0.02, charts: 0.05 } as const;

/** Slides chosen to cover the renderer: [name, deck in corpus/external, slide index, class]. */
const CASES: [string, string, number, keyof typeof LIMIT][] = [
  ['title-photos', 'poi/test-data~slideshow~ca.ubc.cs.people_~emhill_presentations_HowWeRefactor.pptx', 0, 'text'],
  ['arrows-code', 'poi/test-data~slideshow~ca.ubc.cs.people_~emhill_presentations_HowWeRefactor.pptx', 4, 'shapes'],
  ['bullets-template', 'poi/test-data~slideshow~2411-Performance_Up.pptx', 6, 'text'],
  ['blue-template', 'poi/test-data~slideshow~45541_Footer.pptx', 2, 'shapes'],
  ['ole-preview', 'poi/test-data~slideshow~45541_Footer.pptx', 7, 'shapes'],
  ['reflections', 'libreoffice/sd~qa~unit~data~pptx~shape-reflection-effect.pptx', 0, 'shapes'],
  ['soft-edges', 'libreoffice/sd~qa~unit~data~pptx~shape-soft-edges.pptx', 0, 'shapes'],
  ['smartart-outline', 'libreoffice/sd~qa~unit~data~pptx~smartart-children.pptx', 0, 'text'],
  ['bar-chart', 'libreoffice/sd~qa~unit~data~pptx~tdf149985.pptx', 0, 'charts'],
  ['strict-area-chart', 'openxml-sdk/test~DocumentFormat.OpenXml.Tests.Assets~assets~TestDataStorage~O14ISOStrict~Graphics~2D Area-O12-PPT-Charts.pptx', 0, 'charts'],
  ['activex-previews', 'libreoffice/sd~qa~unit~data~pptx~activex_checkbox.pptx', 0, 'shapes'],
  ['smartart-turned', 'poi/test-data~slideshow~ae.ac.uaeu.faculty_nafaachbili_GeomLec1.pptx', 8, 'text'],
  ['tiff-picture', 'poi/test-data~slideshow~ArtisticEffectSample.pptx', 1, 'shapes'],
  ['emf-picture', 'libreoffice/sd~qa~unit~data~pptx~tdf143222.pptx', 0, 'shapes'],
];

const cases = CASES.filter(([, deck]) => existsSync(join(ROOT, 'corpus/external', deck))).map(([name, deck, index, kind]) => ({ name, deck, index, kind }));
let harness: Harness;

beforeAll(async () => { if (cases.length) harness = await startHarness(); }, 120_000);
afterAll(async () => { await harness?.close(); });

describe.skipIf(cases.length === 0)('renders match their approved baselines', () => {
  test.for(cases)('$name', { timeout: 60_000 }, async ({ name, deck, index, kind }, ctx) => {
    const file = join(GOLDEN, `${name}.png`);
    if (!process.env.UPDATE_GOLDEN && !existsSync(file)) ctx.skip();
    const [r] = await harness.render(new Uint8Array(readFileSync(join(ROOT, 'corpus/external', deck))), WIDTH, [index]);
    expect(r?.error, r?.error).toBeUndefined();
    const actual = PNG.sync.read(Buffer.from(r!.png!, 'base64'));
    if (process.env.UPDATE_GOLDEN) {
      mkdirSync(GOLDEN, { recursive: true });
      writeFileSync(file, PNG.sync.write(actual));
      return;
    }
    const expected = PNG.sync.read(readFileSync(file));
    expect([actual.width, actual.height]).toEqual([expected.width, expected.height]);
    const diff = new PNG({ width: actual.width, height: actual.height });
    const bad = pixelmatch(expected.data, actual.data, diff.data, actual.width, actual.height, { threshold: 0.1 });
    const ratio = bad / (actual.width * actual.height);
    if (ratio > LIMIT[kind]) {
      mkdirSync(DIFF, { recursive: true });
      writeFileSync(join(DIFF, `${name}-diff.png`), PNG.sync.write(diff));
      writeFileSync(join(DIFF, `${name}-actual.png`), PNG.sync.write(actual));
    }
    expect(ratio, `${name}: ${(ratio * 100).toFixed(2)}% of pixels differ (limit ${LIMIT[kind] * 100}%)`).toBeLessThanOrEqual(LIMIT[kind]);
  });
});
