// apps/server/scripts/corpus.ts  -- the server's phase 6 exit run: every corpus package uploaded, every slide drawn, a PDF
//
//   node apps/server/scripts/corpus.ts [--pages 4] [--width 960] [--concurrency 4] [--limit N] [--out out/server-corpus.json] [name filter...]
//
// A server in this process, made as the server tests make one, with the parser workers' heap capped at 1 GB (the
// exit criterion) and a fresh build of the render page, so it draws with the packages as they are now. Each package
// is uploaded through POST /api/decks. The reviewed refusals (corpus/expected-refusals.json, plus the two the
// server's own check adds, as in apps/server/test/decks.test.ts) must be refused with their code, and encrypted
// packages stored as uploaded; every other deck must give a PNG of each slide and a PDF with a page per slide.
// Failures are printed; every deck's outcome and timings go to --out.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { build } from 'vite';
import { cleanup, readRepoFile, ROOT, tempDir, testServer, upload } from '../test/helpers.ts';
import { corpusName, opcFiles } from '../../../tools/corpus/list.ts';

const argv = process.argv.slice(2);
const flag = (name: string, fallback: string) => {
  const i = argv.indexOf(`--${name}`);
  return i >= 0 ? argv.splice(i, 2)[1]! : fallback;
};
const pages = Number(flag('pages', '4'));
const width = Number(flag('width', '960'));
const concurrency = Number(flag('concurrency', '4'));
const limit = Number(flag('limit', '0'));
const out = flag('out', 'out/server-corpus.json');
const filters = argv;

const refusals = JSON.parse(new TextDecoder().decode(readRepoFile('corpus/expected-refusals.json'))) as Record<string, { code: string }>;
const expected: Record<string, string> = {
  ...Object.fromEntries(Object.entries(refusals).map(([f, { code }]) => [f, code])),
  'libreoffice/sd~qa~unit~data~pptx~fail~ofz35597-1.pptx': 'invalid-part',
  'poi/test-data~slideshow~testPPT.thmx': 'not-a-presentation',
};

interface Outcome {
  name: string;
  expected: string;
  got: string;
  slides?: number;
  failedSlides?: { n: number; status: number; error?: string }[];
  pdf?: { status: number; pages?: number; bytes?: number; error?: string };
  ms: { upload: number; png: number; pdf: number };
  ok: boolean;
}

const pageDir = tempDir('pptx-render-page-');
await build({ configFile: join(ROOT, 'apps/server/render-page/vite.config.ts'), logLevel: 'error', build: { outDir: pageDir, emptyOutDir: true } });
const app = await testServer({ workers: { size: concurrency, heapMb: 1024 }, render: { pages, pageDir } });

let files = opcFiles.filter(f => !filters.length || filters.some(x => corpusName(f).includes(x)));
if (limit) files = files.slice(0, limit);
const results: Outcome[] = [];
const started = Date.now();
let done = 0;

const isPdf = (b: Buffer) => b.subarray(0, 5).toString() === '%PDF-';

async function run(file: string): Promise<Outcome> {
  const name = corpusName(file);
  const want = expected[name] ?? 'accepted';
  const ms = { upload: 0, png: 0, pdf: 0 };
  let t = Date.now();
  const res = await upload(app, new Uint8Array(readFileSync(file)), { filename: name.slice(name.lastIndexOf('~') + 1) });
  ms.upload = Date.now() - t;
  const body = res.json() as { id?: string; slideCount?: number | null; encrypted?: boolean; error?: string };
  if (res.statusCode !== 201) {
    const got = body.error ?? `http-${res.statusCode}`;
    return { name, expected: want, got, ms, ok: got === want };
  }
  if (want === 'encrypted' || body.encrypted) return { name, expected: want, got: 'encrypted', ms, ok: want === 'encrypted' };
  if (want !== 'accepted') return { name, expected: want, got: 'accepted', ms, ok: false };

  const slides = body.slideCount ?? 0;
  const failedSlides: Outcome['failedSlides'] = [];
  t = Date.now();
  for (let n = 1; n <= slides; n++) {
    const png = await app.inject({ url: `/api/decks/${body.id}/slides/${n}.png?width=${width}` });
    const isPng = png.statusCode === 200 && png.rawPayload.subarray(1, 4).toString() === 'PNG';
    if (!isPng) failedSlides.push({ n, status: png.statusCode, error: png.statusCode === 200 ? 'not-png' : (png.json() as { error?: string }).error });
  }
  ms.png = Date.now() - t;
  t = Date.now();
  const pdfRes = await app.inject({ url: `/api/decks/${body.id}/export.pdf` });
  ms.pdf = Date.now() - t;
  let pdf: Outcome['pdf'];
  if (pdfRes.statusCode === 200) {
    const text = pdfRes.rawPayload.toString('latin1');
    pdf = { status: 200, pages: (text.match(/\/Type\s*\/Page(?![a-zA-Z])/g) ?? []).length, bytes: pdfRes.rawPayload.length };
  } else {
    pdf = { status: pdfRes.statusCode, error: (pdfRes.json() as { error?: string }).error };
  }
  const pdfOk = slides === 0 ? pdf.error === 'no-slides' : pdf.status === 200 && isPdf(pdfRes.rawPayload) && pdf.pages === slides;
  await app.inject({ method: 'DELETE', url: `/api/decks/${body.id}` });
  return { name, expected: want, got: 'accepted', slides, failedSlides, pdf, ms, ok: !failedSlides.length && pdfOk };
}

let next = 0;
await Promise.all(Array.from({ length: concurrency }, async () => {
  while (next < files.length) {
    const file = files[next++]!;
    let r: Outcome;
    try {
      r = await run(file);
    } catch (e) {
      r = { name: corpusName(file), expected: expected[corpusName(file)] ?? 'accepted', got: `threw: ${(e as Error).message}`, ms: { upload: 0, png: 0, pdf: 0 }, ok: false };
    }
    results.push(r);
    done++;
    if (!r.ok) console.log(`FAIL  ${r.name}: expected ${r.expected}, got ${r.got}${r.failedSlides?.length ? `, slides ${JSON.stringify(r.failedSlides)}` : ''}${r.pdf && r.pdf.pages !== r.slides ? `, pdf ${JSON.stringify(r.pdf)} for ${r.slides} slides` : ''}`);
    if (done % 50 === 0) console.log(`${done}/${files.length} decks, ${Math.round((Date.now() - started) / 1000)} s`);
  }
}));

results.sort((a, b) => a.name.localeCompare(b.name));
const accepted = results.filter(r => r.got === 'accepted');
const summary = {
  decks: results.length,
  ok: results.filter(r => r.ok).length,
  failed: results.filter(r => !r.ok).map(r => r.name),
  accepted: accepted.length,
  refusedAsExpected: results.filter(r => r.expected !== 'accepted' && r.ok).length,
  slides: accepted.reduce((n, r) => n + (r.slides ?? 0), 0),
  pdfPages: accepted.reduce((n, r) => n + (r.pdf?.pages ?? 0), 0),
  seconds: Math.round((Date.now() - started) / 1000),
  slowestPdf: [...accepted].sort((a, b) => b.ms.pdf - a.ms.pdf).slice(0, 5).map(r => ({ name: r.name, slides: r.slides, ms: r.ms.pdf })),
  settings: { workerHeapMb: 1024, renderPages: pages, pngWidth: width, pdfWidth: app.services.config.render.pdfWidth, concurrency },
};
mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, JSON.stringify({ summary, results }, null, 1));
console.log(JSON.stringify(summary, null, 1));
await app.close();
cleanup();
process.exit(summary.failed.length ? 1 : 0);
