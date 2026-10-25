// Every EMF and WMF part in the corpus parses, and plays to the end without an exception.
import { expect, test } from 'vitest';
import { canvasSize, drawMetafile, hasDrawing, MetafileError, parseMetafile, type Metafile } from '../src/index.ts';
import { corpusMetafiles } from './corpus.ts';
import { MockCanvas, RecordingContext } from './mock.ts';

const pictures = corpusMetafiles();

// Parts named .emf that are something else (a PNG in LibreOffice's graphic-stroke test deck).
const NOT_METAFILES = new Set(['libreoffice/oox~qa~unit~data~graphic-stroke.pptx!ppt/media/image1.emf']);

test('the corpus has the metafiles this suite expects', () => {
  if (!pictures.length) return;            // corpus not downloaded
  const kinds = pictures.filter(p => !NOT_METAFILES.has(p.name)).map(p => parseMetafile(p.bytes).format);
  expect(kinds.length).toBeGreaterThanOrEqual(139);
  expect(kinds.filter(k => k === 'emf').length).toBeGreaterThanOrEqual(30);
  expect(kinds.filter(k => k === 'wmf').length).toBeGreaterThanOrEqual(100);
});

test.each(pictures.map(p => [p.name, p] as const))('%s', (name, p) => {
  if (NOT_METAFILES.has(name)) {
    expect(() => parseMetafile(p.bytes)).toThrow(expect.objectContaining({ code: 'unsupported' }));
    return;
  }
  const mf: Metafile = parseMetafile(p.bytes);
  expect(mf.records.length).toBeGreaterThan(0);
  expect(mf.skipped).toBe(0);
  expect(Number.isFinite(mf.width) && Number.isFinite(mf.height)).toBe(true);
  const { width, height } = canvasSize(mf);
  expect(width * height).toBeLessThanOrEqual(16 * 1024 * 1024);
  if (!hasDrawing(mf)) return;
  expect(Math.max(width, height)).toBeGreaterThanOrEqual(1024);
  const ctx = new RecordingContext(width, height);
  try {
    drawMetafile(ctx, mf, { x: 0, y: 0, w: width, h: height }, { createCanvas: (w, h) => new MockCanvas(w, h), snap: true });
  } catch (e) {
    expect.fail(`${e instanceof MetafileError ? `MetafileError ${e.code}` : 'exception'}: ${String(e)}`);
  }
  expect(ctx.draws.length).toBeGreaterThan(0);
  expect(ctx.depth).toBe(0);
});
