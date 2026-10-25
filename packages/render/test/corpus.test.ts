// Every slide of every corpus deck builds and draws without an exception, and nothing is drawn with a non-finite
// transform (a NaN there draws nothing, or everything). Charts are drawn by @pptx/charts when it is wired in.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { FontRegistry } from '@pptx/text';
import { corpusFiles, corpusName } from '../../../tools/corpus/list.ts';
import { expectedRefusals } from '../../opc/test/corpus-helpers.ts';
import { buildSlide, drawSlide } from '../src/index.ts';
import { installStubs, recorder } from './canvas-stubs.ts';

installStubs();
const FONTS = join(import.meta.dirname, '../../../assets/fonts');
const fonts = new FontRegistry();
for (const dir of ['carlito', 'caladea', 'liberation']) {
  for (const f of readdirSync(join(FONTS, dir))) if (f.endsWith('.ttf')) fonts.addBytes(new Uint8Array(readFileSync(join(FONTS, dir, f))));
}

const MALFORMED_PARTS = new Set(['libreoffice/sd~qa~unit~data~pptx~fail~ofz35597-1.pptx']);
const decks = corpusFiles.filter(f => !expectedRefusals[corpusName(f)] && !MALFORMED_PARTS.has(corpusName(f)));

describe.skipIf(decks.length === 0)('every corpus slide builds and draws', () => {
  test.each(decks.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    const doc = new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync(file))));
    for (const part of doc.slideParts()) {
      const slide = buildSlide(resolveSlide(doc, part), { fonts });
      const { ctx, calls } = recorder();
      drawSlide(ctx as never, slide, { scale: 1280 / slide.size.cx, image: () => undefined });
      for (const c of calls) for (const v of c.transform) expect(Number.isFinite(v), `${part} ${c.op}`).toBe(true);
    }
  });
});
