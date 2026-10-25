// Every slide, layout and master of every corpus deck resolves without an exception, and every number that comes
// out is finite: a NaN here would draw nothing, or draw everywhere, in the renderer.
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { resolveSlide, type ResolvedShape } from '../src/index.ts';
import { corpusFiles, corpusName } from '../../../tools/corpus/list.ts';
import { expectedRefusals } from '../../opc/test/corpus-helpers.ts';

const MALFORMED_PARTS = new Set(['libreoffice/sd~qa~unit~data~pptx~fail~ofz35597-1.pptx']);
const decks = corpusFiles.filter(f => !expectedRefusals[corpusName(f)] && !MALFORMED_PARTS.has(corpusName(f)));

function checkShape(s: ResolvedShape, where: string): void {
  for (const v of [...s.transform, ...s.textTransform]) expect(Number.isFinite(v), `${where} ${s.name} transform`).toBe(true);
  for (const p of s.text?.paragraphs ?? []) {
    for (const r of p.runs) expect(Number.isFinite(r.props.size) && r.props.size >= 0, `${where} ${s.name} size`).toBe(true);
  }
  for (const c of s.children) checkShape(c, where);
}

describe.skipIf(decks.length === 0)('every corpus slide, layout and master resolves', () => {
  test.each(decks.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    const doc = new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync(file))));
    const parts = new Set(doc.slideParts());
    for (const m of doc.masterParts()) {
      parts.add(m);
      for (const l of doc.layoutsOf(m)) parts.add(l);
    }
    for (const part of parts) {
      const r = resolveSlide(doc, part);
      expect(Number.isFinite(r.size.cx) && Number.isFinite(r.size.cy)).toBe(true);
      for (const s of r.shapes) checkShape(s, part);
    }
  });
});
