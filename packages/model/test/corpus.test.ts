// The model over the whole corpus: every part, shape and property must be readable without an exception, and a
// typical edit (moving every shape on the first slide) must change that slide part only, in valid schema order.
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  checkOrder, child, effectsIn, fillIn, paragraphsOf, parseFill, parseBodyProps, parseLine, parseListStyle, parseParaProps, parseRunProps,
  parseStyle, parseTable, PresentationDocument, runsOf, type ShapeView,
} from '../src/index.ts';
import { corpusFiles, corpusName } from '../../../tools/corpus/list.ts';
import { expectedRefusals, packageDifferences } from '../../opc/test/corpus-helpers.ts';

// Decks whose XML parts are not well-formed: the package round-trips (untouched parts keep their bytes), but the
// model cannot read the broken part and says so with OpcError 'invalid-part', as PowerPoint would refuse it.
const MALFORMED_PARTS = new Set(['libreoffice/sd~qa~unit~data~pptx~fail~ofz35597-1.pptx']);
const decks = corpusFiles.filter(f => !expectedRefusals[corpusName(f)] && !MALFORMED_PARTS.has(corpusName(f)));

test.each([...MALFORMED_PARTS])('%s: a malformed part is reported, not misread', name => {
  const file = corpusFiles.find(f => corpusName(f) === name);
  if (!file) return;
  const doc = new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync(file))));
  expect(() => doc.slideParts().map(p => doc.slide(p).shapes)).toThrow(/is not well-formed|expected|offset/);
});

function readShape(s: ShapeView, counts: { shapes: number; runs: number }): void {
  counts.shapes++;
  void [s.id, s.name, s.placeholder, s.xfrm, s.childSpace, s.geometry, s.media, s.hidden, s.txXfrm];
  parseStyle(s.style);
  fillIn(s.spPr);
  parseLine(child(s.spPr, 'a:ln'));
  effectsIn(s.spPr);
  if (s.blipFill) parseFill(s.blipFill);
  const body = s.txBody;
  if (body) {
    parseBodyProps(child(body, 'a:bodyPr'));
    parseListStyle(child(body, 'a:lstStyle'));
    for (const p of paragraphsOf(body)) {
      parseParaProps(child(p, 'a:pPr'));
      for (const r of runsOf(p)) {
        parseRunProps(r.rPr);
        counts.runs++;
      }
    }
  }
  const gd = s.graphicData;
  const tbl = gd && child(gd.el, 'a:tbl');
  if (tbl) parseTable(tbl);
  for (const c of s.children) readShape(c, counts);
}

describe.skipIf(decks.length === 0)('the model reads every corpus deck', () => {
  test.each(decks.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    const doc = new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync(file))));
    const counts = { shapes: 0, runs: 0 };
    const pres = doc.presentation;
    void [pres.slideSize, pres.notesSize, pres.sections, pres.embeddedFonts, pres.firstSlideNum];
    parseListStyle(pres.defaultTextStyle);
    const parts = new Set<string>();
    for (const slide of doc.slideParts()) {
      parts.add(slide);
      const layout = doc.layoutOf(slide);
      if (layout) parts.add(layout);
      const notes = doc.notesOf(slide);
      if (notes) parts.add(notes);
    }
    for (const master of doc.masterParts()) {
      parts.add(master);
      for (const l of doc.layoutsOf(master)) parts.add(l);
      const theme = doc.themeOf(master);
      if (theme) doc.theme(theme);
    }
    for (const part of parts) {
      const view = doc.slide(part);
      void [view.clrMap, view.clrMapOverride, view.txStyles, view.showMasterSp, view.layoutType, view.transition(), view.bg];
      for (const s of view.shapes) readShape(s, counts);
    }
    expect(counts.shapes).toBeGreaterThanOrEqual(0);
  });
});

describe.skipIf(decks.length === 0)('moving every shape on the first slide changes only that part, in schema order', () => {
  test.each(decks.map(f => [corpusName(f), f]))('%s', (_name, file) => {
    const original = new Uint8Array(readFileSync(file));
    const doc = new PresentationDocument(OpcPackage.open(original));
    const first = doc.slideParts()[0];
    if (!first) return;
    const slide = doc.slide(first);
    const before = checkOrder(slide.root).length;
    const moved = new Map<number, { x: number; y: number }>();
    for (const s of slide.shapes) {
      const x = s.xfrm;
      if (!x || s.kind === 'contentPart') continue;
      s.setXfrm({ ...x, x: x.x + 12700, y: x.y + 12700 });
      moved.set(s.id, { x: x.x + 12700, y: x.y + 12700 });
    }
    if (moved.size === 0) return;
    const saved = doc.pkg.save();
    const changed = packageDifferences(original, saved).map(d => d.split(':')[0]);
    expect(new Set(changed)).toEqual(new Set([first.slice(1)]));
    const reopened = new PresentationDocument(OpcPackage.open(saved));
    const again = reopened.slide(first);
    for (const s of again.shapes) {
      const want = moved.get(s.id);
      if (want) expect({ x: s.xfrm?.x, y: s.xfrm?.y }).toEqual(want);
    }
    expect(checkOrder(again.root).length).toBeLessThanOrEqual(before);
  });
});
