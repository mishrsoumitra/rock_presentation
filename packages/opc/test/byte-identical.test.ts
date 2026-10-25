// Plan section 1.5, second test: parse then serialize is byte-identical for parts PowerPoint wrote. Stricter than
// preservation needs (untouched parts keep their bytes anyway); it catches serializer drift early.
//
// "PowerPoint wrote it" is decided per part, not per package: test corpora often hold decks whose app.xml says
// PowerPoint but whose parts were pretty-printed or edited by hand or by another library. PowerPoint's own XML is a
// declaration, CR LF, then the whole tree on one line, and it never writes "<x />" with a space.
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { decodeXmlBytes, parseXml, serializeXml } from '@pptx/xml';
import { OpcPackage } from '../src/index.ts';
import { corpusName, opcFiles } from '../../../tools/corpus/list.ts';
import { expectedRefusals, isXmlEntry } from './corpus-helpers.ts';

/** Parts that look machine-written by PowerPoint but were edited by hand, with what was changed. */
const HAND_EDITED: Record<string, string> = {
  'libreoffice/sd~qa~unit~data~pptx~tdf173712.pptx:/ppt/slides/slide1.xml': 'a raw ">" typed into <a:t>; PowerPoint writes &gt;',
};

const looksPowerPointWritten = (text: string) => {
  const decl = /^<\?xml[^?]*\?>\r\n/.exec(text);
  return decl !== null && !/[\r\n]/.test(text.slice(decl[0].length)) && !text.includes(' />');
};

function producedByPowerPoint(pkg: OpcPackage): boolean {
  const app = pkg.getPart('/docProps/app.xml');
  return app !== undefined && /<Application>[^<]*PowerPoint[^<]*<\/Application>/.test(new TextDecoder().decode(app.bytes));
}

const candidates = opcFiles.filter(f => !expectedRefusals[corpusName(f)]);

describe.skipIf(candidates.length === 0)('parse then serialize is byte-identical for PowerPoint-written parts', () => {
  let checked = 0;

  test.each(candidates.map(f => [corpusName(f), f]))('%s', (name, file) => {
    const pkg = OpcPackage.open(new Uint8Array(readFileSync(file)));
    if (!producedByPowerPoint(pkg)) return;
    for (const part of pkg.partNames()) {
      if (!isXmlEntry(part) || HAND_EDITED[`${name}:${part}`]) continue;
      let text: string;
      try {
        text = decodeXmlBytes(pkg.getPart(part)!.bytes).text;
      } catch {
        continue;
      }
      if (!looksPowerPointWritten(text)) continue;
      let doc;
      try {
        doc = parseXml(text);
      } catch {
        continue;            // not well-formed (fuzzer files), so not PowerPoint's output
      }
      expect(serializeXml(doc), part).toBe(text);
      checked++;
    }
  });

  test('a meaningful number of parts was checked', () => {
    expect(checked).toBeGreaterThan(20_000);
  });
});
