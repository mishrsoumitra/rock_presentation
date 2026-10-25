// Mutation fuzzing: random edits of real parts must either parse or fail with XmlParseError, never anything else,
// and whatever parses must survive serialize -> parse unchanged. Seeded, so a failure reproduces.
import { describe, expect, test } from 'vitest';
import { canonical, parseXml, serializeXml, XmlParseError } from '../src/index.ts';
import { xmlSamples } from '../../../tools/corpus/samples.ts';
import { mutate, rng } from './mutate.ts';

const BUILT_IN = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1" descr="a&#xA;b &gt; c"/></p:nvSpPr><p:txBody><a:bodyPr/>' +
  '<a:p><a:r><a:rPr lang="en-US"/><a:t>Q&amp;A "x" \'y\'</a:t></a:r><a:br/><a:fld id="{1}" type="slidenum"><a:t>‹#›</a:t></a:fld></a:p>' +
  '</p:txBody></p:sp></p:spTree></p:cSld><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition p14:dur="2000"/></mc:Choice>' +
  '<mc:Fallback><p:transition/></mc:Fallback></mc:AlternateContent></p:sld>',
  '<r><!-- c --><![CDATA[x < y]]>\r\ntext&#13;<e a="1\t2"/></r>',
];

describe('fuzzing the tokenizer', () => {
  const seeds = [...BUILT_IN, ...xmlSamples({ files: 40, perFile: 3, maxBytes: 20_000 }).map(s => s.text)];

  test(`${seeds.length} seed documents x mutations: only XmlParseError, and accepted input round-trips`, () => {
    const rnd = rng(0x5eed);
    let accepted = 0;
    let rejected = 0;
    const deadline = performance.now() + 8000;
    for (let round = 0; performance.now() < deadline && round < 40; round++) {
      for (const seed of seeds) {
        const input = mutate(seed, rnd);
        let doc;
        try {
          doc = parseXml(input);
        } catch (e) {
          if (!(e instanceof XmlParseError)) throw new Error(`non-XmlParseError ${String(e)} for input ${JSON.stringify(input.slice(0, 300))}`);
          rejected++;
          continue;
        }
        accepted++;
        const written = serializeXml(doc);
        const again = parseXml(written);
        expect(canonical(again), `round trip of ${JSON.stringify(input.slice(0, 200))}`).toBe(canonical(doc));
        expect(serializeXml(again)).toBe(written);
      }
    }
    expect(accepted).toBeGreaterThan(100);
    expect(rejected).toBeGreaterThan(100);
  }, 30_000);
});
