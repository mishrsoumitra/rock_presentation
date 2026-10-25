// Differential test against saxes, a conformant XML parser used here only as a reference (never in the product).
// Where both parsers accept a document they must report the same elements, namespaces, attributes and text.
// Where they disagree, the only allowed difference is that ours refuses a DOCTYPE.
import { SaxesParser } from 'saxes';
import { describe, expect, test } from 'vitest';
import { parseXml, XmlParseError, XMLNS_NS, type XmlElement, type XmlMisc } from '../src/index.ts';
import { xmlSamples } from '../../../tools/corpus/samples.ts';
import { mutate, rng } from './mutate.ts';

type Events = { ok: true; events: string[] } | { ok: false; error: string };

function attrList(attrs: { uri: string; local: string; value: string }[]): string {
  return attrs.map(a => `{${a.uri}}${a.local}=${JSON.stringify(a.value)}`).sort().join(' ');
}

const piEvent = (target: string, data: string) => `P${target} ${JSON.stringify(data)}`;

function saxesEvents(src: string): Events & { doctype?: boolean } {
  const parser = new SaxesParser({ xmlns: true });
  const events: string[] = [];
  let text = '';
  let depth = 0;
  let doctype = false;
  const flush = () => {
    if (text) events.push('T' + JSON.stringify(text));
    text = '';
  };
  parser.on('opentag', node => {
    flush();
    depth++;
    const attrs = Object.values(node.attributes).filter(a => a.uri !== XMLNS_NS);
    events.push(`<{${node.uri}}${node.local} ${attrList(attrs)}`);
  });
  parser.on('closetag', () => {
    flush();
    depth--;
    events.push('>');
  });
  parser.on('text', t => { if (depth > 0) text += t; });
  parser.on('cdata', t => { text += t; });
  parser.on('doctype', () => { doctype = true; });
  parser.on('processinginstruction', pi => {
    flush();
    events.push(piEvent(pi.target, pi.body));
  });
  try {
    parser.write(src).close();
    return { ok: true, events, doctype };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e), doctype };
  }
}

function ourEvents(src: string): Events {
  let doc;
  try {
    doc = parseXml(src);
  } catch (e) {
    if (!(e instanceof XmlParseError)) throw e;
    return { ok: false, error: e.message };
  }
  const events: string[] = [];
  // saxes trims namespace declarations; Namespaces in XML 1.0 (2.3) compares namespace names character for
  // character, so we keep " urn:x" as written. Trim here so the rest of the event stream is still compared exactly.
  const walk = (el: XmlElement) => {
    events.push(`<{${el.ns.trim()}}${el.local} ${attrList(el.attrs.filter(a => a.ns !== XMLNS_NS).map(a => ({ uri: a.ns.trim(), local: a.local, value: a.value })))}`);
    let text = '';
    for (const c of el.children) {
      if (c.kind === 'element' || c.kind === 'pi') {
        if (text) events.push('T' + JSON.stringify(text));
        text = '';
        if (c.kind === 'element') walk(c);
        else events.push(piEvent(c.target, c.data));
      } else if (c.kind !== 'comment') text += c.value;
    }
    if (text) events.push('T' + JSON.stringify(text));
    events.push('>');
  };
  const pis = (misc: XmlMisc[]) => { for (const m of misc) if (m.kind === 'pi') events.push(piEvent(m.target, m.data)); };
  pis(doc.prolog);
  walk(doc.root);
  pis(doc.epilog);
  return { ok: true, events };
}

// saxes 6.0.0 lets an unpaired high surrogate through in text and attribute values (it pairs it with whatever
// follows). XML's Char production excludes U+D800-U+DFFF, so such input must be refused, whatever saxes says.
const UNPAIRED_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

function compare(src: string): 'both-accept' | 'both-reject' | 'we-are-stricter' {
  const ours = ourEvents(src);
  if (UNPAIRED_SURROGATE.test(src)) {
    expect(ours.ok, `accepted an unpaired surrogate: ${JSON.stringify(src.slice(0, 200))}`).toBe(false);
    return 'we-are-stricter';
  }
  const theirs = saxesEvents(src);
  if (theirs.ok && ours.ok) {
    expect(ours.events, `events for ${JSON.stringify(src.slice(0, 200))}`).toEqual(theirs.events);
    return 'both-accept';
  }
  if (!theirs.ok && !ours.ok) return 'both-reject';
  if (theirs.ok && !ours.ok) {
    // allowed only when the document has a DOCTYPE, which OOXML parts never contain
    expect(theirs.doctype, `we rejected (${ours.error}) what saxes accepted: ${JSON.stringify(src.slice(0, 300))}`).toBe(true);
    return 'we-are-stricter';
  }
  throw new Error(`saxes rejected (${(theirs as { error: string }).error}) what we accepted: ${JSON.stringify(src.slice(0, 300))}`);
}

describe('differential test against saxes', () => {
  const samples = xmlSamples({ files: 120, perFile: 10 });

  test.skipIf(samples.length === 0)(`${samples.length} corpus parts parse to the same events`, () => {
    for (const s of samples) expect(compare(s.text), s.name).toBe('both-accept');
  });

  test('mutated corpus parts are accepted and rejected alike', () => {
    const rnd = rng(0xd1ff);
    const seeds = samples.filter(s => s.text.length < 20_000).slice(0, 150).map(s => s.text);
    seeds.push('<r xmlns:a="urn:a"><a:x a:y="1">t<![CDATA[c]]></a:x><!-- k --></r>', '<?xml version="1.0"?>\n<r a="&#9;\t&#10;\r\n"/>');
    const tally: Record<string, number> = {};
    // at least 2,500 comparisons: without the corpus (CI) the two hand-written seeds take many rounds each
    const rounds = Math.max(25, Math.ceil(2500 / seeds.length));
    for (let round = 0; round < rounds; round++) {
      for (const seed of seeds) {
        const kind = compare(mutate(seed, rnd));
        tally[kind] = (tally[kind] ?? 0) + 1;
      }
    }
    expect(tally['both-accept']).toBeGreaterThan(50);
    expect(tally['both-reject']).toBeGreaterThan(50);
  }, 60_000);

  test('hand-picked edge cases agree', () => {
    for (const src of [
      '<r>a\r\nb\rc&#13;</r>',
      '<r a="x\ty\nz\r\nw" b="&#9;&#10;&#13;"/>',
      '<r xmlns="urn:d"><x xmlns=""><y/></x></r>',
      '<r>&lt;&gt;&amp;&quot;&apos;&#x10FFFF;</r>',
      '﻿<r/>',
      '<r/>\n<!-- after -->\n',
      '<r><![CDATA[]]]]><![CDATA[>]]></r>',
      '<データ 属性="値"/>',
      '<?mso-contentType?><FormTemplates xmlns="http://schemas.microsoft.com/sharepoint/v3/contenttype/forms"/>',
      '<r>a<?pi  x\r\ny ?>b</r><?after?>',
    ]) expect(compare(src)).toBe('both-accept');
    for (const src of ['<r>&nbsp;</r>', '<r>]]></r>', '<r a="1" a="2"/>', '<p:r/>', '<r><!-- -- --></r>', '<r>&#0;</r>', '<?xml version="1.0" foo="x"?><r/>']) {
      expect(compare(src)).toBe('both-reject');
    }
    expect(compare('<!DOCTYPE r><r/>')).toBe('we-are-stricter');
    expect(compare('<r/><?xml version="1.0"?>')).toBe('both-reject');
  });
});
