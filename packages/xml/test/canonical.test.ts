import { describe, expect, test } from 'vitest';
import { canonical, canonicalDiff, parseXml } from '../src/index.ts';

const canon = (s: string) => canonical(parseXml(s));
const same = (a: string, b: string) => expect(canon(a)).toBe(canon(b));
const differ = (a: string, b: string) => expect(canon(a)).not.toBe(canon(b));

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const P_STRICT = 'http://purl.oclc.org/ooxml/presentationml/main';

describe('canonical', () => {
  test('prefixes do not matter, namespaces do', () => {
    same(`<p:sld xmlns:p="${P}"><p:cSld/></p:sld>`, `<x:sld xmlns:x="${P}"><x:cSld/></x:sld>`);
    same(`<p:sld xmlns:p="${P}"><p:cSld/></p:sld>`, `<sld xmlns="${P}"><cSld/></sld>`);
    differ(`<p:sld xmlns:p="${P}"/>`, `<p:sld xmlns:p="${P_STRICT}"/>`);   // Strict is not Transitional
  });

  test('prefixed attributes compare by namespace; unprefixed ones have none', () => {
    same('<r xmlns:r="urn:r" r:id="1"/>', '<r xmlns:q="urn:r" q:id="1"/>');
    differ('<r xmlns:r="urn:r" r:id="1"/>', '<r id="1"/>');
  });

  test('attribute order and namespace declarations do not matter', () => {
    same('<r a="1" b="2" xmlns:x="urn:x"/>', '<r b="2" a="1"/>');
    differ('<r a="1"/>', '<r a="2"/>');
  });

  test('indentation between elements does not matter, text in leaf elements does', () => {
    same('<r>\n  <a/>\r\n  <b/>\n</r>', '<r><a/><b/></r>');
    differ('<r><t> </t></r>', '<r><t></t></r>');
    differ('<r><t>a</t></r>', '<r><t>a </t></r>');
  });

  test('comments, CDATA versus escaped text, and self-closing form do not matter', () => {
    same('<r>a<!-- note -->b</r>', '<r>ab</r>');
    same('<r><![CDATA[<x> & y]]></r>', '<r>&lt;x&gt; &amp; y</r>');
    same('<r><a/></r>', '<r><a></a></r>');
  });

  test('processing instructions matter, before, inside and after the root', () => {
    differ('<?mso-contentType?><r/>', '<r/>');
    differ('<r><?pi a?></r>', '<r><?pi b?></r>');
    differ('<r/><?after?>', '<r/>');
    same('<?pi  a?><r/>', '<?pi a?><r/>');
  });

  test('element order and nesting do matter', () => {
    differ('<r><a/><b/></r>', '<r><b/><a/></r>');
    differ('<r><a><b/></a></r>', '<r><a/><b/></r>');
  });

  test('canonicalDiff points at the first difference', () => {
    expect(canonicalDiff('abc', 'abc')).toBeNull();
    expect(canonicalDiff(canon('<r><a x="1"/></r>'), canon('<r><a x="2"/></r>'))).toContain('first difference at');
  });
});
