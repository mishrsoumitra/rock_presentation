import { describe, expect, test } from 'vitest';
import { childElements, getAttr, parseXml, XmlParseError, XML_NS, XMLNS_NS, type XmlElement, type XmlText } from '../src/index.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P_STRICT = 'http://purl.oclc.org/ooxml/presentationml/main';

const SLIDE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree>` +
  '<p:pic><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic>' +
  '<p:sp><p:txBody><a:p><a:r><a:t>Q&amp;A &lt;draft&gt; "quoted"</a:t></a:r></a:p></p:txBody></p:sp>' +
  '</p:spTree></p:cSld></p:sld>';

const first = (el: XmlElement, local: string): XmlElement => {
  const found = childElements(el).find(c => c.local === local);
  if (!found) throw new Error(`no ${local} under ${el.local}`);
  return found;
};
const textOf = (el: XmlElement) => el.children.map(c => (c.kind === 'text' || c.kind === 'cdata' ? c.value : '')).join('');

/** Assert that parsing fails with an XmlParseError whose message contains `fragment`. */
function rejects(src: string, fragment: string, options?: { maxDepth?: number }) {
  let error: unknown;
  try {
    parseXml(src, options);
  } catch (e) {
    error = e;
  }
  expect(error, `expected ${JSON.stringify(src.slice(0, 60))} to be rejected`).toBeInstanceOf(XmlParseError);
  expect((error as XmlParseError).message).toContain(fragment);
  expect((error as XmlParseError).offset).toBeGreaterThanOrEqual(0);
}

describe('parseXml: what PowerPoint writes', () => {
  test('keeps the declaration, the CRLF after it, and the tree', () => {
    const doc = parseXml(SLIDE);
    expect(doc.declaration).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(doc.prolog).toEqual([{ kind: 'text', value: '\r\n' }]);
    expect(doc.root.local).toBe('sld');
    expect(doc.root.prefix).toBe('p');
    expect(doc.root.parent).toBeNull();
    const spTree = first(first(doc.root, 'cSld'), 'spTree');
    expect(childElements(spTree).map(c => c.local)).toEqual(['pic', 'sp']);
  });

  test('resolves element and attribute namespaces', () => {
    const doc = parseXml(SLIDE);
    expect(doc.root.ns).toBe(P);
    const blip = first(first(first(first(first(doc.root, 'cSld'), 'spTree'), 'pic'), 'blipFill'), 'blip');
    expect(blip.ns).toBe(A);
    expect(blip.attrs).toEqual([{ prefix: 'r', local: 'embed', value: 'rId2', ns: R }]);
    expect(getAttr(blip, 'embed', R)).toBe('rId2');
    expect(getAttr(blip, 'embed')).toBeUndefined();
    expect(doc.root.attrs.every(a => a.ns === XMLNS_NS)).toBe(true);
  });

  test('decodes the five predefined entities and character references', () => {
    const doc = parseXml('<t a="&quot;&apos;&lt;&gt;&amp;">&#x41;&#66;&#x1F600;&apos;</t>');
    expect(getAttr(doc.root, 'a')).toBe(`"'<>&`);
    expect(textOf(doc.root)).toBe('AB\u{1F600}\'');
    const t = first(first(first(first(first(first(parseXml(SLIDE).root, 'cSld'), 'spTree'), 'sp'), 'txBody'), 'p'), 'r');
    expect(textOf(first(t, 't'))).toBe('Q&A <draft> "quoted"');
  });

  test('keeps self-closing versus empty element forms', () => {
    const doc = parseXml('<r><a/><b></b></r>');
    expect(childElements(doc.root).map(c => [c.local, c.selfClosing])).toEqual([['a', true], ['b', false]]);
  });

  test('sets parent pointers', () => {
    const doc = parseXml('<r><a><b/></a></r>');
    const a = first(doc.root, 'a');
    expect(a.parent).toBe(doc.root);
    expect(first(a, 'b').parent).toBe(a);
  });

  test('default namespace, undeclaring it, and redeclaring a prefix deeper in the tree', () => {
    const doc = parseXml(`<Types xmlns="urn:ct"><Default/><x xmlns=""><y/></x><p:q xmlns:p="urn:one"><p:q xmlns:p="urn:two"/></p:q></Types>`);
    expect(doc.root.ns).toBe('urn:ct');
    expect(first(doc.root, 'Default').ns).toBe('urn:ct');
    expect(first(first(doc.root, 'x'), 'y').ns).toBe('');
    const outer = first(doc.root, 'q');
    expect(outer.ns).toBe('urn:one');
    expect(first(outer, 'q').ns).toBe('urn:two');
  });

  test('mc:AlternateContent with namespaces declared on the Choice, as PowerPoint 2010 and later write it', () => {
    const doc = parseXml(
      `<p:sld xmlns:p="${P}"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
      `<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition p14:dur="2000"/></mc:Choice>` +
      `<mc:Fallback><p:transition/></mc:Fallback></mc:AlternateContent></p:sld>`);
    const choice = first(first(doc.root, 'AlternateContent'), 'Choice');
    const transition = first(choice, 'transition');
    expect(transition.attrs[0]).toEqual({ prefix: 'p14', local: 'dur', value: '2000', ns: 'http://schemas.microsoft.com/office/powerpoint/2010/main' });
  });

  test('Strict OOXML namespaces resolve like any other', () => {
    expect(parseXml(`<p:presentation xmlns:p="${P_STRICT}"/>`).root.ns).toBe(P_STRICT);
  });

  test('the xml prefix is always bound', () => {
    const doc = parseXml('<a:t xmlns:a="urn:a" xml:space="preserve"> x </a:t>');
    expect(doc.root.attrs[1]).toEqual({ prefix: 'xml', local: 'space', value: 'preserve', ns: XML_NS });
    expect(textOf(doc.root)).toBe(' x ');
  });

  test('normalizes line ends in text the way every XML reader does, but keeps a CR written as a reference', () => {
    const doc = parseXml('<t>a\r\nb\rc&#xD;d&#13;\r\n</t>');
    expect(textOf(doc.root)).toBe('a\nb\nc\rd\r\n');
  });

  test('normalizes whitespace in attribute values, but keeps whitespace written as references', () => {
    const doc = parseXml('<t a="x\ty\nz\r\nw" b="1&#xA;2&#9;3&#xD;"/>');
    expect(getAttr(doc.root, 'a')).toBe('x y z w');
    expect(getAttr(doc.root, 'b')).toBe('1\n2\t3\r');
  });

  test('accepts single-quoted attributes and whitespace around =', () => {
    const doc = parseXml(`<?xml version='1.0' encoding='UTF-8' standalone='yes'?><t a = 'say "hi"' b="it's"/>`);
    expect(getAttr(doc.root, 'a')).toBe('say "hi"');
    expect(getAttr(doc.root, 'b')).toBe("it's");
  });

  test('keeps CDATA, comments, prolog and epilog', () => {
    const doc = parseXml('<!-- before -->\n<r><![CDATA[<a> & b\r\n]]><!-- inside --></r>\n<!-- after -->\n');
    expect(doc.declaration).toBeNull();
    expect(doc.prolog).toEqual([{ kind: 'comment', value: ' before ' }, { kind: 'text', value: '\n' }]);
    expect(doc.root.children).toEqual([{ kind: 'cdata', value: '<a> & b\n' }, { kind: 'comment', value: ' inside ' }]);
    expect(doc.epilog).toEqual([{ kind: 'text', value: '\n' }, { kind: 'comment', value: ' after ' }, { kind: 'text', value: '\n' }]);
  });

  test('keeps processing instructions, which Office writes at the top of SharePoint customXml parts', () => {
    const doc = parseXml('<?mso-contentType?><FormTemplates xmlns="http://schemas.microsoft.com/sharepoint/v3/contenttype/forms"><?pi  a\r\nb ?></FormTemplates>');
    expect(doc.prolog).toEqual([{ kind: 'pi', target: 'mso-contentType', data: '' }]);
    expect(doc.root.children).toEqual([{ kind: 'pi', target: 'pi', data: 'a\nb ' }]);
    expect(parseXml('<?xml version="1.0"?><?xml-stylesheet href="a.xsl"?><r/>').prolog).toEqual([{ kind: 'pi', target: 'xml-stylesheet', data: 'href="a.xsl"' }]);
  });

  test('skips a byte order mark left in the string', () => {
    expect(parseXml('﻿<?xml version="1.0"?><r/>').declaration).toBe('<?xml version="1.0"?>');
  });

  test('accepts names outside ASCII, as custom XML parts may use them', () => {
    const doc = parseXml('<データ 属性="値"><élément/></データ>');
    expect(doc.root.local).toBe('データ');
    expect(getAttr(doc.root, '属性')).toBe('値');
  });

  test('accepts whitespace before > in a close tag', () => {
    expect(parseXml('<a:t xmlns:a="urn:a">x</a:t  >').root.local).toBe('t');
  });

  test('nesting up to maxDepth is accepted', () => {
    expect(() => parseXml('<a>'.repeat(256) + '</a>'.repeat(256))).not.toThrow();
  });
});

describe('parseXml: what it refuses', () => {
  test('DOCTYPE, so entity expansion and XXE cannot happen', () => {
    rejects('<!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;">]><lolz>&lol2;</lolz>', 'DOCTYPE');
    rejects('<?xml version="1.0"?>\n<!DOCTYPE r SYSTEM "file:///etc/passwd"><r/>', 'DOCTYPE');
    rejects('<r><!ELEMENT x ANY></r>', 'DOCTYPE');
  });

  test('an XML declaration anywhere but the very start, and malformed processing instructions', () => {
    rejects('<r/><?xml version="1.0"?>', 'only allowed at the very start');
    rejects(' <?xml version="1.0"?><r/>', 'only allowed at the very start');
    rejects('<r><?XmL x?></r>', 'only allowed at the very start');
    rejects('<r><?a:b x?></r>', 'cannot contain a colon');
    rejects('<r><?pi', 'unterminated processing instruction');
    rejects('<r><?pi"x"?></r>', 'expected whitespace after a processing instruction target');
    rejects('<r><? pi?></r>', 'where a name was expected');
  });

  test('entities other than the five predefined ones, including Object.prototype names', () => {
    rejects('<r>&nbsp;</r>', 'undefined entity &nbsp;');
    rejects('<r>&constructor;</r>', 'undefined entity &constructor;');
    rejects('<r a="&toString;"/>', 'undefined entity &toString;');
    rejects('<r>&__proto__;</r>', 'bare &');
  });

  test('bare ampersands', () => {
    rejects('<r>Q&A</r>', 'bare &');
    rejects('<r a="Q&A"/>', 'bare &');
    rejects('<r>&#;</r>', 'bare &');
    rejects('<r>&#xZZ;</r>', 'bare &');
  });

  test('character references to characters XML does not allow', () => {
    rejects('<r>&#0;</r>', 'not an XML character');
    rejects('<r>&#x1;</r>', 'not an XML character');
    rejects('<r>&#xD800;</r>', 'not an XML character');
    rejects('<r>&#xFFFE;</r>', 'not an XML character');
    rejects('<r>&#x110000;</r>', 'not an XML character');
    rejects('<r>&#99999999999999999999;</r>', 'not an XML character');
  });

  test('raw characters XML does not allow', () => {
    rejects('<r>\u0001</r>', 'U+0001');
    rejects('<r a="\u0000"/>', 'U+0000');
    rejects('<r>￿</r>', 'U+FFFF');
    rejects('<r>\uD800x</r>', 'U+D800');
    rejects('<r>x\uDC00</r>', 'U+DC00');
  });

  test('broken structure', () => {
    rejects('<a><b></a></b>', 'mismatched close tag </a> for <b>');
    rejects('<a>', 'unclosed element <a>');
    rejects('</a>', 'close tag without an open tag');
    rejects('<a></a></a>', 'close tag without an open tag');
    rejects('<a/><b/>', 'more than one root element');
    rejects('x<a/>', 'text before the root element');
    rejects('<a/>x', 'text after the root element');
    rejects('', 'no root element');
    rejects('<?xml version="1.0"?>', 'no root element');
    rejects('<![CDATA[x]]><a/>', 'CDATA outside the root element');
    rejects('<a></ a>', 'where a name was expected');
    rejects('<a></a x>', 'expected > at the end of a close tag');
  });

  test('broken tags and attributes', () => {
    rejects('<a b=c/>', 'attribute value must be quoted');
    rejects('<a b="1" b="2"/>', 'duplicate attribute b');
    rejects('<a b="1"c="2"/>', 'expected whitespace before an attribute');
    rejects('<a b="x<y"/>', '< in an attribute value');
    rejects('<a b/>', 'expected = after an attribute name');
    rejects('<a b="1', 'unterminated attribute value');
    rejects('<a b="1"', 'unterminated open tag');
    rejects('<a /x>', 'expected />');
    rejects(
      '<a xmlns:p="urn:same" xmlns:q="urn:same" p:x="1" q:x="2"/>',
      'have the same expanded name',
    );
  });

  test('names that are not namespace QNames', () => {
    rejects('<1a/>', 'where a name was expected');
    rejects('< a/>', 'where a name was expected');
    rejects('<:a/>', 'where a name was expected');
    rejects('<a:b:c xmlns:a="urn:a"/>', 'expected whitespace before an attribute');
  });

  test('namespace errors', () => {
    rejects('<p:sld/>', 'undeclared namespace prefix p');
    rejects('<r r:id="rId1"/>', 'undeclared namespace prefix r');
    rejects('<p:a xmlns:p="urn:p"><p:b/></p:a><p:c/>', 'undeclared namespace prefix p');   // the scope ended with the root
    rejects('<a xmlns:p=""/>', 'xmlns:p cannot be empty');
    rejects('<a xmlns:xml="urn:other"/>', 'cannot be bound');
    rejects('<a xmlns:x="http://www.w3.org/XML/1998/namespace"/>', 'cannot be bound');
    rejects('<a xmlns="http://www.w3.org/2000/xmlns/"/>', 'cannot be bound');
    rejects('<xmlns:a xmlns:xmlns="urn:x"/>', 'cannot be bound');
  });

  test('text, comment and CDATA rules', () => {
    rejects('<r>a ]]> b</r>', "']]>' is not allowed in text");
    rejects('<r><!-- a -- b --></r>', "'--' inside a comment");
    rejects('<r><!-- a ---></r>', "'--' inside a comment");
    rejects('<r><!-- open</r>', 'unterminated comment');
    rejects('<r><![CDATA[open</r>', 'unterminated CDATA section');
    rejects('<?xml version="1.0"', 'unterminated XML declaration');
  });

  test('declarations that do not follow the XMLDecl grammar', () => {
    rejects('<?xml foo?><r/>', 'malformed XML declaration');
    rejects('<?xml encoding="UTF-8"?><r/>', 'malformed XML declaration');
    rejects('<?xml version="2.0"?><r/>', 'malformed XML declaration');
    rejects('<?xml version="1.0" standalone="yes" encoding="UTF-8"?><r/>', 'malformed XML declaration');
    rejects('<?xml version="1.0" standalone="maybe"?><r/>', 'malformed XML declaration');
    rejects(`<?xml version="1.0'?><r/>`, 'malformed XML declaration');
    for (const ok of ['<?xml version="1.0"?>', `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>`, '<?xml version = "1.0"\n encoding="utf-16" ?>']) {
      expect(parseXml(ok + '<r/>').declaration).toBe(ok);
    }
  });

  test('nesting deeper than the limit', () => {
    rejects('<a>'.repeat(257) + '</a>'.repeat(257), 'nested deeper than 256 levels');
    rejects('<a><b><c/></b></a>', 'nested deeper than 2 levels', { maxDepth: 2 });
  });

  test('the error offset points at the problem', () => {
    try {
      parseXml('<r><a></b></r>');
      expect.unreachable();
    } catch (e) {
      expect((e as XmlParseError).offset).toBe(6);
    }
  });
});

describe('parseXml: text nodes', () => {
  test('text is split only where markup splits it', () => {
    const doc = parseXml('<r>a<!--c-->b<![CDATA[c]]>d</r>');
    expect(doc.root.children.map(c => c.kind)).toEqual(['text', 'comment', 'text', 'cdata', 'text']);
    expect((doc.root.children[0] as XmlText).value).toBe('a');
  });
});
