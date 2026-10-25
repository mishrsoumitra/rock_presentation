import { describe, expect, test } from 'vitest';
import { canonical, parseXml, serializeElement, serializeXml, XmlSerializeError, type XmlDocument } from '../src/index.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Written the way PowerPoint writes: CRLF after the declaration, &gt; in text and attributes, &#xA; in alt text,
// raw quotes in text, raw apostrophes in attributes, self-closing empty elements, one empty element written long.
const POWERPOINT_SLIDE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree>` +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture 3" descr="A chart&#xA;&#xA;Description &gt; automatically generated, it&apos;s fine"/>' +
  '</p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic>' +
  '<p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Q&amp;A: "a &lt; b &gt; c"</a:t></a:r>' +
  '<a:r><a:t></a:t></a:r></a:p></p:txBody></p:sp>' +
  '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';

/** What a conforming reader sees in a serialized document must be what it saw in the original. */
function semanticallyStable(src: string): void {
  const once = parseXml(src);
  const written = serializeXml(once);
  expect(canonical(parseXml(written))).toBe(canonical(once));
  expect(serializeXml(parseXml(written))).toBe(written);   // and serialization is a fixed point
}

describe('serializeXml', () => {
  test('a part written the way PowerPoint writes it comes back byte for byte', () => {
    // &apos; in an attribute is the one thing above PowerPoint itself does not write; it comes back as a raw '
    const expected = POWERPOINT_SLIDE.replace('&apos;', "'");
    expect(serializeXml(parseXml(POWERPOINT_SLIDE))).toBe(expected);
    expect(serializeXml(parseXml(expected))).toBe(expected);
  });

  test('escapes text the way PowerPoint does', () => {
    const doc = parseXml('<t/>');
    doc.root.selfClosing = false;
    doc.root.children.push({ kind: 'text', value: `a & b < c > d "e" 'f' \t\n` });
    expect(serializeXml(doc)).toBe(`<t>a &amp; b &lt; c &gt; d "e" 'f' \t\n</t>`);
  });

  test('escapes attribute values the way PowerPoint does', () => {
    const doc = parseXml('<t a=""/>');
    doc.root.attrs[0]!.value = `"x" & <y> 'z'\t\n\r`;
    expect(serializeXml(doc)).toBe(`<t a="&quot;x&quot; &amp; &lt;y&gt; 'z'&#x9;&#xA;&#xD;"/>`);
  });

  test('a CR inside text survives a round trip', () => {
    const doc = parseXml('<t>a&#xD;b</t>');
    expect(serializeXml(doc)).toBe('<t>a&#xD;b</t>');
    semanticallyStable('<t>a&#13;&#10;b\r\nc\rd</t>');
  });

  test('whitespace inside attribute values survives a round trip', () => {
    semanticallyStable('<t a="x&#9;y&#10;z&#13;w" b="raw\ttab\nand\r\nnewlines"/>');
  });

  test('keeps self-closing and long empty forms', () => {
    expect(serializeXml(parseXml('<r><a/><b></b></r>'))).toBe('<r><a/><b></b></r>');
  });

  test('an element marked self-closing but given children is written long', () => {
    const doc = parseXml('<r/>');
    doc.root.children.push({ kind: 'text', value: 'x' });
    expect(serializeXml(doc)).toBe('<r>x</r>');
  });

  test('keeps CDATA, comments, prolog and epilog, and splits a CDATA section that contains ]]>', () => {
    const src = '<?xml version="1.0"?>\n<!-- c --><r><![CDATA[a<b]]><!--x--></r>\n';
    expect(serializeXml(parseXml(src))).toBe(src);
    const doc = parseXml('<r><![CDATA[x]]></r>');
    doc.root.children[0] = { kind: 'cdata', value: 'a]]>b' };
    expect(serializeXml(doc)).toBe('<r><![CDATA[a]]]]><![CDATA[>b]]></r>');
    expect(canonical(parseXml(serializeXml(doc)))).toBe(canonical(parseXml('<r>a]]&gt;b</r>')));
  });

  test('keeps processing instructions in and around the root', () => {
    const src = '<?xml version="1.0"?><?mso-contentType?><r><?pi a b?></r><?after?>';
    expect(serializeXml(parseXml(src))).toBe(src);
    const doc = parseXml('<r/>');
    doc.prolog.push({ kind: 'pi', target: 'xml', data: '' });
    expect(() => serializeXml(doc)).toThrow(XmlSerializeError);
    const data = parseXml('<r/>');
    data.prolog.push({ kind: 'pi', target: 'pi', data: 'a ?> b' });
    expect(() => serializeXml(data)).toThrow(XmlSerializeError);
  });

  test('round trips are semantically stable for awkward inputs', () => {
    for (const src of [
      '<r>&amp;&lt;&gt;&quot;&apos;&#x1F600;</r>',
      `<r a='&apos;"' b="&quot;'" c="&#x20;&#32; "/>`,
      '<r xmlns="urn:d"><x xmlns=""><y/></x></r>',
      '<r>]]&gt;</r>',
      '<r>   ﻿</r>',
      '<r><![CDATA[\r\n]]>\r\n</r>',
      '﻿<r/>',
    ]) semanticallyStable(src);
  });

  test('refuses to write what would not parse back', () => {
    const withText = (value: string): XmlDocument => {
      const doc = parseXml('<t>x</t>');
      doc.root.children[0] = { kind: 'text', value };
      return doc;
    };
    expect(() => serializeXml(withText('bell\u0007'))).toThrow(XmlSerializeError);
    expect(() => serializeXml(withText('lone \uD83D surrogate'))).toThrow(XmlSerializeError);
    const attr = parseXml('<t a=""/>');
    attr.root.attrs[0]!.value = 'nul\u0000';
    expect(() => serializeXml(attr)).toThrow(/U\+0000/);
    const comment = parseXml('<t><!--x--></t>');
    comment.root.children[0] = { kind: 'comment', value: 'a--b' };
    expect(() => serializeXml(comment)).toThrow(XmlSerializeError);
    const prolog = parseXml('<t/>');
    prolog.prolog.push({ kind: 'text', value: 'junk' });
    expect(() => serializeXml(prolog)).toThrow(XmlSerializeError);
  });

  test('serializeElement writes one subtree', () => {
    const doc = parseXml(POWERPOINT_SLIDE);
    expect(serializeElement(doc.root.children[1] as never)).toBe('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>');
  });
});
