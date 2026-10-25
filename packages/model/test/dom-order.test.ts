import { describe, expect, test } from 'vitest';
import { canonical, parseXml, serializeElement, serializeXml } from '@pptx/xml';
import {
  attrNs, checkOrder, child, children, clone, createElement, declareUsedNamespaces, effectiveChildren, ensureChild, insertInOrder,
  is, lname, removeChildren, schemaType, setAttr, setAttrNs,
} from '../src/index.ts';
import { A, find, NSDECL, P, parse, slideXml } from './helpers.ts';

describe('matching by namespace, not prefix', () => {
  test('a document using other prefixes, or the default namespace, matches the same names', () => {
    const odd = parse(`<x:sld xmlns:x="${P}" xmlns:d="${A}"><x:cSld><x:spTree><x:sp><x:spPr><d:prstGeom prst="rect"/></x:spPr></x:sp></x:spTree></x:cSld></x:sld>`);
    expect(lname(odd)).toBe('p:sld');
    expect(is(find(odd, 'prstGeom'), 'a:prstGeom')).toBe(true);
    const deflt = parse(`<sld xmlns="${P}"><cSld/></sld>`);
    expect(child(deflt, 'p:cSld')).toBeDefined();
  });

  test('Strict documents match the same logical names', () => {
    const strict = parse('<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"><p:cSld/></p:sld>');
    expect(lname(strict)).toBe('p:sld');
    expect(child(strict, 'p:cSld')).toBeDefined();
  });

  test('unknown namespaces get a {uri}local name', () => {
    expect(lname(parse('<q:x xmlns:q="urn:q"/>'))).toBe('{urn:q}x');
  });
});

describe('creating elements', () => {
  test('reuses the prefix bound in scope', () => {
    const sld = parse(slideXml(''));
    const el = createElement(find(sld, 'spTree'), 'a:solidFill');
    expect([el.prefix, el.ns, el.attrs]).toEqual(['a', A, []]);
  });

  test('declares a namespace that is not bound, in the flavour of the document', () => {
    const sld = parse(`<p:sld xmlns:p="${P}"><p:cSld/></p:sld>`);
    const el = createElement(child(sld, 'p:cSld')!, 'a:t');
    expect(serializeElement(el)).toBe(`<a:t xmlns:a="${A}"/>`);
    const strict = parse('<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"><p:cSld/></p:sld>');
    expect(createElement(child(strict, 'p:cSld')!, 'a:t').ns).toBe('http://purl.oclc.org/ooxml/drawingml/main');
  });

  test('a prefix bound to something else is not reused', () => {
    const root = parse('<p:sld xmlns:p="urn:not-pml" xmlns:a="urn:not-dml"/>');
    const el = createElement(root, 'a:t');
    expect(el.prefix).toBe('a2');
    expect(el.ns).toBe(A);
  });

  test('setAttrNs declares r: when needed and reuses it when bound', () => {
    const blip = find(parse(slideXml('<p:pic><p:blipFill><a:blip/></p:blipFill></p:pic>')), 'blip');
    setAttrNs(blip, 'r', 'embed', 'rId7');
    expect(serializeElement(blip)).toBe('<a:blip r:embed="rId7"/>');
    const lone = parse(`<a:blip xmlns:a="${A}"/>`);
    setAttrNs(lone, 'r', 'embed', 'rId1');
    expect(attrNs(lone, 'r', 'embed')).toBe('rId1');
    expect(attrNs(parseXml(serializeElement(lone)).root, 'r', 'embed')).toBe('rId1');    // the declaration was written
  });

  test('declareUsedNamespaces makes a moved subtree self-contained', () => {
    const sld = parse(slideXml('<p:sp><p:nvSpPr><p:cNvPr id="2" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'));
    const sp = clone(find(sld, 'sp'));
    declareUsedNamespaces(sp);
    const reparsed = parseXml(serializeElement(sp)).root;
    expect(lname(reparsed)).toBe('p:sp');
    expect(is(find(reparsed, 'prstGeom'), 'a:prstGeom')).toBe(true);
  });
});

describe('schema order', () => {
  const doc = () => parse(slideXml(
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill><a:ln w="12700"/></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>a</a:t></a:r></a:p></p:txBody></p:sp>'));

  test('finds the type of any element by walking down from the root', () => {
    const sld = doc();
    expect(schemaType(sld)).toBe('p:CT_Slide');
    expect(schemaType(find(sld, 'spPr'))).toBe('a:CT_ShapeProperties');
    expect(schemaType(find(sld, 'off'))).toBe('a:CT_Point2D');
    expect(schemaType(find(sld, 't'))).toBe(null);           // a:t is text, not a complex type
  });

  test('a fill replaces the other fill, in its place between geometry and line', () => {
    const spPr = find(doc(), 'spPr');
    insertInOrder(spPr, createElement(spPr, 'a:solidFill'));
    expect(children(spPr).map(lname)).toEqual(['a:xfrm', 'a:prstGeom', 'a:solidFill', 'a:ln']);
  });

  test('an element goes before everything ranked after it', () => {
    const spPr = find(doc(), 'spPr');
    removeChildren(spPr, 'a:gradFill');
    insertInOrder(spPr, createElement(spPr, 'a:effectLst'));
    insertInOrder(spPr, createElement(spPr, 'a:noFill'));
    expect(children(spPr).map(lname)).toEqual(['a:xfrm', 'a:prstGeom', 'a:noFill', 'a:ln', 'a:effectLst']);
  });

  test('repeatable elements are appended after their kind', () => {
    const p = find(doc(), 'p');
    insertInOrder(p, createElement(p, 'a:endParaRPr'));
    insertInOrder(p, createElement(p, 'a:br'));
    insertInOrder(p, createElement(p, 'a:pPr'));
    expect(children(p).map(lname)).toEqual(['a:pPr', 'a:r', 'a:br', 'a:endParaRPr']);
  });

  test('the editor cannot put an element where the schema does not allow it', () => {
    const spPr = find(doc(), 'spPr');
    expect(() => insertInOrder(spPr, createElement(spPr, 'a:t'))).toThrow(/not allowed/);
    expect(() => insertInOrder(find(doc(), 't'), createElement(spPr, 'a:t'))).toThrow(/no schema order/);
  });

  test('ensureChild creates once', () => {
    const sp = find(doc(), 'sp');
    const style = ensureChild(sp, 'p:style');
    expect(ensureChild(sp, 'p:style')).toBe(style);
    expect(children(sp).map(lname)).toEqual(['p:nvSpPr', 'p:spPr', 'p:style', 'p:txBody']);
  });

  test('mc wrappers are transparent for typing and ordering', () => {
    const sld = parse(slideXml('', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="slow" p14:dur="2000"><p14:prism/></p:transition></mc:Choice>' +
      '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'));
    expect(schemaType(find(sld, 'transition'))).toBe('p:CT_SlideTransition');
    expect(checkOrder(sld)).toEqual([]);
  });

  test('checkOrder reports misplaced and duplicated children', () => {
    const sld = doc();
    const spPr = find(sld, 'spPr');
    spPr.children.reverse();
    expect(checkOrder(sld).length).toBeGreaterThan(0);
    const fixed = doc();
    const sp = find(fixed, 'spPr');
    sp.children.push(createElement(sp, 'a:noFill'));
    expect(checkOrder(fixed)).toEqual(['p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:noFill: out of schema order in a:CT_ShapeProperties']);
  });

  test('round trip of an edited tree stays well-formed and namespace-correct', () => {
    const sld = doc();
    const spPr = find(sld, 'spPr');
    const fill = createElement(spPr, 'a:solidFill');
    const clr = createElement(fill, 'a:schemeClr', { val: 'accent2' });
    fill.children.push(clr);
    insertInOrder(spPr, fill);
    setAttr(find(sld, 'off'), 'x', 99);
    const written = serializeXml({ declaration: null, prolog: [], root: sld, epilog: [] });
    expect(written).toContain('<a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:ln w="12700"/>');
    expect(canonical(parseXml(written))).toContain(`{${A}}schemeClr`);
  });
});

describe('markup compatibility', () => {
  const xml = `<p:sld ${NSDECL} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    '<mc:AlternateContent><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition p14:dur="2000"/></mc:Choice>' +
    '<mc:Fallback><p:transition spd="med"/></mc:Fallback></mc:AlternateContent></p:sld>';

  test('takes the Fallback when the Choice needs a namespace this build does not implement', () => {
    const [t] = effectiveChildren(parse(xml), new Set());
    expect(t!.attrs.map(a => a.local)).toEqual(['spd']);
  });

  test('takes the Choice when it does', () => {
    const [t] = effectiveChildren(parse(xml), new Set(['p14']));
    expect(t!.attrs.map(a => a.local)).toEqual(['dur']);
  });

  test('a Choice whose Requires prefix is unbound is skipped', () => {
    const odd = `<p:sld ${NSDECL} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent>` +
      '<mc:Choice Requires="zz"><p:cSld name="choice"/></mc:Choice><mc:Fallback><p:cSld name="fallback"/></mc:Fallback></mc:AlternateContent></p:sld>';
    expect(effectiveChildren(parse(odd), new Set(['p14']))[0]!.attrs[0]!.value).toBe('fallback');
  });
});
