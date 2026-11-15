// Equations: typed on one line, stored as Office math, shown on one line; inserted and changed.
import { OpcPackage } from '@pptx/opc';
import { checkOrder } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { parseXml, serializeElement } from '@pptx/xml';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import {
  deleteShapes, equationOf, equationPieces, insertEquation, insertTextBox, linearOf, moveBy, ommlOf, parseEquation, setEquation, type Editor,
} from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const BOX = { x: 1000000, y: 2000000, w: 5000000, h: 600000 };
const QUADRATIC = 'x = (-b ± sqrt(b^2 - 4ac)) / (2a)';
const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
/** The structure of Office math, without its run properties: easier to read in an expectation. */
const shape = (text: string) => ommlOf(parseEquation(text)).replace(/<a:rPr.*?<\/a:rPr>/g, '').replace(/<\/?m:/g, m => m.replace('m:', '')).replace(/ m:val=/g, ' val=');
const math = (xml: string) => parseXml(`<m:oMath ${M} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</m:oMath>`).root;
const drawn = (ed: Editor, id: number) => resolveSlide(ed.doc, SLIDE).shapes.find(s => s.id === id && s.layer === 'slide')!.text!.paragraphs[0]!.runs;

describe('an equation typed on one line', () => {
  test('exponents, indices, fractions and roots become structure; brackets that only group are dropped', () => {
    expect(shape('x^2')).toBe('<sSup><e><r><t>x</t></r></e><sup><r><t>2</t></r></sup></sSup>');
    expect(shape('a_1')).toBe('<sSub><e><r><t>a</t></r></e><sub><r><t>1</t></r></sub></sSub>');
    expect(shape('a_i^2')).toBe('<sSubSup><e><r><t>a</t></r></e><sub><r><t>i</t></r></sub><sup><r><t>2</t></r></sup></sSubSup>');
    expect(shape('1/2')).toBe('<f><num><r><t>1</t></r></num><den><r><t>2</t></r></den></f>');
    expect(shape('(a+b)/(c)')).toBe('<f><num><r><t>a</t></r><r><t>+</t></r><r><t>b</t></r></num><den><r><t>c</t></r></den></f>');
    expect(shape('sqrt(x+1)')).toBe('<rad><radPr><degHide val="1"/></radPr><deg/><e><r><t>x</t></r><r><t>+</t></r><r><t>1</t></r></e></rad>');
    expect(shape('e^(i π)')).toBe('<sSup><e><r><t>e</t></r></e><sup><r><t>i</t></r><r><t>π</t></r></sup></sSup>');
    expect(shape('x^-1')).toBe('<sSup><e><r><t>x</t></r></e><sup><r><t>-</t></r><r><t>1</t></r></sup></sSup>');
    expect(shape('2(a+b)')).toBe('<r><t>2</t></r><d><e><r><t>a</t></r><r><t>+</t></r><r><t>b</t></r></e></d>');   // brackets that are part of it stay
    expect(shape('3.14 r^2')).toBe('<r><t>3.14</t></r><sSup><e><r><t>r</t></r></e><sup><r><t>2</t></r></sup></sSup>');
  });

  test('the quadratic formula: a fraction over 2a, with a root inside its top', () => {
    const s = shape(QUADRATIC);
    expect(s.startsWith('<r><t>x</t></r><r><t>=</t></r><f><num><r><t>-</t></r><r><t>b</t></r><r><t>±</t></r><rad>')).toBe(true);
    expect(s).toContain('<e><sSup><e><r><t>b</t></r></e><sup><r><t>2</t></r></sup></sSup><r><t>-</t></r><r><t>4</t></r><r><t>ac</t></r></e></rad></num>');
    expect(s.endsWith('<den><r><t>2</t></r><r><t>a</t></r></den></f>')).toBe(true);
  });

  test('nothing typed is an error: stray signs and unclosed brackets stay as they can', () => {
    for (const odd of ['x^', '/', ')', '((a', 'a_^b', '^^', 'sqrt', '   ', 'a < b & c > "d"']) {
      const xml = ommlOf(parseEquation(odd));
      expect(() => math(xml), odd).not.toThrow();
    }
    expect(shape('a < b')).toBe('<r><t>a</t></r><r><t>&lt;</t></r><r><t>b</t></r>');
    expect(shape('sqrtx')).toBe('<r><t>sqrtx</t></r>');                   // a name, not a root
  });

  test('a math run is written as PowerPoint writes one: Cambria Math at the size given', () => {
    expect(ommlOf(parseEquation('x'), 3200)).toBe('<m:r><a:rPr lang="en-US" sz="3200" i="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t>x</m:t></m:r>');
  });
});

describe('Office math back on one line', () => {
  test('what is written reads back as text that parses to the same structure', () => {
    for (const text of [QUADRATIC, 'a^2 + b^2 = c^2', 'E = m c^2', 'A = π r^2', 'x_1 + x_2', 'a_i^2', '1/2 + 3/(x+1)', 'sqrt(2)/2', 'e^(i π) + 1 = 0', '2(a+b)', 'x^-1']) {
      const back = linearOf(math(ommlOf(parseEquation(text))));
      expect(shape(back), `${text} -> ${back}`).toBe(shape(text));
    }
    expect(linearOf(math(ommlOf(parseEquation(QUADRATIC))))).toBe('x=(-b±sqrt(b^2-4ac))/(2a)');
    // no more brackets than it takes: an exponent or a root as half of a fraction stands bare
    expect(linearOf(math(ommlOf(parseEquation('E = m c^2 / sqrt(1 - v^2/c^2)'))))).toBe('E=m c^2/sqrt(1-v^2/c^2)');
    expect(linearOf(math(ommlOf(parseEquation('x^(a_1)'))))).toBe('x^(a_1)');
  });

  test('math this editor does not write still reads as its symbols: a sum with limits, other brackets', () => {
    const sum = math('<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
    expect(linearOf(sum)).toBe('∑_(i=1)^n i');
    const abs = math('<m:d><m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>');
    expect(linearOf(abs)).toBe('|x|');
  });
});

describe('an equation shown on one line', () => {
  test('pieces in order, exponents raised and indices lowered, a fraction with a slash, a root with its sign', () => {
    const pieces = (text: string) => equationPieces(parseEquation(text)).map(p => (p.script ? `${p.script > 0 ? '^' : '_'}${p.text}` : p.text)).join(' ');
    expect(pieces('x^2 + a_1')).toBe('x ^2 + a _1');
    expect(pieces('a_i^2')).toBe('a _i ^2');
    expect(pieces('1/2')).toBe('1 / 2');
    expect(pieces('(a+b)/c')).toBe('( a + b ) / c');
    expect(pieces('sqrt(x+1)')).toBe('√ ( x + 1 )');
    expect(pieces('sqrt(2)')).toBe('√ 2');
    expect(pieces('2(a+b)')).toBe('2 ( a + b )');
    expect(pieces('e^(i π)')).toBe('e ^i ^π');
    expect(pieces('x^(1/2)')).toBe('x ^1 ^/ ^2');
  });
});

describe('inserting an equation', () => {
  test('a text box holding Office math, beside a fallback with the text; drawn on one line; undone exactly', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => insertEquation(ed, SLIDE, QUADRATIC, BOX).command)).toEqual([]);
    const q = insertEquation(ed, SLIDE, QUADRATIC, BOX);
    ed.run(q.command);
    const saved = new TextDecoder().decode(OpcPackage.open(ed.save()).getPart(SLIDE)!.bytes);
    expect(saved).toMatch(/<mc:AlternateContent[^>]*><mc:Choice[^>]* Requires="a14"[^>]*><p:sp>/);
    expect(saved).toContain(`<a14:m><m:oMathPara ${M}><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath ${M}><m:r>`);
    // the fallback: the same box with the equation as text, the exponent a raised run
    expect(saved).toMatch(/<mc:Fallback><p:sp>.*<a:t>x<\/a:t>.*<a:rPr lang="en-US" sz="2400" baseline="30000">.*?<\/a:rPr><a:t>2<\/a:t>.*<\/p:sp><\/mc:Fallback>/);
    expect(() => parseXml(saved)).not.toThrow();
    // the shape this editor reads is the fallback; it knows its equation, and is drawn as it on one line
    const view = ed.shape(SLIDE, q.id)!;
    expect(equationOf(view)).toBe('x=(-b±sqrt(b^2-4ac))/(2a)');
    const runs = drawn(ed, q.id);
    expect(runs.map(r => r.text).join('')).toBe('x=(-b±√(b2-4ac))/(2a)');
    const two = runs.find(r => r.text === '2')!;
    expect(two.props.baseline).toBeGreaterThan(0);                        // the exponent is raised
    expect(runs.find(r => r.text === 'x')!.props.italic).toBe(true);
    expect(runs.find(r => r.text.includes('='))!.props.italic).toBe(false);
    expect(runs[0]!.props.size).toBe(24);
    expect(checkOrder(ed.doc.root(SLIDE))).toEqual([]);
    expect(ed.log.undoLabel).toBe('Insert Equation');
    expect(() => insertEquation(ed, SLIDE, '   ', BOX)).toThrow(/needs something/);
  });

  test('the box moves and is deleted as one shape, both branches together', () => {
    const ed = openEditor(deckBytes());
    const q = insertEquation(ed, SLIDE, 'a^2', BOX);
    ed.run(q.command);
    ed.run(moveBy(ed, SLIDE, [ed.shape(SLIDE, q.id)!], 100000, 50000));
    const xml = serializeElement(ed.doc.root(SLIDE));
    expect(xml.match(/<a:off x="1100000" y="2050000"\/>/g)?.length).toBe(2);
    ed.run(deleteShapes(ed, SLIDE, [ed.shape(SLIDE, q.id)!]));
    expect(serializeElement(ed.doc.root(SLIDE))).not.toContain('AlternateContent');
  });
});

describe('changing an equation', () => {
  test('the math is written anew at the size it had, and the fallback\'s text with it; undone exactly', () => {
    const ed = openEditor(deckBytes());
    const q = insertEquation(ed, SLIDE, 'a^2', BOX, 3200);
    ed.run(q.command);
    ed.log.clear();
    expect(cycle(ed, () => setEquation(ed, SLIDE, ed.shape(SLIDE, q.id)!, 'E = m c^2'))).toEqual([]);
    ed.run(setEquation(ed, SLIDE, ed.shape(SLIDE, q.id)!, 'E = m c^2'));
    expect(equationOf(ed.shape(SLIDE, q.id)!)).toBe('E=m c^2');
    const xml = serializeElement(ed.doc.root(SLIDE));
    expect(drawn(ed, q.id).map(r => r.text).join('')).toBe('E=mc2');
    expect(drawn(ed, q.id).at(-1)!.props.baseline).toBeGreaterThan(0);
    expect(xml.match(/<a14:m[ >]/g)?.length).toBe(1);
    expect(drawn(ed, q.id)[0]!.props.size).toBe(32);
    expect(ed.log.undoLabel).toBe('Edit Equation');
    const back = openEditor(ed.save());
    expect(equationOf(back.shape(SLIDE, q.id)!)).toBe('E=m c^2');
  });

  test('a shape with no equation has none to read or change', () => {
    const ed = openEditor(deckBytes());
    const t = insertTextBox(ed, SLIDE, BOX, 'plain');
    ed.run(t.command);
    expect(equationOf(ed.shape(SLIDE, t.id)!)).toBeUndefined();
    expect(() => setEquation(ed, SLIDE, ed.shape(SLIDE, t.id)!, 'x')).toThrow(/no equation/);
  });
});
