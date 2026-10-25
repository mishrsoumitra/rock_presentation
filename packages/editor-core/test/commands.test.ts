// Editing commands on a small deck: each one checked by the oracle and by exact undo and redo (see helpers.cycle).
import { OpcPackage, readZip } from '@pptx/opc';
import { attr, child, children, paragraphText, runsOf } from '@pptx/model';
import { resolveSlide } from '@pptx/resolve';
import { serializeElement } from '@pptx/xml';
import { findPreset, listEffects, slideTransition } from '@pptx/anim';
import { describe, expect, test } from 'vitest';
import { deckBytes, deckFiles } from '../../resolve/test/deck.ts';
import { zipFiles } from '../../opc/test/fixtures.ts';
import {
  addSlide, align, applyLayout, copyShapes, cropPicture, deleteShapes, deleteSlide, deleteText, distribute, duplicateShapes, duplicateSlide,
  effectiveXfrm, flip, formatBody, formatParagraphs, formatRuns, group, importMasters, insertPicture, insertShape, insertTextBox,
  layoutToModel, modelToLayout, moveBy, moveSlide, pasteShapes, prepareForSave, prepareImage, prepareTemplate, replacePicture,
  rotateQuarter, setFill, setLine, setSlideHidden, setSlideSize, themePalette, typeText, ungroup, visualBounds, parentMatrix, zOrder, type Editor,
  addAnimation, moveAnimation, removeAnimation, setAnimationTiming, setSlideTransition, setTransitionOnAll,
} from '../src/index.ts';
import { cycle, diffStates, openEditor, parseAll, state } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const LAYOUT = '/ppt/slideLayouts/slideLayout1.xml';

const shape = (ed: Editor, id: number) => ed.shape(SLIDE, id)!;
const xml = (el: { el: Parameters<typeof serializeElement>[0] } | Parameters<typeof serializeElement>[0]) => serializeElement('el' in el ? el.el : el);
const ids = (ed: Editor) => ed.slide(SLIDE).shapes.map(s => s.id);

/** A tiny but real PNG: 2x1 pixels. */
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0, 123, 64, 232, 221, 0, 0, 0, 15, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 86, 199, 45, 179, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);
const PNG2 = PNG.slice();
PNG2[PNG2.length - 13] = PNG2[PNG2.length - 13]! ^ 1;

describe('command log', () => {
  test('commands with one key merge within the window; undo and redo seal the step', () => {
    const ed = openEditor(deckBytes());
    const s = shape(ed, 5);
    ed.run(moveBy(ed, SLIDE, [s], 100, 0));
    ed.run(moveBy(ed, SLIDE, [s], 100, 0));
    ed.run(moveBy(ed, SLIDE, [s], 100, 0));
    expect(effectiveXfrm(ed, SLIDE, s).x).toBe(400);
    ed.undo();
    expect(effectiveXfrm(ed, SLIDE, s).x).toBe(100);           // all three nudges were one step
    expect(ed.log.canUndo).toBe(false);
    ed.redo();
    ed.run(moveBy(ed, SLIDE, [s], 50, 0));                      // after a redo: a new step, not merged
    ed.undo();
    expect(effectiveXfrm(ed, SLIDE, s).x).toBe(400);
    expect(ed.log.redoLabel).toBe('Move');
  });

  test('a new command clears redo; different keys never merge; seal() ends a run', () => {
    const ed = openEditor(deckBytes());
    const s = shape(ed, 5);
    ed.run(moveBy(ed, SLIDE, [s], 100, 0));
    ed.undo();
    ed.run(rotateQuarter(ed, SLIDE, [s], 1));
    expect(ed.log.canRedo).toBe(false);
    ed.run(moveBy(ed, SLIDE, [s], 100, 0));
    ed.log.seal();
    ed.run(moveBy(ed, SLIDE, [s], 100, 0));
    ed.undo();
    ed.undo();
    expect(effectiveXfrm(ed, SLIDE, s)).toMatchObject({ x: 100, rot: 5400000 });
  });

  test('a command that throws leaves the tree as it was', () => {
    const ed = openEditor(deckBytes());
    parseAll(ed);
    const before = state(ed);
    expect(() => ed.run(group(ed, SLIDE, [shape(ed, 5)]))).toThrow(/two or more/);
    expect(diffStates(before, state(ed))).toEqual([]);
    expect(ed.log.canUndo).toBe(false);
  });
});

describe('transforms', () => {
  test('moving a placeholder that inherits its position gives it its own xfrm there; undo takes it away', () => {
    const ed = openEditor(deckBytes());
    const title = shape(ed, 2);
    expect(title.xfrm).toBeNull();
    expect(cycle(ed, () => moveBy(ed, SLIDE, [title], 1000, 2000))).toEqual([]);
    ed.run(moveBy(ed, SLIDE, [title], 1000, 2000));
    expect(title.xfrm).toMatchObject({ x: 838200 + 1000, y: 365125 + 2000, cx: 10515600, cy: 1325563 });
    expect(xml(title)).toContain('<p:spPr><a:xfrm><a:off x="839200" y="367125"/>');
  });

  test('a member of a rotated, scaled group moves the same distance on the slide', () => {
    const ed = openEditor(deckBytes());
    const member = shape(ed, 8);
    const before = visualBounds(effectiveXfrm(ed, SLIDE, member), parentMatrix(ed, SLIDE, member.el));
    ed.run(moveBy(ed, SLIDE, [member], 300000, 0));
    const after = visualBounds(effectiveXfrm(ed, SLIDE, member), parentMatrix(ed, SLIDE, member.el));
    expect(after.x - before.x).toBeCloseTo(300000, 3);
    expect(after.y - before.y).toBeCloseTo(0, 3);
  });

  test('rotate, flip, align, distribute', () => {
    const ed = openEditor(deckBytes());
    const s = shape(ed, 5);
    expect(cycle(ed, () => rotateQuarter(ed, SLIDE, [s], -1))).toEqual([]);
    ed.run(rotateQuarter(ed, SLIDE, [s], -1));
    expect(s.xfrm!.rot).toBe(16200000);
    ed.run(flip(ed, SLIDE, [s], 'h'));
    expect(s.xfrm).toMatchObject({ flipH: true, rot: 5400000 });
    expect(cycle(ed, () => align(ed, SLIDE, [s], 'right'))).toEqual([]);
    ed.run(align(ed, SLIDE, [s], 'right'));
    const b = visualBounds(s.xfrm!);
    expect(b.x + b.w).toBeCloseTo(12192000, 0);
    const a = insertShape(ed, SLIDE, 'rect', { x: 0, y: 0, w: 100, h: 100 });
    ed.run(a.command);
    const c = insertShape(ed, SLIDE, 'rect', { x: 1000, y: 0, w: 300, h: 100 });
    ed.run(c.command);
    const d = insertShape(ed, SLIDE, 'rect', { x: 400, y: 0, w: 100, h: 100 });
    ed.run(d.command);
    const three = [a.id, c.id, d.id].map(id => shape(ed, id));
    expect(cycle(ed, () => distribute(ed, SLIDE, three, 'h'))).toEqual([]);
    ed.run(distribute(ed, SLIDE, three, 'h'));
    // gaps: 1300 total span, 500 of shapes, two gaps of 400
    expect(three.map(x => x.xfrm!.x)).toEqual([0, 1000, 500]);
  });
});

describe('arrange', () => {
  test('z-order moves units among the tree\'s shapes and keeps the rest in place', () => {
    const ed = openEditor(deckBytes());
    expect(ids(ed)).toEqual([2, 3, 4, 5, 6]);
    expect(cycle(ed, () => zOrder(ed, SLIDE, [shape(ed, 3)], 'front'))).toEqual([]);
    ed.run(zOrder(ed, SLIDE, [shape(ed, 3)], 'front'));
    expect(ids(ed)).toEqual([2, 4, 5, 6, 3]);
    ed.run(zOrder(ed, SLIDE, [shape(ed, 3), shape(ed, 5)], 'back'));
    expect(ids(ed)).toEqual([5, 3, 2, 4, 6]);
    ed.run(zOrder(ed, SLIDE, [shape(ed, 5)], 'forward'));
    expect(ids(ed)).toEqual([3, 5, 2, 4, 6]);
    ed.run(zOrder(ed, SLIDE, [shape(ed, 6)], 'backward'));
    expect(ids(ed)).toEqual([3, 5, 2, 6, 4]);
  });

  test('grouping keeps every member where it was; ungrouping the rotated, scaled group too', () => {
    const ed = openEditor(deckBytes());
    const boundsOf = (id: number) => {
      const s = resolveSlide(ed.doc, SLIDE);
      const find = (list: typeof s.shapes): (typeof s.shapes)[number] | undefined => {
        for (const x of list) { if (x.id === id) return x; const y = find(x.children); if (y) return y; }
        return undefined;
      };
      const r = find(s.shapes)!;
      const m = r.transform;
      const pts = [[0, 0], [r.xfrm.cx, 0], [r.xfrm.cx, r.xfrm.cy], [0, r.xfrm.cy]].map(([u, v]) => [m[0] * u! + m[2] * v! + m[4], m[1] * u! + m[3] * v! + m[5]]);
      return [Math.min(...pts.map(p => p[0]!)), Math.min(...pts.map(p => p[1]!)), Math.max(...pts.map(p => p[0]!)), Math.max(...pts.map(p => p[1]!))].map(Math.round);
    };
    const before5 = boundsOf(5), before8 = boundsOf(8);
    expect(cycle(ed, () => group(ed, SLIDE, [shape(ed, 5), shape(ed, 3)]))).toEqual([]);
    ed.run(group(ed, SLIDE, [shape(ed, 5), shape(ed, 3)]));
    expect(ids(ed)).toEqual([2, 4, 9, 6]);                      // the group takes the topmost member's place
    expect(boundsOf(5)).toEqual(before5);
    const g = shape(ed, 9);
    expect(g.children.map(c => c.id)).toEqual([3, 5]);
    expect(cycle(ed, () => ungroup(ed, SLIDE, shape(ed, 6)))).toEqual([]);
    ed.run(ungroup(ed, SLIDE, shape(ed, 6)));
    expect(ids(ed)).toEqual([2, 4, 9, 8]);
    const after8 = boundsOf(8);
    after8.forEach((v, i) => expect(Math.abs(v - before8[i]!)).toBeLessThanOrEqual(2));
  });
});

describe('shapes', () => {
  test('a new shape is written as PowerPoint writes one', () => {
    const ed = openEditor(deckBytes());
    const { command, id } = insertShape(ed, SLIDE, 'ellipse', { x: 10, y: 20, w: 300, h: 200 });
    expect(id).toBe(9);
    expect(cycle(ed, () => insertShape(ed, SLIDE, 'ellipse', { x: 10, y: 20, w: 300, h: 200 }).command)).toEqual([]);
    ed.run(command);
    expect(xml(shape(ed, 9))).toBe('<p:sp><p:nvSpPr><p:cNvPr id="9" name="Oval 8"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
      + '<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="200"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr>'
      + '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
      + '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'
      + '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>');
    const line = insertShape(ed, SLIDE, 'straightConnector1', { x: 0, y: 0, w: 100, h: 0 });
    ed.run(line.command);
    expect(shape(ed, line.id).kind).toBe('cxnSp');
    expect(cycle(ed, () => insertTextBox(ed, SLIDE, { x: 0, y: 0, w: 100, h: 100 }, 'one\ntwo').command)).toEqual([]);
    const tb = insertTextBox(ed, SLIDE, { x: 0, y: 0, w: 100, h: 100 }, 'one\ntwo');
    ed.run(tb.command);
    expect(children(shape(ed, tb.id).txBody, 'a:p').map(paragraphText)).toEqual(['one', 'two']);
  });

  test('deleting a shape drops the relationships only it used', () => {
    const ed = openEditor(deckBytes());
    expect(ed.pkg.relationships(SLIDE).map(r => r.id)).toEqual(['rId1', 'rId2']);
    expect(cycle(ed, () => deleteShapes(ed, SLIDE, [shape(ed, 3)]))).toEqual([]);
    ed.run(deleteShapes(ed, SLIDE, [shape(ed, 3)]));
    expect(ids(ed)).toEqual([2, 4, 5, 6]);
    expect(ed.pkg.relationships(SLIDE).map(r => r.id)).toEqual(['rId1']);     // the hyperlink went with its text
    ed.run(deleteShapes(ed, SLIDE, [shape(ed, 8)]));                          // the group's only member: the group goes
    expect(ids(ed)).toEqual([2, 4, 5]);
  });

  test('duplicate and paste give fresh ids; paste into another deck brings pictures along', async () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => duplicateShapes(ed, SLIDE, [shape(ed, 6)]).command)).toEqual([]);
    const dup = duplicateShapes(ed, SLIDE, [shape(ed, 6)]);
    ed.run(dup.command);
    expect(dup.ids).toEqual([9]);
    expect(ids(ed)).toEqual([2, 3, 4, 5, 6, 9]);
    expect(shape(ed, 9).children.map(c => c.id)).toEqual([10]);
    expect(shape(ed, 9).xfrm).toMatchObject({ x: 1152400, y: 1152400 });

    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 2, height: 1 });
    const pic = insertPicture(ed, SLIDE, img);
    ed.run(pic.command);
    const clip = copyShapes(ed, SLIDE, [shape(ed, pic.id), shape(ed, 3)]);
    const other = openEditor(deckBytes());
    const hashes = await prepareTemplate(other, ed.pkg);
    expect(cycle(other, () => pasteShapes(other, SLIDE, clip, hashes).command)).toEqual([]);
    const pasted = pasteShapes(other, SLIDE, clip, hashes);
    other.run(pasted.command);
    expect(pasted.ids).toEqual([9, 10]);                       // after the largest id, as PowerPoint numbers
    expect(other.pkg.partNames()).toContain('/ppt/media/image1.png');
    // the picture got a relationship; the link was the same as one the slide already had, so it is shared
    expect(other.pkg.relationships(SLIDE).map(r => r.type.split('/').pop())).toEqual(['slideLayout', 'hyperlink', 'image']);
  });
});

describe('text', () => {
  const body = (ed: Editor) => shape(ed, 3).txBody!;
  const texts = (ed: Editor) => children(body(ed), 'a:p').map(paragraphText);

  test('typing goes into the run before the caret and coalesces; "\\n" splits paragraphs', () => {
    const ed = openEditor(deckBytes());
    const key = 'type:3';
    let t = typeText(ed, SLIDE, body(ed), { p: 0, offset: 7 }, ' sharply', { coalesceKey: key });
    expect(cycle(ed, () => typeText(ed, SLIDE, body(ed), { p: 0, offset: 7 }, ' sharply').command)).toEqual([]);
    ed.run(t.command);
    expect(t.caret).toEqual({ p: 0, offset: 15 });
    t = typeText(ed, SLIDE, body(ed), t.caret, '!', { coalesceKey: key });
    ed.run(t.command);
    expect(texts(ed)[0]).toBe('Revenue sharply! up');
    t = typeText(ed, SLIDE, body(ed), { p: 1, offset: 8 }, 'x\ny', { coalesceKey: key });
    ed.run(t.command);
    expect(texts(ed)).toEqual(['Revenue sharply! up', 'in everyx', 'y region', 'details']);
    // the new paragraph keeps the level and the bold red run formatting
    const third = children(body(ed), 'a:p')[2]!;
    expect(attr(child(third, 'a:pPr'), 'lvl')).toBe('1');
    expect(attr(runsOf(third)[0]!.rPr, 'b')).toBe('1');
    ed.undo();
    expect(texts(ed)).toEqual(['Revenue up', 'in every region', 'details']);
  });

  test('backspace joins paragraphs; deleting a selection across paragraphs merges them', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => deleteText(ed, SLIDE, body(ed), { p: 1, offset: 0 }).command)).toEqual([]);
    let d = deleteText(ed, SLIDE, body(ed), { p: 1, offset: 0 });
    ed.run(d.command);
    expect(texts(ed)).toEqual(['Revenue upin every region', 'details']);
    expect(d.caret).toEqual({ p: 0, offset: 10 });
    d = deleteText(ed, SLIDE, body(ed), { start: { p: 0, offset: 3 }, end: { p: 1, offset: 4 } });
    expect(cycle(ed, () => deleteText(ed, SLIDE, body(ed), { start: { p: 0, offset: 3 }, end: { p: 1, offset: 4 } }).command)).toEqual([]);
    ed.run(d.command);
    expect(texts(ed)).toEqual(['Revils']);
  });

  test('formatting a range splits runs at its edges', () => {
    const ed = openEditor(deckBytes());
    const range = { start: { p: 0, offset: 2 }, end: { p: 0, offset: 5 } };
    expect(cycle(ed, () => formatRuns(ed, SLIDE, body(ed), range, { b: true, sz: 2400, fill: { kind: 'scheme', val: 'accent1', lumMod: 75000 } }))).toEqual([]);
    ed.run(formatRuns(ed, SLIDE, body(ed), range, { b: true, sz: 2400, fill: { kind: 'scheme', val: 'accent1', lumMod: 75000 } }));
    const p = children(body(ed), 'a:p')[0]!;
    expect(runsOf(p).map(r => r.text)).toEqual(['Re', 'ven', 'ue up']);
    expect(serializeElement(runsOf(p)[1]!.el)).toBe('<a:r><a:rPr lang="en-US" b="1" sz="2400"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:rPr><a:t>ven</a:t></a:r>');
    expect(cycle(ed, () => formatParagraphs(ed, SLIDE, body(ed), range, { algn: 'ctr', bullet: { kind: 'autonum', scheme: 'arabicPeriod' }, spcBef: { pts: 600 } }))).toEqual([]);
    ed.run(formatParagraphs(ed, SLIDE, body(ed), range, { algn: 'ctr', bullet: { kind: 'autonum', scheme: 'arabicPeriod' }, spcBef: { pts: 600 } }));
    expect(serializeElement(child(p, 'a:pPr')!)).toBe('<a:pPr algn="ctr"><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buAutoNum type="arabicPeriod"/></a:pPr>');
    expect(cycle(ed, () => formatBody(ed, SLIDE, body(ed), { anchor: 'b', autofit: { kind: 'norm', fontScale: 85000, lnSpcReduction: 10000 } }))).toEqual([]);
    ed.run(formatBody(ed, SLIDE, body(ed), { anchor: 'b', autofit: { kind: 'shape' } }));
    expect(serializeElement(child(body(ed), 'a:bodyPr')!)).toBe('<a:bodyPr anchor="b"><a:spAutoFit/></a:bodyPr>');
  });

  test('layout and file offsets differ after a field', () => {
    const ed = openEditor(deckBytes());
    const resolved = resolveSlide(ed.doc, SLIDE);
    const num = resolved.shapes.find(s => s.id === 4 && s.layer === 'slide')!;
    const p = children(shape(ed, 4).txBody, 'a:p')[0]!;
    const rp = num.text!.paragraphs[0]!;
    expect(rp.runs[0]!.text).toBe('1');              // shown as the slide number; stored as ‹#›
    expect(layoutToModel(p, rp, 1)).toBe(3);
    expect(modelToLayout(p, rp, 3)).toBe(1);
    expect(modelToLayout(p, rp, 1)).toBe(1);         // inside the field: its end
  });
});

describe('pictures', () => {
  test('insert reuses identical media; delete keeps media still shown elsewhere', async () => {
    const ed = openEditor(deckBytes());
    const img = await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 200, height: 100, dpiX: 96, dpiY: 96 });
    expect(cycle(ed, () => insertPicture(ed, SLIDE, img).command)).toEqual([]);
    const a = insertPicture(ed, SLIDE, img);
    ed.run(a.command);
    const b = insertPicture(ed, SLIDE, await prepareImage(ed, { bytes: PNG.slice(), contentType: 'image/png', ext: 'png', width: 200, height: 100 }));
    ed.run(b.command);
    expect(ed.pkg.partNames().filter(n => n.startsWith('/ppt/media/'))).toEqual(['/ppt/media/image1.png']);
    expect(shape(ed, a.id).xfrm).toMatchObject({ cx: 1905000, cy: 952500, x: (12192000 - 1905000) / 2 });
    const types = new TextDecoder().decode(readZip(ed.pkg.save()).find(e => e.name === '[Content_Types].xml')!.data);
    expect(types).toContain('<Default Extension="png" ContentType="image/png"/>');
    ed.run(deleteShapes(ed, SLIDE, [shape(ed, a.id)]));
    expect(ed.pkg.getPart('/ppt/media/image1.png')).toBeDefined();
    ed.run(deleteShapes(ed, SLIDE, [shape(ed, b.id)]));
    expect(ed.pkg.getPart('/ppt/media/image1.png')).toBeUndefined();
    expect(ed.pkg.relationships(SLIDE).map(r => r.id)).toEqual(['rId1', 'rId2']);
  });

  test('crop and replace', async () => {
    const ed = openEditor(deckBytes());
    const pic = insertPicture(ed, SLIDE, await prepareImage(ed, { bytes: PNG, contentType: 'image/png', ext: 'png', width: 2, height: 1 }));
    ed.run(pic.command);
    expect(cycle(ed, () => cropPicture(ed, SLIDE, shape(ed, pic.id), { l: 10000, t: 0, r: 5000, b: 0 }))).toEqual([]);
    ed.run(cropPicture(ed, SLIDE, shape(ed, pic.id), { l: 10000, t: 0, r: 5000, b: 0 }));
    expect(xml(child(shape(ed, pic.id).el, 'p:blipFill')!)).toContain('<a:srcRect l="10000" r="5000"/><a:stretch>');
    const other = await prepareImage(ed, { bytes: PNG2, contentType: 'image/png', ext: 'png', width: 2, height: 1 });
    expect(cycle(ed, () => replacePicture(ed, SLIDE, shape(ed, pic.id), other))).toEqual([]);
    ed.run(replacePicture(ed, SLIDE, shape(ed, pic.id), other));
    expect(ed.pkg.partNames().filter(n => n.startsWith('/ppt/media/'))).toEqual(['/ppt/media/image2.png']);
  });
});

describe('slides', () => {
  test('a new slide gets the layout\'s placeholders, empty, and comes after the one given', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => addSlide(ed, LAYOUT).command)).toEqual([]);
    const { command, part } = addSlide(ed, LAYOUT, 0);
    ed.run(command);
    expect(part).toBe('/ppt/slides/slide2.xml');
    expect(ed.doc.slideParts()).toEqual([part, SLIDE]);
    expect(ed.doc.presentation.slideIds.map(e => e.id)).toEqual([257, 256]);
    const shapes = ed.slide(part).shapes;
    expect(shapes.map(s => [s.id, s.name, s.placeholder?.type, s.placeholder?.idx])).toEqual([[2, 'Title 1', 'title', 0], [3, 'Content Placeholder 2', 'obj', 1]]);
    expect(ed.pkg.getPart(part)!.contentType).toBe('application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    expect(ed.doc.layoutOf(part)).toBe(LAYOUT);
  });

  test('delete, duplicate, move, hide; apply layout', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => duplicateSlide(ed, SLIDE).command)).toEqual([]);
    const dup = duplicateSlide(ed, SLIDE);
    ed.run(dup.command);
    expect(ed.doc.slideParts()).toEqual([SLIDE, '/ppt/slides/slide2.xml']);
    expect(cycle(ed, () => moveSlide(ed, SLIDE, 1))).toEqual([]);
    ed.run(moveSlide(ed, SLIDE, 1));
    expect(ed.doc.slideParts()).toEqual(['/ppt/slides/slide2.xml', SLIDE]);
    expect(cycle(ed, () => setSlideHidden(ed, SLIDE, true))).toEqual([]);
    expect(cycle(ed, () => deleteSlide(ed, SLIDE))).toEqual([]);
    ed.run(deleteSlide(ed, SLIDE));
    expect(ed.doc.slideParts()).toEqual(['/ppt/slides/slide2.xml']);
    expect(ed.pkg.getPart(SLIDE)).toBeUndefined();
    const two = '/ppt/slides/slide2.xml';
    // moving the title first, then re-applying the layout puts it back at the layout's position
    ed.run(moveBy(ed, two, [ed.shape(two, 2)!], 5000, 0));
    expect(cycle(ed, () => applyLayout(ed, two, LAYOUT))).toEqual([]);
    ed.run(applyLayout(ed, two, LAYOUT));
    expect(ed.shape(two, 2)!.xfrm).toBeNull();
  });

  test('importing a template\'s masters copies master, layouts and theme with fresh ids', async () => {
    const ed = openEditor(deckBytes());
    const tpl = OpcPackage.open(deckBytes());
    const hashes = await prepareTemplate(ed, tpl);
    expect(cycle(ed, () => importMasters(ed, tpl, hashes).command)).toEqual([]);
    const imp = importMasters(ed, tpl, hashes);
    ed.run(imp.command);
    expect(imp.masters).toEqual(['/ppt/slideMasters/slideMaster2.xml']);
    expect(ed.doc.masterParts()).toEqual(['/ppt/slideMasters/slideMaster1.xml', '/ppt/slideMasters/slideMaster2.xml']);
    expect(ed.doc.layoutsOf(imp.masters[0]!)).toEqual(['/ppt/slideLayouts/slideLayout2.xml']);
    expect(ed.doc.themeOf(imp.masters[0]!)).toBe('/ppt/theme/theme2.xml');
    const ids = [...ed.doc.presentation.masterIds.map(m => m.id), ...ed.doc.masterParts().flatMap(m => children(child(ed.doc.root(m), 'p:sldLayoutIdLst'), 'p:sldLayoutId').map(l => Number(attr(l, 'id'))))];
    expect(new Set(ids).size).toBe(ids.length);
    expect(Math.min(...ids)).toBeGreaterThanOrEqual(2147483648);
    // the new layout is usable for a new slide
    ed.run(addSlide(ed, '/ppt/slideLayouts/slideLayout2.xml').command);
  });

  test('slide size: ensure fit scales and centres content and font sizes', () => {
    const ed = openEditor(deckBytes());
    expect(cycle(ed, () => setSlideSize(ed, 9144000, 6858000, 'ensureFit'))).toEqual([]);
    ed.run(setSlideSize(ed, 9144000, 6858000, 'ensureFit'));
    expect(ed.doc.presentation.slideSize).toEqual({ cx: 9144000, cy: 6858000, type: 'screen4x3' });
    expect(shape(ed, 5).xfrm).toMatchObject({ x: 75, y: Math.round(200 * 0.75 + (6858000 - 6858000 * 0.75) / 2), cx: 2250000, cy: 750000 });
  });

  test('saving updates docProps/app.xml counts', () => {
    const files = deckFiles();
    const appXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Slides>9</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides></Properties>';
    files['_rels/.rels'] = files['_rels/.rels']!.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
    files['[Content_Types].xml'] = files['[Content_Types].xml']!.replace('</Types>', '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
    const ed = openEditor(zipFiles({ ...files, 'docProps/app.xml': appXml }));
    prepareForSave(ed);
    expect(serializeElement(ed.pkg.xml('/docProps/app.xml').root)).toContain('<Slides>1</Slides>');
  });
});

describe('fill and outline', () => {
  test('fill and line are written in schema order and removed with null', () => {
    const ed = openEditor(deckBytes());
    const s = shape(ed, 5);
    expect(cycle(ed, () => setFill(ed, SLIDE, [s], { kind: 'scheme', val: 'accent2', lumMod: 75000 }))).toEqual([]);
    ed.run(setFill(ed, SLIDE, [s], { kind: 'scheme', val: 'accent2', lumMod: 75000 }));
    ed.run(setLine(ed, SLIDE, [s], { color: { kind: 'srgb', hex: '#ff0000' }, width: 25400, dash: 'dash' }));
    expect(xml(child(s.el, 'p:spPr')!)).toBe('<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000000" cy="1000000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
      + '<a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr>');
    expect(cycle(ed, () => setLine(ed, SLIDE, [s], { color: null, width: null, dash: null }))).toEqual([]);
    ed.run(setLine(ed, SLIDE, [s], { color: null, width: null, dash: null }));
    ed.run(setFill(ed, SLIDE, [s], null));
    expect(xml(child(s.el, 'p:spPr')!)).toBe('<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000000" cy="1000000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr>');
    // a placeholder with no spPr content gets one, in place
    expect(cycle(ed, () => setFill(ed, SLIDE, [shape(ed, 2)], 'none'))).toEqual([]);
  });
});

describe('animations and transitions', () => {
  test('add, retime, reorder and remove effects; deleting a shape takes its effects along', () => {
    const ed = openEditor(deckBytes());
    const fade = findPreset('entr', 10)!, spin = findPreset('emph', 8)!;
    expect(cycle(ed, () => addAnimation(ed, SLIDE, shape(ed, 5), fade))).toEqual([]);
    ed.run(addAnimation(ed, SLIDE, shape(ed, 5), fade));
    ed.run(addAnimation(ed, SLIDE, shape(ed, 3), spin, { trigger: 'after' }));
    const list = () => listEffects(ed.slide(SLIDE)).map(e => [e.name, e.spid, e.trigger]);
    expect(list()).toEqual([['Fade', 5, 'click'], ['Spin', 3, 'after']]);
    expect(cycle(ed, () => setAnimationTiming(ed, SLIDE, 0, { duration: 2000, delay: 250 }))).toEqual([]);
    ed.run(setAnimationTiming(ed, SLIDE, 0, { duration: 2000, delay: 250 }));
    expect(listEffects(ed.slide(SLIDE))[0]).toMatchObject({ duration: 2000, delay: 250 });
    expect(cycle(ed, () => moveAnimation(ed, SLIDE, 1, 0))).toEqual([]);
    expect(cycle(ed, () => removeAnimation(ed, SLIDE, 0))).toEqual([]);
    expect(cycle(ed, () => deleteShapes(ed, SLIDE, [shape(ed, 5)]))).toEqual([]);
    ed.run(deleteShapes(ed, SLIDE, [shape(ed, 5)]));
    // what followed keeps its start: After Previous now plays as the slide opens, as in PowerPoint
    expect(list()).toEqual([['Spin', 3, 'after']]);
    ed.run(deleteShapes(ed, SLIDE, [shape(ed, 3)]));
    expect(child(ed.doc.root(SLIDE), 'p:timing')).toBeUndefined();          // the last effect took p:timing with it
  });

  test('transitions: set one, apply to all, remove', () => {
    const ed = openEditor(deckBytes());
    ed.run(addSlide(ed, LAYOUT).command);
    expect(cycle(ed, () => setSlideTransition(ed, SLIDE, { kind: 'push', dir: 'u', duration: 1200 }))).toEqual([]);
    ed.run(setSlideTransition(ed, SLIDE, { kind: 'push', dir: 'u', duration: 1200 }));
    expect(slideTransition(ed.slide(SLIDE))).toMatchObject({ kind: 'push', dir: 'u', duration: 1200 });
    expect(cycle(ed, () => setTransitionOnAll(ed, { kind: 'fade', speed: 'med' }))).toEqual([]);
    ed.run(setTransitionOnAll(ed, { kind: 'fade', speed: 'med' }));
    expect(ed.doc.slideParts().map(p => slideTransition(ed.slide(p))?.kind)).toEqual(['fade', 'fade']);
    expect(cycle(ed, () => setSlideTransition(ed, SLIDE, null))).toEqual([]);
  });
});

describe('theme colours', () => {
  test('accents get 80/60/40% lighter and 25/50% darker; white and black their own steps', () => {
    const ed = openEditor(deckBytes());
    const cols = themePalette(ed, SLIDE);
    expect(cols.map(c => c.base.hex)).toEqual(['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']);
    expect(cols[4]!.variants.map(v => v.label.split(', ')[1])).toEqual(['Lighter 80%', 'Lighter 60%', 'Lighter 40%', 'Darker 25%', 'Darker 50%']);
    expect(cols[0]!.variants.map(v => v.label.split(', ')[1])).toEqual(['Darker 5%', 'Darker 15%', 'Darker 25%', 'Darker 35%', 'Darker 50%']);
    expect(cols[1]!.variants.map(v => v.label.split(', ')[1])).toEqual(['Lighter 50%', 'Lighter 35%', 'Lighter 25%', 'Lighter 15%', 'Lighter 5%']);
    expect(cols[4]!.variants[0]!.color).toEqual({ kind: 'scheme', val: 'accent1', lumMod: 20000, lumOff: 80000 });
    expect(cols[4]!.variants[0]!.hex).toBe('#DAE3F3');                 // PowerPoint shows #DAE3F3 for Office accent 1, lighter 80%
  });
});

