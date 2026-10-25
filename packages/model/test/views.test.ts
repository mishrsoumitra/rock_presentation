import { describe, expect, test } from 'vitest';
import { serializeElement } from '@pptx/xml';
import { checkOrder, children, lname, PresentationView, ShapeView, SlideView, shapesIn, type PartOwner } from '../src/index.ts';
import { find, NSDECL, parse, slideXml } from './helpers.ts';

const owner = (): PartOwner & { touched: number } => {
  const o = { part: '/ppt/slides/slide1.xml', touched: 0, touch: () => { o.touched++; } };
  return o;
};

const SHAPES =
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1" descr="alt"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
  '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm rot="5400000" flipH="1"><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm>' +
  '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom></p:spPr></p:sp>' +
  '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="1" y="1"/><a:ext cx="2" cy="2"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="200" cy="200"/><a:chOff x="0" y="0"/><a:chExt cx="400" cy="400"/></a:xfrm></p:grpSpPr>' +
  '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="6" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp></p:grpSp>' +
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
  '<p:xfrm><a:off x="5" y="6"/><a:ext cx="7" cy="8"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl/></a:graphicData></a:graphic></p:graphicFrame>';

describe('ShapeView', () => {
  test('reads kinds, ids, names, placeholders, transforms and geometry', () => {
    const o = owner();
    const shapes = shapesIn(find(parse(slideXml(SHAPES)), 'spTree'), o);
    expect(shapes.map(s => [s.kind, s.id, s.name])).toEqual([['sp', 2, 'Title 1'], ['sp', 3, 'Content'], ['pic', 4, 'Picture'], ['grpSp', 5, 'Group'], ['graphicFrame', 7, 'Table']]);
    const [title, content, pic, group, frame] = shapes as [ShapeView, ShapeView, ShapeView, ShapeView, ShapeView];
    expect(title.placeholder).toEqual({ type: 'title', idx: 0, hasIdx: false, orient: 'horz', sz: 'full', hasCustomPrompt: false });
    expect(title.descr).toBe('alt');
    expect(title.xfrm).toBeNull();
    expect(content.placeholder?.type).toBe('obj');
    expect(content.placeholder?.idx).toBe(1);
    expect(content.xfrm).toEqual({ x: 10, y: 20, cx: 300, cy: 400, rot: 5400000, flipH: true, flipV: false });
    expect(content.geometry).toEqual({ kind: 'preset', prst: 'roundRect', av: [{ name: 'adj', fmla: 'val 25000' }] });
    expect(pic.blipFill).toBeDefined();
    expect(group.childSpace).toEqual({ chOff: { x: 0, y: 0 }, chExt: { cx: 400, cy: 400 } });
    expect(group.children.map(c => c.kind)).toEqual(['cxnSp']);
    expect(frame.xfrm).toEqual({ x: 5, y: 6, cx: 7, cy: 8, rot: 0, flipH: false, flipV: false });
    expect(frame.graphicData?.uri).toBe('http://schemas.openxmlformats.org/drawingml/2006/table');
    expect(frame.spPr).toBeUndefined();
  });

  test('setXfrm writes where each kind keeps its transform, in schema order, and touches the part', () => {
    const o = owner();
    const sld = parse(slideXml(SHAPES));
    const shapes = shapesIn(find(sld, 'spTree'), o);
    const to = { x: 1000, y: 2000, cx: 3000, cy: 4000, rot: 0, flipH: false, flipV: true };
    for (const s of shapes) s.setXfrm(to);
    for (const s of shapes) expect(s.xfrm).toEqual(to);
    expect(o.touched).toBe(5);
    // the title placeholder had an empty spPr: the xfrm is created inside it
    expect(serializeElement(shapes[0]!.spPr!)).toBe('<p:spPr><a:xfrm flipV="1"><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm></p:spPr>');
    // the old rotation and flip are gone, the geometry stays after the xfrm
    expect(children(shapes[1]!.spPr).map(lname)).toEqual(['a:xfrm', 'a:prstGeom']);
    // the group keeps its child coordinate space
    expect(shapes[3]!.childSpace).toEqual({ chOff: { x: 0, y: 0 }, chExt: { cx: 400, cy: 400 } });
    // the graphic frame writes p:xfrm, never a p:spPr it is not allowed to have
    expect(children(shapes[4]!.el).map(lname)).toEqual(['p:nvGraphicFramePr', 'p:xfrm', 'a:graphic']);
    expect(checkOrder(sld)).toEqual([]);
  });

  test('a group without a child space gets one equal to its extent', () => {
    const sld = parse(slideXml('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:grpSp>'));
    const g = shapesIn(find(sld, 'spTree'), null)[0]!;
    g.setXfrm({ x: 1, y: 2, cx: 3, cy: 4, rot: 0, flipH: false, flipV: false });
    expect(g.childSpace).toEqual({ chOff: { x: 1, y: 2 }, chExt: { cx: 3, cy: 4 } });
    expect(checkOrder(sld)).toEqual([]);
  });

  test('mc:AlternateContent in a shape tree yields its fallback shape', () => {
    const sld = parse(slideXml('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">' +
      '<p:contentPart r:id="rId9"/></mc:Choice><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="8" name="Ink"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></mc:Fallback></mc:AlternateContent>'));
    expect(shapesIn(find(sld, 'spTree'), null).map(s => [s.kind, s.name])).toEqual([['pic', 'Ink']]);
  });
});

describe('SlideView and PresentationView', () => {
  test('a master: colour map, text styles; a layout: type and showMasterSp', () => {
    const o = owner();
    const master = new SlideView(parse(`<p:sldMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree/></p:cSld>` +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'), o);
    expect(master.kind).toBe('master');
    expect(master.clrMap?.bg1).toBe('lt1');
    expect(Object.keys(master.txStyles)).toEqual(['title', 'body', 'other']);
    expect(master.bg).toBeDefined();
    const layout = new SlideView(parse(`<p:sldLayout ${NSDECL} type="twoObj" showMasterSp="0"><p:cSld name="Two Content"><p:spTree/></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1"/></p:clrMapOvr></p:sldLayout>`), o);
    expect([layout.kind, layout.layoutType, layout.showMasterSp, layout.name]).toEqual(['layout', 'twoObj', false, 'Two Content']);
    expect(layout.clrMapOverride).toEqual({ bg1: 'dk1', tx1: 'lt1' });
  });

  test('slide id list: add, move, remove, and sections', () => {
    const o = owner();
    const pres = new PresentationView(parse(`<p:presentation ${NSDECL}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="260" r:id="rId3"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>' +
      '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
      '<p14:section name="Intro" id="{A}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>'), o);
    expect(pres.slideSize).toEqual({ cx: 12192000, cy: 6858000 });
    expect(pres.sections).toEqual([{ name: 'Intro', id: '{A}', slideIds: [256] }]);
    const id = pres.addSlideId('rId9', 1);
    expect(id).toBe(261);
    expect(pres.slideIds).toEqual([{ id: 256, rId: 'rId2' }, { id: 261, rId: 'rId9' }, { id: 260, rId: 'rId3' }]);
    pres.moveSlideId(256, 2);
    expect(pres.slideIds.map(e => e.id)).toEqual([261, 260, 256]);
    pres.removeSlideId(260);
    expect(pres.slideIds.map(e => e.id)).toEqual([261, 256]);
    expect(serializeElement(find(pres.root, 'sldIdLst'))).toBe('<p:sldIdLst><p:sldId id="261" r:id="rId9"/><p:sldId id="256" r:id="rId2"/></p:sldIdLst>');
    expect(checkOrder(pres.root)).toEqual([]);
    expect(o.touched).toBe(3);
  });
});
