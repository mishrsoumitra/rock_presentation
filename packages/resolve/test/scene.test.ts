import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { applyMatrix, boxMatrix, resolveSlide, smartArtTextBox, toHex, type Paint, type ResolvedShape } from '../src/index.ts';
import { deckBytes, deckFiles } from './deck.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const open = (overrides: Record<string, string> = {}) => new PresentationDocument(OpcPackage.open(deckBytes(overrides)));
const colorOf = (p: Paint) => (p.kind === 'solid' ? toHex(p.color) : p.kind);
const byName = (shapes: ResolvedShape[], name: string) => {
  const s = shapes.find(x => x.name === name);
  if (!s) throw new Error(`no shape ${name}: ${shapes.map(x => x.name).join(', ')}`);
  return s;
};

describe('a slide through its layout and master', () => {
  const slide = resolveSlide(open(), SLIDE);

  test('the background comes from the master bgRef: background fill 1 with bg1', () => {
    expect(colorOf(slide.background)).toBe('#FFFFFF');
    expect(slide.size).toEqual({ cx: 12192000, cy: 6858000 });
  });

  test('drawing order: master non-placeholders, then the slide; inherited placeholders are not drawn', () => {
    expect(slide.shapes.map(s => `${s.layer}:${s.name}`)).toEqual([
      'master:Logo', 'slide:Title 1', 'slide:Content Placeholder 2', 'slide:Slide Number Placeholder 3', 'slide:Callout', 'slide:Group',
    ]);
    expect(colorOf(byName(slide.shapes, 'Logo').fill)).toBe('#ED7D31');
  });

  test('the title takes its position from the master and its text from the title style', () => {
    const t = byName(slide.shapes, 'Title 1');
    expect(t.xfrm).toMatchObject({ x: 838200, y: 365125, cx: 10515600, cy: 1325563 });
    const run = t.text!.paragraphs[0]!.runs[0]!;
    expect(run.text).toBe('Quarterly results');
    expect(run.props).toMatchObject({ size: 44, latin: 'Calibri Light', bold: false });
    expect(colorOf(run.props.fill)).toBe('#000000');
    expect(t.text!.body.anchor).toBe('ctr');
    expect(t.text!.paragraphs[0]!.props.lineSpacing).toEqual({ kind: 'pct', val: 0.9 });
    expect(t.text!.paragraphs[0]!.props.bullet).toBeNull();
  });

  test('the content placeholder matches by index, levels come from the body style, and its own formatting wins', () => {
    const c = byName(slide.shapes, 'Content Placeholder 2');
    expect(c.xfrm).toMatchObject({ x: 838200, y: 2000000, cx: 10515600, cy: 4000000 });
    expect(c.text!.body.autofit).toEqual({ kind: 'norm', fontScale: 0.925, lnSpcReduction: 0.1 });
    const [p1, p2, p3] = c.text!.paragraphs;
    expect(p1!.props).toMatchObject({ level: 0, marL: 228600, indent: -228600, spaceBefore: { kind: 'pts', val: 1000 } });
    expect(p1!.props.bullet).toMatchObject({ kind: 'char', char: '•', font: 'Arial' });
    expect(p1!.runs[0]!.props).toMatchObject({ size: 28, latin: 'Calibri' });
    expect(p2!.props.level).toBe(1);
    expect(p2!.runs[0]!.props).toMatchObject({ size: 24, bold: true });
    expect(colorOf(p2!.runs[0]!.props.fill)).toBe('#C00000');
    expect(colorOf(p3!.runs[0]!.props.fill)).toBe('#0563C1');
    expect(p3!.runs[0]!.props.underline).toBe('sng');
    expect(p3!.runs[0]!.props.hyperlink).toEqual({ url: 'https://example.com/' });
  });

  test('the slide number field shows the slide number in the master style', () => {
    const n = byName(slide.shapes, 'Slide Number Placeholder 3');
    expect(n.xfrm).toMatchObject({ x: 8610600, y: 6356350 });
    const run = n.text!.paragraphs[0]!.runs[0]!;
    expect(run).toMatchObject({ kind: 'field', text: '1', fieldType: 'slidenum' });
    expect(run.props.size).toBe(12);
    expect(colorOf(run.props.fill)).toBe('#898989');
    expect(n.text!.paragraphs[0]!.props.align).toBe('r');
  });

  test('a shape with a style: fill, line and text colour from the theme style matrix', () => {
    const s = byName(slide.shapes, 'Callout');
    expect(colorOf(s.fill)).toBe('#4472C4');
    expect(s.line).toMatchObject({ width: 12700, join: 'miter', cap: 'flat' });
    expect(colorOf(s.line!.paint)).toBe('#2F528F');
    expect(s.effects).toEqual({});
    const run = s.text!.paragraphs[0]!.runs[0]!;
    expect(colorOf(run.props.fill)).toBe('#FFFFFF');          // fontRef lt1 wins over the master's otherStyle tx1
    expect(run.props).toMatchObject({ size: 18, latin: 'Calibri' });
    expect(run.props.scriptFonts).toMatchObject({ Jpan: '游ゴシック' });
    expect(s.text!.paragraphs[0]!.props.align).toBe('ctr');
  });

  test('a rotated, scaled group places its member and passes on its fill', () => {
    const g = byName(slide.shapes, 'Group');
    const m = g.children[0]!;
    expect(colorOf(m.fill)).toBe('#00FF00');
    const near = ([x, y]: [number, number], [ex, ey]: [number, number]) => {
      expect(x).toBeCloseTo(ex, 3);
      expect(y).toBeCloseTo(ey, 3);
    };
    near(applyMatrix(m.transform, 0, 0), [2500000, 500000]);
    near(applyMatrix(m.transform, 4000000, 2000000), [1500000, 2500000]);
  });
});

describe('layers and options', () => {
  test('showMasterSp="0" on the slide hides the master and layout shapes', () => {
    const files = deckFiles();
    const slide = files['ppt/slides/slide1.xml']!.replace('<p:sld ', '<p:sld showMasterSp="0" ');
    const r = resolveSlide(open({ 'ppt/slides/slide1.xml': slide }), SLIDE);
    expect(r.shapes.some(s => s.layer === 'master')).toBe(false);
  });

  test('a layout on its own draws its placeholders over the master shapes', () => {
    const r = resolveSlide(open(), '/ppt/slideLayouts/slideLayout1.xml');
    expect(r.kind).toBe('layout');
    expect(r.shapes.map(s => `${s.layer}:${s.name}`)).toEqual(['master:Logo', 'layout:Title 1', 'layout:Content Placeholder 2', 'layout:Slide Number Placeholder 3']);
    expect(byName(r.shapes, 'Title 1').xfrm.x).toBe(838200);
  });

  test('a slide colour map override changes what tx1 and bg1 mean', () => {
    const files = deckFiles();
    const slide = files['ppt/slides/slide1.xml']!.replace('<a:masterClrMapping/>', '<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>');
    const r = resolveSlide(open({ 'ppt/slides/slide1.xml': slide }), SLIDE);
    expect(colorOf(r.background)).toBe('#000000');
    expect(colorOf(byName(r.shapes, 'Title 1').text!.paragraphs[0]!.runs[0]!.props.fill)).toBe('#FFFFFF');
  });
});

describe('SmartArt text boxes (dsp:txXfrm)', () => {
  const xf = (x: number, y: number, cx: number, cy: number, rot = 0) => ({ x, y, cx, cy, rot, flipH: false, flipV: false });
  // the direction text runs in on the slide, in degrees
  const angle = (shape: ReturnType<typeof xf>, text: ReturnType<typeof xf>, hasRot: boolean) => {
    const m = boxMatrix(smartArtTextBox(shape, text, hasRot), true);
    return Math.round(((Math.atan2(m[1], m[0]) * 180) / Math.PI + 360) % 360);
  };

  test('rot turns the text from the shape\'s angle: upright in arrows turned over, and in a turned matrix', () => {
    // the arrows of a "Vertical Arrow List" pointing left (a deck of geometry lectures)
    expect(angle(xf(1566730, 1165, 5602582, 918921, 10800000), xf(1566730, 1165, 5602582, 918921, 10800000), true)).toBe(0);
    // tdf145528's matrix: quarters turned by 270, 180 and 90 degrees, their text upright
    expect(angle(xf(5040000, 720000, 2160000, 3600000, 16200000), xf(4320000, 1710000, 3600000, 1620000, 5400000), true)).toBe(0);
    expect(angle(xf(5040000, 720000, 2160000, 3600000, 5400000), xf(4320000, 1710000, 3600000, 1620000, -5400000), true)).toBe(0);
  });

  test('without rot, the text turns with the shape if its box is the shape\'s, and stays upright if it fits the turned shape', () => {
    // a rectangle the size of its shape's box, turned 29.5 degrees (Mac PowerPoint)
    expect(angle(xf(802940, 562816, 330990, 177951, 1770845), xf(802940, 562816, 330990, 177951), false)).toBe(30);
    // a bar turned 90 degrees, with a wide text box across it (PowerPoint 2010)
    expect(angle(xf(5111415, -2065717, 865792, 5218176, 5400000), xf(2935223, 152739, 5175912, 781264), false)).toBe(0);
    // the text box is placed by its own rectangle, not carried around the shape's centre
    const m = boxMatrix(smartArtTextBox(xf(5111415, -2065717, 865792, 5218176, 5400000), xf(2935223, 152739, 5175912, 781264), false), true);
    expect(applyMatrix(m, 0, 0)).toEqual([2935223, 152739]);
  });

  test('a shape turned upside down by flipV turns its text over, as other shapes do', () => {
    expect(angle({ ...xf(0, 0, 100, 50), flipV: true }, xf(0, 0, 100, 50, 0), true)).toBe(180);
  });
});
