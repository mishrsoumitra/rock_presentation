// The GDI player against a recording context: mapping modes and transforms, the device context and object
// tables, pens and brushes, paths, clipping, text layout and bitmaps. Pictures are built record by record.
import { describe, expect, test } from 'vitest';
import { drawMetafile, parseMetafile, type DrawOptions } from '../src/index.ts';
import { dibParts, E, emf, emr, packedDib, W, wmf } from './build.ts';
import { box, MockCanvas, MockGradient, MockPattern, points, RecordingContext, type Draw } from './mock.ts';

/** Draw records into a 200x200 destination: the default EMF frame is 100x100 device pixels, so output = 2 x device. */
function play(bytes: Uint8Array, opts: DrawOptions = {}, dest = { x: 0, y: 0, w: 200, h: 200 }): RecordingContext {
  const ctx = new RecordingContext(200, 200);
  drawMetafile(ctx, parseMetafile(bytes), dest, { createCanvas: (w, h) => new MockCanvas(w, h), ...opts });
  return ctx;
}

const at = (d: Draw, x = d.x!, y = d.y!): [number, number] => {
  const m = d.matrix!;
  const r = (v: number) => Math.round(v * 1000) / 1000 || 0;
  return [r(m[0] * x + m[2] * y + m[4]), r(m[1] * x + m[3] * y + m[5])];
};

const NULL_PEN = 0x80000008, NULL_BRUSH = 0x80000005, BLACK_BRUSH = 0x80000004;

describe('mapping modes and transforms', () => {
  test('MM_TEXT: logical units are device pixels; the frame maps onto the destination', () => {
    const fills = play(emf([E.rectangle(10, 20, 30, 40)])).of('fill');
    expect(box(fills[0]!.path)).toEqual([20, 40, 60, 80]);
  });

  test('the frame, not the bounds, maps onto the destination (offset and scaled)', () => {
    // frame 10..20 mm is device pixels 39.37..78.74 at 100 dpi: device 0 lands at -100, device 100 at 154
    const fills = play(emf([E.rectangle(0, 0, 100, 100)], { frame: [1000, 1000, 2000, 2000] }), {}, { x: 0, y: 0, w: 100, h: 100 }).of('fill');
    expect(box(fills[0]!.path)).toEqual([-100, -100, 154, 154]);
  });

  test('MM_ANISOTROPIC maps the window onto the viewport', () => {
    const d = play(emf([E.mapMode(8), E.windowExt(1000, 1000), E.viewportOrg(10, 10), E.viewportExt(50, 50), E.rectangle(100, 200, 300, 400)])).of('fill');
    expect(box(d[0]!.path)).toEqual([30, 40, 50, 60]);
  });

  test('MM_ISOTROPIC keeps units square: the smaller scale wins', () => {
    const d = play(emf([E.mapMode(7), E.windowExt(100, 200), E.viewportExt(100, 100), E.rectangle(0, 0, 100, 100)])).of('fill');
    expect(box(d[0]!.path)).toEqual([0, 0, 100, 100]);
  });

  test('window extents are ignored in MM_TEXT, as GDI ignores them', () => {
    const d = play(emf([E.windowExt(1000, 1000), E.viewportExt(10, 10), E.rectangle(10, 10, 20, 20)])).of('fill');
    expect(box(d[0]!.path)).toEqual([20, 20, 40, 40]);
  });

  test('MM_LOMETRIC: tenths of a millimetre, y up, through the reference device resolution', () => {
    const d = play(emf([E.mapMode(2), E.rectangle(0, 0, 100, -100)])).of('fill');
    const mm10 = 3.937007874015748;          // device pixels in 10 x 0.1 mm at 100 dpi, doubled below
    expect(box(d[0]!.path)).toEqual([0, 0, +(mm10 * 10 * 2).toFixed(3), +(mm10 * 10 * 2).toFixed(3)]);
  });

  test('switching from a metric mode to MM_ANISOTROPIC keeps its scale', () => {
    const d = play(emf([E.mapMode(2), E.mapMode(8), E.rectangle(0, 0, 100, -100)])).of('fill');
    expect(box(d[0]!.path)[2]).toBeCloseTo(78.74, 2);
  });

  test('SetWorldTransform, and ModifyWorldTransform left, right, identity', () => {
    const corner = (recs: Uint8Array[]) => points(play(emf([...recs, E.polygon([0, 0, 1, 0, 1, 1])])).of('fill')[0]!.path)[0];
    expect(corner([E.worldTransform([2, 0, 0, 2, 10, 0])])).toEqual([20, 0]);
    // left multiply: the new matrix applies first (0,0) -> (5,0) -> (2*5+10, 0)
    expect(corner([E.worldTransform([2, 0, 0, 2, 10, 0]), E.modifyWorldTransform([1, 0, 0, 1, 5, 0], 2)])).toEqual([40, 0]);
    // right multiply: the new matrix applies last (0,0) -> (10,0) -> (15,0)
    expect(corner([E.worldTransform([2, 0, 0, 2, 10, 0]), E.modifyWorldTransform([1, 0, 0, 1, 5, 0], 3)])).toEqual([30, 0]);
    expect(corner([E.worldTransform([2, 0, 0, 2, 10, 0]), E.modifyWorldTransform([1, 0, 0, 1, 0, 0], 1)])).toEqual([0, 0]);
  });

  test('a singular or non-finite world transform is ignored', () => {
    const d = play(emf([E.worldTransform([0, 0, 0, 0, 5, 5]), E.worldTransform([NaN, 0, 0, 1, 0, 0]), E.rectangle(1, 1, 2, 2)])).of('fill');
    expect(box(d[0]!.path)).toEqual([2, 2, 4, 4]);
  });

  test('SaveDC and RestoreDC bring back transforms, objects and modes; nested and relative', () => {
    const ctx = play(emf([
      E.saveDC(), E.windowOrg(50, 50), E.createBrush(1, 0, 0x0000ff), E.select(1),
      E.saveDC(), E.windowOrg(10, 10),
      E.restoreDC(-2), E.rectangle(0, 0, 10, 10),
    ]));
    const fill = ctx.of('fill')[0]!;
    expect(box(fill.path)).toEqual([0, 0, 20, 20]);
    expect(fill.style).toBe('rgb(255,255,255)');
  });

  test('a WMF without a placeable header takes its frame from the first window', () => {
    const bytes = wmf([W.mapMode(8), W.windowOrg(100, 100), W.windowExt(400, 200), W.rectangle(100, 100, 300, 200)]);
    const mf = parseMetafile(bytes);
    expect(mf.frame).toEqual({ l: 100, t: 100, r: 500, b: 300 });
    expect([mf.width, mf.height]).toEqual([400, 200]);
    const d = play(bytes, {}, { x: 0, y: 0, w: 400, h: 200 }).of('fill');
    expect(box(d[0]!.path)).toEqual([0, 0, 200, 100]);
  });

  test('a placeable WMF: the box sets the frame and the size; a later window rescales within it', () => {
    const bytes = wmf([W.windowOrg(0, 0), W.windowExt(2880, 1440), W.rectangle(0, 0, 1440, 720)], { l: 0, t: 0, r: 1440, b: 720, inch: 1440 });
    const mf = parseMetafile(bytes);
    expect([mf.width, mf.height]).toEqual([96, 48]);
    const d = play(bytes, {}, { x: 0, y: 0, w: 96, h: 48 }).of('fill');
    expect(box(d[0]!.path)).toEqual([0, 0, 48, 24]);
  });
});

describe('objects', () => {
  test('WMF objects take the lowest free slot', () => {
    const bytes = wmf([
      W.windowExt(100, 100),
      W.createPen(0, 0, 0x0000ff),            // slot 0
      W.createBrush(0, 0x00ff00),             // slot 1
      W.createBrush(0, 0xff0000),             // slot 2
      W.deleteObject(1),
      W.createBrush(0, 0x00ffff),             // reuses slot 1
      W.select(1), W.rectangle(0, 0, 10, 10),
      W.select(2), W.rectangle(0, 0, 10, 10),
      W.select(0), W.rectangle(0, 0, 10, 10),
    ]);
    const ctx = play(bytes);
    expect(ctx.of('fill').map(d => d.style)).toEqual(['rgb(255,255,0)', 'rgb(0,0,255)', 'rgb(0,0,255)']);
    expect(ctx.of('stroke').map(d => d.style)).toEqual(['rgb(0,0,0)', 'rgb(0,0,0)', 'rgb(255,0,0)']);
  });

  test('EMF stock objects: the null brush and pen draw nothing, the black brush fills black', () => {
    const ctx = play(emf([E.select(NULL_BRUSH), E.rectangle(0, 0, 10, 10), E.select(NULL_PEN), E.select(BLACK_BRUSH), E.rectangle(0, 0, 10, 10)]));
    expect(ctx.of('fill').map(d => d.style)).toEqual(['rgb(0,0,0)']);
    expect(ctx.of('stroke')).toHaveLength(1);
  });

  test('an object deleted while selected stays in use; selecting a deleted slot changes nothing', () => {
    const ctx = play(emf([E.createBrush(1, 0, 0x0000ff), E.select(1), E.deleteObject(1), E.select(1), E.rectangle(0, 0, 1, 1)]));
    expect(ctx.of('fill')[0]!.style).toBe('rgb(255,0,0)');
  });

  test('palette-index colours go through the selected palette', () => {
    const bytes = wmf([W.windowExt(100, 100), W.createPalette([0x00ff00, 0x123456]), W.select(0), W.createBrush(0, 0x01000001), W.select(1), W.rectangle(0, 0, 1, 1)]);
    expect(play(bytes).of('fill')[0]!.style).toBe('rgb(86,52,18)');
  });
});

describe('pens and brushes', () => {
  test('a cosmetic pen is one natural pixel wide; a geometric one scales with the transform', () => {
    // natural size 96 px drawn at 200: one pixel is 200/96
    const ctx = play(emf([E.rectangle(0, 0, 1, 1), E.extCreatePen(1, 0x10000, 10, 0), E.select(1), E.rectangle(0, 0, 1, 1)]));
    const widths = ctx.of('stroke').map(d => +d.lineWidth!.toFixed(3));
    expect(widths).toEqual([+(200 / 96).toFixed(3), 20]);
  });

  test('styled pens: cosmetic dashes in pixels, geometric in pen widths; gaps filled in OPAQUE mode', () => {
    const px = 200 / 96;
    const ctx = play(emf([E.bkMode(1), E.createPen(1, 1, 0, 0), E.select(1), E.lineTo(10, 0)]));
    expect(ctx.of('stroke')[0]!.dash!.map(v => +v.toFixed(3))).toEqual([18 * px, 6 * px].map(v => +v.toFixed(3)));
    const geo = play(emf([E.bkMode(1), E.extCreatePen(1, 0x10000 | 2, 4, 0), E.select(1), E.lineTo(10, 0)]));
    expect(geo.of('stroke')[0]!.dash).toEqual([8, 8]);
    // OPAQUE (the default): a solid background-coloured stroke first
    const opaque = play(emf([E.createPen(1, 2, 0, 0x0000ff), E.select(1), E.lineTo(10, 0)])).of('stroke');
    expect(opaque.map(d => [d.style, d.dash!.length > 0])).toEqual([['rgb(255,255,255)', false], ['rgb(255,0,0)', true]]);
  });

  test('a wide CreatePen pen with a dash style draws solid, as GDI does', () => {
    expect(play(emf([E.createPen(1, 1, 5, 0), E.select(1), E.lineTo(10, 0)])).of('stroke').at(-1)!.dash).toEqual([]);
  });

  test('user-styled geometric pens scale their dashes with the transform', () => {
    const ctx = play(emf([E.bkMode(1), E.extCreatePen(1, 0x10000 | 7, 2, 0, [3, 1]), E.select(1), E.lineTo(10, 0)]));
    expect(ctx.of('stroke')[0]!.dash).toEqual([6, 2]);
  });

  test('a hatch brush fills with an 8x8 pattern in device pixels, transparent in TRANSPARENT mode', () => {
    const ctx = play(emf([E.bkMode(1), E.createBrush(1, 2, 0x0000ff, 0), E.select(1), E.rectangle(0, 0, 10, 10)]));
    const p = ctx.of('fill')[0]!.style as MockPattern;
    expect(p).toBeInstanceOf(MockPattern);
    expect((p.image as MockCanvas).width).toBe(8);
    expect(p.transform?.a).toBeCloseTo(200 / 96);
    const tile = (p.image as MockCanvas).ctx.images[0]!.data;
    // HS_HORIZONTAL: row 6 is the line, the rest transparent
    expect(Array.from(tile.subarray((6 * 8) * 4, (6 * 8) * 4 + 4))).toEqual([255, 0, 0, 255]);
    expect(tile[3]).toBe(0);
  });

  test('a DIB pattern brush tiles its bitmap; mono brushes take the text and background colours', () => {
    const mono = E.patternBrush(93, 1, dibParts({ width: 2, height: 2, bitCount: 1, colors: [0, 0xffffff], rows: [[1, 0], [0, 1]] }));
    const ctx = play(emf([E.textColor(0x0000ff), E.bkColor(0x00ff00), mono, E.select(1), E.rectangle(0, 0, 4, 4)]));
    const p = ctx.of('fill')[0]!.style as MockPattern;
    const px = (p.image as MockCanvas).ctx.images[0]!.data;
    expect(Array.from(px.subarray(0, 8))).toEqual([0, 255, 0, 255, 255, 0, 0, 255]);
  });

  test('without a canvas factory, patterns fall back to their colour', () => {
    const bytes = emf([E.createBrush(1, 2, 0x0000ff, 4), E.select(1), E.rectangle(0, 0, 10, 10)]);
    const ctx = new RecordingContext();
    drawMetafile(ctx, parseMetafile(bytes), { x: 0, y: 0, w: 200, h: 200 });
    expect(ctx.of('fill')[0]!.style).toBe('rgb(255,0,0)');
  });

  test('the fill mode decides the rule; ROP2 changes how pens and brushes paint', () => {
    const ctx = play(emf([
      E.polygon([0, 0, 10, 0, 0, 10]), E.polyFillMode(2), E.polygon([0, 0, 10, 0, 0, 10]),
      E.rop2(7), E.rectangle(0, 0, 1, 1), E.rop2(11), E.rectangle(0, 0, 1, 1), E.rop2(1), E.rectangle(0, 0, 1, 1),
    ]));
    const fills = ctx.of('fill');
    expect(fills.slice(0, 2).map(d => d.rule)).toEqual(['evenodd', 'nonzero']);
    expect(fills[2]!.composite).toBe('difference');
    expect(fills[3]!.style).toBe('rgb(0,0,0)');                  // R2_NOP drew nothing; R2_BLACK fills black
    expect(fills).toHaveLength(4);
  });
});

describe('shapes and paths', () => {
  test('Arc goes counter-clockwise from the start radial to the end radial; SetArcDirection reverses it', () => {
    const quarter = play(emf([E.arc([0, 0, 100, 100], [100, 50], [50, 0])])).of('stroke')[0]!;
    const pts = points(quarter.path);
    expect(pts[0]).toEqual([200, 100]);
    expect(pts.at(-1)![0]).toBeCloseTo(100);
    expect(pts.at(-1)![1]).toBeCloseTo(0);
    expect(pts.every(([x, y]) => x >= 99.9 && y <= 100.1)).toBe(true);        // stays in the upper right
    const long = points(play(emf([E.arcDirection(2), E.arc([0, 0, 100, 100], [100, 50], [50, 0])])).of('stroke')[0]!.path);
    expect(long.some(([x, y]) => x < 50 && y > 100)).toBe(true);               // goes round through the lower left
  });

  test('the arc direction is the device\'s: with y up in logical units, the same quarter is drawn', () => {
    const yUp = [E.mapMode(8), E.windowOrg(0, 100), E.windowExt(100, -100), E.viewportExt(100, 100)];
    const pie = (records: Uint8Array[], end: number[]) => box(play(emf([...records, emr(47, b => b.i32(0, 0, 100, 100, 100, 50, ...end))])).of('fill')[0]!.path);
    expect(pie([], [50, 0])).toEqual([100, 0, 200, 100]);
    expect(pie(yUp, [50, 100])).toEqual([100, 0, 200, 100]);
  });

  test('Ellipse is four Beziers inside its box', () => {
    const fill = play(emf([E.ellipse(0, 0, 100, 50)])).of('fill')[0]!;
    expect(fill.path!.filter(p => p[0] === 'C')).toHaveLength(4);
    expect(box(fill.path)).toEqual([0, 0, 200, 100]);
  });

  test('a path bracket collects figures; FillPath fills them once, closed, and does not stroke', () => {
    const ctx = play(emf([E.beginPath(), E.moveTo(0, 0), E.lineTo(10, 0), E.lineTo(10, 10), E.endPath(), E.fillPath()]));
    expect(ctx.of('fill')).toHaveLength(1);
    expect(ctx.of('stroke')).toHaveLength(0);
    expect(ctx.of('fill')[0]!.path!.at(-1)).toEqual(['Z']);
  });

  test('PolyBezierTo continues from the current position and moves it', () => {
    const ctx = play(emf([E.moveTo(1, 1), E.polyBezierTo([2, 2, 3, 3, 4, 4]), E.lineTo(5, 5)]));
    const [curve, line] = ctx.of('stroke');
    expect(curve!.path![0]).toEqual(['M', 2, 2]);
    expect(line!.path).toEqual([['M', 8, 8], ['L', 10, 10]]);
  });

  test('PolyPolygon fills all its polygons as one path', () => {
    const fills = play(emf([E.polyPolygon([[0, 0, 10, 0, 10, 10], [20, 20, 30, 20, 30, 30]])])).of('fill');
    expect(fills).toHaveLength(1);
    expect(fills[0]!.path!.filter(p => p[0] === 'M')).toHaveLength(2);
  });
});

describe('clipping', () => {
  test('IntersectClipRect clips later drawing; SaveDC/RestoreDC brings the old clip back', () => {
    const ctx = play(emf([E.saveDC(), E.intersectClip(10, 10, 50, 50), E.rectangle(0, 0, 1, 1), E.restoreDC(-1), E.rectangle(0, 0, 1, 1)]));
    const clipsBeforeFirstFill = ctx.draws.slice(0, ctx.draws.findIndex(d => d.op === 'fill')).filter(d => d.op === 'clip');
    expect(clipsBeforeFirstFill.at(-1)!.path).toEqual([['R', 20, 20, 100, 100]]);
    // the second rectangle's clip is only the destination's (set once at the start)
    const i2 = ctx.draws.findLastIndex(d => d.op === 'fill');
    const i1 = ctx.draws.findIndex(d => d.op === 'fill');
    expect(ctx.draws.slice(i1 + 1, i2).filter(d => d.op === 'clip')).toHaveLength(0);
  });

  test('ExcludeClipRect cuts a hole: the region becomes four rectangles around it', () => {
    const ctx = play(emf([E.intersectClip(0, 0, 50, 50), E.excludeClip(10, 10, 20, 20), E.rectangle(0, 0, 1, 1)]));
    const clip = ctx.of('clip').at(-1)!;
    expect(clip.path!.filter(p => p[0] === 'R')).toHaveLength(4);
  });

  test('EXTSELECTCLIPRGN: rectangles in device units; RGN_COPY with no data resets', () => {
    const ctx = play(emf([E.clipRegion(5, [[0, 0, 10, 10], [20, 20, 30, 30]]), E.rectangle(0, 0, 1, 1), E.clipRegion(5, null), E.rectangle(0, 0, 1, 1)]));
    const clips = ctx.of('clip');
    expect(clips[1]!.path).toEqual([['R', 0, 0, 20, 20], ['R', 40, 40, 60, 60]]);
    expect(clips).toHaveLength(2);
  });

  test('SelectClipPath intersects with the path', () => {
    const ctx = play(emf([E.beginPath(), E.moveTo(0, 0), E.lineTo(10, 0), E.lineTo(0, 10), E.closeFigure(), E.endPath(), E.selectClipPath(1), E.rectangle(0, 0, 50, 50)]));
    const clip = ctx.of('clip').at(-1)!;
    expect(points(clip.path)).toEqual([[0, 0], [20, 0], [0, 20]]);
  });
});

describe('text', () => {
  // EMF: logical = device; output = 2 x device. Font em 20 logical = 40 output. The mock measures half an em per
  // character, with ascent 0.8 and descent 0.2 of the em.
  const text = (recs: Uint8Array[]) => play(emf([E.createFont(1, -20, 'Arial'), E.select(1), E.bkMode(1), ...recs]));

  test('the font becomes a CSS font with the face and fallbacks', () => {
    const d = text([E.extTextOutW(10, 10, 'A')]).of('fillText')[0]!;
    expect(d.font).toMatch(/^400 [\d.]+px "Arial", "Liberation Sans", "Helvetica", sans-serif$/);
  });

  test('TA_TOP (the default) puts the reference point at the top of the cell', () => {
    expect(at(text([E.extTextOutW(10, 10, 'A')]).of('fillText')[0]!)).toEqual([20, 52]);
  });

  test('TA_BASELINE and TA_BOTTOM', () => {
    expect(at(text([E.textAlign(24), E.extTextOutW(10, 10, 'A')]).of('fillText')[0]!)).toEqual([20, 20]);
    expect(at(text([E.textAlign(8), E.extTextOutW(10, 10, 'A')]).of('fillText')[0]!)).toEqual([20, 12]);
  });

  test('TA_CENTER and TA_RIGHT move the string by half and all of its width', () => {
    expect(at(text([E.textAlign(24 | 6), E.extTextOutW(50, 10, 'abcd')]).of('fillText')[0]!)).toEqual([60, 20]);
    expect(at(text([E.textAlign(24 | 2), E.extTextOutW(50, 10, 'abcd')]).of('fillText')[0]!)).toEqual([20, 20]);
  });

  test('advances from the record place each character', () => {
    const ds = text([E.textAlign(24), E.extTextOutW(10, 10, 'abc', { dx: [10, 15, 20] })]).of('fillText');
    expect(ds.map(d => [d.text, ...at(d)])).toEqual([['a', 20, 20], ['b', 40, 20], ['c', 70, 20]]);
  });

  test('TA_UPDATECP draws at the current position and moves it by the advances', () => {
    const ctx = text([E.textAlign(24 | 1), E.moveTo(5, 5), E.extTextOutW(99, 99, 'ab', { dx: [7, 8] }), E.lineTo(20, 50)]);
    expect(at(ctx.of('fillText')[0]!)).toEqual([10, 10]);
    expect(ctx.of('stroke')[0]!.path).toEqual([['M', 40, 10], ['L', 40, 100]]);
  });

  test('escapement turns the baseline counter-clockwise', () => {
    const ctx = play(emf([E.createFont(1, -20, 'Arial', { escapement: 900 }), E.select(1), E.bkMode(1), E.textAlign(24), E.extTextOutW(50, 50, 'a')]));
    const m = ctx.of('fillText')[0]!.matrix!;
    expect(m[0]).toBeCloseTo(0);
    expect(m[1]).toBeLessThan(0);        // the glyphs' x axis points up the page
  });

  test('OPAQUE mode paints the text cell in the background colour first; ETO_OPAQUE paints the rectangle', () => {
    const ctx = play(emf([E.createFont(1, -20, 'Arial'), E.select(1), E.bkColor(0x00ff00), E.extTextOutW(10, 10, 'ab')]));
    expect(ctx.of('fillRect')[0]!.style).toBe('rgb(0,255,0)');
    const eto = play(emf([E.createFont(1, -20, 'Arial'), E.select(1), E.bkMode(1), E.bkColor(0x0000ff), E.extTextOutW(10, 10, 'ab', { options: 2, rect: [0, 0, 40, 30] })]));
    expect(eto.of('fill')[0]!.style).toBe('rgb(255,0,0)');
    expect(box(eto.of('fill')[0]!.path)).toEqual([0, 0, 80, 60]);
  });

  test('ETO_CLIPPED clips the string to the rectangle', () => {
    const ctx = text([E.extTextOutW(10, 10, 'ab', { options: 4, rect: [0, 0, 15, 15] })]);
    const i = ctx.draws.findIndex(d => d.op === 'fillText');
    expect(box(ctx.draws.slice(0, i).filter(d => d.op === 'clip').at(-1)!.path)).toEqual([0, 0, 30, 30]);
  });

  test('glyph-index text is not drawn (its font is not available)', () => {
    expect(text([E.extTextOutW(10, 10, '$', { options: 0x10 })]).of('fillText')).toHaveLength(0);
  });

  test('Symbol font text comes out as Unicode Greek and operators', () => {
    const ctx = play(emf([E.createFont(1, -20, 'Symbol', { charSet: 2 }), E.select(1), E.bkMode(1), E.extTextOutW(0, 0, 'a\uf0e5')]));
    expect(ctx.of('fillText')[0]!.text).toBe('\u03b1\u2211');
    expect(ctx.of('fillText')[0]!.font).not.toContain('"Symbol"');
  });

  test('8-bit WMF text is decoded in the font\'s character set; DEFAULT_CHARSET goes by the face name', () => {
    const cyr = wmf([W.windowExt(100, 100), W.createFont(-10, 'Arial', { charSet: 204 }), W.select(0), W.bkMode(1), W.textOut(0, 0, '\u00c0\u00c1')]);
    expect(play(cyr).of('fillText')[0]!.text).toBe('\u0410\u0411');
    const gbk = wmf([W.windowExt(100, 100), W.createFont(-10, 'Microsoft YaHei UI', { charSet: 1 }), W.select(0), W.bkMode(1),
      W.extTextOut(0, 0, [208, 194, 189, 168])]);
    expect(play(gbk).of('fillText')[0]!.text).toBe('\u65b0\u5efa');
  });

  test('advances given per byte are summed per character for double-byte text', () => {
    const gbk = wmf([W.windowExt(100, 100), W.createFont(-10, 'SimSun', { charSet: 134 }), W.select(0), W.bkMode(1), W.textAlign(24),
      W.extTextOut(0, 0, [208, 194, 189, 168], { dx: [6, 6, 7, 7] })]);
    const ds = play(gbk).of('fillText');
    expect(ds.map(d => at(d)[0])).toEqual([0, 24]);
  });

  test('a stand-in font wider than the advances is narrowed instead of overlapping', () => {
    const ds = text([E.textAlign(24), E.extTextOutW(0, 0, 'ab', { dx: [5, 5] })]).of('fillText');
    // natural width 2 x 10 logical = 20 against 10 of advances: glyphs drawn at half width, still 5 logical apart
    expect(ds[0]!.matrix![0] / Math.abs(ds[0]!.matrix![3])).toBeCloseTo(0.5);
    expect(ds.map(d => at(d)[0])).toEqual([0, 10]);
  });
});

describe('bitmaps and raster operations', () => {
  const dib = dibParts({ width: 2, height: 2, bitCount: 24, rows: [[0xff0000, 0x00ff00], [0x0000ff, 0xffffff]] });

  test('BitBlt SRCCOPY draws the bitmap over the destination rectangle', () => {
    const ctx = play(emf([E.bitBlt(10, 10, 2, 2, 0xcc0020, dib)]));
    const d = ctx.of('drawImage')[0]!;
    expect(d.src).toEqual([0, 0, 2, 2]);
    expect(at(d, 0, 0)).toEqual([20, 20]);
    expect(at(d, 2, 2)).toEqual([24, 24]);
    expect((d.image as MockCanvas).ctx.images[0]!.data.subarray(0, 4)).toEqual(Uint8ClampedArray.from([255, 0, 0, 255]));
  });

  test('a stretched DIB with a source rectangle, measured from the bottom of a bottom-up DIB', () => {
    const ctx = play(emf([E.stretchDIBits(0, 0, 50, 50, [1, 0, 1, 1], dib)]));
    // the bottom-left 1x1 source starting at x=1: in rows from the top, row 1
    expect(ctx.of('drawImage')[0]!.src).toEqual([1, 1, 1, 1]);
  });

  test('PatBlt-style BitBlt fills with the brush; BLACKNESS, WHITENESS and DSTINVERT', () => {
    const ctx = play(emf([E.createBrush(1, 0, 0x0000ff), E.select(1), E.bitBlt(0, 0, 10, 10, 0xf00021), E.bitBlt(0, 0, 10, 10, 0x000042),
      E.bitBlt(0, 0, 10, 10, 0xff0062), E.bitBlt(0, 0, 10, 10, 0x550009), E.bitBlt(0, 0, 10, 10, 0xaa0029)]));
    expect(ctx.of('fill').map(d => [d.style, d.composite])).toEqual([
      ['rgb(255,0,0)', 'source-over'], ['#000', 'source-over'], ['#fff', 'source-over'], ['#fff', 'difference'],
    ]);
  });

  test('the mask idiom: SRCAND multiplies, SRCPAINT lightens, pixel-exact', () => {
    const ctx = play(emf([E.bitBlt(0, 0, 2, 2, 0x8800c6, dib), E.bitBlt(0, 0, 2, 2, 0xee0086, dib)]));
    expect(ctx.of('drawImage').map(d => [d.composite, d.smoothing])).toEqual([['multiply', false], ['lighten', false]]);
  });

  test('the XOR idiom (SRCINVERT, SRCAND mask, SRCINVERT) draws the image where the mask is black', () => {
    const mask = dibParts({ width: 2, height: 2, bitCount: 1, colors: [0, 0xffffff], rows: [[0, 1], [1, 0]] });
    const ctx = play(emf([E.bitBlt(0, 0, 2, 2, 0x660046, dib), E.bitBlt(0, 0, 2, 2, 0x8800c6, mask), E.bitBlt(0, 0, 2, 2, 0x660046, dib)]));
    const draws = ctx.of('drawImage');
    expect(draws).toHaveLength(1);
    const px = (draws[0]!.image as MockCanvas).ctx.images[0]!.data;
    expect([px[3], px[7], px[11], px[15]]).toEqual([255, 0, 0, 255]);
  });

  test('SetDIBitsToDevice copies pixel for pixel: its size is in device pixels, not logical units', () => {
    const ctx = play(emf([E.mapMode(8), E.windowExt(50, 50), E.viewportExt(100, 100), E.setDIBitsToDevice(10, 10, 2, 2, dib)]));
    const d = ctx.of('drawImage')[0]!;
    expect(at(d, 0, 0)).toEqual([40, 40]);
    expect(at(d, 2, 2)).toEqual([44, 44]);
  });

  test('ALPHABLEND: constant alpha, and per-pixel premultiplied alpha divided back out', () => {
    const argb = dibParts({ width: 1, height: 1, bitCount: 32, rows: [[0x80400000]] });
    const ctx = play(emf([E.blend(114, 0, 0, 10, 10, (0x01 << 24) | (0x80 << 16), argb)]));
    const d = ctx.of('drawImage')[0]!;
    expect(d.alpha).toBeCloseTo(128 / 255);
    expect(Array.from((d.image as MockCanvas).ctx.images[0]!.data)).toEqual([128, 0, 0, 128]);
  });

  test('TRANSPARENTBLT leaves out the transparent colour', () => {
    const ctx = play(emf([E.blend(116, 0, 0, 2, 2, 0x0000ff, dib)]));
    const px = (ctx.of('drawImage')[0]!.image as MockCanvas).ctx.images[0]!.data;
    expect([px[3], px[7], px[11], px[15]]).toEqual([0, 255, 255, 255]);
  });

  test('MASKBLT draws the source where the mask is 1', () => {
    const mask = dibParts({ width: 2, height: 2, bitCount: 1, colors: [0, 0xffffff], rows: [[1, 0], [0, 1]] });
    const px = (play(emf([E.maskBlt(0, 0, 2, 2, dib, mask)])).of('drawImage')[0]!.image as MockCanvas).ctx.images[0]!.data;
    expect([px[3], px[7], px[11], px[15]]).toEqual([255, 0, 0, 255]);
  });

  test('PLGBLT maps the source onto a parallelogram', () => {
    const d = play(emf([E.plgBlt([10, 0, 20, 10, 0, 10], 2, 2, dib)])).of('drawImage')[0]!;
    expect(at(d, 0, 0)).toEqual([20, 0]);
    expect(at(d, 2, 0)).toEqual([40, 20]);
    expect(at(d, 0, 2)).toEqual([0, 20]);
  });

  test('WMF DIBBITBLT draws its packed DIB', () => {
    const bytes = wmf([W.windowExt(100, 100), W.dibBitBlt(10, 20, 2, 2, 0, 0, 0xcc0020, packedDib({ width: 2, height: 2, bitCount: 24, rows: [[0, 0], [0, 0]] }))]);
    const d = play(bytes, {}, { x: 0, y: 0, w: 100, h: 100 }).of('drawImage')[0]!;
    expect(at(d, 0, 0)).toEqual([10, 20]);
  });

  test('without a canvas factory bitmaps are skipped, not failed', () => {
    const ctx = new RecordingContext();
    drawMetafile(ctx, parseMetafile(emf([E.bitBlt(0, 0, 2, 2, 0xcc0020, dib)])), { x: 0, y: 0, w: 100, h: 100 });
    expect(ctx.of('drawImage')).toHaveLength(0);
  });
});

describe('playback as a whole', () => {
  test('the context is left as it was found: saves and restores balance, the destination clips', () => {
    const ctx = play(emf([E.saveDC(), E.intersectClip(0, 0, 5, 5), E.rectangle(0, 0, 9, 9), E.createFont(1, -10, 'Arial'), E.select(1), E.extTextOutW(0, 0, 'x')]));
    expect(ctx.depth).toBe(0);
    expect(ctx.calls.filter(c => c === 'save').length).toBe(ctx.calls.filter(c => c === 'restore').length);
    expect(ctx.of('clip')[0]!.path).toEqual([['R', 0, 0, 200, 200]]);
  });

  test('records after EOF are not played', () => {
    const bytes = emf([E.rectangle(0, 0, 1, 1)]);
    const extra = new Uint8Array(bytes.length + E.rectangle(0, 0, 2, 2).length);
    extra.set(bytes);
    extra.set(E.rectangle(0, 0, 2, 2), bytes.length);
    expect(play(extra).of('fill')).toHaveLength(1);
  });

  test('the work limit stops a picture that asks for too much', () => {
    const many = Array.from({ length: 50 }, () => E.rectangle(0, 0, 1, 1));
    expect(() => play(emf(many), { maxWork: 100 })).toThrow(expect.objectContaining({ code: 'limit' }));
  });

  test('gradients: rectangle mode fills with a linear gradient between its two corners', () => {
    const ctx = play(emf([gradient()]));
    const g = ctx.of('fill')[0]!.style as MockGradient;
    expect(g.coords).toEqual([0, 0, 20, 0]);
    expect(g.stops).toEqual([[0, 'rgb(255,0,0)'], [1, 'rgb(0,0,255)']]);
  });
});

/** EMR_GRADIENTFILL, one horizontal rectangle from red at (0,0) to blue at (10,10). */
function gradient(): Uint8Array {
  const b: number[] = [];
  const u32 = (...v: number[]) => { for (const x of v) b.push(x & 255, (x >>> 8) & 255, (x >>> 16) & 255, (x >>> 24) & 255); };
  const u16 = (...v: number[]) => { for (const x of v) b.push(x & 255, (x >>> 8) & 255); };
  u32(118, 0, 0, 0, 0, 0, 2, 1, 0);
  u32(0, 0); u16(0xff00, 0, 0, 0xff00);
  u32(10, 10); u16(0, 0, 0xff00, 0xff00);
  u32(0, 1);
  const out = Uint8Array.from(b);
  new DataView(out.buffer).setUint32(4, out.length, true);
  return out;
}
