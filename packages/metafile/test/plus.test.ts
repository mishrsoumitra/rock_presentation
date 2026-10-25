// The EMF+ player on hand-made EMF+-only pictures, and how EMF and EMF+ records are chosen in dual pictures.
import { describe, expect, test } from 'vitest';
import { drawMetafile, parseMetafile, type DrawOptions } from '../src/index.ts';
import { plusImages } from '../src/plus.ts';
import { Bytes, E, emf, plus, plusComment, plusHeader } from './build.ts';
import { box, MockCanvas, MockGradient, points, RecordingContext } from './mock.ts';
import { plusObjects, plusOnly } from './samples.ts';

/** Draw into a 200x200 destination: the frame is 100x100 device pixels at 100 dpi, so output = 2 x pixels. */
function play(bytes: Uint8Array, opts: DrawOptions = {}): RecordingContext {
  const ctx = new RecordingContext(200, 200);
  drawMetafile(ctx, parseMetafile(bytes), { x: 0, y: 0, w: 200, h: 200 }, { createCanvas: (w, h) => new MockCanvas(w, h), ...opts });
  return ctx;
}

describe('EMF+ only pictures', () => {
  const ctx = play(plusOnly());
  const fills = ctx.of('fill');
  const strokes = ctx.of('stroke');

  test('the header says EMF+ only, so the EMF+ records are played', () => {
    expect(parseMetafile(plusOnly()).plus).toBe('only');
    expect(fills.length).toBeGreaterThan(5);
  });

  test('FillRects with a colour in the record', () => {
    expect(fills[0]!.style).toBe('rgba(255,0,0,1)');
    expect(box(fills[0]!.path)).toEqual([20, 20, 80, 60]);
  });

  test('DrawRects with a pen object: world-unit width, its dash style and brush colour', () => {
    expect(strokes[0]!.lineWidth).toBe(8);
    expect(strokes[0]!.dash).toEqual([24, 8]);
    expect(strokes[0]!.style).toBe('rgba(0,0,255,1)');
  });

  test('FillPath and DrawPath with a path object; DrawLines closed when the L flag is set', () => {
    expect(fills[1]!.style).toBe('rgba(0,255,0,1)');
    expect(points(fills[1]!.path)).toEqual([[0, 0], [100, 0], [0, 100]]);
    expect(fills[1]!.path!.at(-1)).toEqual(['Z']);
    expect(points(strokes[1]!.path)).toEqual([[0, 0], [100, 0], [0, 100]]);
    expect(strokes[2]!.path!.at(-1)).toEqual(['Z']);
  });

  test('a path object fills by the alternate rule, the GDI+ default that a serialized path cannot change', () => {
    expect(fills[1]!.rule).toBe('evenodd');
  });

  test('FillEllipse and FillPie', () => {
    expect(fills[2]!.path!.filter(p => p[0] === 'C')).toHaveLength(4);
    expect(box(fills[2]!.path)).toEqual([120, 120, 160, 140]);
    // the pie starts at the centre, and covers the quarter clockwise from the x axis (down on the page)
    expect(fills[3]!.path![0]).toEqual(['M', 40, 40]);
    expect(box(fills[3]!.path)).toEqual([40, 40, 80, 80]);
  });

  test('DrawImage through the world transform: the bitmap object decoded, 32-bpp ARGB', () => {
    const img = ctx.of('drawImage')[0]!;
    const m = img.matrix!;
    expect([m[4], m[5]]).toEqual([150, 10]);
    expect(m[0] * 2).toBeCloseTo(40);
    expect(Array.from((img.image as MockCanvas).ctx.images[0]!.data.subarray(12, 16))).toEqual([255, 255, 255, 128]);
  });

  test('DrawString in a clip: the font object, the brush colour, near alignment at the top of the layout box', () => {
    const t = ctx.of('fillText')[0]!;
    expect(t.text).toBe('Hi');
    expect(t.font).toMatch(/^700 [\d.]+px "Arial"/);
    const i = ctx.draws.indexOf(t);
    expect(box(ctx.draws.slice(0, i).filter(d => d.op === 'clip').at(-1)!.path)).toEqual([0, 0, 120, 120]);
  });

  test('Save, TranslateWorldTransform, Restore', () => {
    const poly = fills.find(f => f.style === 'rgba(18,52,86,1)')!;
    expect(points(poly.path)[0]).toEqual([20, 20]);
  });

  test('SetClipRegion with a region object: rectangles combined exactly', () => {
    const last = fills.at(-1)!;
    const i = ctx.draws.indexOf(last);
    expect(box(ctx.draws.slice(0, i).filter(d => d.op === 'clip').at(-1)!.path)).toEqual([50, 50, 100, 100]);
  });

  test('a linear gradient brush runs across its rectangle', () => {
    const g = fills.at(-1)!.style as MockGradient;
    expect(g).toBeInstanceOf(MockGradient);
    expect(g.coords).toEqual([0, 0, 200, 0]);
    expect(g.stops).toEqual([[0, 'rgba(255,0,0,1)'], [1, 'rgba(0,0,255,1)']]);
  });

  test('saves and restores balance', () => {
    expect(ctx.depth).toBe(0);
  });
});

describe('pens and containers', () => {
  // a 2-point pen in pixel units, and one in points; the world transform scales by 3
  const pen = (id: number, unit: number, width: number) =>
    plus(0x4008, (2 << 8) | id, b => b.u32(0xdbc01002, 0, 0, unit).f32(width).u32(0xdbc01002, 0, 0xff000000));
  const line = (id: number) => plus(0x400d, id, b => b.u32(2).f32(0, 0, 10, 0));

  test('a pixel pen keeps its device width; a point pen is a physical width scaled by the world transform', () => {
    const ctx = play(plusOnly([plus(0x402e, 0, b => b.f32(3, 3)), pen(20, 2, 2), pen(21, 3, 72), line(20), line(21)]));
    const [px, pt] = ctx.of('stroke').slice(-2);
    expect(px!.lineWidth).toBeCloseTo(2 * 2);                       // 2 device pixels; output is 2 x pixels
    expect(pt!.lineWidth).toBeCloseTo(100 * 3 * 2);                 // 1 inch at 100 dpi, x 3 for the world, x 2
  });

  test('BeginContainer maps its source rectangle onto the destination; EndContainer puts the transform back', () => {
    const ctx = play(plusOnly([
      plus(0x4027, 2, b => b.f32(10, 10, 20, 20).f32(0, 0, 10, 10).u32(9)),     // pixels: (0,0,10,10) -> (10,10,20,20)
      plus(0x400a, 0x8000, b => b.u32(0xff010203, 1).f32(0, 0, 10, 10)),
      plus(0x4029, 0, b => b.u32(9)),
      plus(0x400a, 0x8000, b => b.u32(0xff040506, 1).f32(0, 0, 10, 10)),
    ]));
    const inner = ctx.of('fill').find(f => f.style === 'rgba(1,2,3,1)')!, outer = ctx.of('fill').find(f => f.style === 'rgba(4,5,6,1)')!;
    expect(box(inner.path)).toEqual([20, 20, 60, 60]);
    expect(box(outer.path)).toEqual([0, 0, 20, 20]);
  });
});

describe('choosing EMF or EMF+ records', () => {
  const dual = emf([
    plusComment([plusHeader(true), ...plusObjects(), plus(0x400a, 0x8000, b => b.u32(0xff00ff00, 1).f32(0, 0, 10, 10))]),
    E.createBrush(1, 0, 0x0000ff), E.select(1), E.rectangle(0, 0, 10, 10),
  ]);

  test('a dual picture is drawn from its EMF records', () => {
    expect(parseMetafile(dual).plus).toBe('dual');
    expect(play(dual).of('fill').map(f => f.style)).toEqual(['rgb(255,0,0)']);
  });

  test('forcing EMF+ draws the same picture from the EMF+ records; a picture without them is drawn from EMF', () => {
    expect(play(dual, { plus: 'emf+' }).of('fill').map(f => f.style)).toEqual(['rgba(0,255,0,1)']);
    expect(play(emf([E.rectangle(0, 0, 10, 10)]), { plus: 'emf+' }).of('fill')).toHaveLength(1);
  });

  test('in an EMF+ only picture, EMF records play only after EmfPlusGetDC, until the next EMF+ record', () => {
    const bytes = emf([
      plusComment([plusHeader(false), plus(0x4004)]),
      E.rectangle(0, 0, 10, 10),
      plusComment([plus(0x400a, 0x8000, b => b.u32(0xff00ff00, 1).f32(0, 0, 10, 10))]),
      E.rectangle(0, 0, 20, 20),
    ]);
    expect(play(bytes).of('fill').map(f => f.style)).toEqual(['rgb(255,255,255)', 'rgba(0,255,0,1)']);
  });
});

describe('objects', () => {
  test('an object split over records with the continue flag is joined', () => {
    // the 2x2 image object from plusObjects, cut in two
    const whole = plusObjects()[4]!;
    const data = whole.subarray(12);
    const total = data.length;
    const first = plus(0x4008, 0x8000 | (5 << 8) | 9, b => b.u32(total).bytes(data.subarray(0, 20)));
    const last = plus(0x4008, (5 << 8) | 9, b => b.bytes(data.subarray(20)));
    const bytes = emf([plusComment([plusHeader(false), first, last, plus(0x401a, 9, b => b.u32(0xffffffff, 2).f32(0, 0, 2, 2).f32(0, 0, 10, 10))])]);
    expect(play(bytes).of('drawImage')).toHaveLength(1);
  });

  test('compressed images are listed for decoding before playback, with stable keys', () => {
    const png = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4];
    const obj = plus(0x4008, (5 << 8) | 1, b => b.u32(0xdbc01002, 1).i32(0, 0, 0).u32(0, 1).bytes(png).pad(4));
    const mf = parseMetafile(emf([plusComment([plusHeader(false), obj])]));
    const found = plusImages(mf);
    expect(found).toHaveLength(1);
    expect(Array.from(found[0]!.bytes.subarray(0, 4))).toEqual([0x89, 0x50, 0x4e, 0x47]);
    expect(plusImages(mf)[0]!.key).toBe(found[0]!.key);
  });

  test('an image object holding a metafile is drawn into the destination, its unit square never snapped', () => {
    const inner = emf([E.createBrush(1, 0, 0x0000ff), E.select(1), E.rectangle(25, 25, 75, 75)]);
    const obj = plus(0x4008, (5 << 8) | 1, b => b.u32(0xdbc01002, 2, 3, inner.length).bytes(inner).pad(4));
    const draw = plus(0x401a, 1, b => b.u32(0xffffffff, 2).f32(0, 0, 100, 100).f32(10, 10, 40, 40));
    const bytes = emf([plusComment([plusHeader(false), obj, draw])]);
    for (const snap of [false, true]) {
      const fill = play(bytes, { snap }).of('fill').find(f => f.style === 'rgb(255,0,0)');
      expect(box(fill?.path)).toEqual([40, 40, 80, 80]);
    }
  });

  test('a bad object takes its slot as nothing, and drawing with it draws nothing', () => {
    const broken = plus(0x4008, (3 << 8) | 2, b => b.u32(0xdbc01002, 0xffffffff, 0));
    const bytes = emf([plusComment([plusHeader(false), broken, plus(0x4014, 0x8000 | 2, b => b.u32(0xff000000))])]);
    expect(play(bytes).of('fill')).toHaveLength(0);
  });

  test('region trees are walked without recursion, however deep', () => {
    // 2000 nested AND nodes: op, op, op, ..., rect, rect, ...
    const depth = 2000;
    const b = new Bytes().u32(0xdbc01002, depth * 2);
    for (let i = 0; i < depth; i++) b.u32(1);
    for (let i = 0; i <= depth; i++) b.u32(0x10000000).f32(0, 0, 10 + i, 10 + i);
    const region = plus(0x4008, (4 << 8) | 1, b);
    const bytes = emf([plusComment([plusHeader(false), region, plus(0x4034, (1 << 8) | 1), plus(0x400a, 0x8000, x => x.u32(0xff000000, 1).f32(0, 0, 50, 50))])]);
    const ctx = play(bytes);
    const i = ctx.draws.findIndex(d => d.op === 'fill');
    expect(box(ctx.draws.slice(0, i).filter(d => d.op === 'clip').at(-1)!.path)).toEqual([0, 0, 20, 20]);
  });
});
