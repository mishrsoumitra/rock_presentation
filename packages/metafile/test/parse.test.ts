// Headers, frames and record parsing; metafileImage as the renderer uses it.
import { describe, expect, test } from 'vitest';
import { canvasSize, decodeImages, hasDrawing, metafileImage, MetafileError, parseMetafile } from '../src/index.ts';
import { Bytes, E, emf, emr, plus, plusComment, plusHeader, W, wmf, wmr } from './build.ts';
import { MockCanvas } from './mock.ts';
import { wmfWithEmf } from './samples.ts';

describe('telling formats apart', () => {
  test('by content, not by name: EMF, placeable WMF, plain WMF; anything else is unsupported', () => {
    expect(parseMetafile(emf([])).format).toBe('emf');
    expect(parseMetafile(wmf([W.rectangle(0, 0, 1, 1)], { l: 0, t: 0, r: 10, b: 10, inch: 96 })).format).toBe('wmf');
    expect(parseMetafile(wmf([W.rectangle(0, 0, 1, 1)])).format).toBe('wmf');
    for (const junk of [new Uint8Array(0), Uint8Array.from([0x89, 0x50, 0x4e, 0x47]), new Uint8Array(200)]) {
      expect(() => parseMetafile(junk)).toThrow(expect.objectContaining({ code: 'unsupported' }));
    }
  });

  test('a WMF whose escapes carry a complete EMF is read as that EMF', () => {
    const mf = parseMetafile(wmfWithEmf());
    expect(mf.format).toBe('emf');
    expect(mf.records.some(r => r.type === 'box')).toBe(true);
  });

  test('an incomplete embedded EMF leaves the WMF as it is', () => {
    const mf = parseMetafile(wmfWithEmf(false));
    expect(mf.format).toBe('wmf');
    expect(mf.records.some(r => r.type === 'box')).toBe(true);
  });
});

describe('EMF headers', () => {
  test('the frame (0.01 mm) becomes device units through the reference device resolution', () => {
    const mf = parseMetafile(emf([], { frame: [0, 0, 5080, 2540], device: [1920, 1080], mm: [480, 270] }));
    // 4 pixels per mm: 50.8 mm is 203.2 pixels
    expect(mf.frame.r).toBeCloseTo(203.2);
    expect(mf.frame.b).toBeCloseTo(101.6);
    expect(mf.width).toBeCloseTo(192);
    expect(mf.height).toBeCloseTo(96);
  });

  test('an empty frame falls back to the bounds', () => {
    const mf = parseMetafile(emf([], { frame: [0, 0, 0, 0], bounds: [10, 10, 59, 29] }));
    expect(mf.frame).toEqual({ l: 10, t: 10, r: 60, b: 30 });
  });

  test('a nonsense device size falls back to 96 dpi', () => {
    const mf = parseMetafile(emf([], { device: [0, 0], mm: [0, 0], frame: [0, 0, 2540, 2540] }));
    expect(mf.dpmmX).toBeCloseTo(96 / 25.4);
  });

  test('a broken header is malformed', () => {
    const bytes = emf([]);
    const bad = bytes.slice();
    new DataView(bad.buffer).setUint32(4, 0xffff, true);
    expect(() => parseMetafile(bad)).toThrow(expect.objectContaining({ code: 'malformed' }));
  });
});

describe('records', () => {
  test('a record that does not fit ends the file; what came before is kept', () => {
    const bytes = emf([E.rectangle(0, 0, 1, 1), E.rectangle(0, 0, 2, 2)]);
    const cut = bytes.slice(0, bytes.length - 20 - 10);
    expect(parseMetafile(cut).records.filter(r => r.type === 'box')).toHaveLength(1);
  });

  test('an inconsistent record is skipped and counted; an object it should create still takes its slot', () => {
    const badPoly = emr(3, b => b.i32(0, 0, 0, 0).u32(1000).i32(1, 2));         // 1000 points, 1 there
    const badPen = emr(95, b => b.u32(7, 0));                                    // EXTCREATEPEN cut short
    const mf = parseMetafile(emf([badPoly, badPen, E.rectangle(0, 0, 1, 1)]));
    expect(mf.skipped).toBe(2);
    expect(mf.records.find(r => r.type === 'create')).toEqual({ type: 'create', index: 7, object: { kind: 'other' } });
  });

  test('WMF objects that cannot be read still take a slot, so later indexes stay right', () => {
    const shortFont = wmr(0x02fb, b => b.i16(-10));
    const mf = parseMetafile(wmf([shortFont, W.createBrush(0, 0xff)]));
    expect(mf.records.filter(r => r.type === 'create').map(r => r.type === 'create' && r.object.kind)).toEqual(['other', 'brush']);
  });

  test('EXTTEXTOUTW: string, advances, options and rectangle', () => {
    const mf = parseMetafile(emf([E.extTextOutW(5, 6, 'Hi', { dx: [7, 8], options: 4, rect: [1, 2, 3, 4], mode: 2 })]));
    expect(mf.records.find(r => r.type === 'text')).toEqual({
      type: 'text', x: 5, y: 6, options: 4, text: 'Hi', dx: Int32Array.from([7, 8]), box: { l: 1, t: 2, r: 3, b: 4 }, advanced: true,
    });
  });

  test('WMF rectangles are stored bottom, right, top, left', () => {
    const mf = parseMetafile(wmf([W.rectangle(1, 2, 3, 4)]));
    expect(mf.records.find(r => r.type === 'box')).toEqual({ type: 'box', shape: 'rectangle', box: { l: 1, t: 2, r: 3, b: 4 } });
  });

  test('a WMF with nothing set takes its frame from what it draws', () => {
    const mf = parseMetafile(wmf([W.rectangle(10, 20, 110, 70)]));
    expect(mf.frame).toEqual({ l: 10, t: 20, r: 110, b: 70 });
  });
});

describe('metafileImage', () => {
  const make = (w: number, h: number) => new MockCanvas(w, h);

  test('a canvas sized from the frame: at least 1024 on the longer side, the aspect kept', async () => {
    const c = await metafileImage(emf([E.rectangle(0, 0, 50, 50)], { frame: [0, 0, 5080, 2540] }), make);
    expect([c?.width, c?.height]).toEqual([1024, 512]);
    expect(c!.ctx.of('fill')).toHaveLength(1);
  });

  test('a picture with nothing in it is a 1x1 transparent canvas', async () => {
    const c = await metafileImage(emf([]), make);
    expect([c?.width, c?.height]).toEqual([1, 1]);
    expect(c!.ctx.draws).toHaveLength(0);
  });

  test('undefined for what is not a picture (no host decoder for rasters in Node)', async () => {
    expect(await metafileImage(Uint8Array.from([1, 2, 3]), make)).toBeUndefined();
    expect(await metafileImage(Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0]), make)).toBeUndefined();
  });

  test('undefined, not an exception, when the picture asks for too much', async () => {
    const many = Array.from({ length: 2000 }, () => E.saveDC());
    const bytes = emf([...many, ...many, ...many, E.rectangle(0, 0, 1, 1)]);
    expect(await metafileImage(bytes, make)).toBeUndefined();
  });

  test('canvasSize caps the pixel count and keeps extreme aspects in range', () => {
    expect(canvasSize({ width: 96, height: 48, frame: { l: 0, t: 0, r: 1, b: 1 } })).toEqual({ width: 1024, height: 512 });
    const big = canvasSize({ width: 10_000, height: 10_000, frame: { l: 0, t: 0, r: 1, b: 1 } });
    expect(big.width * big.height).toBeLessThanOrEqual(16 * 1024 * 1024);
    const thin = canvasSize({ width: 1, height: 100_000, frame: { l: 0, t: 0, r: 1, b: 1 } });
    expect(thin.height).toBeLessThanOrEqual(16_384);
    expect(thin.width).toBeGreaterThanOrEqual(1);
    expect(canvasSize({ width: NaN, height: 0, frame: { l: 0, t: 0, r: 0, b: 0 } })).toEqual({ width: 1024, height: 1024 });
  });

  test('hasDrawing tells header-only pictures apart', () => {
    expect(hasDrawing(parseMetafile(emf([E.saveDC(), E.restoreDC(-1)])))).toBe(false);
    expect(hasDrawing(parseMetafile(emf([E.lineTo(1, 1)])))).toBe(true);
  });

  test('decodeImages decodes the EMF+ images only when the EMF+ records are the ones played', async () => {
    const decode = async (bytes: Uint8Array) => ({ width: bytes.length, height: 1 }) as unknown as CanvasImageSource;
    const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const image = plus(0x4008, (5 << 8) | 1, b => b.u32(0xdbc01002, 1).i32(0, 0, 0).u32(0, 1).bytes(png));
    const dual = parseMetafile(emf([plusComment([plusHeader(true), image])]));
    const only = parseMetafile(emf([plusComment([plusHeader(false), image])]));
    expect((await decodeImages(dual, 'auto', decode)).size).toBe(0);
    expect((await decodeImages(dual, 'emf+', decode)).size).toBe(1);
    expect((await decodeImages(only, 'auto', decode)).size).toBe(1);
    expect((await decodeImages(only, 'emf', decode)).size).toBe(0);
  });

  test('the error type is exported for callers that parse themselves', () => {
    expect(new MetafileError('limit', 'x').code).toBe('limit');
    expect(new Bytes().u8(1).length).toBe(1);
  });
});
