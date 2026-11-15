// Cropping with handles: whatever handle is dragged and however the picture is turned or flipped, the picture
// itself must not move; only the frame's edge moves over it, within the picture.
import type { Xfrm } from '@pptx/model';
import { describe, expect, test } from 'vitest';
import { CROP_HANDLES, dragCrop, fullCorners, handlePoint, type Cropped } from '../src/ui/crop.ts';

const frame = (over: Partial<Xfrm> = {}): Xfrm => ({ x: 1000, y: 2000, cx: 4000, cy: 2000, rot: 0, flipH: false, flipV: false, ...over });
const whole = (over: Partial<Xfrm> = {}): Cropped => ({ xfrm: frame(over), crop: { l: 0, t: 0, r: 0, b: 0 } });
const close = (a: number[], b: number[]) => a.forEach((v, i) => expect(v, `at ${i}`).toBeCloseTo(b[i]!, 6));
const R = { x: 1, y: 0 } as const, L = { x: -1, y: 0 } as const, T = { x: 0, y: -1 } as const, B = { x: 0, y: 1 } as const;

describe('dragging one edge of an upright picture', () => {
  test('the right edge in: the frame gets narrower from the right, a quarter of the picture is cut off there', () => {
    const c = dragCrop(whole(), R, -1000, 0);
    expect(c.xfrm).toMatchObject({ x: 1000, y: 2000, cx: 3000, cy: 2000 });
    expect(c.crop).toEqual({ l: 0, t: 0, r: 0.25, b: 0 });
  });

  test('the left edge in: the frame starts further right, and the picture under it has not moved', () => {
    const c = dragCrop(whole(), L, 1000, 0);
    expect(c.xfrm).toMatchObject({ x: 2000, cx: 3000 });
    expect(c.crop).toEqual({ l: 0.25, t: 0, r: 0, b: 0 });
    close(fullCorners(c), fullCorners(whole()));
  });

  test('top and bottom the same way; movement across an edge does nothing to it', () => {
    expect(dragCrop(whole(), T, 500, 500).crop).toEqual({ l: 0, t: 0.25, r: 0, b: 0 });
    expect(dragCrop(whole(), B, 500, -500).xfrm).toMatchObject({ x: 1000, y: 2000, cx: 4000, cy: 1500 });
  });

  test('an edge stops at the edge of the picture going out, and leaves a sliver showing going in', () => {
    expect(dragCrop(whole(), R, 5000, 0)).toEqual(whole());                    // nothing beyond the picture to show
    const all = dragCrop(whole(), R, -99999, 0);
    expect(all.xfrm.cx).toBeCloseTo(80);                                        // 2% of 4000
    expect(all.crop.r).toBeCloseTo(0.98);
    const cropped = dragCrop(whole(), L, 1000, 0);
    expect(dragCrop(cropped, L, -5000, 0).crop.l).toBe(0);                      // back out to the picture's edge, no further
  });
});

describe('corners, and cropping again', () => {
  test('a corner moves two edges at once', () => {
    const c = dragCrop(whole(), { x: 1, y: 1 }, -2000, -500);
    expect(c.xfrm).toMatchObject({ x: 1000, y: 2000, cx: 2000, cy: 1500 });
    expect(c.crop).toEqual({ l: 0, t: 0, r: 0.5, b: 0.25 });
  });

  test('a second crop is measured on the whole picture, not on what was showing', () => {
    const once = dragCrop(whole(), R, -2000, 0);                                // half showing: 2000 of 4000
    const twice = dragCrop(once, L, 1000, 0);                                   // another quarter of the whole from the left
    expect(twice.crop).toEqual({ l: 0.25, t: 0, r: 0.5, b: 0 });
    expect(twice.xfrm).toMatchObject({ x: 2000, cx: 1000 });
    close(fullCorners(twice), fullCorners(whole()));
  });
});

describe('the picture never moves', () => {
  const cases: [string, Partial<Xfrm>][] = [
    ['upright', {}], ['turned 90°', { rot: 5400000 }], ['turned 30°', { rot: 1800000 }], ['turned 180°', { rot: 10800000 }],
    ['flipped left to right', { flipH: true }], ['flipped top to bottom', { flipV: true }], ['turned 45° and flipped both ways', { rot: 2700000, flipH: true, flipV: true }],
  ];
  for (const [name, over] of cases) {
    test(`${name}: every handle, dragged every way, leaves the whole picture where it was`, () => {
      const start: Cropped = { xfrm: frame(over), crop: { l: 0.1, t: 0.05, r: 0.2, b: 0.15 } };
      const before = fullCorners(start);
      for (const h of CROP_HANDLES) {
        for (const [dx, dy] of [[300, 0], [-300, 0], [0, 200], [0, -200], [250, -150], [-9000, 9000]] as const) {
          const after = dragCrop(start, h, dx, dy);
          close(fullCorners(after), before);
          for (const k of ['l', 't', 'r', 'b'] as const) {
            expect(after.crop[k]).toBeGreaterThanOrEqual(0);
            expect(after.crop[k]).toBeLessThan(1);
          }
          expect(after.xfrm.cx).toBeGreaterThan(0);
          expect(after.xfrm.cy).toBeGreaterThan(0);
        }
      }
    });
  }

  test('on a picture turned 90° clockwise, dragging its right-hand handle (which is at the bottom on screen) up crops it', () => {
    const start = whole({ rot: 5400000 });
    const at = handlePoint(start.xfrm, R);
    expect(at.x).toBeCloseTo(3000);                                             // under the centre
    expect(at.y).toBeCloseTo(5000);
    const c = dragCrop(start, R, 0, -1000);                                     // up the screen is "in" along the picture's width
    expect(c.crop.r).toBeCloseTo(0.25);
    expect(c.xfrm.cx).toBeCloseTo(3000);
  });

  test('on a picture flipped left to right, the handle on the left of the screen crops the picture\'s right side', () => {
    const start = whole({ flipH: true });
    expect(handlePoint(start.xfrm, R).x).toBeCloseTo(1000);                     // the picture's right edge is on the left
    const c = dragCrop(start, R, 1000, 0);                                      // dragged right, into the picture
    expect(c.crop).toEqual({ l: 0, t: 0, r: 0.25, b: 0 });
    expect(c.xfrm).toMatchObject({ x: 2000, cx: 3000 });
  });
});
