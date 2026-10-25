// Where shapes' Konva nodes sit, and back from a moved, resized or rotated node to an a:xfrm.
import { describe, expect, test } from 'vitest';
import type { Xfrm } from '@pptx/model';
import type { RenderShape } from '@pptx/render';
import { boxMatrix, childSpaceMatrix, multiply, type Matrix } from '@pptx/resolve';
import { drawBase, parentOf, placementOf, restMatrix, xfrmFor } from '../src/ui/placement.ts';

const X = (x: Partial<Xfrm>): Xfrm => ({ x: 100, y: 200, cx: 3000, cy: 1000, rot: 0, flipH: false, flipV: false, ...x });

/** A render shape as the renderer makes it: its transform is the parent's matrix times its own box matrix. */
function shape(xfrm: Xfrm, parent: Matrix = [1, 0, 0, 1, 0, 0]): RenderShape {
  return { transform: multiply(parent, boxMatrix(xfrm)), box: { w: xfrm.cx, h: xfrm.cy }, source: { xfrm } } as unknown as RenderShape;
}

const close = (a: number, b: number, eps = 1e-6) => expect(Math.abs(a - b)).toBeLessThan(eps);

describe('placementOf', () => {
  test('a rotated shape: its box centre, its size, its rotation in degrees', () => {
    const p = placementOf(shape(X({ rot: 30 * 60000 })));
    expect(p).toMatchObject({ w: 3000, h: 1000 });
    close(p.x, 1600);
    close(p.y, 700);
    close(p.rotation, 30);
  });

  test('a shape flipped both ways keeps its own rotation (not turned half round)', () => {
    close(placementOf(shape(X({ rot: 30 * 60000, flipH: true, flipV: true }))).rotation, 30);
    close(placementOf(shape(X({ flipH: true }))).rotation, 0);
  });

  test('a group member: through the group child space, with the group rotation added', () => {
    const group = X({ x: 1000, y: 1000, cx: 2000, cy: 1000, rot: 90 * 60000 });
    const parent = multiply(boxMatrix(group), childSpaceMatrix(group, { chOff: { x: 0, y: 0 }, chExt: { cx: 4000, cy: 2000 } }));
    const p = placementOf(shape(X({ x: 0, y: 0, cx: 4000, cy: 2000 }), parent));
    expect(p.w).toBeCloseTo(2000, 6);
    expect(p.h).toBeCloseTo(1000, 6);
    close(p.rotation, 90);
    close(p.x, 2000);
    close(p.y, 1500);
  });
});

describe('xfrmFor', () => {
  test('a placement read back gives the xfrm it came from, at the top and inside a scaled, rotated group', () => {
    const group = X({ x: 1000, y: 1000, cx: 2000, cy: 1000, rot: 45 * 60000 });
    const parent = multiply(boxMatrix(group), childSpaceMatrix(group, { chOff: { x: 10, y: 20 }, chExt: { cx: 4000, cy: 2000 } }));
    for (const [x, m] of [[X({ rot: 15 * 60000 }), undefined], [X({ x: 50, y: 60, rot: 10 * 60000 }), parent]] as const) {
      const s = shape(x, m);
      const back = xfrmFor(placementOf(s), parentOf(s), x);
      for (const k of ['x', 'y', 'cx', 'cy', 'rot'] as const) expect(back[k]).toBeCloseTo(x[k], 3);
    }
  });

  test('a node resized and turned gives the box it shows', () => {
    const s = shape(X({}));
    const back = xfrmFor({ x: 2000, y: 1000, w: 600, h: 400, rotation: -90 }, parentOf(s), X({}));
    expect(back).toMatchObject({ x: 1700, y: 800, cx: 600, cy: 400, rot: 270 * 60000 });
  });
});

test('drawBase undoes the node at rest: drawing in slide coordinates lands where the stage puts the slide', () => {
  const p = placementOf(shape(X({ rot: 20 * 60000 })));
  const stage: Matrix = [0.5, 0, 0, 0.5, 10, 20];
  const ctm = multiply(stage, restMatrix(p));
  const k = 2;
  const base = drawBase(new DOMMatrixStub(ctm) as unknown as DOMMatrix, p, k);
  const device = multiply(base, [k, 0, 0, k, 0, 0]);
  for (let i = 0; i < 6; i++) close(device[i]!, stage[i]!);
});

class DOMMatrixStub {
  a: number; b: number; c: number; d: number; e: number; f: number;
  constructor(m: Matrix) { [this.a, this.b, this.c, this.d, this.e, this.f] = m; }
}
