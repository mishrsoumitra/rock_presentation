// Just enough of Path2D, DOMMatrix and ImageData for drawing in Node, and a context that records what is drawn.
type M6 = [number, number, number, number, number, number];

class StubMatrix {
  a: number; b: number; c: number; d: number; e: number; f: number;
  constructor(init: number[] = [1, 0, 0, 1, 0, 0]) {
    [this.a, this.b, this.c, this.d, this.e, this.f] = init as M6;
  }
  multiply(o: StubMatrix): StubMatrix {
    return new StubMatrix([
      this.a * o.a + this.c * o.b, this.b * o.a + this.d * o.b, this.a * o.c + this.c * o.d, this.b * o.c + this.d * o.d,
      this.a * o.e + this.c * o.f + this.e, this.b * o.e + this.d * o.f + this.f,
    ]);
  }
  inverse(): StubMatrix {
    const det = this.a * this.d - this.b * this.c || 1;
    return new StubMatrix([this.d / det, -this.b / det, -this.c / det, this.a / det, (this.c * this.f - this.d * this.e) / det, (this.b * this.e - this.a * this.f) / det]);
  }
}

export class StubPath {
  readonly ops: (string | number)[][] = [];
  moveTo(x: number, y: number) { this.ops.push(['M', x, y]); }
  lineTo(x: number, y: number) { this.ops.push(['L', x, y]); }
  quadraticCurveTo(...a: number[]) { this.ops.push(['Q', ...a]); }
  bezierCurveTo(...a: number[]) { this.ops.push(['C', ...a]); }
  arc(...a: number[]) { this.ops.push(['A', ...a]); }
  ellipse(...a: number[]) { this.ops.push(['E', ...a]); }
  rect(...a: number[]) { this.ops.push(['R', ...a]); }
  closePath() { this.ops.push(['Z']); }
  addPath(p: StubPath) { this.ops.push(...p.ops); }
}

class StubImageData {
  data: Uint8ClampedArray; width: number; height: number;
  constructor(data: Uint8ClampedArray | number, w: number, h?: number) {
    if (typeof data === 'number') { this.width = data; this.height = w; this.data = new Uint8ClampedArray(data * w * 4); }
    else { this.data = data; this.width = w; this.height = h ?? data.length / 4 / w; }
  }
}

export function installStubs(): void {
  const g = globalThis as Record<string, unknown>;
  g.Path2D ??= StubPath;
  g.DOMMatrix ??= StubMatrix;
  g.ImageData ??= StubImageData;
}

export interface Call { op: string; args: unknown[]; transform: M6; alpha: number; fillStyle: unknown; }

/** A 2D context that records every drawing call with the transform and alpha it was made under. */
export function recorder(): { ctx: Record<string, unknown>; calls: Call[] } {
  const calls: Call[] = [];
  let t: M6 = [1, 0, 0, 1, 0, 0];
  const stack: { t: M6; alpha: number; fill: unknown }[] = [];
  const ctx: Record<string, unknown> = {
    globalAlpha: 1, fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, lineCap: 'butt', lineJoin: 'miter', miterLimit: 10,
    globalCompositeOperation: 'source-over', shadowBlur: 0, shadowColor: '', shadowOffsetX: 0, shadowOffsetY: 0, font: '', textAlign: 'start', textBaseline: 'alphabetic',
    save() { stack.push({ t, alpha: ctx.globalAlpha as number, fill: ctx.fillStyle }); },
    restore() { const s = stack.pop(); if (s) { t = s.t; ctx.globalAlpha = s.alpha; ctx.fillStyle = s.fill; } },
    setTransform(a: number, b: number, c: number, d: number, e: number, f: number) { t = [a, b, c, d, e, f]; },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      t = [t[0] * a + t[2] * b, t[1] * a + t[3] * b, t[0] * c + t[2] * d, t[1] * c + t[3] * d, t[0] * e + t[2] * f + t[4], t[1] * e + t[3] * f + t[5]];
    },
    getTransform() { return new StubMatrix(t); },
    setLineDash() {},
    createLinearGradient: () => ({ addColorStop() {} }),
    createRadialGradient: () => ({ addColorStop() {} }),
    createPattern: () => ({ setTransform() {} }),
    getImageData: (_x: number, _y: number, w: number, h: number) => new StubImageData(w, h),
    putImageData() {},
  };
  for (const op of ['fill', 'stroke', 'clip', 'fillRect', 'strokeRect', 'fillText', 'drawImage', 'beginPath', 'rect']) {
    ctx[op] = (...args: unknown[]) => calls.push({ op, args, transform: t, alpha: ctx.globalAlpha as number, fillStyle: ctx.fillStyle });
  }
  return { ctx, calls };
}
