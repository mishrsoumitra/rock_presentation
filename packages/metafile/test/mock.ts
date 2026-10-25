// A 2D context that records what is drawn, in device coordinates, for tests of the players.
import type { MetaCtx } from '../src/index.ts';

type M = [number, number, number, number, number, number];
const mul = (a: M, b: M): M => [
  a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
  a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
];
const at = (m: M, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

/** A path command, with points already in device coordinates. */
export type PathOp = ['M' | 'L', number, number] | ['C', number, number, number, number, number, number] | ['Z'] | ['R', number, number, number, number];

export interface Draw {
  op: 'fill' | 'stroke' | 'clip' | 'fillRect' | 'fillText' | 'drawImage';
  path?: PathOp[];
  rule?: string;
  style?: unknown;
  lineWidth?: number;
  dash?: number[];
  cap?: string;
  composite?: string;
  alpha?: number;
  smoothing?: boolean;
  text?: string;
  font?: string;
  /** fillText and fillRect: the position in user space and the transform at the time. */
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  matrix?: M;
  image?: unknown;
  src?: number[];
}

interface State {
  matrix: M;
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  lineCap: CanvasLineCap;
  lineJoin: CanvasLineJoin;
  miterLimit: number;
  globalAlpha: number;
  globalCompositeOperation: GlobalCompositeOperation;
  imageSmoothingEnabled: boolean;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  dash: number[];
  clips: number;
}

export class MockPattern {
  readonly image: unknown;
  readonly repetition: string | null;
  transform: DOMMatrix2DInit | undefined;
  constructor(image: unknown, repetition: string | null) { this.image = image; this.repetition = repetition; }
  setTransform(m?: DOMMatrix2DInit): void { this.transform = m; }
}

export class MockGradient {
  readonly kind: string;
  readonly coords: number[];
  stops: [number, string][] = [];
  constructor(kind: string, coords: number[]) { this.kind = kind; this.coords = coords; }
  addColorStop(offset: number, color: string): void { this.stops.push([offset, color]); }
}

export class RecordingContext implements MetaCtx {
  draws: Draw[] = [];
  /** Every call in order, for tests of save/restore balance and such. */
  calls: string[] = [];
  images: { width: number; height: number; data: Uint8ClampedArray }[] = [];
  private s: State = {
    matrix: [1, 0, 0, 1, 0, 0], fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1, lineCap: 'butt', lineJoin: 'miter',
    miterLimit: 10, globalAlpha: 1, globalCompositeOperation: 'source-over', imageSmoothingEnabled: true, font: '10px sans-serif',
    textAlign: 'start', textBaseline: 'alphabetic', dash: [], clips: 0,
  };
  private stack: State[] = [];
  private path: PathOp[] = [];
  readonly width: number;
  readonly height: number;

  constructor(width = 100, height = 100) { this.width = width; this.height = height; }

  get depth(): number { return this.stack.length; }
  get fillStyle() { return this.s.fillStyle; }
  set fillStyle(v) { this.s.fillStyle = v; }
  get strokeStyle() { return this.s.strokeStyle; }
  set strokeStyle(v) { this.s.strokeStyle = v; }
  get lineWidth() { return this.s.lineWidth; }
  set lineWidth(v) { this.s.lineWidth = v; }
  get lineCap() { return this.s.lineCap; }
  set lineCap(v) { this.s.lineCap = v; }
  get lineJoin() { return this.s.lineJoin; }
  set lineJoin(v) { this.s.lineJoin = v; }
  get miterLimit() { return this.s.miterLimit; }
  set miterLimit(v) { this.s.miterLimit = v; }
  get globalAlpha() { return this.s.globalAlpha; }
  set globalAlpha(v) { this.s.globalAlpha = v; }
  get globalCompositeOperation() { return this.s.globalCompositeOperation; }
  set globalCompositeOperation(v) { this.s.globalCompositeOperation = v; }
  get imageSmoothingEnabled() { return this.s.imageSmoothingEnabled; }
  set imageSmoothingEnabled(v) { this.s.imageSmoothingEnabled = v; }
  get font() { return this.s.font; }
  set font(v) { this.s.font = v; }
  get textAlign() { return this.s.textAlign; }
  set textAlign(v) { this.s.textAlign = v; }
  get textBaseline() { return this.s.textBaseline; }
  set textBaseline(v) { this.s.textBaseline = v; }

  save(): void { this.calls.push('save'); this.stack.push({ ...this.s, dash: [...this.s.dash] }); }
  restore(): void {
    this.calls.push('restore');
    const s = this.stack.pop();
    if (!s) throw new Error('restore without save');
    this.s = s;
  }
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void { this.s.matrix = mul(this.s.matrix, [a, b, c, d, e, f]); }
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void { this.s.matrix = [a, b, c, d, e, f]; }
  get matrix(): M { return this.s.matrix; }

  beginPath(): void { this.path = []; }
  moveTo(x: number, y: number): void { this.path.push(['M', ...at(this.s.matrix, x, y)]); }
  lineTo(x: number, y: number): void { this.path.push(['L', ...at(this.s.matrix, x, y)]); }
  bezierCurveTo(a: number, b: number, c: number, d: number, x: number, y: number): void {
    this.path.push(['C', ...at(this.s.matrix, a, b), ...at(this.s.matrix, c, d), ...at(this.s.matrix, x, y)]);
  }
  closePath(): void { this.path.push(['Z']); }
  rect(x: number, y: number, w: number, h: number): void {
    const [x0, y0] = at(this.s.matrix, x, y);
    const [x1, y1] = at(this.s.matrix, x + w, y + h);
    this.path.push(['R', x0, y0, x1, y1]);
  }

  private record(op: Draw['op'], extra: Partial<Draw>): void {
    this.calls.push(op);
    this.draws.push({ op, composite: this.s.globalCompositeOperation, alpha: this.s.globalAlpha, ...extra });
  }

  fill(rule?: CanvasFillRule): void { this.record('fill', { path: [...this.path], rule: rule ?? 'nonzero', style: this.s.fillStyle }); }
  stroke(): void {
    this.record('stroke', { path: [...this.path], style: this.s.strokeStyle, lineWidth: this.s.lineWidth, dash: [...this.s.dash], cap: this.s.lineCap });
  }
  clip(rule?: CanvasFillRule): void { this.s.clips++; this.record('clip', { path: [...this.path], rule: rule ?? 'nonzero' }); }
  fillRect(x: number, y: number, w: number, h: number): void { this.record('fillRect', { x, y, w, h, matrix: this.s.matrix, style: this.s.fillStyle }); }
  fillText(text: string, x: number, y: number): void {
    this.record('fillText', { text, x, y, matrix: this.s.matrix, font: this.s.font, style: this.s.fillStyle });
  }
  /** Width half the font size per character, ascent 0.8 and descent 0.2 of it: easy numbers for layout tests. */
  measureText(text: string): { width: number; fontBoundingBoxAscent: number; fontBoundingBoxDescent: number } {
    const size = Number(this.s.font.split(' ').find(t => t.endsWith('px'))?.slice(0, -2) || 10);
    return { width: Array.from(text).length * size * 0.5, fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2 };
  }
  drawImage(image: CanvasImageSource, sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number): void {
    this.record('drawImage', { image, src: [sx, sy, sw, sh], x: dx, y: dy, w: dw, h: dh, matrix: this.s.matrix, smoothing: this.s.imageSmoothingEnabled });
  }
  createPattern(image: CanvasImageSource, repetition: string | null): CanvasPattern | null {
    return new MockPattern(image, repetition) as unknown as CanvasPattern;
  }
  createLinearGradient(x0: number, y0: number, x1: number, y1: number): CanvasGradient {
    return new MockGradient('linear', [x0, y0, x1, y1]) as unknown as CanvasGradient;
  }
  createRadialGradient(x0: number, y0: number, r0: number, x1: number, y1: number, r1: number): CanvasGradient {
    return new MockGradient('radial', [x0, y0, r0, x1, y1, r1]) as unknown as CanvasGradient;
  }
  setLineDash(segments: number[]): void { this.s.dash = [...segments]; }
  createImageData(w: number, h: number): ImageData {
    return { width: w, height: h, data: new Uint8ClampedArray(w * h * 4), colorSpace: 'srgb' } as ImageData;
  }
  putImageData(data: ImageData): void { this.images.push({ width: data.width, height: data.height, data: data.data }); }

  /** The draws of one kind. */
  of(op: Draw['op']): Draw[] { return this.draws.filter(d => d.op === op); }
}

/** A canvas whose 2D context records; what createCanvas makes in tests. */
export class MockCanvas {
  readonly width: number;
  readonly height: number;
  readonly ctx: RecordingContext;
  constructor(width: number, height: number) {
    this.width = width;
    this.height = height;
    this.ctx = new RecordingContext(width, height);
  }
  getContext(_id: '2d'): RecordingContext { return this.ctx; }
}

/** The points of a recorded path (control points left out). */
export function points(path: PathOp[] | undefined): [number, number][] {
  const out: [number, number][] = [];
  for (const p of path ?? []) {
    if (p[0] === 'M' || p[0] === 'L') out.push([p[1], p[2]]);
    else if (p[0] === 'C') out.push([p[5], p[6]]);
    else if (p[0] === 'R') out.push([p[1], p[2]], [p[3], p[4]]);
  }
  return out;
}

/** The bounding box of a recorded path, rounded to 3 decimals. */
export function box(path: PathOp[] | undefined): [number, number, number, number] {
  const pts = points(path);
  const r = (v: number) => Math.round(v * 1000) / 1000 || 0;
  return [r(Math.min(...pts.map(p => p[0]))), r(Math.min(...pts.map(p => p[1]))), r(Math.max(...pts.map(p => p[0]))), r(Math.max(...pts.map(p => p[1])))];
}
