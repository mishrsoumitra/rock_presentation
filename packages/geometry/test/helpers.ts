// Shared by the geometry tests.
import { parseXml } from '@pptx/xml';
import { expect } from 'vitest';
import { parseCustomGeometry, type EvaluatedGeometry, type GeometryPath, type Guide, type ShapeGeometry } from '../src/index.ts';

export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** A custom geometry from the children of an a:custGeom element. */
export const custGeom = (inner: string): ShapeGeometry =>
  parseCustomGeometry(parseXml(`<a:custGeom xmlns:a="${A}">${inner}</a:custGeom>`).root);

/** A geometry with only guides, for testing formulas. */
export const guidesOnly = (gd: [string, string][], av: [string, string][] = []): ShapeGeometry => ({
  avLst: av.map(([name, fmla]) => ({ name, fmla })),
  gdLst: gd.map(([name, fmla]) => ({ name, fmla })),
  ahLst: [], cxnLst: [], rect: null, pathLst: [],
});

export const adj = (values: Record<string, number>): Guide[] => Object.entries(values).map(([name, v]) => ({ name, fmla: `val ${v}` }));

/** End points of a path's commands (M, L, Q, C), in order. */
export const ends = (path: GeometryPath): [number, number][] =>
  path.commands.flatMap(c => (c.op === 'Z' ? [] : [[c.x, c.y] as [number, number]]));

/** Rounded to 6 decimals, with -0 as 0, for comparing computed values. */
export const fixed = (v: number) => +v.toFixed(6) || 0;
export const fixedRect = (r: { x: number; y: number; w: number; h: number }) => [r.x, r.y, r.w, r.h].map(fixed);

/** Assert two points are equal to within `tol` (absolute, EMU). */
export function near(actual: readonly [number, number] | { x: number; y: number }, expected: readonly [number, number], tol = 1e-6): void {
  const [x, y] = 'x' in actual ? [actual.x, actual.y] : actual;
  expect(Math.abs(x - expected[0]), `x ${x} vs ${expected[0]}`).toBeLessThanOrEqual(tol);
  expect(Math.abs(y - expected[1]), `y ${y} vs ${expected[1]}`).toBeLessThanOrEqual(tol);
}

/** The point at t on a cubic Bezier. */
export function cubicAt(p0: [number, number], c: { x1: number; y1: number; x2: number; y2: number; x: number; y: number }, t: number): [number, number] {
  const u = 1 - t;
  const f = (a: number, b: number, cc: number, d: number) => u * u * u * a + 3 * u * u * t * b + 3 * u * t * t * cc + t * t * t * d;
  return [f(p0[0], c.x1, c.x2, c.x), f(p0[1], c.y1, c.y2, c.y)];
}

/** Every number in an evaluated geometry. */
export function numbersOf(e: EvaluatedGeometry): number[] {
  const out = [...e.guides.values(), e.textRect.x, e.textRect.y, e.textRect.w, e.textRect.h];
  for (const p of e.paths) for (const c of p.commands) for (const v of Object.values(c)) if (typeof v === 'number') out.push(v);
  for (const h of e.handles) out.push(h.x, h.y);
  for (const s of e.connectionSites) out.push(s.x, s.y, s.angle);
  return out;
}

/** The value of a limit or coordinate attribute: a literal, or a guide's value. */
export function valueOf(e: EvaluatedGeometry, token: string): number {
  if (/^[+-]?[0-9.]+$/.test(token)) return Number(token);
  const v = e.guides.get(token);
  if (v === undefined) throw new Error(`no guide ${token}`);
  return v;
}
