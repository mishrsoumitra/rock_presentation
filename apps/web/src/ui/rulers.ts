// apps/web/src/ui/rulers.ts  -- what a ruler shows: where its marks fall for an origin and a zoom

export interface Tick {
  /** Pixels along the ruler. */
  at: number;
  size: 'inch' | 'half' | 'eighth';
  /** Inches from the slide's edge, on the marks that are numbered. */
  label?: string;
}

/**
 * The marks of a ruler `length` pixels long, when the slide's edge (zero) is `origin` pixels along it and an inch is
 * `pxPerInch` pixels. Every inch is marked; halves and eighths appear as the zoom makes room for them, and inches
 * are numbered as closely as their numbers fit (every inch, every other, every fifth...). Marks left of or above
 * the slide count away from it, without a sign, as on PowerPoint's ruler.
 */
export function rulerTicks(origin: number, pxPerInch: number, length: number): Tick[] {
  if (!(pxPerInch > 0) || !(length > 0)) return [];
  const step = pxPerInch >= 80 ? 1 / 8 : pxPerInch >= 20 ? 1 / 2 : 1;
  const every = [1, 2, 5, 10, 20, 50, 100].find(n => n * pxPerInch >= 28) ?? 100;
  const first = Math.ceil((0 - origin) / (pxPerInch * step)), last = Math.floor((length - origin) / (pxPerInch * step));
  const ticks: Tick[] = [];
  for (let i = first; i <= last; i++) {
    const inches = i * step;
    const whole = Number.isInteger(inches);
    ticks.push({
      at: origin + inches * pxPerInch,
      size: whole ? 'inch' : Number.isInteger(inches * 2) ? 'half' : 'eighth',
      ...(whole && inches % every === 0 ? { label: String(Math.abs(inches)) } : {}),
    });
  }
  return ticks;
}

/** Where a new guide goes: the middle of the slide, or the first half-inch step from it that no guide is on yet. */
export function newGuideAt(existing: number[], size: number): number {
  const step = 457200;
  for (let n = 0; n < 200; n++) {
    for (const at of n ? [size / 2 + n * step, size / 2 - n * step] : [size / 2]) {
      if (at >= 0 && at <= size && !existing.some(g => Math.abs(g - at) < 1)) return Math.round(at);
    }
  }
  return Math.round(size / 2);
}
