// packages/anim/src/path.ts  -- motion paths of p:animMotion ("M 0 0 L 0.25 0 E"), sampled by distance along the path
//
// Coordinates are fractions of the slide (x of its width, y of its height), offsets from where the shape is. The
// commands are M, L, C, Z and E, lower case for relative coordinates (ECMA-376 Part 1, 19.5.4). A shape moves at a
// steady speed along the path, as PowerPoint plays it, so curves are flattened and measured.

export interface MotionPath {
  /** The path flattened to points, with the distance from the start to each. */
  points: { x: number; y: number }[];
  lengths: number[];
  length: number;
}

const MAX_NUMBERS = 20_000;
const CURVE_STEPS = 16;

/** Parse a path; undefined when it holds no point. Unknown commands end the path where they start. */
export function parseMotionPath(src: string): MotionPath | undefined {
  const tokens = src.slice(0, 200_000).match(/[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?|[A-Za-z]/g) ?? [];
  // a command is a lone letter; "5E-6" is a number
  const isCommand = (t: string) => t.length === 1 && /[A-Za-z]/.test(t);
  const points: { x: number; y: number }[] = [];
  let cur = { x: 0, y: 0 };
  let start = cur;
  let i = 0;
  let numbers = 0;
  let cmd = '';
  const nums = (n: number): number[] | undefined => {
    const out: number[] = [];
    for (let k = 0; k < n; k++) {
      const t = tokens[i];
      if (t === undefined || isCommand(t) || ++numbers > MAX_NUMBERS) return undefined;
      out.push(Number(t));
      i++;
    }
    return out.every(Number.isFinite) ? out : undefined;
  };
  const add = (p: { x: number; y: number }) => {
    points.push(p);
    cur = p;
  };
  while (i < tokens.length) {
    const t = tokens[i]!;
    if (isCommand(t)) {
      cmd = t;
      i++;
    } else if (!cmd || cmd === 'Z' || cmd === 'z' || cmd === 'E' || cmd === 'e') {
      break;                                            // a number where no command takes one
    }
    const rel = cmd === cmd.toLowerCase();
    const at = (x: number, y: number) => (rel ? { x: cur.x + x, y: cur.y + y } : { x, y });
    switch (cmd.toUpperCase()) {
      case 'M': {
        const v = nums(2);
        if (!v) return finish(points);
        add(at(v[0]!, v[1]!));
        start = cur;
        if (!rel) cmd = 'L';                            // further pairs after a moveto are linetos, as in SVG
        else cmd = 'l';
        break;
      }
      case 'L': {
        const v = nums(2);
        if (!v) return finish(points);
        add(at(v[0]!, v[1]!));
        break;
      }
      case 'C': {
        const v = nums(6);
        if (!v) return finish(points);
        const p0 = cur;
        const p1 = at(v[0]!, v[1]!);
        const p2 = at(v[2]!, v[3]!);
        const p3 = at(v[4]!, v[5]!);
        for (let k = 1; k <= CURVE_STEPS; k++) {
          const s = k / CURVE_STEPS;
          const u = 1 - s;
          const b0 = u * u * u, b1 = 3 * u * u * s, b2 = 3 * u * s * s, b3 = s * s * s;
          points.push({
            x: b0 * p0.x + b1 * p1.x + b2 * p2.x + b3 * p3.x,
            y: b0 * p0.y + b1 * p1.y + b2 * p2.y + b3 * p3.y,
          });
        }
        cur = p3;
        break;
      }
      case 'Z':
        add(start);
        break;
      case 'E':
        return finish(points);
      default:
        return finish(points);
    }
  }
  return finish(points);
}

function finish(points: { x: number; y: number }[]): MotionPath | undefined {
  if (!points.length) return undefined;
  const lengths = [0];
  for (let k = 1; k < points.length; k++) {
    lengths.push(lengths[k - 1]! + Math.hypot(points[k]!.x - points[k - 1]!.x, points[k]!.y - points[k - 1]!.y));
  }
  return { points, lengths, length: lengths.at(-1)! };
}

/** The point a fraction t (0..1) of the way along the path, by distance. */
export function pointAt(path: MotionPath, t: number): { x: number; y: number } {
  const { points, lengths, length } = path;
  if (length <= 0 || points.length === 1 || t <= 0) return points[0]!;
  if (t >= 1) return points.at(-1)!;
  const d = t * length;
  let lo = 0;
  let hi = lengths.length - 1;
  while (hi - lo > 1) {
    const mid = (lo + hi) >> 1;
    if (lengths[mid]! <= d) lo = mid;
    else hi = mid;
  }
  const seg = lengths[hi]! - lengths[lo]!;
  const f = seg > 0 ? (d - lengths[lo]!) / seg : 0;
  const a = points[lo]!;
  const b = points[hi]!;
  return { x: a.x + (b.x - a.x) * f, y: a.y + (b.y - a.y) * f };
}
