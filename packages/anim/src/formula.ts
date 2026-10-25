// packages/anim/src/formula.ts  -- the expression language of p:tav fmla and formula values ("#ppt_x+.1"), evaluated safely
//
// ECMA-376 Part 1, 19.5.79: numbers, $ (the formula's input), the variables ppt_x, ppt_y, ppt_w and ppt_h (with or
// without the # PowerPoint puts before them in values; exit effects write them bare), pi and e, + - * / % ^, unary
// signs, and abs acos asin atan ceil cos cosh deg exp floor ln max min rad rand sin sinh sqrt tan tanh. Formulas come
// from files, so this is a small hand-written parser that compiles to closures: no eval, no Function, and input
// length, token count and nesting depth are bounded.

export class FormulaError extends Error {
  override name = 'FormulaError';
}

/** What a formula can read. Positions and sizes are fractions of the slide (0..1 across it). */
export interface FormulaScope {
  /** $: the input, the value interpolated between keyframes. */
  input?: number;
  ppt_x: number;
  ppt_y: number;
  ppt_w: number;
  ppt_h: number;
  /** rand(x) returns random() * x; deterministic by default so the same frame always looks the same. */
  random?: () => number;
}

export type Formula = (scope: FormulaScope) => number;

export const MAX_FORMULA_LENGTH = 1024;
const MAX_TOKENS = 512;
const MAX_DEPTH = 32;

type Token = { kind: 'num'; value: number } | { kind: 'name'; value: string } | { kind: 'op'; value: string };

function tokenize(src: string): Token[] {
  if (src.length > MAX_FORMULA_LENGTH) throw new FormulaError(`formula longer than ${MAX_FORMULA_LENGTH} characters`);
  const out: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i]!;
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') {
      i++;
      continue;
    }
    if (out.length >= MAX_TOKENS) throw new FormulaError(`formula with more than ${MAX_TOKENS} tokens`);
    const number = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(src.slice(i, i + 64));
    if (number) {
      out.push({ kind: 'num', value: Number(number[0]) });
      i += number[0].length;
      continue;
    }
    const name = /^#?[A-Za-z_][A-Za-z0-9_]*/.exec(src.slice(i, i + 64));
    if (name) {
      out.push({ kind: 'name', value: name[0].replace(/^#/, '').toLowerCase() });
      i += name[0].length;
      continue;
    }
    if ('+-*/%^(),$'.includes(c)) {
      out.push(c === '$' ? { kind: 'name', value: '$' } : { kind: 'op', value: c });
      i++;
      continue;
    }
    throw new FormulaError(`unexpected "${c}" at ${i}`);
  }
  return out;
}

// Maps, not object literals: a name from a file must never reach Object.prototype ("constructor", "__proto__").
const UNARY = new Map<string, (x: number) => number>([
  ['abs', Math.abs], ['acos', Math.acos], ['asin', Math.asin], ['atan', Math.atan], ['ceil', Math.ceil], ['cos', Math.cos],
  ['cosh', Math.cosh], ['deg', x => (x * 180) / Math.PI], ['exp', Math.exp], ['floor', Math.floor], ['ln', Math.log],
  ['rad', x => (x * Math.PI) / 180], ['sin', Math.sin], ['sinh', Math.sinh], ['sqrt', Math.sqrt], ['tan', Math.tan], ['tanh', Math.tanh],
]);
const BINARY = new Map<string, (x: number, y: number) => number>([['max', Math.max], ['min', Math.min]]);
const VARIABLES = new Map<string, Formula>([
  ['$', s => s.input ?? 0], ['ppt_x', s => s.ppt_x], ['ppt_y', s => s.ppt_y], ['ppt_w', s => s.ppt_w],
  ['ppt_h', s => s.ppt_h], ['pi', () => Math.PI], ['e', () => Math.E],
]);

/** A fixed pseudo-random sequence (mulberry32), used when the scope brings no random source. */
function seeded(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

class Parser {
  private pos = 0;
  private depth = 0;
  private readonly tokens: Token[];
  constructor(tokens: Token[]) { this.tokens = tokens; }

  parse(): Formula {
    if (!this.tokens.length) throw new FormulaError('empty formula');
    const f = this.sum();
    if (this.pos < this.tokens.length) throw new FormulaError(`unexpected "${this.tokens[this.pos]!.value}"`);
    return f;
  }

  private peek(op: string): boolean {
    const t = this.tokens[this.pos];
    return t?.kind === 'op' && t.value === op;
  }

  private expect(op: string): void {
    if (!this.peek(op)) throw new FormulaError(`expected "${op}"`);
    this.pos++;
  }

  private enter(): void {
    if (++this.depth > MAX_DEPTH) throw new FormulaError(`formula nested deeper than ${MAX_DEPTH}`);
  }

  private sum(): Formula {
    this.enter();
    let f = this.product();
    while (this.peek('+') || this.peek('-')) {
      const op = this.tokens[this.pos++]!.value;
      const l = f;
      const r = this.product();
      f = op === '+' ? s => l(s) + r(s) : s => l(s) - r(s);
    }
    this.depth--;
    return f;
  }

  private product(): Formula {
    let f = this.power();
    while (this.peek('*') || this.peek('/') || this.peek('%')) {
      const op = this.tokens[this.pos++]!.value;
      const l = f;
      const r = this.power();
      f = op === '*' ? s => l(s) * r(s) : op === '/' ? s => l(s) / r(s) : s => l(s) % r(s);
    }
    return f;
  }

  /** "^" binds tighter than * and /, looser than a sign: -2^2 is (-2)^2, as the spec's grammar has it. */
  private power(): Formula {
    const base = this.unary();
    if (!this.peek('^')) return base;
    this.pos++;
    this.enter();
    const exp = this.power();
    this.depth--;
    return s => base(s) ** exp(s);
  }

  private unary(): Formula {
    if (this.peek('-') || this.peek('+')) {
      const minus = this.tokens[this.pos++]!.value === '-';
      this.enter();
      const f = this.unary();
      this.depth--;
      return minus ? s => -f(s) : f;
    }
    return this.factor();
  }

  private factor(): Formula {
    const t = this.tokens[this.pos];
    if (!t) throw new FormulaError('unexpected end of formula');
    if (t.kind === 'num') {
      this.pos++;
      const v = t.value;
      return () => v;
    }
    if (t.kind === 'op' && t.value === '(') {
      this.pos++;
      const f = this.sum();
      this.expect(')');
      return f;
    }
    if (t.kind !== 'name') throw new FormulaError(`unexpected "${t.value}"`);
    this.pos++;
    const name = t.value;
    if (this.peek('(')) return this.call(name);
    const v = VARIABLES.get(name);
    if (!v) throw new FormulaError(`unknown name "${name}"`);
    return v;
  }

  private call(name: string): Formula {
    this.pos++;
    const args = [this.sum()];
    while (this.peek(',')) {
      this.pos++;
      args.push(this.sum());
    }
    this.expect(')');
    const [a, b] = args;
    if (name === 'rand' && args.length === 1) {
      const fallback = seeded(0x5eed);
      return s => (s.random ?? fallback)() * a!(s);
    }
    const one = UNARY.get(name);
    if (one && args.length === 1) return s => one(a!(s));
    const two = BINARY.get(name);
    if (two && args.length === 2) return s => two(a!(s), b!(s));
    throw new FormulaError(`unknown function ${name} with ${args.length} argument${args.length === 1 ? '' : 's'}`);
  }
}

const cache = new Map<string, Formula | FormulaError>();

/** Compile a formula once; throws FormulaError for anything outside the language or its limits. */
export function compileFormula(src: string): Formula {
  let f = cache.get(src);
  if (!f) {
    try {
      f = new Parser(tokenize(src)).parse();
    } catch (e) {
      f = e instanceof FormulaError ? e : new FormulaError(String(e));
    }
    if (cache.size > 2000) cache.clear();
    cache.set(src, f);
  }
  if (f instanceof FormulaError) throw f;
  return f;
}

/** Evaluate a formula; NaN when it cannot be compiled. Results can be infinite (1/0), as in the spec's arithmetic. */
export function evaluateFormula(src: string, scope: FormulaScope): number {
  try {
    return compileFormula(src)(scope);
  } catch (e) {
    if (e instanceof FormulaError) return Number.NaN;
    throw e;
  }
}
