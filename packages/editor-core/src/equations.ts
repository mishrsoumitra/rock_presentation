// packages/editor-core/src/equations.ts  -- equations: typed on one line, stored as Office math
import type { XmlElement } from '@pptx/xml';
import { child, children, type ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { topSlot, type Rect } from './shapes.ts';
import { element, esc, fragment, insertAt, nextShapeId, twins, unitOf } from './xml.ts';

const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

/** An equation as a tree: what the one-line form parses to and what Office math is written from. */
export type MathNode =
  | { t: 'run'; text: string }
  | { t: 'seq'; items: MathNode[] }
  | { t: 'sup'; base: MathNode; sup: MathNode }
  | { t: 'sub'; base: MathNode; sub: MathNode }
  | { t: 'subsup'; base: MathNode; sub: MathNode; sup: MathNode }
  | { t: 'frac'; num: MathNode; den: MathNode }
  | { t: 'rad'; body: MathNode }
  | { t: 'paren'; body: MathNode };

const TOKEN = /sqrt(?![\p{L}])|\d+(?:\.\d+)?|[\p{L}]+|\S/gu;

/**
 * Parse an equation typed on one line, the way Word's linear format reads: `^` raises and `_` lowers what follows,
 * `/` makes a fraction of what is on either side of it, `sqrt(...)` (or √) is a root, and brackets group. Brackets
 * that only group an exponent, an index, a fraction's halves or a root's inside are not shown. So
 * `x = (-b ± sqrt(b^2 - 4ac)) / (2a)` is the quadratic formula. Nothing typed is ever an error: what does not
 * parse as structure stays as plain symbols.
 */
export function parseEquation(text: string): MathNode {
  const tokens = text.match(TOKEN) ?? [];
  let at = 0;
  const peek = () => tokens[at];
  const seq = (items: MathNode[]): MathNode => (items.length === 1 ? items[0]! : { t: 'seq', items });
  /** Brackets that only group are dropped. */
  const bare = (n: MathNode): MathNode => (n.t === 'paren' ? n.body : n);
  const atom = (): MathNode => {
    const tok = tokens[at++];
    if (tok === undefined) return { t: 'seq', items: [] };
    if (tok === '(') {
      const body = sequence(true);
      if (peek() === ')') at++;
      return { t: 'paren', body };
    }
    if (tok === 'sqrt' || tok === '√') return { t: 'rad', body: bare(atom()) };
    return { t: 'run', text: tok };
  };
  /** What a `^`, `_` or `/` applies to: an atom, with a sign in front of it taken along (x^-1). */
  const operand = (): MathNode => {
    if (peek() === '-' || peek() === '+' || peek() === '−') { const sign = tokens[at++]!; return seq([{ t: 'run', text: sign }, bare(scripted())]); }
    return bare(scripted());
  };
  function scripted(): MathNode {
    let base = atom();
    for (;;) {
      if (peek() === '^') {
        at++;
        const sup = bare(atomWithSign());
        base = base.t === 'sub' ? { t: 'subsup', base: base.base, sub: base.sub, sup } : { t: 'sup', base, sup };
      } else if (peek() === '_') {
        at++;
        const sub = bare(atomWithSign());
        base = base.t === 'sup' ? { t: 'subsup', base: base.base, sub, sup: base.sup } : { t: 'sub', base, sub };
      } else return base;
    }
  }
  function atomWithSign(): MathNode {
    if (peek() === '-' || peek() === '+' || peek() === '−') { const sign = tokens[at++]!; return seq([{ t: 'run', text: sign }, bare(atom())]); }
    return atom();
  }
  const term = (): MathNode => {
    let left = scripted();
    while (peek() === '/') {
      at++;
      left = { t: 'frac', num: bare(left), den: operand() };
    }
    return left;
  };
  function sequence(inBrackets: boolean): MathNode {
    const items: MathNode[] = [];
    while (at < tokens.length && !(inBrackets && peek() === ')')) items.push(term());
    return seq(items);
  }
  return sequence(false);
}

/** A math run as PowerPoint writes one: Cambria Math, with the text's size. */
const run = (text: string, size: number) =>
  `<m:r><a:rPr lang="en-US" sz="${size}" i="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t>${esc(text)}</m:t></m:r>`;

/** Office math (the content of an m:oMath) for an equation. */
export function ommlOf(n: MathNode, size = 2400): string {
  const x = (m: MathNode) => ommlOf(m, size);
  switch (n.t) {
    case 'run': return run(n.text, size);
    case 'seq': return n.items.map(x).join('');
    case 'sup': return `<m:sSup><m:e>${x(n.base)}</m:e><m:sup>${x(n.sup)}</m:sup></m:sSup>`;
    case 'sub': return `<m:sSub><m:e>${x(n.base)}</m:e><m:sub>${x(n.sub)}</m:sub></m:sSub>`;
    case 'subsup': return `<m:sSubSup><m:e>${x(n.base)}</m:e><m:sub>${x(n.sub)}</m:sub><m:sup>${x(n.sup)}</m:sup></m:sSubSup>`;
    case 'frac': return `<m:f><m:num>${x(n.num)}</m:num><m:den>${x(n.den)}</m:den></m:f>`;
    case 'rad': return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${x(n.body)}</m:e></m:rad>`;
    case 'paren': return `<m:d><m:e>${x(n.body)}</m:e></m:d>`;
  }
}

/** A piece of an equation written on one line as ordinary text: raised (1), lowered (-1) or on the line. */
interface Piece { text: string; script: -1 | 0 | 1; }

/**
 * An equation on one line: a fraction with a slash, a root with its sign, exponents raised and indices lowered.
 * This is not typesetting; it is what a program that does not read Office math shows in the equation's place.
 */
export function equationPieces(n: MathNode, script: -1 | 0 | 1 = 0): Piece[] {
  const x = (m: MathNode, s: -1 | 0 | 1 = script) => equationPieces(m, s);
  /** In brackets, unless it is one number or name. */
  const bracketed = (m: MathNode, inFraction = false): Piece[] => {
    const bare = m.t === 'run' || m.t === 'paren' || (inFraction && (m.t === 'rad' || m.t === 'sup' || m.t === 'sub' || m.t === 'subsup'));
    return bare ? x(m) : [{ text: '(', script }, ...x(m), { text: ')', script }];
  };
  switch (n.t) {
    case 'run': return n.text ? [{ text: n.text, script }] : [];
    case 'seq': return n.items.flatMap(m => x(m));
    case 'sup': return [...x(n.base), ...x(n.sup, 1)];
    case 'sub': return [...x(n.base), ...x(n.sub, -1)];
    case 'subsup': return [...x(n.base), ...x(n.sub, -1), ...x(n.sup, 1)];
    case 'frac': return [...bracketed(n.num, true), { text: '/', script }, ...bracketed(n.den, true)];
    case 'rad': return [{ text: '√', script }, ...bracketed(n.body)];
    case 'paren': return [{ text: '(', script }, ...x(n.body), { text: ')', script }];
  }
}

/** The fallback's runs: names italic, numbers and signs upright, as mathematics is set; scripts raised and lowered. */
function fallbackRuns(text: string, size: number): string {
  const runs: { text: string; script: number; italic: boolean }[] = [];
  for (const piece of equationPieces(parseEquation(text))) {
    const italic = /^\p{L}+$/u.test(piece.text);
    const last = runs[runs.length - 1];
    if (last && last.script === piece.script && last.italic === italic) last.text += piece.text;
    else runs.push({ ...piece, italic });
  }
  return runs.map(r => `<a:r><a:rPr lang="en-US" sz="${size}"${r.italic ? ' i="1"' : ''}${r.script ? ` baseline="${r.script > 0 ? 30000 : -25000}"` : ''}><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><a:t>${esc(r.text)}</a:t></a:r>`).join('');
}

const kids = (e: XmlElement | undefined): XmlElement[] => (e ? e.children.filter((c): c is XmlElement => c.kind === 'element') : []);
const kid = (e: XmlElement | undefined, local: string) => kids(e).find(c => c.local === local);
const SIMPLE = /^(?:\d+(?:\.\d+)?|[\p{L}]+|.)?$/su;

/**
 * Office math back on one line, in the form `parseEquation` reads. Structures this editor does not write (matrices,
 * integrals with limits, accents) come out as their symbols in order, which is readable but loses their structure.
 */
export function linearOf(e: XmlElement | undefined): string {
  if (!e) return '';
  const part = (name: string) => linearOf(kid(e, name));
  /**
   * A part that the one-line form must keep together: bare when it is one number or name, a root or a bracket (and,
   * for the halves of a fraction, an exponent or index, which bind tighter than the slash); in brackets otherwise.
   */
  const grouped = (name: string, inFraction = false) => {
    const s = part(name);
    const inner = kids(kid(e, name)).filter(c => !c.local.endsWith('Pr'));
    const one = inner.length === 1 ? inner[0]!.local : '';
    const bare = (SIMPLE.test(s) && s !== 'sqrt') || one === 'rad' || one === 'd' || (inFraction && (one === 'sSup' || one === 'sSub' || one === 'sSubSup'));
    return bare ? s : `(${s})`;
  };
  switch (e.local) {
    case 'r': return kids(e).filter(c => c.local === 't').map(t => t.children.map(c => (c.kind === 'text' ? c.value : '')).join('')).join('');
    case 'f': return `${grouped('num', true)}/${grouped('den', true)}`;
    case 'sSup': return `${part('e')}^${grouped('sup')}`;
    case 'sSub': return `${part('e')}_${grouped('sub')}`;
    case 'sSubSup': return `${part('e')}_${grouped('sub')}^${grouped('sup')}`;
    case 'rad': return `sqrt(${part('e')})`;
    case 'd': {
      const chr = (name: string, dflt: string) => kid(kid(e, 'dPr'), name)?.attrs.find(a => a.local === 'val')?.value ?? dflt;
      return `${chr('begChr', '(')}${kids(e).filter(c => c.local === 'e').map(linearOf).join(chr('sepChr', '|'))}${chr('endChr', ')')}`;
    }
    case 'nary': return `${kid(kid(e, 'naryPr'), 'chr')?.attrs.find(a => a.local === 'val')?.value ?? '∫'}${part('sub') ? `_${grouped('sub')}` : ''}${part('sup') ? `^${grouped('sup')}` : ''} ${part('e')}`;
    default: return kids(e).filter(c => !c.local.endsWith('Pr')).map(linearOf).reduce((all, next) => {
      // two names, or two numbers, side by side keep a space between them, so they read back as two
      const gap = /[\p{L}]$/u.test(all) && /^[\p{L}]/u.test(next) || /\d$/.test(all) && /^\d/.test(next);
      return all + (gap ? ' ' : '') + next;
    }, '');
  }
}

/** The equation (a14:m) of a shape's text, if it holds one: the first, in the first paragraph that has any. */
function mathOf(shape: XmlElement): XmlElement | undefined {
  for (const p of children(child(shape, 'p:txBody'), 'a:p')) {
    const m = p.children.find((c): c is XmlElement => c.kind === 'element' && c.local === 'm' && c.prefix === 'a14');
    if (m) return m;
  }
  return undefined;
}

/** The equation a shape holds, on one line, or undefined when it holds none. */
export function equationOf(shape: ShapeView | XmlElement): string | undefined {
  const m = mathIn('el' in shape ? shape.el : shape);
  return m ? linearOf(m) : undefined;
}

/** The math of a shape or of the same shape in another branch of its alternate content (where PowerPoint keeps it). */
const mathIn = (el: XmlElement): XmlElement | undefined => twins(el).map(mathOf).find(m => m !== undefined);

const mathXml = (text: string, size: number) =>
  `<a14:m><m:oMathPara xmlns:m="${M_NS}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath xmlns:m="${M_NS}">${ommlOf(parseEquation(text), size)}</m:oMath></m:oMathPara></a14:m>`;

/**
 * Insert an equation as PowerPoint stores one: a text box whose paragraph holds Office math, for every program
 * that reads it, beside a fallback of the same box with the equation as ordinary text on one line, for those that
 * do not. PowerPoint typesets the math; this editor, like PowerPoint 2007, shows the fallback.
 */
export function insertEquation(ed: Editor, part: string, text: string, box: Rect, size = 2400): { command: Command; id: number } {
  if (!text.trim()) throw new Error('an equation needs something in it');
  const tree = ed.spTree(part);
  const id = nextShapeId(ed.doc.root(part));
  const r = Math.round;
  const shell = (body: string) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr><a:xfrm><a:off x="${r(box.x)}" y="${r(box.y)}"/><a:ext cx="${r(box.w)}" cy="${r(box.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
    + `<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/>${body}</a:p></p:txBody></p:sp>`;
  const xml = `<mc:AlternateContent><mc:Choice Requires="a14">${shell(`${mathXml(text, size)}<a:endParaRPr lang="en-US" sz="${size}"/>`)}</mc:Choice>`
    + `<mc:Fallback>${shell(`${fallbackRuns(text, size)}<a:endParaRPr lang="en-US" sz="${size}"/>`)}</mc:Fallback></mc:AlternateContent>`;
  const command = ed.edit({
    label: 'Insert Equation', parts: [part], scope: () => [tree],
    apply: () => { insertAt(tree, topSlot(tree), element(tree, xml, ed.strict)); },
  });
  return { command, id };
}

/** Change the equation a shape holds: its math is written anew from the text, and its fallback's text with it. */
export function setEquation(ed: Editor, part: string, shape: ShapeView, text: string): Command {
  if (!text.trim()) throw new Error('an equation needs something in it');
  if (!mathIn(shape.el)) throw new Error('this shape holds no equation');
  return ed.edit({
    label: 'Edit Equation', parts: [part], scope: () => [unitOf(shape.el)],
    apply: () => {
      const old = mathIn(shape.el)!;
      // the size of the math there, so a changed equation is no bigger or smaller than it was
      let size = 2400;
      const find = (e: XmlElement): boolean => {
        if (e.local === 'rPr' && e.prefix === 'a') { const sz = Number(e.attrs.find(a => a.local === 'sz')?.value); if (sz > 0) { size = sz; return true; } }
        return kids(e).some(find);
      };
      find(old);
      const p = old.parent!;
      const fresh = element(p, mathXml(text, size), ed.strict);
      p.children[p.children.indexOf(old)] = fresh;
      fresh.parent = p;
      // the same box in the other branches shows the equation as text: its first paragraph's runs are written anew
      for (const twin of twins(shape.el)) {
        if (mathOf(twin)) continue;
        const para = children(child(twin, 'p:txBody'), 'a:p')[0];
        if (!para) continue;
        const runs = fragment(para, fallbackRuns(text, size), ed.strict);
        para.children = para.children.filter(c => !(c.kind === 'element' && ['r', 'br', 'fld'].includes(c.local)));
        const end = para.children.findIndex(c => c.kind === 'element' && c.local === 'endParaRPr');
        para.children.splice(end < 0 ? para.children.length : end, 0, ...runs);
        para.selfClosing = false;
      }
    },
  });
}
