// apps/web/src/ui/caret.ts  -- caret movement and hit testing over laid-out text, in file offsets
import { layoutToModel, modelToLayout, paragraphLength, paragraphs, type TextPos } from '@pptx/editor-core';
import { toHex, type Paint, type ResolvedRunProps, type ResolvedTextBody } from '@pptx/resolve';
import { caretAt, hitTest, type TextLayout } from '@pptx/text';
import type { XmlElement } from '@pptx/xml';

export interface TextView {
  txBody: XmlElement;
  resolved: ResolvedTextBody | null;
  layout: TextLayout<ResolvedRunProps>;
}

const paras = (v: TextView) => paragraphs(v.txBody);
const resolvedPara = (v: TextView, p: number) => v.resolved?.paragraphs[p];

export const toLayout = (v: TextView, pos: TextPos) => {
  const p = paras(v)[pos.p];
  return { paragraph: pos.p, offset: p ? modelToLayout(p, resolvedPara(v, pos.p), pos.offset) : pos.offset };
};

export const toModel = (v: TextView, paragraph: number, offset: number): TextPos => {
  const p = paras(v)[paragraph];
  return { p: paragraph, offset: p ? layoutToModel(p, resolvedPara(v, paragraph), offset) : offset };
};

/** The caret's rectangle in text box coordinates. */
export function caretRect(v: TextView, pos: TextPos): { x: number; top: number; height: number } | undefined {
  const l = toLayout(v, pos);
  return caretAt(v.layout, l.paragraph, l.offset);
}

const colourOf = (p: Paint | undefined): string | undefined =>
  p?.kind === 'solid' ? toHex(p.color)
    : p?.kind === 'linear' || p?.kind === 'path' ? (p.stops[0] ? toHex(p.stops[0].color) : undefined)
      : p?.kind === 'pattern' ? toHex(p.fg) : undefined;

/**
 * The caret's colour: the colour of the text it types into (the character before it, else the one after it, else
 * the paragraph mark's), as PowerPoint draws it. It shows wherever the text shows: white on a dark slide.
 */
export function caretColour(v: TextView, pos: TextPos): string {
  const l = toLayout(v, pos);
  const runs = v.layout.lines.filter(x => x.paragraph === l.paragraph).flatMap(x => x.runs).filter(r => !r.bullet);
  const run = runs.find(r => r.start < l.offset && l.offset <= r.end) ?? runs.find(r => r.start <= l.offset && l.offset < r.end);
  return colourOf(run?.props.fill) ?? colourOf(resolvedPara(v, l.paragraph)?.endProps.fill) ?? '#000000';
}

/** The position under a point of the text box. */
export function hit(v: TextView, x: number, y: number): TextPos {
  const r = hitTest(v.layout, x, y);
  return toModel(v, r.paragraph, r.offset);
}

/**
 * Where a click at a point of the text box puts the caret: the nearest position on the line under it, and, for a
 * click in the empty space below the last line, the end of that line (as in Google Slides).
 */
export function clickPos(v: TextView, x: number, y: number): TextPos {
  const last = v.layout.lines[v.layout.lines.length - 1];
  if (last && y > last.top + last.height) return toModel(v, last.paragraph, last.end);
  return hit(v, x, y);
}

/** Whether a point of the text box is on one of its lines of text (a little padded), rather than empty space. */
export function onTextLine(v: TextView, x: number, y: number, pad: number): boolean {
  for (const line of v.layout.lines) {
    const runs = line.runs.filter(r => !r.bullet);
    if (!runs.length) continue;
    const x0 = Math.min(...runs.map(g => g.x)) - pad, x1 = Math.max(...runs.map(g => g.x + g.advance)) + pad;
    if (x >= x0 && x <= x1 && y >= line.top - pad && y <= line.top + line.height + pad) return true;
  }
  return false;
}

export type Move = 'left' | 'right' | 'up' | 'down' | 'home' | 'end' | 'start' | 'finish' | 'wordLeft' | 'wordRight';

const isWord = (ch: string | undefined) => !!ch && /[\p{L}\p{N}_]/u.test(ch);

function textOfPara(p: XmlElement): string {
  let s = '';
  for (const c of p.children) {
    if (c.kind !== 'element') continue;
    if (c.local === 'br') s += '\v';
    else if (c.local === 'r' || c.local === 'fld') {
      for (const t of c.children) if (t.kind === 'element' && t.local === 't') for (const x of t.children) if (x.kind === 'text') s += x.value;
    }
  }
  return s;
}

/** Where the caret goes for a key. Up and down keep the caret's x across lines, as editors do. */
export function move(v: TextView, pos: TextPos, how: Move, goalX?: number): { pos: TextPos; goalX?: number } {
  const ps = paras(v);
  const len = (i: number) => (ps[i] ? paragraphLength(ps[i]!) : 0);
  switch (how) {
    case 'left':
      if (pos.offset > 0) return { pos: { p: pos.p, offset: pos.offset - 1 } };
      return { pos: pos.p > 0 ? { p: pos.p - 1, offset: len(pos.p - 1) } : pos };
    case 'right':
      if (pos.offset < len(pos.p)) return { pos: { p: pos.p, offset: pos.offset + 1 } };
      return { pos: pos.p + 1 < ps.length ? { p: pos.p + 1, offset: 0 } : pos };
    case 'wordLeft':
    case 'wordRight': {
      const text = textOfPara(ps[pos.p]!);
      let i = pos.offset;
      if (how === 'wordLeft') {
        if (i === 0) return move(v, pos, 'left');
        while (i > 0 && !isWord(text[i - 1])) i--;
        while (i > 0 && isWord(text[i - 1])) i--;
      } else {
        if (i >= text.length) return move(v, pos, 'right');
        while (i < text.length && isWord(text[i])) i++;
        while (i < text.length && !isWord(text[i])) i++;
      }
      return { pos: { p: pos.p, offset: i } };
    }
    case 'start': return { pos: { p: 0, offset: 0 } };
    case 'finish': return { pos: { p: ps.length - 1, offset: len(ps.length - 1) } };
    case 'home':
    case 'end': {
      const l = toLayout(v, pos);
      const line = v.layout.lines.find(x => x.paragraph === l.paragraph && l.offset >= x.start && l.offset <= x.end);
      if (!line) return { pos };
      return { pos: toModel(v, line.paragraph, how === 'home' ? line.start : line.end) };
    }
    case 'up':
    case 'down': {
      const l = toLayout(v, pos);
      const lines = v.layout.lines;
      const idx = lines.findIndex(x => x.paragraph === l.paragraph && l.offset >= x.start && l.offset <= x.end);
      const here = caretAt(v.layout, l.paragraph, l.offset);
      const x = goalX ?? here?.x ?? 0;
      const target = lines[idx + (how === 'up' ? -1 : 1)];
      if (idx < 0 || !target) return { pos: how === 'up' ? { p: 0, offset: 0 } : { p: ps.length - 1, offset: len(ps.length - 1) }, goalX: x };
      return { pos: hit(v, x, target.top + target.height / 2), goalX: x };
    }
  }
}

/** The word around a position (double-click). */
export function wordAt(v: TextView, pos: TextPos): { start: TextPos; end: TextPos } {
  const text = textOfPara(paras(v)[pos.p]!);
  let a = pos.offset, b = pos.offset;
  while (a > 0 && isWord(text[a - 1])) a--;
  while (b < text.length && isWord(text[b])) b++;
  return { start: { p: pos.p, offset: a }, end: { p: pos.p, offset: b } };
}

export const samePos = (a: TextPos, b: TextPos) => a.p === b.p && a.offset === b.offset;
export const before = (a: TextPos, b: TextPos) => a.p < b.p || (a.p === b.p && a.offset < b.offset);
