// packages/editor-core/src/text.ts  -- editing text: typing, deleting, paragraphs, character and paragraph formatting
import { child, children, clone, createElement, ensureChild, insertInOrder, is, removeChildren, runsOf, setAttr, setText } from '@pptx/model';
import type { ResolvedParagraph } from '@pptx/resolve';
import { FORBIDDEN_CHAR, type XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { element, esc, insertAt } from './xml.ts';

/**
 * A caret position: paragraph index and character offset in the paragraph as the file stores it. A run counts its
 * text, a line break (a:br) one character, a field (a:fld) its cached text; fields are never entered.
 */
export interface TextPos { p: number; offset: number; }
export interface TextRange { start: TextPos; end: TextPos; }

export type ColorSpec = { kind: 'srgb'; hex: string } | { kind: 'scheme'; val: string; lumMod?: number; lumOff?: number };

/** Character formatting to set: undefined leaves a property alone, null removes it (the text inherits it again). */
export interface RunPatch {
  b?: boolean | null;
  i?: boolean | null;
  u?: string | null;
  strike?: string | null;
  sz?: number | null;           // hundredths of a point
  baseline?: number | null;     // thousandths of a percent: 30000 superscript
  cap?: string | null;
  spc?: number | null;
  fill?: ColorSpec | null;
  highlight?: ColorSpec | null;
  latin?: string | null;
  ea?: string | null;
  cs?: string | null;
}

export type BulletSpec = { kind: 'none' } | { kind: 'char'; char: string; font?: string } | { kind: 'autonum'; scheme: string; startAt?: number };
export type SpacingSpec = { pct: number } | { pts: number };

/** Paragraph formatting to set; undefined leaves alone, null removes. */
export interface ParaPatch {
  algn?: string | null;
  lvl?: number | null;
  marL?: number | null;
  indent?: number | null;
  rtl?: boolean | null;
  lnSpc?: SpacingSpec | null;
  spcBef?: SpacingSpec | null;
  spcAft?: SpacingSpec | null;
  bullet?: BulletSpec | null;
}

export interface BodyPatch {
  anchor?: string | null;
  anchorCtr?: boolean | null;
  wrap?: 'square' | 'none' | null;
  vert?: string | null;
  lIns?: number | null; tIns?: number | null; rIns?: number | null; bIns?: number | null;
  numCol?: number | null;
  autofit?: { kind: 'none' } | { kind: 'shape' } | { kind: 'norm'; fontScale?: number; lnSpcReduction?: number } | null;
}

const runLength = (r: { kind: string; text: string }) => (r.kind === 'br' ? 1 : r.text.length);

export function paragraphs(txBody: XmlElement): XmlElement[] { return children(txBody, 'a:p'); }

export function paragraphLength(p: XmlElement): number {
  return runsOf(p).reduce((n, r) => n + runLength(r), 0);
}

/** Text of a range, paragraphs joined with "\n" and line breaks as U+000B. */
export function textIn(txBody: XmlElement, range: TextRange): string {
  const ps = paragraphs(txBody);
  const out: string[] = [];
  for (let i = range.start.p; i <= range.end.p && i < ps.length; i++) {
    const text = runsOf(ps[i]!).map(r => (r.kind === 'br' ? '\u000b' : r.text)).join('');
    const from = i === range.start.p ? range.start.offset : 0;
    const to = i === range.end.p ? range.end.offset : text.length;
    out.push(text.slice(from, to));
  }
  return out.join('\n');
}

const order = (a: TextPos, b: TextPos) => a.p - b.p || a.offset - b.offset;
const sorted = (r: TextRange): TextRange => (order(r.start, r.end) <= 0 ? r : { start: r.end, end: r.start });

/** Make sure a run boundary falls at `offset` (splitting a text run), and return the child index it falls before. */
function boundaryAt(p: XmlElement, offset: number): number {
  let at = 0;
  for (const r of runsOf(p)) {
    const len = runLength(r);
    const idx = p.children.indexOf(r.el);
    if (offset <= at) return idx;
    if (offset < at + len) {
      if (r.kind !== 'r') return idx + 1;                    // inside a field: the boundary after it
      const t = child(r.el, 'a:t');
      const head = r.text.slice(0, offset - at), tail = r.text.slice(offset - at);
      const second = clone(r.el, p);
      setText(t!, head);
      setText(child(second, 'a:t')!, tail);
      insertAt(p, idx + 1, second);
      return idx + 1;
    }
    at += len;
  }
  return endIndex(p);
}

/** Where runs end in a paragraph: before a:endParaRPr and any extLst. */
function endIndex(p: XmlElement): number {
  const end = p.children.findIndex(c => c.kind === 'element' && (is(c, 'a:endParaRPr') || is(c, 'a:extLst')));
  return end < 0 ? p.children.length : end;
}

/** The run elements of a paragraph between two child indexes. */
const runElsBetween = (p: XmlElement, from: number, to: number) =>
  p.children.slice(from, to).filter((c): c is XmlElement => c.kind === 'element' && (is(c, 'a:r') || is(c, 'a:br') || is(c, 'a:fld')));

/** Properties for new text at a position: the run before it, else the one after, else the paragraph mark's. */
function propsAt(p: XmlElement, offset: number): XmlElement | undefined {
  let at = 0;
  let after: XmlElement | undefined;
  let seenAfter = false;
  for (const r of runsOf(p)) {
    const len = runLength(r);
    if (offset > at && offset <= at + len) return r.rPr ?? child(p, 'a:endParaRPr');
    if (!seenAfter && at >= offset) {
      after = r.rPr;
      seenAfter = true;
    }
    at += len;
  }
  return after ?? child(p, 'a:endParaRPr');
}

/** A new a:rPr for `parent` from run or paragraph-mark properties (a:endParaRPr becomes a:rPr), minus dirty flags. */
function rPrFrom(parent: XmlElement, source: XmlElement | undefined): XmlElement {
  const rPr = source ? clone(source, parent) : createElement(parent, 'a:rPr', { lang: 'en-US' });
  rPr.local = 'rPr';
  rPr.attrs = rPr.attrs.filter(a => !(a.ns === '' && (a.local === 'dirty' || a.local === 'err' || a.local === 'smtClean' || a.local === 'smtId')));
  return rPr;
}

const clean = (s: string) => [...s].filter(ch => !FORBIDDEN_CHAR.test(ch) || ch === '\u000b').join('');

/** Insert text without paragraph breaks into one paragraph; "\v" becomes a line break. Returns the new offset. */
function insertInParagraph(ed: Editor, p: XmlElement, offset: number, text: string, patch?: RunPatch): number {
  if (!text) return offset;
  const pieces = text.split('\u000b');
  // the simple, common case: typing into a text run, with no formatting change
  if (!patch && pieces.length === 1) {
    let at = 0;
    for (const r of runsOf(p)) {
      const len = runLength(r);
      if (r.kind === 'r' && offset >= at && offset <= at + len && (offset > at || at === 0)) {
        const t = child(r.el, 'a:t') ?? ensureChild(r.el, 'a:t');
        setText(t, r.text.slice(0, offset - at) + text + r.text.slice(offset - at));
        return offset + text.length;
      }
      at += len;
    }
  }
  const props = propsAt(p, offset);
  let idx = boundaryAt(p, offset);
  pieces.forEach((piece, i) => {
    if (i > 0) {
      const br = createElement(p, 'a:br');
      const brPr = rPrFrom(br, props);
      br.children.push(brPr);
      br.selfClosing = false;
      insertAt(p, idx++, br);
    }
    if (!piece) return;
    const run = createElement(p, 'a:r');
    const rPr = rPrFrom(run, props);
    if (patch) patchRunProps(ed, rPr, patch);
    run.children.push(rPr, element(run, `<a:t>${esc(piece)}</a:t>`, ed.strict));
    for (const c of run.children) if (c.kind === 'element') c.parent = run;
    run.selfClosing = false;
    insertAt(p, idx++, run);
  });
  return offset + text.length;
}

/** Split paragraph `p` at an offset: the new paragraph after it gets the runs after the offset and the same pPr. */
function splitAt(txBody: XmlElement, p: XmlElement, offset: number): XmlElement {
  const idx = boundaryAt(p, offset);
  const end = endIndex(p);
  const moved = runElsBetween(p, idx, end);
  const next = createElement(txBody, 'a:p');
  const pPr = child(p, 'a:pPr');
  if (pPr) next.children.push(clone(pPr, next));
  for (const r of moved) {
    p.children.splice(p.children.indexOf(r), 1);
    next.children.push(r);
    r.parent = next;
  }
  const endPr = child(p, 'a:endParaRPr');
  const mark = endPr ? clone(endPr, next) : (() => {
    const m = rPrFrom(next, runsOf(p).at(-1)?.rPr ?? moved.map(r => child(r, 'a:rPr')).find(x => x));
    m.local = 'endParaRPr';
    return m;
  })();
  next.children.push(mark);
  next.selfClosing = false;
  insertAt(txBody, txBody.children.indexOf(p) + 1, next);
  return next;
}

/** Remove the characters [from, to) of one paragraph; a field or line break the range touches goes entirely. */
function deleteInParagraph(p: XmlElement, from: number, to: number): void {
  if (to <= from) return;
  const runs = runsOf(p);
  const firstRPr = (() => {
    let at = 0;
    for (const r of runs) {
      if (at + runLength(r) > from) return r.rPr;
      at += runLength(r);
    }
    return undefined;
  })();
  let at = 0;
  for (const r of runs) {
    const len = runLength(r);
    const s = Math.max(from, at), e = Math.min(to, at + len);
    if (e > s || (len === 0 && at >= from && at < to)) {
      if (r.kind === 'r' && (s > at || e < at + len)) setText(child(r.el, 'a:t')!, r.text.slice(0, s - at) + r.text.slice(e - at));
      else p.children.splice(p.children.indexOf(r.el), 1);
    }
    at += len;
  }
  // a paragraph emptied of text keeps the formatting of what was deleted for what is typed next
  if (runsOf(p).length === 0 && firstRPr) {
    const old = child(p, 'a:endParaRPr');
    const mark = clone(firstRPr, p);
    mark.local = 'endParaRPr';
    if (old) p.children[p.children.indexOf(old)] = mark;
    else insertInOrder(p, mark);
  }
}

function deleteRange(txBody: XmlElement, range: TextRange): void {
  const { start, end } = sorted(range);
  const ps = paragraphs(txBody);
  const first = ps[start.p], last = ps[end.p];
  if (!first || !last) return;
  if (first === last) {
    deleteInParagraph(first, start.offset, end.offset);
    return;
  }
  deleteInParagraph(first, start.offset, paragraphLength(first));
  deleteInParagraph(last, 0, end.offset);
  for (let i = start.p + 1; i < end.p; i++) txBody.children.splice(txBody.children.indexOf(ps[i]!), 1);
  // merge: the first paragraph keeps its pPr, takes the last one's runs and paragraph mark
  const at = endIndex(first);
  const moved = runElsBetween(last, 0, endIndex(last));
  first.children.splice(at, 0, ...moved);
  for (const r of moved) r.parent = first;
  const mark = child(last, 'a:endParaRPr');
  removeChildren(first, 'a:endParaRPr');
  if (mark) {
    const m = clone(mark, first);
    insertInOrder(first, m);
  }
  txBody.children.splice(txBody.children.indexOf(last), 1);
}

/** Make sure the body has at least one paragraph (a:txBody requires one). */
function ensureParagraph(ed: Editor, txBody: XmlElement): void {
  if (paragraphs(txBody).length === 0) insertInOrder(txBody, element(txBody, '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>', ed.strict));
}

export interface TextEdit { command: Command; caret: TextPos; }

/**
 * Options every text command takes. `after` runs inside the command, after the edit: the editor refits autofit text
 * there (PowerPoint's shrink-on-overflow and resize-shape-to-fit), so typing and refitting are one undo step;
 * `scope` widens what the command may change to the whole shape for that (a resized text box).
 */
export interface TextOptions { coalesceKey?: string; scope?: XmlElement; after?: () => void; label?: string; }

/**
 * Type text at a caret or over a selection. "\n" starts a new paragraph, "\v" is a line break. Consecutive typing
 * in one text body coalesces into one undo step (the UI passes the same `coalesceKey` and seals on caret moves).
 */
export function typeText(ed: Editor, part: string, txBody: XmlElement, at: TextPos | TextRange, text: string, options: TextOptions & { patch?: RunPatch } = {}): TextEdit {
  const range = 'start' in at ? sorted(at) : { start: at, end: at };
  let caret = range.start;
  const command = ed.edit({
    label: options.label ?? 'Typing', parts: [part], scope: () => [options.scope ?? txBody], ...(options.coalesceKey ? { coalesceKey: options.coalesceKey } : {}),
    apply: () => {
      ensureParagraph(ed, txBody);
      if (order(range.start, range.end) < 0) deleteRange(txBody, range);
      const lines = clean(text).split(/\r\n|\r|\n/);
      let pos = range.start;
      lines.forEach((line, i) => {
        if (i > 0) {
          const p = paragraphs(txBody)[pos.p]!;
          splitAt(txBody, p, pos.offset);
          pos = { p: pos.p + 1, offset: 0 };
        }
        const p = paragraphs(txBody)[pos.p];
        if (!p) return;
        pos = { p: pos.p, offset: insertInParagraph(ed, p, pos.offset, line, options.patch) };
      });
      caret = pos;
      options.after?.();
    },
  });
  return { command, get caret() { return caret; } };
}

/** Delete a selection, or the character before (backspace) or after (delete) a caret, joining paragraphs at edges. */
export function deleteText(ed: Editor, part: string, txBody: XmlElement, at: TextPos | TextRange, direction: 'back' | 'forward' = 'back', options: TextOptions = {}): TextEdit {
  let caret = 'start' in at ? sorted(at).start : at;
  const command = ed.edit({
    label: options.label ?? 'Delete', parts: [part], scope: () => [options.scope ?? txBody], ...(options.coalesceKey ? { coalesceKey: options.coalesceKey } : {}),
    apply: () => {
      let range: TextRange;
      if ('start' in at && order(at.start, at.end) !== 0) range = sorted(at);
      else {
        const c = 'start' in at ? at.start : at;
        const ps = paragraphs(txBody);
        if (direction === 'back') {
          if (c.offset > 0) range = { start: { p: c.p, offset: c.offset - 1 }, end: c };
          else if (c.p > 0) range = { start: { p: c.p - 1, offset: paragraphLength(ps[c.p - 1]!) }, end: c };
          else return;
        } else {
          const len = ps[c.p] ? paragraphLength(ps[c.p]!) : 0;
          if (c.offset < len) range = { start: c, end: { p: c.p, offset: c.offset + 1 } };
          else if (c.p + 1 < ps.length) range = { start: c, end: { p: c.p + 1, offset: 0 } };
          else return;
        }
      }
      deleteRange(txBody, range);
      ensureParagraph(ed, txBody);
      caret = range.start;
      options.after?.();
    },
  });
  return { command, get caret() { return caret; } };
}

/** Replace all text with plain text, keeping the first paragraph's and first run's formatting. */
export function replaceText(ed: Editor, part: string, txBody: XmlElement, text: string): Command {
  return ed.edit({
    label: 'Replace Text', parts: [part], scope: () => [txBody],
    apply: () => {
      ensureParagraph(ed, txBody);
      const ps = paragraphs(txBody);
      deleteRange(txBody, { start: { p: 0, offset: 0 }, end: { p: ps.length - 1, offset: paragraphLength(ps[ps.length - 1]!) } });
      const lines = clean(text).split(/\r\n|\r|\n/);
      let pos: TextPos = { p: 0, offset: 0 };
      lines.forEach((line, i) => {
        if (i > 0) {
          splitAt(txBody, paragraphs(txBody)[pos.p]!, pos.offset);
          pos = { p: pos.p + 1, offset: 0 };
        }
        pos = { p: pos.p, offset: insertInParagraph(ed, paragraphs(txBody)[pos.p]!, pos.offset, line) };
      });
    },
  });
}

function colorXml(c: ColorSpec): string {
  if (c.kind === 'srgb') return `<a:srgbClr val="${esc(c.hex.replace(/^#/, '').toUpperCase())}"/>`;
  const mods = `${c.lumMod !== undefined ? `<a:lumMod val="${c.lumMod}"/>` : ''}${c.lumOff !== undefined ? `<a:lumOff val="${c.lumOff}"/>` : ''}`;
  return mods ? `<a:schemeClr val="${esc(c.val)}">${mods}</a:schemeClr>` : `<a:schemeClr val="${esc(c.val)}"/>`;
}

const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

/** Apply a character formatting patch to an a:rPr (or a:endParaRPr, a:defRPr), keeping schema order. */
export function patchRunProps(ed: Editor, rPr: XmlElement, patch: RunPatch): void {
  const flag = (v: boolean | null | undefined) => (v === null ? undefined : v ? '1' : '0');
  if (patch.b !== undefined) setAttr(rPr, 'b', flag(patch.b));
  if (patch.i !== undefined) setAttr(rPr, 'i', flag(patch.i));
  if (patch.u !== undefined) setAttr(rPr, 'u', patch.u ?? undefined);
  if (patch.strike !== undefined) setAttr(rPr, 'strike', patch.strike ?? undefined);
  if (patch.sz !== undefined) setAttr(rPr, 'sz', patch.sz === null ? undefined : Math.round(patch.sz));
  if (patch.baseline !== undefined) setAttr(rPr, 'baseline', patch.baseline ?? undefined);
  if (patch.cap !== undefined) setAttr(rPr, 'cap', patch.cap ?? undefined);
  if (patch.spc !== undefined) setAttr(rPr, 'spc', patch.spc ?? undefined);
  if (patch.fill !== undefined) {
    removeChildren(rPr, ...FILLS);
    if (patch.fill) insertInOrder(rPr, element(rPr, `<a:solidFill>${colorXml(patch.fill)}</a:solidFill>`, ed.strict));
  }
  if (patch.highlight !== undefined) {
    removeChildren(rPr, 'a:highlight');
    if (patch.highlight) insertInOrder(rPr, element(rPr, `<a:highlight>${colorXml(patch.highlight)}</a:highlight>`, ed.strict));
  }
  for (const script of ['latin', 'ea', 'cs'] as const) {
    const v = patch[script];
    if (v === undefined) continue;
    removeChildren(rPr, `a:${script}`);
    if (v) insertInOrder(rPr, element(rPr, `<a:${script} typeface="${esc(v)}"/>`, ed.strict));
  }
  rPr.selfClosing = rPr.children.length === 0;
}

/** The run elements a range covers, splitting runs at its edges; and the paragraphs whose mark it covers. */
function runsIn(txBody: XmlElement, range: TextRange): { runs: XmlElement[]; marks: XmlElement[] } {
  const { start, end } = sorted(range);
  const ps = paragraphs(txBody);
  const runs: XmlElement[] = [];
  const marks: XmlElement[] = [];
  for (let i = start.p; i <= end.p && i < ps.length; i++) {
    const p = ps[i]!;
    const len = paragraphLength(p);
    const from = i === start.p ? start.offset : 0;
    const to = i === end.p ? end.offset : len;
    const a = boundaryAt(p, from);
    const b = boundaryAt(p, to);
    runs.push(...runElsBetween(p, a, b));
    if (i < end.p || to >= len) marks.push(p);
  }
  return { runs, marks };
}

/**
 * Character formatting over a range (bold, size, colour, font...). Runs are split at the range's edges; the
 * paragraph mark (a:endParaRPr) of each paragraph the range covers to its end is formatted too, as PowerPoint does,
 * so new text typed there looks the same. A bare caret in an empty paragraph formats the mark.
 */
export function formatRuns(ed: Editor, part: string, txBody: XmlElement, range: TextRange, patch: RunPatch, options: TextOptions = {}): Command {
  return ed.edit({
    label: options.label ?? 'Font', parts: [part], scope: () => [options.scope ?? txBody], ...(options.coalesceKey ? { coalesceKey: options.coalesceKey } : {}),
    apply: () => {
      ensureParagraph(ed, txBody);
      const { runs, marks } = runsIn(txBody, range);
      for (const r of runs) {
        let rPr = child(r, 'a:rPr');
        if (!rPr) {
          rPr = createElement(r, 'a:rPr', { lang: 'en-US' });
          insertInOrder(r, rPr);
        }
        patchRunProps(ed, rPr, patch);
      }
      for (const p of marks) {
        const mark = child(p, 'a:endParaRPr') ?? (() => {
          const m = createElement(p, 'a:endParaRPr', { lang: 'en-US' });
          insertInOrder(p, m);
          return m;
        })();
        patchRunProps(ed, mark, patch);
      }
      options.after?.();
    },
  });
}

function spacingXml(name: string, s: SpacingSpec): string {
  return 'pct' in s ? `<a:${name}><a:spcPct val="${Math.round(s.pct)}"/></a:${name}>` : `<a:${name}><a:spcPts val="${Math.round(s.pts)}"/></a:${name}>`;
}

const BULLETS = ['a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip'];

/** Paragraph formatting (alignment, level, indents, spacing, bullets) for every paragraph a range touches. */
export function formatParagraphs(ed: Editor, part: string, txBody: XmlElement, range: TextRange, patch: ParaPatch, options: TextOptions = {}): Command {
  return ed.edit({
    label: options.label ?? 'Paragraph', parts: [part], scope: () => [options.scope ?? txBody], ...(options.coalesceKey ? { coalesceKey: options.coalesceKey } : {}),
    apply: () => {
      ensureParagraph(ed, txBody);
      const { start, end } = sorted(range);
      const ps = paragraphs(txBody);
      for (let i = start.p; i <= end.p && i < ps.length; i++) {
        const p = ps[i]!;
        const pPr = child(p, 'a:pPr') ?? (() => {
          const e = createElement(p, 'a:pPr');
          insertInOrder(p, e);
          return e;
        })();
        if (patch.algn !== undefined) setAttr(pPr, 'algn', patch.algn ?? undefined);
        if (patch.lvl !== undefined) setAttr(pPr, 'lvl', patch.lvl ? Math.max(0, Math.min(8, patch.lvl)) : undefined);
        if (patch.marL !== undefined) setAttr(pPr, 'marL', patch.marL ?? undefined);
        if (patch.indent !== undefined) setAttr(pPr, 'indent', patch.indent ?? undefined);
        if (patch.rtl !== undefined) setAttr(pPr, 'rtl', patch.rtl === null ? undefined : patch.rtl ? '1' : '0');
        for (const name of ['lnSpc', 'spcBef', 'spcAft'] as const) {
          const v = patch[name];
          if (v === undefined) continue;
          removeChildren(pPr, `a:${name}`);
          if (v) insertInOrder(pPr, element(pPr, spacingXml(name, v), ed.strict));
        }
        if (patch.bullet !== undefined) {
          removeChildren(pPr, ...BULLETS);
          const b = patch.bullet;
          if (b?.kind === 'char') {
            if (b.font) {
              removeChildren(pPr, 'a:buFont', 'a:buFontTx');
              insertInOrder(pPr, element(pPr, `<a:buFont typeface="${esc(b.font)}"/>`, ed.strict));
            }
            insertInOrder(pPr, element(pPr, `<a:buChar char="${esc(b.char)}"/>`, ed.strict));
          } else if (b?.kind === 'autonum') {
            insertInOrder(pPr, element(pPr, `<a:buAutoNum type="${esc(b.scheme)}"${b.startAt && b.startAt !== 1 ? ` startAt="${b.startAt}"` : ''}/>`, ed.strict));
          } else if (b?.kind === 'none') {
            insertInOrder(pPr, element(pPr, '<a:buNone/>', ed.strict));
          }
        }
        pPr.selfClosing = pPr.children.length === 0;
      }
      options.after?.();
    },
  });
}

/** Text box properties: anchoring, wrapping, insets, columns, autofit. */
export function formatBody(ed: Editor, part: string, txBody: XmlElement, patch: BodyPatch, options: TextOptions = {}): Command {
  return ed.edit({
    label: options.label ?? 'Text Options', parts: [part], scope: () => [options.scope ?? txBody], ...(options.coalesceKey ? { coalesceKey: options.coalesceKey } : {}),
    apply: () => {
      const bodyPr = ensureChild(txBody, 'a:bodyPr');
      for (const k of ['anchor', 'wrap', 'vert'] as const) if (patch[k] !== undefined) setAttr(bodyPr, k, patch[k] ?? undefined);
      if (patch.anchorCtr !== undefined) setAttr(bodyPr, 'anchorCtr', patch.anchorCtr === null ? undefined : patch.anchorCtr ? '1' : '0');
      for (const k of ['lIns', 'tIns', 'rIns', 'bIns', 'numCol'] as const) if (patch[k] !== undefined) setAttr(bodyPr, k, patch[k] ?? undefined);
      if (patch.autofit !== undefined) {
        removeChildren(bodyPr, 'a:noAutofit', 'a:normAutofit', 'a:spAutoFit');
        const a = patch.autofit;
        if (a?.kind === 'none') insertInOrder(bodyPr, element(bodyPr, '<a:noAutofit/>', ed.strict));
        else if (a?.kind === 'shape') insertInOrder(bodyPr, element(bodyPr, '<a:spAutoFit/>', ed.strict));
        else if (a?.kind === 'norm') {
          const attrs = `${a.fontScale !== undefined && a.fontScale !== 100000 ? ` fontScale="${Math.round(a.fontScale)}"` : ''}${a.lnSpcReduction ? ` lnSpcReduction="${Math.round(a.lnSpcReduction)}"` : ''}`;
          insertInOrder(bodyPr, element(bodyPr, `<a:normAutofit${attrs}/>`, ed.strict));
        }
      }
      bodyPr.selfClosing = bodyPr.children.length === 0;
      options.after?.();
    },
  });
}

/**
 * Offsets in the laid-out text (where fields show their computed text, a slide number "12") and in the file (the
 * field's cached text, "‹#›") differ after a field. These convert between them with the resolved paragraph, whose
 * runs say which source run they came from. Positions inside a field snap to its end.
 */
export function layoutToModel(p: XmlElement, resolved: ResolvedParagraph | undefined, offset: number): number {
  if (!resolved) return offset;
  const model = runsOf(p);
  let lay = 0, mod = 0;
  for (const rr of resolved.runs) {
    const mr = model[rr.source.r];
    const lLen = rr.kind === 'break' ? 1 : rr.text.length;
    const mLen = mr ? runLength(mr) : lLen;
    if (offset < lay + lLen) return rr.kind === 'field' ? mod + mLen : mod + Math.min(offset - lay, mLen);
    lay += lLen;
    mod += mLen;
  }
  return mod;
}

export function modelToLayout(p: XmlElement, resolved: ResolvedParagraph | undefined, offset: number): number {
  if (!resolved) return offset;
  const model = runsOf(p);
  let lay = 0, mod = 0;
  for (const rr of resolved.runs) {
    const mr = model[rr.source.r];
    const lLen = rr.kind === 'break' ? 1 : rr.text.length;
    const mLen = mr ? runLength(mr) : lLen;
    if (offset < mod + mLen) return rr.kind === 'field' ? lay + lLen : lay + Math.min(offset - mod, lLen);
    lay += lLen;
    mod += mLen;
  }
  return lay;
}

/** The text of a whole body: paragraphs joined by "\n". */
export function bodyPlainText(txBody: XmlElement): string {
  return paragraphs(txBody).map(p => runsOf(p).map(r => (r.kind === 'br' ? '\u000b' : r.text)).join('')).join('\n');
}

