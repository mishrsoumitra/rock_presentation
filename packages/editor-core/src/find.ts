// packages/editor-core/src/find.ts  -- find and replace across the slides of a deck
import { runsOf, type ShapeView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import { compound, type Command } from './history.ts';
import { paragraphs, typeText, type TextOptions, type TextRange } from './text.ts';

/** One place the text was found: a stretch of one paragraph of one shape, in file offsets (as carets are). */
export interface TextMatch {
  part: string;
  /** Where the match is in reading order: the slide's number from 0, then the shape's place among the slide's shapes. */
  slideIndex: number;
  shapeIndex: number;
  shape: XmlElement;
  txBody: XmlElement;
  p: number;
  start: number;
  end: number;
}

export interface FindOptions {
  /** "Total" does not find "total". Off by default, as in PowerPoint and Google Slides. */
  matchCase?: boolean;
  /** "art" does not find the middle of "party": the match must not touch a letter or digit on either side. */
  wholeWords?: boolean;
}

const WORD_CHAR = /[\p{L}\p{N}_]/u;

/** Where `query` occurs in `text`, left to right, never overlapping. Offsets are UTF-16, as everywhere in the editor. */
export function findInText(text: string, query: string, options: FindOptions = {}): { start: number; end: number }[] {
  if (!query) return [];
  const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.matchCase ? 'gu' : 'giu');
  const out: { start: number; end: number }[] = [];
  for (const m of text.matchAll(re)) {
    const start = m.index, end = start + m[0].length;
    if (options.wholeWords) {
      const before = start > 0 ? String.fromCodePoint(text.codePointAt(start - 1)!) : '';
      const after = end < text.length ? String.fromCodePoint(text.codePointAt(end)!) : '';
      if ((before && WORD_CHAR.test(before)) || (after && WORD_CHAR.test(after))) continue;
    }
    out.push({ start, end });
  }
  return out;
}

/** The text of a paragraph as carets count it: a line break is one character, which no query matches. */
const paragraphText = (p: XmlElement) => runsOf(p).map(r => (r.kind === 'br' ? '\u000b' : r.text)).join('');

/**
 * Every match in the deck, in reading order: slide by slide, each slide's shapes in drawing order (the shapes of a
 * group where the group is), paragraph by paragraph. A match never spans two paragraphs. Layouts and masters are
 * not searched: their text is not the deck's content.
 */
export function findText(ed: Editor, query: string, options: FindOptions = {}): TextMatch[] {
  const out: TextMatch[] = [];
  if (!query) return out;
  ed.doc.slideParts().forEach((part, slideIndex) => {
    let shapeIndex = 0;
    const visit = (shapes: ShapeView[]) => {
      for (const s of shapes) {
        if (s.kind === 'grpSp') { visit(s.children); continue; }
        const at = shapeIndex++;
        const txBody = s.txBody;
        if (!txBody) continue;
        paragraphs(txBody).forEach((p, pi) => {
          for (const m of findInText(paragraphText(p), query, options)) out.push({ part, slideIndex, shapeIndex: at, shape: s.el, txBody, p: pi, ...m });
        });
      }
    };
    visit(ed.doc.slide(part).shapes);
  });
  return out;
}

export const matchRange = (m: TextMatch): TextRange => ({ start: { p: m.p, offset: m.start }, end: { p: m.p, offset: m.end } });

/** Order of two places in the deck: negative when `a` comes first in reading order. */
export function compareMatches(a: Pick<TextMatch, 'slideIndex' | 'shapeIndex' | 'p' | 'start'>, b: Pick<TextMatch, 'slideIndex' | 'shapeIndex' | 'p' | 'start'>): number {
  return a.slideIndex - b.slideIndex || a.shapeIndex - b.shapeIndex || a.p - b.p || a.start - b.start;
}

/** Replace one match. The new text takes the formatting of the first character it replaces. */
export function replaceMatch(ed: Editor, m: TextMatch, replacement: string, options: TextOptions = {}): Command {
  return typeText(ed, m.part, m.txBody, matchRange(m), replacement, { label: 'Replace', likeReplaced: true, ...options }).command;
}

/**
 * Replace every match in the deck as one undo step. Within a shape the last match is replaced first, so the
 * offsets of the matches before it stay true. `options` gives each shape's text options (the app refits autofit
 * text there).
 */
export function replaceAll(
  ed: Editor, query: string, replacement: string, find: FindOptions = {}, options: (m: TextMatch) => TextOptions = () => ({}),
): { command: Command; count: number } {
  const matches = findText(ed, query, find);
  const last = [...matches].sort((a, b) => compareMatches(b, a));
  return { command: compound('Replace All', last.map(m => replaceMatch(ed, m, replacement, options(m)))), count: matches.length };
}
