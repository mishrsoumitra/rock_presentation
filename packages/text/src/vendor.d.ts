// Types for the two Unicode-algorithm libraries the plan names, which ship none usable from here.

declare module 'linebreak' {
  /** UAX #14 line breaking (foliojs/linebreak). */
  export default class LineBreaker {
    constructor(text: string);
    /** The next break opportunity: `position` is the UTF-16 index the next line may start at. */
    nextBreak(): { position: number; required: boolean } | null;
  }
}

// bidi-js points its "module" field at a file it does not ship, so its ES module sources are imported directly.
declare module 'bidi-js/src/index.js' {
  export interface EmbeddingLevels {
    levels: Uint8Array;
    paragraphs: { start: number; end: number; level: number }[];
  }
  /** UAX #9 embedding levels; odd levels are right-to-left. */
  export function getEmbeddingLevels(text: string, baseDirection?: 'ltr' | 'rtl' | 'auto'): EmbeddingLevels;
  /** Ranges [start, end] (inclusive) to reverse, in order, to get visual order for text[start..end]. */
  export function getReorderSegments(text: string, levels: EmbeddingLevels, start?: number, end?: number): [number, number][];
  export function getMirroredCharacter(char: string): string | null;
}
