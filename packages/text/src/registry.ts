// packages/text/src/registry.ts  -- which face draws a run: family and style matching, substitutes, fallbacks
import { FontFace } from './font.ts';

export interface FontMatch {
  face: FontFace;
  /** The face is regular but bold was asked for: draw it emboldened, as PowerPoint does. */
  syntheticBold: boolean;
  /** The face is upright but italic was asked for: draw it slanted. */
  syntheticItalic: boolean;
  /** The family asked for was not available and another one is used. */
  substituted: boolean;
}

/**
 * Metric-compatible open substitutes for Office's core fonts (plan section 3.4): same advance widths, so line breaks
 * fall where they would with the original. Aptos, the Office default since 2023, has none; Carlito stands in and
 * the match is reported as substituted.
 */
export const SUBSTITUTES: Record<string, string[]> = {
  'calibri': ['Carlito'],
  'calibri light': ['Carlito'],
  'cambria': ['Caladea'],
  'cambria math': ['Caladea'],
  'arial': ['Liberation Sans'],
  'helvetica': ['Liberation Sans'],
  'arial narrow': ['Liberation Sans Narrow', 'Liberation Sans'],
  'times new roman': ['Liberation Serif'],
  'times': ['Liberation Serif'],
  'courier new': ['Liberation Mono'],
  'courier': ['Liberation Mono'],
  'aptos': ['Carlito'],
  'aptos display': ['Carlito'],
  'segoe ui': ['Carlito'],
  'tahoma': ['Liberation Sans'],
  'verdana': ['Liberation Sans'],
  'georgia': ['Liberation Serif'],
};

const key = (family: string) => family.trim().toLowerCase();

export class FontRegistry {
  private readonly families = new Map<string, FontFace[]>();
  private readonly order: FontFace[] = [];
  /** The family used when nothing else matches. */
  defaultFamily = 'Carlito';

  add(face: FontFace, family = face.family): void {
    const k = key(family);
    const list = this.families.get(k) ?? [];
    list.push(face);
    this.families.set(k, list);
    if (!this.order.includes(face)) this.order.push(face);
  }

  /** Register font bytes; returns the face, or undefined when the data is not a usable font. */
  addBytes(bytes: Uint8Array, options: { family?: string; index?: number } = {}): FontFace | undefined {
    try {
      const face = new FontFace(bytes, options.index ?? 0, options.family);
      this.add(face);
      return face;
    } catch {
      return undefined;
    }
  }

  has(family: string): boolean {
    return this.families.has(key(family));
  }

  familyNames(): string[] {
    return [...this.families.keys()];
  }

  /** The face of a known family closest to the asked style, or undefined when the family is not registered. */
  find(family: string, bold: boolean, italic: boolean): FontMatch | undefined {
    const faces = this.families.get(key(family));
    if (!faces?.length) return undefined;
    const score = (f: FontFace) => (f.metrics.bold === bold ? 0 : 2) + (f.metrics.italic === italic ? 0 : 1);
    const face = faces.reduce((best, f) => (score(f) < score(best) ? f : best));
    return { face, syntheticBold: bold && !face.metrics.bold, syntheticItalic: italic && !face.metrics.italic, substituted: false };
  }

  /** The best face for a family: itself, its substitutes, then the default family, then anything registered. */
  resolve(family: string, bold: boolean, italic: boolean): FontMatch | undefined {
    const own = family ? this.find(family, bold, italic) : undefined;
    if (own) return own;
    for (const sub of SUBSTITUTES[key(family)] ?? []) {
      const m = this.find(sub, bold, italic);
      if (m) return { ...m, substituted: true };
    }
    const d = this.find(this.defaultFamily, bold, italic);
    if (d) return { ...d, substituted: true };
    const any = this.order[0];
    return any ? { face: any, syntheticBold: bold && !any.metrics.bold, syntheticItalic: italic && !any.metrics.italic, substituted: true } : undefined;
  }

  /** A face with a glyph for `cp`, preferring the given families in order, for characters the run's font lacks. */
  fallback(cp: number, bold: boolean, italic: boolean, preferred: string[] = []): FontMatch | undefined {
    for (const fam of preferred) {
      const m = this.resolve(fam, bold, italic);
      if (m?.face.hasGlyph(cp)) return m;
    }
    for (const face of this.order) {
      if (face.hasGlyph(cp)) {
        const m = this.find(face.family, bold, italic);
        const hit = m?.face.hasGlyph(cp) ? m : { face, syntheticBold: false, syntheticItalic: false, substituted: true };
        return { ...hit, substituted: true };
      }
    }
    return undefined;
  }
}
