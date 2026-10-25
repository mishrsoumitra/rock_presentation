// packages/text/src/font.ts  -- one font face: HarfBuzz for shaping and outlines, metrics from the font's own tables
import * as hb from 'harfbuzzjs';

export interface FontMetrics {
  upem: number;
  /** Distance from baseline to the top of the line, positive, in font units. */
  ascent: number;
  /** Distance from baseline to the bottom of the line, positive, in font units. */
  descent: number;
  lineGap: number;
  capHeight: number;
  xHeight: number;
  underlinePosition: number;           // font units, negative below the baseline
  underlineThickness: number;
  strikeoutPosition: number;           // font units above the baseline
  strikeoutSize: number;
  weight: number;
  italic: boolean;
  bold: boolean;
}

export type PathOp = { op: 'M' | 'L'; x: number; y: number } | { op: 'Q'; x1: number; y1: number; x: number; y: number } | { op: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number } | { op: 'Z' };

const u16 = (b: Uint8Array, o: number) => ((b[o] ?? 0) << 8) | (b[o + 1] ?? 0);
const i16 = (b: Uint8Array, o: number) => {
  const v = u16(b, o);
  return v & 0x8000 ? v - 0x10000 : v;
};

/**
 * The line metrics PowerPoint lays text out with. PowerPoint is a Windows-first application: GDI and DirectWrite
 * place lines by the OS/2 usWinAscent and usWinDescent of a font, so those are used when present, else hhea.
 * This is the single place to change if the plan's text-metrics oracle says otherwise.
 */
function lineMetrics(os2: Uint8Array | undefined, hhea: Uint8Array | undefined): { ascent: number; descent: number; lineGap: number } {
  if (os2 && os2.length >= 78) {
    const winAscent = u16(os2, 74);
    const winDescent = u16(os2, 76);
    if (winAscent + winDescent > 0) return { ascent: winAscent, descent: winDescent, lineGap: 0 };
  }
  if (hhea && hhea.length >= 10) return { ascent: i16(hhea, 4), descent: -i16(hhea, 6), lineGap: i16(hhea, 8) };
  return { ascent: 800, descent: 200, lineGap: 0 };
}

export class FontFace {
  readonly family: string;
  readonly subfamily: string;
  readonly metrics: FontMetrics;
  readonly hbFace: hb.Face;
  readonly hbFont: hb.Font;
  private readonly outlines = new Map<number, PathOp[]>();
  private readonly coverage = new Map<number, boolean>();

  /** `bytes` is a TrueType or OpenType font (or a collection, with `index`). Throws on data HarfBuzz cannot read. */
  constructor(bytes: Uint8Array, index = 0, familyOverride?: string) {
    this.hbFace = new hb.Face(new hb.Blob(bytes), index);
    if (!this.hbFace.upem) throw new Error('not a font HarfBuzz can read');
    this.hbFont = new hb.Font(this.hbFace);
    const upem = this.hbFace.upem;
    this.hbFont.setScale(upem, upem);
    const os2 = this.hbFace.referenceTable('OS/2');
    const hhea = this.hbFace.referenceTable('hhea');
    const head = this.hbFace.referenceTable('head');
    const post = this.hbFace.referenceTable('post');
    const lm = lineMetrics(os2, hhea);
    const fsSelection = os2 && os2.length >= 64 ? u16(os2, 62) : 0;
    const macStyle = head && head.length >= 46 ? u16(head, 44) : 0;
    const weight = os2 && os2.length >= 6 ? u16(os2, 4) : 400;
    this.metrics = {
      upem, ...lm,
      capHeight: os2 && os2.length >= 90 ? i16(os2, 88) : Math.round(lm.ascent * 0.7),
      xHeight: os2 && os2.length >= 88 ? i16(os2, 86) : Math.round(lm.ascent * 0.5),
      underlinePosition: post && post.length >= 12 ? i16(post, 8) : -Math.round(upem * 0.1),
      underlineThickness: post && post.length >= 12 ? Math.max(1, i16(post, 10)) : Math.round(upem * 0.05),
      strikeoutPosition: os2 && os2.length >= 30 ? i16(os2, 28) : Math.round(lm.ascent * 0.3),
      strikeoutSize: os2 && os2.length >= 30 ? Math.max(1, i16(os2, 26)) : Math.round(upem * 0.05),
      weight,
      bold: (fsSelection & 0x20) !== 0 || (macStyle & 1) !== 0 || weight >= 600,
      italic: (fsSelection & 0x01) !== 0 || (macStyle & 2) !== 0,
    };
    const name = (id: number) => {
      try {
        return this.hbFace.getName(id, 'en') || '';
      } catch {
        return '';
      }
    };
    this.family = familyOverride ?? (name(16) || name(1));
    this.subfamily = name(17) || name(2);
  }

  hasGlyph(cp: number): boolean {
    let has = this.coverage.get(cp);
    if (has === undefined) {
      const g = this.hbFont.nominalGlyph(cp);
      has = g !== undefined && g !== 0;
      this.coverage.set(cp, has);
    }
    return has;
  }

  /** Glyph outline in font units, y up, cached. */
  outline(glyph: number): PathOp[] {
    let ops = this.outlines.get(glyph);
    if (!ops) {
      ops = [];
      for (const c of this.hbFont.glyphToJson(glyph)) {
        const v = c.values;
        switch (c.type) {
          case 'M': ops.push({ op: 'M', x: v[0]!, y: v[1]! }); break;
          case 'L': ops.push({ op: 'L', x: v[0]!, y: v[1]! }); break;
          case 'Q': ops.push({ op: 'Q', x1: v[0]!, y1: v[1]!, x: v[2]!, y: v[3]! }); break;
          case 'C': ops.push({ op: 'C', x1: v[0]!, y1: v[1]!, x2: v[2]!, y2: v[3]!, x: v[4]!, y: v[5]! }); break;
          case 'Z': ops.push({ op: 'Z' }); break;
        }
      }
      this.outlines.set(glyph, ops);
    }
    return ops;
  }
}
