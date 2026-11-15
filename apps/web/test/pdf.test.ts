// The PDF the browser exports: a valid file, one page per picture, each the size given, with every object where
// the cross-reference table says it is.
import { describe, expect, test } from 'vitest';
import { pdfOfImages, type PdfPage } from '../src/engine/pdf.ts';

/** Bytes that stand for a JPEG (the writer stores them as they are, never reading them). */
const jpeg = (n: number, fill: number) => new Uint8Array(n).fill(fill);
const page = (n: number, fill: number, over: Partial<PdfPage> = {}): PdfPage => ({ jpeg: jpeg(n, fill), width: 1920, height: 1080, widthPt: 960, heightPt: 540, ...over });
const text = (pdf: Uint8Array) => Array.from(pdf, b => String.fromCharCode(b)).join('');

/** Read the cross-reference table: each object's offset, by number. */
function xref(pdf: Uint8Array): { offsets: number[]; size: number } {
  const s = text(pdf);
  const start = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(s)![1]);
  expect(s.slice(start, start + 5)).toBe('xref\n');
  const lines = s.slice(start).split('\n');
  const size = Number(lines[1]!.split(' ')[1]);
  const offsets = lines.slice(2, 2 + size).map(l => Number(l.slice(0, 10)));
  for (const l of lines.slice(2, 2 + size)) expect(l.length).toBe(19);         // each entry is 20 bytes with its line end
  return { offsets, size };
}

describe('a PDF of pictures', () => {
  test('starts and ends as a PDF, with one page per picture in order, each page the size given', () => {
    const pdf = pdfOfImages([page(100, 1), page(250, 2, { widthPt: 720, heightPt: 540, width: 1440 })], 'Quarterly review');
    const s = text(pdf);
    expect(s.startsWith('%PDF-1.4\n')).toBe(true);
    expect(s.endsWith('%%EOF\n')).toBe(true);
    expect(s.match(/\/Type \/Page /g)?.length).toBe(2);
    expect(s).toContain('/Type /Pages /Count 2 /Kids [4 0 R 7 0 R]');
    expect(s).toContain('/MediaBox [0 0 960 540]');
    expect(s).toContain('/MediaBox [0 0 720 540]');
    expect(s).toContain('q 960 0 0 540 0 0 cm /Im0 Do Q');
    expect(s).toContain('/Title (Quarterly review)');
    expect(s.indexOf('/MediaBox [0 0 960 540]')).toBeLessThan(s.indexOf('/MediaBox [0 0 720 540]'));
  });

  test('every object is exactly where the cross-reference table says', () => {
    const pdf = pdfOfImages([page(100, 1), page(250, 2), page(7, 3)]);
    const s = text(pdf);
    const { offsets, size } = xref(pdf);
    expect(size).toBe(4 + 3 * 3);
    expect(offsets[0]).toBe(0);
    for (let n = 1; n < size; n++) expect(s.slice(offsets[n]!, offsets[n]! + `${n} 0 obj\n`.length), `object ${n}`).toBe(`${n} 0 obj\n`);
    expect(s).toContain(`/Size ${size} /Root 1 0 R`);
  });

  test('each picture is stored byte for byte, with its length and pixel size declared', () => {
    const a = jpeg(300, 0xaa), b = jpeg(41, 0x0a);                              // the second is all line feeds: lengths, not delimiters, find the end
    const pdf = pdfOfImages([{ ...page(0, 0), jpeg: a, width: 800, height: 450 }, { ...page(0, 0), jpeg: b }]);
    const s = text(pdf);
    const { offsets } = xref(pdf);
    for (const [obj, bytes, w] of [[6, a, 800], [9, b, 1920]] as const) {
      const head = /^(\d+) 0 obj\n<< ([^>]*) >>\nstream\n/.exec(s.slice(offsets[obj]!))!;
      expect(head[2]).toContain(`/Width ${w} `);
      expect(head[2]).toContain('/Filter /DCTDecode');
      expect(head[2]).toContain(`/Length ${bytes.length}`);
      const from = offsets[obj]! + head[0].length;
      expect(Array.from(pdf.subarray(from, from + bytes.length))).toEqual(Array.from(bytes));
      expect(s.slice(from + bytes.length, from + bytes.length + 17)).toBe('\nendstream\nendobj');
    }
  });

  test('a title with brackets or a backslash is escaped; one outside Latin letters is written as UTF-16', () => {
    expect(text(pdfOfImages([page(1, 1)], 'Costs (2026) \\ plan'))).toContain('/Title (Costs \\(2026\\) \\\\ plan)');
    expect(text(pdfOfImages([page(1, 1)], 'Résumé'))).toContain('/Title <FEFF005200E900730075006D00E9>');
    expect(text(pdfOfImages([page(1, 1)]))).toContain('/Title ()');
  });

  test('sizes that are not whole numbers are written without exponents; no pages is still a valid, empty PDF', () => {
    const s = text(pdfOfImages([page(1, 1, { widthPt: 959.999999, heightPt: 0.00001234 })]));
    expect(s).toContain('/MediaBox [0 0 960 0]');
    const empty = pdfOfImages([]);
    expect(text(empty)).toContain('/Count 0 /Kids []');
    expect(xref(empty).size).toBe(4);
  });
});
