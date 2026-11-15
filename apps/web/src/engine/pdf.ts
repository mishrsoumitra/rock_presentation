// apps/web/src/engine/pdf.ts  -- a PDF of pictures: one JPEG per page, each page the size of its slide

/** One page: a JPEG of the slide, its pixel size, and the page size in points (1/72 inch). */
export interface PdfPage { jpeg: Uint8Array; width: number; height: number; widthPt: number; heightPt: number; }

const enc = new TextEncoder();
/** A number as PDF writes one: no exponent, at most three decimals, no trailing zeros. */
const num = (n: number) => String(Math.round(n * 1000) / 1000);
/** Text as a PDF string: in parentheses, with its own delimiters escaped; characters outside Latin-1 as UTF-16 hex. */
function pdfString(s: string): string {
  if (/^[\x20-\x7e]*$/.test(s)) return `(${s.replace(/[\\()]/g, '\\$&')})`;
  let hex = 'FEFF';
  for (let i = 0; i < s.length; i++) hex += s.charCodeAt(i).toString(16).padStart(4, '0').toUpperCase();
  return `<${hex}>`;
}

/**
 * A PDF whose pages are pictures. Each page is exactly its picture: a JPEG stored as it is (PDF reads JPEG itself),
 * stretched over a page of the size given. There is no text in it to select or search: it is what the slides look
 * like, for printing and sending, like the PDF the server makes.
 */
export function pdfOfImages(pages: PdfPage[], title = ''): Uint8Array {
  const chunks: Uint8Array[] = [];
  const offsets: number[] = [];
  let at = 0;
  const put = (data: string | Uint8Array) => {
    const bytes = typeof data === 'string' ? enc.encode(data) : data;
    chunks.push(bytes);
    at += bytes.length;
  };
  /** Start object `n` (numbered from 1) here. */
  const obj = (n: number, body: string) => { offsets[n] = at; put(`${n} 0 obj\n${body}\nendobj\n`); };

  put('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');
  // 1 the catalog, 2 the page tree, 3 the document information, then three objects per page: page, contents, image
  const pageObj = (i: number) => 4 + i * 3;
  obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
  obj(2, `<< /Type /Pages /Count ${pages.length} /Kids [${pages.map((_, i) => `${pageObj(i)} 0 R`).join(' ')}] >>`);
  obj(3, `<< /Producer (Presentations) /Title ${pdfString(title)} >>`);
  pages.forEach((p, i) => {
    const n = pageObj(i);
    obj(n, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${num(p.widthPt)} ${num(p.heightPt)}] /Resources << /XObject << /Im0 ${n + 2} 0 R >> >> /Contents ${n + 1} 0 R >>`);
    const draw = `q ${num(p.widthPt)} 0 0 ${num(p.heightPt)} 0 0 cm /Im0 Do Q`;
    obj(n + 1, `<< /Length ${draw.length} >>\nstream\n${draw}\nendstream`);
    offsets[n + 2] = at;
    put(`${n + 2} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${p.width} /Height ${p.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`);
    put(p.jpeg);
    put('\nendstream\nendobj\n');
  });
  const count = 4 + pages.length * 3;
  const xref = at;
  put(`xref\n0 ${count}\n0000000000 65535 f \n`);
  for (let n = 1; n < count; n++) put(`${String(offsets[n]).padStart(10, '0')} 00000 n \n`);
  put(`trailer\n<< /Size ${count} /Root 1 0 R /Info 3 0 R >>\nstartxref\n${xref}\n%%EOF\n`);

  const out = new Uint8Array(at);
  let o = 0;
  for (const c of chunks) { out.set(c, o); o += c.length; }
  return out;
}
