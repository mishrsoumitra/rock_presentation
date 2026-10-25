// Packages and hand-built ZIP archives for the package tests.
import { deflateSync, strToU8, zipSync } from 'fflate';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const REL_TYPE = {
  officeDocument: `${REL}/officeDocument`,
  slide: `${REL}/slide`,
  image: `${REL}/image`,
  hyperlink: `${REL}/hyperlink`,
  notesSlide: `${REL}/notesSlide`,
};
export const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52]);

/** A small presentation-shaped package: enough structure for the OPC layer, not a deck PowerPoint would open. */
export function deckFiles(): Record<string, string | Uint8Array> {
  return {
    '[Content_Types].xml': DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '</Types>',
    '_rels/.rels': DECL +
      `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL_TYPE.officeDocument}" Target="ppt/presentation.xml"/></Relationships>`,
    'ppt/presentation.xml': DECL +
      `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}" ` +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
      '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': DECL +
      `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId2" Type="${REL_TYPE.slide}" Target="slides/slide1.xml"/></Relationships>`,
    'ppt/slides/slide1.xml': DECL +
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}" ` +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' +
      '<p:pic><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>',
    'ppt/slides/_rels/slide1.xml.rels': DECL +
      `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId2" Type="${REL_TYPE.image}" Target="../media/image1.png"/>` +
      `<Relationship Id="rId3" Type="${REL_TYPE.hyperlink}" Target="https://example.com/a?b=1&amp;c=2" TargetMode="External"/></Relationships>`,
    'ppt/media/image1.png': PNG_BYTES,
  };
}

export function zipFiles(files: Record<string, string | Uint8Array>): Uint8Array {
  const input: Record<string, Uint8Array> = {};
  for (const [name, content] of Object.entries(files)) input[name] = typeof content === 'string' ? strToU8(content) : content;
  return zipSync(input);
}

export const deck = () => zipFiles(deckFiles());

// ---------------------------------------------------------------------------------------------------------------
// Raw ZIP writer for hostile archives. Every size, count and method can be set to a value a real writer never uses.

export interface RawEntry {
  name: string;
  content?: Uint8Array;        // uncompressed content; deflated when method is 8
  method?: number;             // 0 stored, 8 deflate (default), anything else to test rejection
  declaredSize?: number;       // uncompressed size claimed in the central directory
  zip64?: boolean;             // sizes as 0xFFFFFFFF plus a ZIP64 extra field carrying the real values
  zip64ExtraMissing?: boolean; // sizes as 0xFFFFFFFF and no ZIP64 extra field: the input that hung fflate 0.8.2
}

class Writer {
  private chunks: number[] = [];
  get length() { return this.chunks.length; }
  u16(v: number) { this.chunks.push(v & 0xff, (v >>> 8) & 0xff); }
  u32(v: number) { this.u16(v & 0xffff); this.u16(Math.floor(v / 0x10000) & 0xffff); }
  u64(v: number) { this.u32(v % 0x100000000); this.u32(Math.floor(v / 0x100000000)); }
  bytes(b: Uint8Array) { for (const x of b) this.chunks.push(x); }
  out() { return new Uint8Array(this.chunks); }
}

/** Build an archive; `zip64Locator` adds a ZIP64 end-of-central-directory record and locator, and `claimedEntries` lies about the count. */
export function rawZip(entries: RawEntry[], options: { zip64Locator?: boolean; claimedEntries?: number } = {}): Uint8Array {
  const w = new Writer();
  const central: { e: RawEntry; data: Uint8Array; offset: number; name: Uint8Array }[] = [];
  for (const e of entries) {
    const content = e.content ?? new Uint8Array(0);
    const method = e.method ?? 8;
    const data = method === 8 ? deflateSync(content) : content;
    const name = strToU8(e.name);
    const offset = w.length;
    w.u32(0x04034b50); w.u16(20); w.u16(0x0800); w.u16(method); w.u16(0); w.u16(0x21);
    w.u32(0); w.u32(data.length); w.u32(content.length); w.u16(name.length); w.u16(0);
    w.bytes(name); w.bytes(data);
    central.push({ e, data, offset, name });
  }
  const cdStart = w.length;
  for (const { e, data, offset, name } of central) {
    const declared = e.declaredSize ?? (e.content?.length ?? 0);
    const big = e.zip64 || e.zip64ExtraMissing;
    const extraLength = e.zip64 ? 20 : 0;
    w.u32(0x02014b50); w.u16(45); w.u16(45); w.u16(0x0800); w.u16(e.method ?? 8); w.u16(0); w.u16(0x21);
    w.u32(0); w.u32(big ? 0xffffffff : data.length); w.u32(big ? 0xffffffff : declared);
    w.u16(name.length); w.u16(extraLength); w.u16(0); w.u16(0); w.u16(0); w.u32(0); w.u32(offset);
    w.bytes(name);
    if (e.zip64) { w.u16(0x0001); w.u16(16); w.u64(declared); w.u64(data.length); }
  }
  const cdSize = w.length - cdStart;
  const count = options.claimedEntries ?? entries.length;
  if (options.zip64Locator) {
    const recordAt = w.length;
    w.u32(0x06064b50); w.u64(44); w.u16(45); w.u16(45); w.u32(0); w.u32(0);
    w.u64(count); w.u64(count); w.u64(cdSize); w.u64(cdStart);
    w.u32(0x07064b50); w.u32(0); w.u64(recordAt); w.u32(1);
  }
  // With ZIP64 records present, the classic end record carries the "see ZIP64" markers, as real ZIP64 writers do
  // (fflate 0.8.2 only switched to ZIP64 parsing when it saw them).
  const z64 = options.zip64Locator === true;
  w.u32(0x06054b50); w.u16(0); w.u16(0);
  w.u16(z64 ? 0xffff : Math.min(count, 0xffff)); w.u16(z64 ? 0xffff : Math.min(count, 0xffff));
  w.u32(z64 ? 0xffffffff : cdSize); w.u32(z64 ? 0xffffffff : cdStart); w.u16(0);
  return w.out();
}

/** A compound file with one directory entry per name, at the 128-byte slots MS-CFB puts them in. */
export function fakeCompoundFile(streamNames: string[]): Uint8Array {
  const out = new Uint8Array(512 + 512);
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  streamNames.forEach((name, i) => {
    const at = 512 + i * 128;
    for (let k = 0; k < name.length; k++) out[at + 2 * k] = name.charCodeAt(k);
    const len = (name.length + 1) * 2;
    out[at + 64] = len & 0xff;
    out[at + 65] = len >> 8;
    out[at + 66] = 2;
  });
  return out;
}
