// apps/server/test/helpers.ts  -- servers on temp directories, multipart bodies, and the decks and archives tests send
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { constants, crc32, deflateRawSync } from 'node:zlib';
import type { FastifyInstance, LightMyRequestResponse } from 'fastify';
import { buildServer } from '../src/app.ts';
import type { ConfigInput } from '../src/config.ts';

export const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../../..');

const temps: string[] = [];
/** A fresh directory under the OS temp dir (TMPDIR), removed by cleanup(). */
export function tempDir(prefix = 'pptx-server-test-'): string {
  const dir = mkdtempSync(join(tmpdir(), prefix));
  temps.push(dir);
  return dir;
}
export function cleanup(): void {
  for (const d of temps.splice(0)) rmSync(d, { recursive: true, force: true });
}

/** A server on a temp DATA_DIR: quiet, no rate limit, no Chromium and no editor unless asked for. */
export function testServer(input: ConfigInput = {}): Promise<FastifyInstance> {
  return buildServer({
    logLevel: 'silent',
    rateLimit: false,
    ...input,
    dataDir: input.dataDir ?? tempDir(),
    webDir: input.webDir ?? tempDir('pptx-no-web-'),
    workers: { size: 2, ...input.workers },
    render: { pages: 0, ...input.render },
  });
}

/** A multipart/form-data body, serialized by the platform's own FormData. */
export async function form(files: Record<string, { data: Uint8Array; filename?: string; type?: string }>, fields: Record<string, string> = {}) {
  const fd = new FormData();
  for (const [k, v] of Object.entries(fields)) fd.set(k, v);
  for (const [k, f] of Object.entries(files)) fd.set(k, new Blob([f.data as BlobPart], { type: f.type ?? 'application/octet-stream' }), f.filename ?? 'deck.pptx');
  const res = new Response(fd);
  return { payload: Buffer.from(await res.arrayBuffer()), headers: { 'content-type': res.headers.get('content-type')! } };
}

export interface UploadOptions {
  filename?: string;
  fields?: Record<string, string>;
  method?: 'POST' | 'PUT';
  url?: string;
  headers?: Record<string, string>;
}

/** Send `data` as the `file` part of a multipart upload (POST /api/decks unless told otherwise). */
export async function upload(app: FastifyInstance, data: Uint8Array, options: UploadOptions = {}): Promise<LightMyRequestResponse> {
  const body = await form({ file: { data, filename: options.filename ?? 'deck.pptx' } }, options.fields);
  return app.inject({
    method: options.method ?? 'POST', url: options.url ?? '/api/decks', payload: body.payload,
    headers: { ...body.headers, ...options.headers },
  });
}

/** A corpus file, or undefined when the corpus has not been fetched (pnpm corpus:fetch). */
export function corpusFile(rel: string): Uint8Array | undefined {
  const p = join(ROOT, 'corpus/external', rel);
  return existsSync(p) ? new Uint8Array(readFileSync(p)) : undefined;
}

export const readRepoFile = (rel: string) => new Uint8Array(readFileSync(join(ROOT, rel)));

// ------------------------------------------------------------------------------------------------------- ZIP

export interface ZipEntry {
  name: string;
  data?: Uint8Array | string;
  deflate?: boolean;              // deflate `data` (default: store it)
  deflated?: Uint8Array;          // content already deflated, for archives whose inflated form is too big to build
  size?: number;                  // the inflated size to declare, true or not
  zip64?: boolean;                // sizes as 0xFFFFFFFF plus a ZIP64 extra field carrying them
}

/**
 * Write a ZIP archive. Real CRCs and sizes unless an entry overrides them; `claimedEntries` lies in the end record,
 * `zip64End` adds the ZIP64 end record and locator. Laid out as packages/opc/test/fixtures.ts's rawZip.
 */
export function zip(entries: ZipEntry[], options: { claimedEntries?: number; zip64End?: boolean } = {}): Buffer {
  const parts: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const raw = typeof e.data === 'string' ? Buffer.from(e.data, 'utf8') : Buffer.from(e.data ?? new Uint8Array(0));
    const method = e.deflated || e.deflate ? 8 : 0;
    const body = e.deflated ? Buffer.from(e.deflated) : e.deflate ? deflateRawSync(raw) : raw;
    const size = e.size ?? raw.length;
    const crc = e.deflated ? 0 : crc32(raw);
    const name = Buffer.from(e.name, 'utf8');
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(0x21, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(Math.min(size, 0xffffffff), 22);
    local.writeUInt16LE(name.length, 26);
    parts.push(local, name, body);
    const extra = e.zip64 ? Buffer.alloc(20) : Buffer.alloc(0);
    if (e.zip64) {
      extra.writeUInt16LE(0x0001, 0);
      extra.writeUInt16LE(16, 2);
      extra.writeBigUInt64LE(BigInt(size), 4);
      extra.writeBigUInt64LE(BigInt(body.length), 12);
    }
    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(45, 4);
    cd.writeUInt16LE(45, 6);
    cd.writeUInt16LE(0x0800, 8);
    cd.writeUInt16LE(method, 10);
    cd.writeUInt16LE(0x21, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(e.zip64 ? 0xffffffff : body.length, 20);
    cd.writeUInt32LE(e.zip64 ? 0xffffffff : Math.min(size, 0xffffffff), 24);
    cd.writeUInt16LE(name.length, 28);
    cd.writeUInt16LE(extra.length, 30);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, name, extra);
    offset += 30 + name.length + body.length;
  }
  const cdStart = offset;
  const cdSize = central.reduce((n, b) => n + b.length, 0);
  const count = options.claimedEntries ?? entries.length;
  const tail: Buffer[] = [];
  if (options.zip64End) {
    const rec = Buffer.alloc(56 + 20);
    rec.writeUInt32LE(0x06064b50, 0);
    rec.writeBigUInt64LE(44n, 4);
    rec.writeUInt16LE(45, 12);
    rec.writeUInt16LE(45, 14);
    rec.writeBigUInt64LE(BigInt(count), 24);
    rec.writeBigUInt64LE(BigInt(count), 32);
    rec.writeBigUInt64LE(BigInt(cdSize), 40);
    rec.writeBigUInt64LE(BigInt(cdStart), 48);
    rec.writeUInt32LE(0x07064b50, 56);
    rec.writeBigUInt64LE(BigInt(cdStart + cdSize), 64);
    rec.writeUInt32LE(1, 72);
    tail.push(rec);
  }
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(options.zip64End ? 0xffff : Math.min(count, 0xffff), 8);
  end.writeUInt16LE(options.zip64End ? 0xffff : Math.min(count, 0xffff), 10);
  end.writeUInt32LE(options.zip64End ? 0xffffffff : cdSize, 12);
  end.writeUInt32LE(options.zip64End ? 0xffffffff : cdStart, 16);
  return Buffer.concat([...parts, ...central, ...tail, end]);
}

/**
 * Raw deflate data for `mib` MiB of zeros, built without ever holding them: one MiB compressed and sync-flushed
 * (byte-aligned, not final), repeated, then an empty final block. Every back-reference has distance 1, so the copies
 * chain into one valid stream. About 1 KB per MiB.
 */
export function zeroBomb(mib: number): Buffer {
  const chunk = deflateRawSync(Buffer.alloc(1 << 20), { finishFlush: constants.Z_SYNC_FLUSH });
  return Buffer.concat([...Array.from({ length: mib }, () => chunk), Buffer.from([0x03, 0x00])]);
}

// ------------------------------------------------------------------------------------------------------- decks

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESENTATION_MAIN = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';

/**
 * A presentation package with `slides` slides, enough for the server's checks (not for PowerPoint: no masters).
 * `note` goes into every slide, to make decks that differ. `presentationXml` replaces the main part, `slideXml`
 * every slide part.
 */
export function deck(slides = 1, options: { note?: string; presentationXml?: string; slideXml?: string; mainType?: string; extra?: ZipEntry[] } = {}): Buffer {
  const ids = Array.from({ length: slides }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join('');
  const entries: ZipEntry[] = [
    {
      name: '[Content_Types].xml',
      data: DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/ppt/presentation.xml" ContentType="${options.mainType ?? PRESENTATION_MAIN}"/>` +
        Array.from({ length: slides }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') +
        '</Types>',
    },
    { name: '_rels/.rels', data: DECL + `<Relationships xmlns="${PKG_R}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>` },
    {
      name: 'ppt/presentation.xml', deflate: true,
      data: options.presentationXml ?? DECL + `<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:sldIdLst>${ids}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    },
    {
      name: 'ppt/_rels/presentation.xml.rels',
      data: DECL + `<Relationships xmlns="${PKG_R}">` +
        Array.from({ length: slides }, (_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/slide" Target="slides/slide${i + 1}.xml"/>`).join('') + '</Relationships>',
    },
    ...Array.from({ length: slides }, (_, i): ZipEntry => ({
      name: `ppt/slides/slide${i + 1}.xml`,
      deflate: true,
      data: options.slideXml ?? DECL + `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><!-- ${options.note ?? ''} ${i + 1} --></p:spTree></p:cSld></p:sld>`,
    })),
    ...(options.extra ?? []),
  ];
  return zip(entries);
}

/** A compound file (MS-CFB, version 3) with these directory entries; `fatCycle` makes sector 2's chain loop. */
export function compoundFile(streams: string[], options: { fatCycle?: boolean } = {}): Uint8Array {
  const out = Buffer.alloc(512 + 3 * 512);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(out, 0);
  out.writeUInt16LE(0x3e, 24);
  out.writeUInt16LE(3, 26);
  out.writeUInt16LE(0xfffe, 28);
  out.writeUInt16LE(9, 30);
  out.writeUInt16LE(6, 32);
  out.writeUInt32LE(1, 44);                // one FAT sector
  out.writeUInt32LE(1, 48);                // the directory starts at sector 1
  out.writeUInt32LE(0x1000, 56);
  out.writeUInt32LE(0xfffffffe, 60);
  out.writeUInt32LE(0xfffffffe, 68);
  for (let i = 0; i < 109; i++) out.writeUInt32LE(i === 0 ? 0 : 0xffffffff, 76 + i * 4);
  const fat = 512;
  for (let i = 0; i < 128; i++) out.writeUInt32LE(0xffffffff, fat + i * 4);
  out.writeUInt32LE(0xfffffffd, fat);                                         // sector 0: the FAT itself
  out.writeUInt32LE(0xfffffffe, fat + 4);                                     // sector 1: the directory, one sector
  out.writeUInt32LE(options.fatCycle ? 2 : 0xfffffffe, fat + 8);              // sector 2: data, or a chain that loops
  streams.forEach((name, i) => {
    const at = 1024 + i * 128;
    for (let k = 0; k < name.length; k++) out.writeUInt16LE(name.charCodeAt(k), at + 2 * k);
    out.writeUInt16LE((name.length + 1) * 2, at + 64);
    out[at + 66] = i === 0 ? 5 : 2;
  });
  return new Uint8Array(out);
}
