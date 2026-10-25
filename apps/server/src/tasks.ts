// apps/server/src/tasks.ts  -- what worker threads do with untrusted bytes: check a deck, round-trip it, compare PNGs
//
// Plain functions, so tests can call them directly; src/worker.ts runs them inside the pool's threads, where a
// resource limit or the deadline can stop them without touching the server. Every refusal is a TaskError, an
// OpcError or an OfficeCryptoError with a stable `code`, which crosses the thread boundary as { code, message }.
import { Buffer } from 'node:buffer';
import pixelmatch from 'pixelmatch';
import pngjs from 'pngjs';
import { assertSafeContainer, isEncryptedPackage, OfficeCryptoError, type ContainerLimits } from '@pptx/crypto';
import { CONTENT_TYPE, PresentationDocument } from '@pptx/model';
import { OpcError, OpcPackage, readZip, type PackageLimits, type ZipLimits } from '@pptx/opc';
import { canonical, canonicalDiff, decodeXmlBytes, parseXml } from '@pptx/xml';

const { PNG } = pngjs;

export class TaskError extends Error {
  readonly code: string;
  constructor(code: string, message: string) {
    super(message);
    this.name = 'TaskError';
    this.code = code;
  }
}

/** The media type and extension a deck is downloaded with, by the content type of its main part. */
export const PACKAGE_TYPES: Record<string, { type: string; ext: string }> = {
  [CONTENT_TYPE.presentation]: { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', ext: 'pptx' },
  [CONTENT_TYPE.slideshow]: { type: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow', ext: 'ppsx' },
  [CONTENT_TYPE.template]: { type: 'application/vnd.openxmlformats-officedocument.presentationml.template', ext: 'potx' },
  [CONTENT_TYPE.macroPresentation]: { type: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12', ext: 'pptm' },
  [CONTENT_TYPE.macroSlideshow]: { type: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12', ext: 'ppsm' },
  [CONTENT_TYPE.macroTemplate]: { type: 'application/vnd.ms-powerpoint.template.macroEnabled.12', ext: 'potm' },
};
export const PPTX_TYPE = PACKAGE_TYPES[CONTENT_TYPE.presentation]!.type;

export interface DeckCheck {
  encrypted: boolean;
  slideCount: number | null;
  slideSize: { cx: number; cy: number } | null;
  contentType: string;
  macros: boolean;
  diagnostics: string[];
}

export interface ValidateInput { bytes: Uint8Array; limits: PackageLimits; container: ContainerLimits; }

/**
 * Parse a part to prove it well-formed and within the XML limits, as OpcPackage.xml would, but without keeping the
 * tree, so that checking a deck of many slides needs the memory of its largest slide rather than of all of them.
 */
function checkXml(pkg: OpcPackage, name: string, maxBytes: number): void {
  const part = pkg.getPart(name)!;
  if (part.bytes.length > maxBytes) {
    throw new OpcError('limit', `${part.name} is ${part.bytes.length} bytes, over the ${maxBytes}-byte limit for XML parts`);
  }
  try {
    parseXml(decodeXmlBytes(part.bytes).text);
  } catch (e) {
    throw new OpcError('invalid-part', `${part.name}: ${e instanceof Error ? e.message : String(e)}`, { cause: e });
  }
}

/**
 * Check an upload. A password-protected package (a compound file with an EncryptedPackage stream) only has its
 * container checked: the server never sees passwords, so it is stored as it came and only a client can open it.
 * Anything else is opened with the server's ZIP limits, which OpcPackage checks before inflating anything, must be
 * a presentation, and has every slide part parsed, which is where the XML limits (depth, part size) apply.
 */
export function validateDeck({ bytes, limits, container }: ValidateInput): DeckCheck {
  if (isEncryptedPackage(bytes)) {
    assertSafeContainer(bytes, container);
    return { encrypted: true, slideCount: null, slideSize: null, contentType: PPTX_TYPE, macros: false, diagnostics: [] };
  }
  let pkg: OpcPackage;
  try {
    pkg = OpcPackage.open(bytes, { limits });
  } catch (e) {
    // encryption streams without the EncryptedPackage one checked above: not something a password could open
    if (!(e instanceof OpcError) || e.code !== 'encrypted') throw e;
    throw new TaskError('malformed', 'This compound file has encryption information but no encrypted package.');
  }
  let doc: PresentationDocument;
  let slides: string[];
  try {
    doc = new PresentationDocument(pkg);
    slides = doc.slideParts();
  } catch (e) {
    if (e instanceof OpcError) throw e;
    throw new TaskError('not-a-presentation', `This Open XML package is not a presentation: ${e instanceof Error ? e.message : String(e)}`);
  }
  const main = pkg.getPart(doc.mainPart)!.contentType;
  const type = PACKAGE_TYPES[main];
  if (!type) throw new TaskError('not-a-presentation', `This Open XML package is not a presentation: its main part is ${main || 'untyped'}.`);
  for (const part of slides) checkXml(pkg, part, limits.maxXmlPartBytes);
  const { cx, cy } = doc.presentation.slideSize;
  return {
    encrypted: false,
    slideCount: slides.length,
    slideSize: { cx, cy },
    contentType: type.type,
    macros: pkg.partNames().some(n => /\/vbaProject\.bin$/i.test(n) || pkg.getPart(n)?.contentType === 'application/vnd.ms-office.vbaProject'),
    diagnostics: pkg.diagnostics.slice(0, 20).map(d => d.slice(0, 200)),
  };
}

// ------------------------------------------------------------------------------------------------- round trip

export interface RoundtripInput { bytes: Uint8Array; limits: PackageLimits; }

export interface RoundtripReport {
  lossless: boolean;               // open -> save kept every entry, every XML part in canonical form, every other part's bytes
  reserializedLossless: boolean;   // the same with every XML part written from its parsed tree, as an edit would
  parts: number;
  xmlParts: number;
  unparseableXmlParts: number;     // kept as bytes
  differences: string[];
  diagnostics: string[];
}

const isXmlEntry = (name: string) => /\.(xml|rels)$/i.test(name);
const sameBytes = (x: Uint8Array, y: Uint8Array) => x.length === y.length && Buffer.compare(x, y) === 0;

/** The archive's file entries by name, read with the same limits as an upload. */
function entries(zip: Uint8Array, limits: ZipLimits): Map<string, Uint8Array> {
  return new Map(readZip(zip, limits).map(e => [e.name, e.data]));
}

function canonicalOf(bytes: Uint8Array, maxBytes: number): string | null {
  if (bytes.length > maxBytes) return null;
  try {
    return canonical(parseXml(decodeXmlBytes(bytes).text));
  } catch {
    return null;   // not well-formed: such a part must come back byte for byte
  }
}

/**
 * Every way `saved` differs from `original` as PowerPoint would read them: the definition of tools/corpus/compare.ts,
 * except that both archives are read through OpcPackage's limit-checked reader rather than an unbounded unzip.
 */
function packageDifferences(original: Map<string, Uint8Array>, saved: Map<string, Uint8Array>, maxXml: number): string[] {
  const problems: string[] = [];
  for (const n of original.keys()) if (!saved.has(n)) problems.push(`missing after save: ${n}`);
  for (const n of saved.keys()) if (!original.has(n)) problems.push(`added by save: ${n}`);
  for (const [n, x] of original) {
    const y = saved.get(n);
    if (!y) continue;
    const cx = isXmlEntry(n) ? canonicalOf(x, maxXml) : null;
    if (cx !== null) {
      const cy = canonicalOf(y, maxXml);
      const diff = cy === null ? 'no longer parses' : canonicalDiff(cx, cy);
      if (diff) problems.push(`${n}: ${diff}`);
    } else if (!sameBytes(x, y)) {
      problems.push(`${n}: bytes differ (${x.length} -> ${y.length})`);
    }
  }
  return problems;
}

/** Open and re-save a deck with OpcPackage, as tools/corpus/roundtrip.ts does for the corpus, and report the losses. */
export function roundtripDeck({ bytes, limits }: RoundtripInput): RoundtripReport {
  if (isEncryptedPackage(bytes)) throw new TaskError('encrypted', 'This presentation is password-protected; the server cannot open it.');
  const pkg = OpcPackage.open(bytes, { limits });
  const original = entries(bytes, limits);
  const differences = packageDifferences(original, entries(pkg.save(), limits), limits.maxXmlPartBytes);
  let xmlParts = 0;
  let unparseable = 0;
  for (const part of pkg.partNames()) {
    if (!isXmlEntry(part)) continue;
    xmlParts++;
    try {
      pkg.xml(part);
      pkg.markDirty(part);
    } catch {
      unparseable++;
    }
  }
  const reserialized = packageDifferences(original, entries(pkg.save(), limits), limits.maxXmlPartBytes);
  return {
    lossless: differences.length === 0,
    reserializedLossless: reserialized.length === 0,
    parts: pkg.partNames().length,
    xmlParts,
    unparseableXmlParts: unparseable,
    differences: [...differences, ...reserialized.map(d => `re-serialized: ${d}`)].slice(0, 50).map(d => d.slice(0, 500)),
    diagnostics: pkg.diagnostics.slice(0, 50).map(d => d.slice(0, 200)),
  };
}

// ------------------------------------------------------------------------------------------------- PNG compare

export interface CompareInput {
  expected: Uint8Array;
  actual: Uint8Array;
  threshold: number;      // pixelmatch's per-pixel colour threshold, 0..1
  includeAA: boolean;     // count anti-aliased pixels as differences
  diff: boolean;          // also return the diff image
  maxPixels: number;
}

export interface CompareResult { width: number; height: number; mismatched: number; ratio: number; diff?: Uint8Array; }

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

/**
 * Decode a PNG after checking its header: pngjs allocates width x height x 4 bytes up front, and inflates an
 * interlaced image with no cap on the output, so huge dimensions and interlacing are refused before it runs.
 */
function decodePng(bytes: Uint8Array, which: string, maxPixels: number): { width: number; height: number; data: Uint8Array } {
  const ihdr = bytes.length >= 33 && PNG_SIGNATURE.every((b, i) => bytes[i] === b) && String.fromCharCode(...bytes.subarray(12, 16)) === 'IHDR';
  if (!ihdr) throw new TaskError('not-a-png', `The ${which} image is not a PNG.`);
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = dv.getUint32(16);
  const height = dv.getUint32(20);
  if (width === 0 || height === 0 || width * height > maxPixels) {
    throw new TaskError('limit', `The ${which} image is ${width} x ${height}; the limit is ${maxPixels} pixels.`);
  }
  if (bytes[28] !== 0) throw new TaskError('unsupported-png', `The ${which} image is interlaced, which is not supported.`);
  try {
    const png = PNG.sync.read(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength));
    return { width: png.width, height: png.height, data: png.data };
  } catch (e) {
    throw new TaskError('invalid-png', `The ${which} image could not be decoded: ${e instanceof Error ? e.message : String(e)}`);
  }
}

/** Oracle O3: pixelmatch's count of differing pixels between two renders of the same slide, and the diff image. */
export function comparePngs(input: CompareInput): CompareResult {
  const a = decodePng(input.expected, 'expected', input.maxPixels);
  const b = decodePng(input.actual, 'actual', input.maxPixels);
  if (a.width !== b.width || a.height !== b.height) {
    throw new TaskError('size-mismatch', `The images differ in size: expected ${a.width} x ${a.height}, actual ${b.width} x ${b.height}.`);
  }
  const out = input.diff ? new PNG({ width: a.width, height: a.height }) : undefined;
  const mismatched = pixelmatch(a.data, b.data, out?.data, a.width, a.height, { threshold: input.threshold, includeAA: input.includeAA });
  const result: CompareResult = { width: a.width, height: a.height, mismatched, ratio: mismatched / (a.width * a.height) };
  if (out) result.diff = new Uint8Array(PNG.sync.write(out));
  return result;
}

// ------------------------------------------------------------------------------------------------- dispatch

/** Stand-ins for a parse that never finishes, one that eats the heap, and a thread that dies. Tests only. */
function testTask(task: string): never {
  if (task === 'test:spin') for (;;) { /* until the deadline terminates the thread */ }
  if (task === 'test:alloc') {
    const hog: number[][] = [];
    for (;;) hog.push(new Array<number>(1 << 16).fill(Math.random()));
  }
  process.exit(70);
}

export interface Tasks {
  validate: { input: ValidateInput; output: DeckCheck };
  roundtrip: { input: RoundtripInput; output: RoundtripReport };
  compare: { input: CompareInput; output: CompareResult };
  'test:spin': { input: null; output: never };
  'test:alloc': { input: null; output: never };
  'test:crash': { input: null; output: never };
}
export type TaskName = keyof Tasks;

/** Run a task by name. Results go back to the pool; their large byte arrays are listed to be transferred, not copied. */
export function runTask(task: string, input: unknown, testTasks: boolean): { result: unknown; transfer: ArrayBuffer[] } {
  switch (task) {
    case 'validate': return { result: validateDeck(input as ValidateInput), transfer: [] };
    case 'roundtrip': return { result: roundtripDeck(input as RoundtripInput), transfer: [] };
    case 'compare': {
      const result = comparePngs(input as CompareInput);
      return { result, transfer: result.diff ? [result.diff.buffer as ArrayBuffer] : [] };
    }
    default:
      if (testTasks && task.startsWith('test:')) testTask(task);
      throw new TaskError('unknown-task', `no task ${task}`);
  }
}

/** An error as it crosses the thread boundary: the code of a known refusal, else 'unexpected'. */
export function describeError(e: unknown): { code: string; message: string } {
  const message = (e instanceof Error ? e.message : String(e)).slice(0, 1000);
  if (e instanceof OpcError || e instanceof OfficeCryptoError || e instanceof TaskError) return { code: e.code, message };
  return { code: 'unexpected', message };
}
