// packages/opc/src/zip.ts  -- the ZIP container (ECMA-376 Part 2, Annex B) and the checks that run before inflation
import { unzipSync, zipSync, type Zippable } from 'fflate';
import { OpcError } from './errors.ts';

export interface ZipLimits {
  maxEntries: number;      // entries in the central directory, folders included
  maxEntryBytes: number;   // declared inflated size of one entry
  maxTotalBytes: number;   // declared inflated size of all entries together
}
export const DEFAULT_LIMITS: ZipLimits = { maxEntries: 20_000, maxEntryBytes: 1 << 30, maxTotalBytes: 2 * (1 << 30) };

export interface ZipEntry { name: string; data: Uint8Array; declaredSize: number; }

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function sniff(bytes: Uint8Array): 'zip' | 'cfb' | 'unknown' {
  // local file header, or the end-of-central-directory record of an empty archive
  if (bytes[0] === 0x50 && bytes[1] === 0x4b && ((bytes[2] === 3 && bytes[3] === 4) || (bytes[2] === 5 && bytes[3] === 6))) return 'zip';
  if (CFB_SIGNATURE.every((b, i) => bytes[i] === b)) return 'cfb';
  return 'unknown';
}

/** True when the compound file has a directory entry with this name. Directory entries are 128 bytes, start with
 *  the UTF-16LE name, and always sit at 128-byte offsets after the 512-byte header (MS-CFB 2.6), so only those
 *  offsets are checked and file content cannot produce a false match. */
function cfbHasEntry(bytes: Uint8Array, name: string): boolean {
  const width = (name.length + 1) * 2;
  outer: for (let at = 512; at + 128 <= bytes.length; at += 128) {
    for (let k = 0; k < name.length; k++) {
      if (bytes[at + 2 * k] !== name.charCodeAt(k) || bytes[at + 2 * k + 1] !== 0) continue outer;
    }
    if (bytes[at + 2 * name.length] === 0 && bytes[at + 2 * name.length + 1] === 0 && (bytes[at + 64]! | (bytes[at + 65]! << 8)) === width) return true;
  }
  return false;
}

/** Explain why a compound file cannot be opened: a password-protected package, a legacy .ppt, or something else. */
export function compoundFileError(bytes: Uint8Array): OpcError {
  const encrypted = () => new OpcError('encrypted', 'This presentation is password-protected. Opening encrypted files is not supported yet.');
  if (cfbHasEntry(bytes, 'EncryptedPackage')) return encrypted();
  if (cfbHasEntry(bytes, 'PowerPoint Document')) {
    return new OpcError('legacy-ppt', 'This is a PowerPoint 97-2003 (.ppt) file. Only .pptx and the related Open XML formats are supported; save it as .pptx in PowerPoint first.');
  }
  if (cfbHasEntry(bytes, 'PP40')) {
    return new OpcError('legacy-ppt', 'This is a PowerPoint 4.0 file. Only .pptx and the related Open XML formats are supported.');
  }
  if (cfbHasEntry(bytes, 'EncryptionInfo')) return encrypted();
  return new OpcError('unsupported-container', 'This is an OLE compound file, not a PowerPoint Open XML package.');
}

/**
 * Read every entry of a ZIP archive, in central directory order. fflate inflates an entry as soon as the filter
 * accepts it, so a single pass could allocate earlier entries before a later one breaks the total limit. The
 * first pass therefore only walks the central directory (the filter declines every entry) and checks every limit;
 * the second pass inflates. Each entry is inflated into a buffer of exactly its declared size, so an entry that
 * lies about its size cannot allocate more than the limits allow.
 */
export function readZip(bytes: Uint8Array, limits: ZipLimits = DEFAULT_LIMITS): ZipEntry[] {
  let count = 0;
  let total = 0;
  const declared = new Map<string, number>();     // file entries in central directory order, with declared sizes
  const folders = new Set<string>();
  const unzip = (filter: Parameters<typeof unzipSync>[1]) => {
    try {
      return unzipSync(bytes, filter);
    } catch (e) {
      if (e instanceof OpcError) throw e;
      throw new OpcError('invalid-zip', `not a readable ZIP archive: ${e instanceof Error ? e.message : String(e)}`, { cause: e });
    }
  };

  unzip({
    filter: f => {
      if (++count > limits.maxEntries) throw new OpcError('limit', `more than ${limits.maxEntries} ZIP entries`);
      // fflate returns entries keyed by name, so a repeated name would silently replace the first one
      if (declared.has(f.name) || folders.has(f.name)) throw new OpcError('invalid-package', `duplicate ZIP entry ${f.name}`);
      if (f.name.endsWith('/')) {                                      // a folder entry, not a part
        folders.add(f.name);
        return false;
      }
      if (f.compression !== 0 && f.compression !== 8) {
        throw new OpcError('invalid-zip', `ZIP entry ${f.name} uses unsupported compression method ${f.compression}`);
      }
      if (f.originalSize > limits.maxEntryBytes) {
        throw new OpcError('limit', `ZIP entry ${f.name} declares ${f.originalSize} bytes inflated, over the ${limits.maxEntryBytes}-byte limit`);
      }
      total += f.originalSize;
      if (total > limits.maxTotalBytes) throw new OpcError('limit', `archive declares more than ${limits.maxTotalBytes} bytes inflated`);
      declared.set(f.name, f.originalSize);
      return false;
    },
  });

  const files = unzip({ filter: f => declared.has(f.name) });
  return [...declared].map(([name, declaredSize]) => ({ name, data: files[name]!, declaredSize }));
}

// DOS timestamps start in 1980. Every entry gets the same one, so saving the same package twice gives the same bytes.
const ZIP_EPOCH = new Date(1980, 0, 1, 0, 0, 0);

/**
 * Write entries in the given order. Entries marked `store` (already-compressed media) are stored rather than
 * deflated again. Sizes and CRCs go in the local headers (no data descriptors), as PowerPoint writes them.
 * One caveat of fflate's object-keyed input: a purely numeric root-level name would be written first.
 */
export function writeZip(entries: { name: string; data: Uint8Array; store: boolean }[]): Uint8Array {
  const files: Zippable = {};
  for (const e of entries) files[e.name] = [e.data, { level: e.store ? 0 : 6 }];
  return zipSync(files, { mtime: ZIP_EPOCH });
}
