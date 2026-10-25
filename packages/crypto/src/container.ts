// packages/crypto/src/container.ts
// The MS-CFB compound file that wraps an encrypted package: reading and writing it through the `cfb` package,
// with a bounded pre-parse guard. cfb's `get_sector_list` follows a stream's FAT chain with no cycle check, so a
// crafted FAT cycle on a >=4096-byte stream would loop forever; this module validates the header and proves the
// whole FAT is acyclic before handing the bytes to cfb. All downloaded/opened files are untrusted data.

import * as CFB from 'cfb';
import { OfficeCryptoError } from './errors.ts';

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const MAXREGSECT = 0xfffffffa; // values above this are reserved markers, not sector numbers

export interface ContainerLimits {
  maxBytes: number;    // the whole compound file
  maxSectors: number;  // bounds FAT/DIFAT traversal and the sector table cfb builds
}
export const DEFAULT_CONTAINER_LIMITS: ContainerLimits = { maxBytes: 256 * (1 << 20), maxSectors: 1 << 20 };

/** True when the bytes begin with the compound file signature (MS-CFB 2.2). */
export function isCompoundFile(bytes: Uint8Array): boolean {
  return bytes.length >= 8 && CFB_SIGNATURE.every((b, i) => bytes[i] === b);
}

/**
 * True when a compound file has a directory entry with this exact name. Directory entries are 128 bytes at
 * 128-byte offsets after the 512-byte header (MS-CFB 2.6), each starting with the UTF-16LE name then its byte
 * length at +64, so only those offsets are checked and stream content cannot forge a match. Mirrors the sniff in
 * packages/opc so `isEncryptedPackage` need not parse the whole container.
 */
export function cfbHasStream(bytes: Uint8Array, name: string): boolean {
  const width = (name.length + 1) * 2;
  outer: for (let at = 512; at + 128 <= bytes.length; at += 128) {
    for (let k = 0; k < name.length; k++) {
      if (bytes[at + 2 * k] !== name.charCodeAt(k) || bytes[at + 2 * k + 1] !== 0) continue outer;
    }
    if (bytes[at + 2 * name.length] === 0 && bytes[at + 2 * name.length + 1] === 0 && (bytes[at + 64]! | (bytes[at + 65]! << 8)) === width) return true;
  }
  return false;
}

/**
 * Validate the header and prove the FAT has no cycles, so cfb cannot loop. Throws OfficeCryptoError on anything
 * malformed or over a limit. Returns nothing; it is a gate run before CFB.read.
 */
export function assertSafeContainer(bytes: Uint8Array, limits: ContainerLimits): void {
  const bad = (m: string): never => { throw new OfficeCryptoError('malformed', m); };
  if (bytes.length > limits.maxBytes) throw new OfficeCryptoError('limit', `compound file is ${bytes.length} bytes, over the ${limits.maxBytes}-byte limit`);
  if (bytes.length < 512) bad(`compound file is only ${bytes.length} bytes`);
  if (!isCompoundFile(bytes)) bad('not a compound file (bad signature)');

  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const major = dv.getUint16(26, true);
  const sectorShift = dv.getUint16(30, true);
  if (major !== 3 && major !== 4) bad(`unsupported CFB major version ${major}`);
  if ((major === 3 && sectorShift !== 9) || (major === 4 && sectorShift !== 12)) bad(`sector shift ${sectorShift} does not match version ${major}`);
  if (dv.getUint16(32, true) !== 6) bad('mini sector shift is not 6');
  if (dv.getUint32(56, true) !== 0x1000) bad('mini stream cutoff is not 0x1000');

  const ssz = 1 << sectorShift;
  // cfb's sectorify numbers sectors from 0 after the 512-byte header; the last may be partial: ceil - 1.
  const nSectors = Math.ceil(bytes.length / ssz) - 1; // sector i occupies bytes[(i+1)*ssz ..]
  if (nSectors < 0) bad('no sectors');
  if (nSectors > limits.maxSectors) throw new OfficeCryptoError('limit', `${nSectors} sectors, over the ${limits.maxSectors}-sector limit`);

  // A FAT/DIFAT entry read from a sector; out-of-file reads (a partial final sector) count as FREESECT, as cfb's.
  const sectorU32 = (sector: number, index: number): number => {
    const at = (sector + 1) * ssz + index * 4;
    return at + 4 <= bytes.length ? dv.getUint32(at, true) >>> 0 : FREESECT;
  };
  const perSector = ssz >>> 2;
  const isRegular = (s: number): boolean => s <= MAXREGSECT && s < nSectors; // a real, in-range sector number

  // Collect FAT sector locations: the 109 in the header DIFAT, then any chained DIFAT sectors (guarded).
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i++) {
    const s = dv.getUint32(76 + i * 4, true) >>> 0;
    if (isRegular(s)) fatSectors.push(s);
  }
  let difat = dv.getUint32(68, true) >>> 0;     // first DIFAT sector
  const numDifat = dv.getUint32(72, true) >>> 0;
  const seenDifat = new Set<number>();
  let difatSteps = 0;
  while (isRegular(difat)) {
    if (seenDifat.has(difat)) bad('DIFAT chain cycles');   // the only real hang risk here
    seenDifat.add(difat);
    if (++difatSteps > numDifat + 1) bad('DIFAT chain longer than declared');
    for (let i = 0; i < perSector - 1; i++) {
      const s = sectorU32(difat, i);
      if (isRegular(s)) fatSectors.push(s);
    }
    difat = sectorU32(difat, perSector - 1); // last entry chains to the next DIFAT sector
  }

  // Build the FAT (next-sector pointer for every sector) from those FAT sectors.
  const fat = new Uint32Array(nSectors).fill(FREESECT);
  let fatIndex = 0;
  for (const fs of fatSectors) {
    for (let i = 0; i < perSector; i++) {
      if (fatIndex < nSectors) fat[fatIndex] = sectorU32(fs, i);
      fatIndex++;
    }
  }

  // Prove the FAT is acyclic, which is the one thing cfb's get_sector_list does not guard against. Anything that
  // is not a real in-range sector (ENDOFCHAIN, FREESECT, markers, or an out-of-range pointer, all of which cfb
  // treats as the end of a chain) terminates the walk; only re-entering the current path is a fault. 3-colour
  // walk: 0 unknown, 1 on the current path, 2 proven to terminate.
  const state = new Uint8Array(nSectors);
  const path: number[] = [];
  for (let start = 0; start < nSectors; start++) {
    if (state[start] !== 0) continue;
    let s = start;
    path.length = 0;
    while (isRegular(s) && state[s] === 0) {
      state[s] = 1;
      path.push(s);
      s = fat[s]!;
    }
    if (isRegular(s) && state[s] === 1) bad('FAT chain cycles');  // re-entered the current path
    for (const p of path) state[p] = 2;
  }
}

/** A readable view over the streams of a compound file. */
export interface Container {
  stream(name: string): Uint8Array | undefined;
}

/** Guard, then read the compound file with cfb. The guard rules out the loops cfb does not defend against. */
export function readContainer(bytes: Uint8Array, limits: ContainerLimits = DEFAULT_CONTAINER_LIMITS): Container {
  assertSafeContainer(bytes, limits);
  let cfb: CFB.CFB$Container;
  try {
    cfb = CFB.read(bytes, { type: 'buffer' });
  } catch (e) {
    throw new OfficeCryptoError('malformed', `could not read the compound file: ${e instanceof Error ? e.message : String(e)}`, { cause: e });
  }
  return {
    stream(name: string): Uint8Array | undefined {
      const entry = CFB.find(cfb, name);
      if (!entry || entry.content == null) return undefined;
      return new Uint8Array(entry.content as ArrayLike<number>);
    },
  };
}

export interface StreamSpec { path: string; content: Uint8Array; }

// cfb seeds a 4-byte throwaway stream named "\u0001Sh33tJ5" into every container (seed_cfb), and CFB.write always
// runs rebuild_cfb -> init_cfb -> seed_cfb, which re-adds it whenever it is absent. So no public cfb call can emit
// a container without it: deleting it before write only makes write put it back. We leave it in. It is inert for
// our purpose -- Office and msoffcrypto locate the payload through the DataSpaces map and the EncryptionInfo and
// EncryptedPackage stream names, never by enumerating root streams -- and the differential test confirms
// msoffcrypto decrypts what we write. It is the one documented deviation from a byte-clean PowerPoint layout.

/** Write a compound file holding the given streams (full paths like "/EncryptedPackage", "/\u0006DataSpaces/Version"). */
export function writeContainer(streams: StreamSpec[]): Uint8Array {
  const cfb = CFB.utils.cfb_new();
  for (const s of streams) CFB.utils.cfb_add(cfb, s.path, s.content);
  const out = CFB.write(cfb, { type: 'buffer' }) as Uint8Array | number[];
  return out instanceof Uint8Array ? new Uint8Array(out) : Uint8Array.from(out);
}
