// packages/metafile/src/view.ts  -- bounds-checked little-endian reads over a window of the file
import { MetafileError } from './errors.ts';

/**
 * A window of the file, usually one record: reads are relative to its start and fail with a 'malformed'
 * MetafileError rather than ever reading past its end, so a record cannot see its neighbours' bytes.
 */
export class View {
  readonly bytes: Uint8Array;
  readonly start: number;
  readonly size: number;
  private readonly dv: DataView;

  constructor(bytes: Uint8Array, start = 0, size = bytes.length - start) {
    if (start < 0 || size < 0 || start + size > bytes.length) throw new MetafileError('malformed', `window ${start}+${size} is outside ${bytes.length} bytes`);
    this.bytes = bytes;
    this.start = start;
    this.size = size;
    this.dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  /** True when `n` bytes at `at` are inside the window. */
  has(at: number, n: number): boolean {
    return Number.isInteger(at) && Number.isInteger(n) && at >= 0 && n >= 0 && at + n <= this.size;
  }

  private check(at: number, n: number): number {
    if (!this.has(at, n)) throw new MetafileError('malformed', `read of ${n} bytes at ${at} outside a ${this.size}-byte record`);
    return this.start + at;
  }

  u8(at: number): number { return this.dv.getUint8(this.check(at, 1)); }
  u16(at: number): number { return this.dv.getUint16(this.check(at, 2), true); }
  i16(at: number): number { return this.dv.getInt16(this.check(at, 2), true); }
  u32(at: number): number { return this.dv.getUint32(this.check(at, 4), true); }
  i32(at: number): number { return this.dv.getInt32(this.check(at, 4), true); }
  f32(at: number): number { return this.dv.getFloat32(this.check(at, 4), true); }

  /** `n` bytes at `at`, shared with the file (not copied). */
  sub(at: number, n: number): Uint8Array {
    const p = this.check(at, n);
    return this.bytes.subarray(p, p + n);
  }

  /** A narrower window, relative to this one. */
  view(at: number, n: number): View {
    return new View(this.bytes, this.check(at, n), n);
  }
}
