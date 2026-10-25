// packages/crypto/src/bytes.ts -- small byte helpers, no Node and no DOM (TextEncoder/atob exist in all targets).

/** Concatenate byte arrays into one. */
export function concatBytes(...parts: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

/** A 4-byte little-endian encoding of an unsigned 32-bit integer (the agile/standard block-key counter). */
export function u32le(n: number): Uint8Array {
  const out = new Uint8Array(4);
  out[0] = n & 0xff;
  out[1] = (n >>> 8) & 0xff;
  out[2] = (n >>> 16) & 0xff;
  out[3] = (n >>> 24) & 0xff;
  return out;
}

/** Read an unsigned 32-bit little-endian integer. */
export function readU32le(b: Uint8Array, at: number): number {
  return (b[at]! | (b[at + 1]! << 8) | (b[at + 2]! << 16) | (b[at + 3]! << 24)) >>> 0;
}

/**
 * An 8-byte little-endian encoding of a size. The EncryptedPackage prefix is a 64-bit value, but a .pptx never
 * approaches 2^53 bytes, so a JS number is exact here and avoids BigInt in the hot path.
 */
export function u64le(n: number): Uint8Array {
  const out = new Uint8Array(8);
  for (let i = 0; i < 8; i++) {
    out[i] = n & 0xff;
    n = Math.floor(n / 256);
  }
  return out;
}

/** Read an unsigned 64-bit little-endian size as a JS number (exact below 2^53). */
export function readU64le(b: Uint8Array, at: number): number {
  let n = 0;
  for (let i = 7; i >= 0; i--) n = n * 256 + b[at + i]!;
  return n;
}

/** UTF-16LE bytes of a password, as MS-OFFCRYPTO hashes it (no BOM, no terminator). */
export function utf16le(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[2 * i] = c & 0xff;
    out[2 * i + 1] = (c >>> 8) & 0xff;
  }
  return out;
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/** Base64-encode bytes (the agile descriptor stores salts and hashes this way). Avoids Node's Buffer and btoa. */
export function toBase64(b: Uint8Array): string {
  let out = '';
  let i = 0;
  for (; i + 3 <= b.length; i += 3) {
    const n = (b[i]! << 16) | (b[i + 1]! << 8) | b[i + 2]!;
    out += B64[(n >>> 18) & 63]! + B64[(n >>> 12) & 63]! + B64[(n >>> 6) & 63]! + B64[n & 63]!;
  }
  const rem = b.length - i;
  if (rem === 1) {
    const n = b[i]! << 16;
    out += B64[(n >>> 18) & 63]! + B64[(n >>> 12) & 63]! + '==';
  } else if (rem === 2) {
    const n = (b[i]! << 16) | (b[i + 1]! << 8);
    out += B64[(n >>> 18) & 63]! + B64[(n >>> 12) & 63]! + B64[(n >>> 6) & 63]! + '=';
  }
  return out;
}

const B64_INV = (() => {
  const inv = new Int16Array(128).fill(-1);
  for (let i = 0; i < B64.length; i++) inv[B64.charCodeAt(i)] = i;
  return inv;
})();

/** Decode standard base64 (ignoring ASCII whitespace); throws on an invalid character. */
export function fromBase64(s: string): Uint8Array {
  const chars: number[] = [];
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) continue;
    if (c === 0x3d) break; // '=' padding ends the data
    const v = c < 128 ? B64_INV[c]! : -1;
    if (v < 0) throw new Error(`invalid base64 character ${JSON.stringify(s[i])}`);
    chars.push(v);
  }
  const outLen = (chars.length * 6) >> 3;
  const out = new Uint8Array(outLen);
  let acc = 0;
  let bits = 0;
  let at = 0;
  for (const v of chars) {
    acc = (acc << 6) | v;
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      out[at++] = (acc >>> bits) & 0xff;
    }
  }
  return out;
}

/** Hex for diagnostics and fixed block keys written in tests. */
export function toHex(b: Uint8Array): string {
  let out = '';
  for (const x of b) out += x.toString(16).padStart(2, '0');
  return out;
}

/** Constant-time equality: verifier and HMAC comparisons must not leak where they first differ. */
export function timingSafeEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a[i]! ^ b[i]!;
  return diff === 0;
}

/** A copy padded with zero bytes up to the next multiple of `block` (MS-OFFCRYPTO pads segments to the block size). */
export function padToBlock(b: Uint8Array, block: number): Uint8Array {
  const rem = b.length % block;
  if (rem === 0) return b;
  const out = new Uint8Array(b.length + (block - rem));
  out.set(b);
  return out;
}
