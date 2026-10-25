// packages/crypto/src/webcrypto.ts
// AES and HMAC through Web Crypto only (`globalThis.crypto.subtle`), which exists in browsers, workers and Node.
// Web Crypto has no raw (unpadded) CBC and no ECB, so both are built here out of the padded AES-CBC it does
// offer. The tricks are explained at each function. Key derivation's hash spin loop lives in sha.ts instead.

import { OfficeCryptoError } from './errors.ts';
import { concatBytes } from './bytes.ts';

// No explicit SubtleCrypto/CryptoKey annotations: those type names are not in every target's lib, so the types
// are inferred from globalThis.crypto instead (both the webworker and the Node libs provide the shape).
function subtle() {
  const c = globalThis.crypto;
  if (!c || !c.subtle) {
    throw new OfficeCryptoError('unsupported', 'Web Crypto (globalThis.crypto.subtle) is not available in this environment.');
  }
  return c.subtle;
}

/** Cryptographically random bytes for salts and the package key (crypto.getRandomValues). */
export function randomBytes(n: number): Uint8Array {
  const c = globalThis.crypto;
  if (!c || !c.getRandomValues) {
    throw new OfficeCryptoError('unsupported', 'crypto.getRandomValues is not available in this environment.');
  }
  const out = new Uint8Array(n);
  c.getRandomValues(out);
  return out;
}

const ZERO_IV = new Uint8Array(16);
const PAD_BLOCK = new Uint8Array(16).fill(0x10); // a full PKCS#7 pad block (16 bytes, each 0x10)

// A Web Crypto view needs its own ArrayBuffer (not a subarray over a larger buffer), or BufferSource may read too much.
function ab(b: Uint8Array): ArrayBuffer {
  return b.buffer.byteLength === b.length && b.byteOffset === 0 ? (b.buffer as ArrayBuffer) : (b.slice().buffer as ArrayBuffer);
}

async function cbcKey(key: Uint8Array) {
  return subtle().importKey('raw', ab(key), { name: 'AES-CBC' }, false, ['encrypt', 'decrypt']);
}

/**
 * AES-CBC encryption with no padding. The input length must be a multiple of 16. Web Crypto always appends one
 * PKCS#7 block; since CBC is causal, the first `data.length` output bytes are exactly the unpadded ciphertext,
 * so the appended block is sliced off.
 */
export async function aesCbcEncrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  if (data.length % 16 !== 0) throw new OfficeCryptoError('malformed', `AES-CBC input must be a multiple of 16 bytes, got ${data.length}`);
  const k = await cbcKey(key);
  const ct = new Uint8Array(await subtle().encrypt({ name: 'AES-CBC', iv: ab(iv) }, k, ab(data)));
  return ct.subarray(0, data.length);
}

/**
 * AES-CBC decryption with no padding. The input length must be a multiple of 16. Web Crypto's decrypt requires a
 * valid PKCS#7 padding on the last block and strips it, so a crafted extra block is appended that decrypts to a
 * full 0x10 pad block: appended block B must satisfy D(B) XOR Clast = 0x10..., i.e. B = E(PAD_BLOCK XOR Clast).
 * Web Crypto then removes exactly those 16 bytes and returns the original data.
 */
export async function aesCbcDecrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  if (data.length === 0) return new Uint8Array(0);
  if (data.length % 16 !== 0) throw new OfficeCryptoError('malformed', `AES-CBC input must be a multiple of 16 bytes, got ${data.length}`);
  const k = await cbcKey(key);
  const cLast = data.subarray(data.length - 16);
  const x = new Uint8Array(16);
  for (let i = 0; i < 16; i++) x[i] = PAD_BLOCK[i]! ^ cLast[i]!;
  // E(x) via a single CBC block with a zero IV; take the first 16 bytes (the second is CBC's own pad block).
  const encX = new Uint8Array(await subtle().encrypt({ name: 'AES-CBC', iv: ab(ZERO_IV) }, k, ab(x)));
  const padded = concatBytes(data, encX.subarray(0, 16));
  const pt = new Uint8Array(await subtle().decrypt({ name: 'AES-CBC', iv: ab(iv) }, k, ab(padded)));
  return pt; // already data.length bytes: the appended block was stripped as padding
}

/**
 * AES-ECB decryption, used by the older "standard" scheme. ECB has no chaining, so decrypting through CBC with a
 * zero IV yields P_i = D(C_i) XOR C_(i-1); XORing each block back with the previous ciphertext block (zero before
 * the first) recovers D(C_i), the true ECB plaintext. Input length must be a multiple of 16.
 */
export async function aesEcbDecrypt(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  if (data.length % 16 !== 0) throw new OfficeCryptoError('malformed', `AES-ECB input must be a multiple of 16 bytes, got ${data.length}`);
  const cbc = await aesCbcDecrypt(key, ZERO_IV, data);
  const out = new Uint8Array(cbc.length);
  for (let block = 0; block < data.length; block += 16) {
    for (let i = 0; i < 16; i++) {
      out[block + i] = cbc[block + i]! ^ (block === 0 ? 0 : data[block - 16 + i]!);
    }
  }
  return out;
}

/** HMAC over the whole message, through Web Crypto. */
export async function hmac(hashName: 'SHA-1' | 'SHA-512', key: Uint8Array, message: Uint8Array): Promise<Uint8Array> {
  const k = await subtle().importKey('raw', ab(key), { name: 'HMAC', hash: hashName }, false, ['sign']);
  return new Uint8Array(await subtle().sign('HMAC', k, ab(message)));
}
