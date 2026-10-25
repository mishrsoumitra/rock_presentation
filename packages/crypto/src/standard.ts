// packages/crypto/src/standard.ts
// Standard (binary) encryption, what Office 2007 wrote (MS-OFFCRYPTO 2.3.4.5-2.3.4.9): version 3.2/4.2 header,
// AES in ECB mode, SHA-1, 50000 iterations. Decryption only; RC4 and XOR obfuscation are reported as unsupported.

import { concatBytes, padToBlock, readU32le, timingSafeEqual, u32le } from './bytes.ts';
import { OfficeCryptoError } from './errors.ts';
import { sha1 } from './sha.ts';
import { aesEcbDecrypt } from './webcrypto.ts';

const ITER_COUNT = 50_000;
const SHA1_SIZE = 20;

// AlgID values for AES (MS-OFFCRYPTO 2.3.4.5). Anything else (RC4 0x6801, etc.) is out of scope.
const ALG_AES = new Map<number, number>([[0x660e, 128], [0x660f, 192], [0x6610, 256]]);
const ALGID_SHA1 = 0x8004;

interface StandardInfo {
  keyBits: number;
  salt: Uint8Array;
  encryptedVerifier: Uint8Array;     // 16 bytes
  encryptedVerifierHash: Uint8Array; // 32 bytes (AES)
}

/** Parse the standard EncryptionInfo stream (binary header + verifier). Throws unsupported for RC4/non-AES. */
export function parseStandardInfo(encryptionInfo: Uint8Array): StandardInfo {
  const bad = (m: string): never => { throw new OfficeCryptoError('malformed', m); };
  if (encryptionInfo.length < 12) bad('standard EncryptionInfo is too short');
  const headerSize = readU32le(encryptionInfo, 8);
  const headerStart = 12;
  if (headerSize < 32 || headerStart + headerSize > encryptionInfo.length) bad(`bad standard header size ${headerSize}`);

  const algId = readU32le(encryptionInfo, headerStart + 8);
  const algIdHash = readU32le(encryptionInfo, headerStart + 12);
  const headerKeySize = readU32le(encryptionInfo, headerStart + 16);

  const aesBits = ALG_AES.get(algId);
  if (aesBits === undefined) {
    throw new OfficeCryptoError('unsupported', `standard encryption with AlgID 0x${algId.toString(16)} is not supported (only AES; RC4 and XOR obfuscation are out of scope)`);
  }
  if (algIdHash !== 0 && algIdHash !== ALGID_SHA1) {
    throw new OfficeCryptoError('unsupported', `standard encryption hash 0x${algIdHash.toString(16)} is not supported (only SHA-1)`);
  }
  const keyBits = headerKeySize !== 0 ? headerKeySize : aesBits;

  // EncryptionVerifier follows the header (MS-OFFCRYPTO 2.3.3).
  let at = headerStart + headerSize;
  const saltSize = readU32le(encryptionInfo, at); at += 4;
  if (saltSize !== 16) bad(`unexpected standard salt size ${saltSize}`);
  const salt = encryptionInfo.subarray(at, at + 16); at += 16;
  const encryptedVerifier = encryptionInfo.subarray(at, at + 16); at += 16;
  const verifierHashSize = readU32le(encryptionInfo, at); at += 4;
  if (verifierHashSize !== SHA1_SIZE) bad(`unexpected standard verifier hash size ${verifierHashSize}`);
  const encryptedVerifierHash = encryptionInfo.subarray(at, at + 32); at += 32;
  if (at > encryptionInfo.length) bad('standard EncryptionInfo verifier is truncated');

  return { keyBits, salt, encryptedVerifier: encryptedVerifier.slice(), encryptedVerifierHash: encryptedVerifierHash.slice() };
}

/**
 * Derive the key from the password: H0 = SHA1(salt + UTF16LE(pw)); Hn = SHA1(LE32(n) + H(n-1)) for 50000
 * iterations; Hfinal = SHA1(H + LE32(0)); then the key is the leading bytes of SHA1(Hfinal^0x36...) ||
 * SHA1(Hfinal^0x5c...) (MS-OFFCRYPTO 2.3.4.7).
 */
function deriveKey(info: StandardInfo, passwordUtf16: Uint8Array): Uint8Array {
  let h = sha1(concatBytes(info.salt, passwordUtf16));
  const buf = new Uint8Array(4 + SHA1_SIZE);
  buf.set(h, 4);
  for (let i = 0; i < ITER_COUNT; i++) {
    buf.set(u32le(i), 0);
    buf.set(h, 4);
    h = sha1(buf);
  }
  const hFinal = sha1(concatBytes(h, u32le(0)));

  const x1 = new Uint8Array(64).fill(0x36);
  const x2 = new Uint8Array(64).fill(0x5c);
  for (let i = 0; i < SHA1_SIZE; i++) { x1[i] = x1[i]! ^ hFinal[i]!; x2[i] = x2[i]! ^ hFinal[i]!; }
  const derived = concatBytes(sha1(x1), sha1(x2));
  return derived.subarray(0, info.keyBits / 8);
}

/** Verify the password, returning the derived key or throwing wrong-password. */
export async function verifyStandardPassword(info: StandardInfo, passwordUtf16: Uint8Array): Promise<Uint8Array> {
  const key = deriveKey(info, passwordUtf16);
  const verifier = await aesEcbDecrypt(key, info.encryptedVerifier);
  const expected = await aesEcbDecrypt(key, info.encryptedVerifierHash);
  const actual = sha1(verifier);
  if (!timingSafeEqual(actual, expected.subarray(0, SHA1_SIZE))) {
    throw new OfficeCryptoError('wrong-password', 'The password is incorrect.');
  }
  return key;
}

/** Decrypt the EncryptedPackage stream (ECB over the whole payload after the 8-byte size prefix). */
export async function decryptStandardPayload(key: Uint8Array, encryptedPackage: Uint8Array): Promise<Uint8Array> {
  if (encryptedPackage.length < 8) throw new OfficeCryptoError('malformed', 'EncryptedPackage stream is too short');
  const totalSize = readU32le(encryptedPackage, 0);
  const body = encryptedPackage.subarray(8);
  if (totalSize > body.length) throw new OfficeCryptoError('malformed', `EncryptedPackage declares ${totalSize} bytes but holds ${body.length}`);
  const decrypted = await aesEcbDecrypt(key, padToBlock(body, 16));
  return decrypted.subarray(0, totalSize);
}
