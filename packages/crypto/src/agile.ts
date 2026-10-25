// packages/crypto/src/agile.ts
// Agile encryption (MS-OFFCRYPTO 2.3.4.10-2.3.4.15), what PowerPoint 2013+ writes: AES-256-CBC, SHA-512,
// spinCount 100000, an XML descriptor in EncryptionInfo, and a data-integrity HMAC over EncryptedPackage.

import { childElements, decodeXmlBytes, getAttr, parseXml, type XmlElement } from '@pptx/xml';
import {
  concatBytes, fromBase64, padToBlock, readU64le, timingSafeEqual, toBase64, u32le, u64le,
} from './bytes.ts';
import { OfficeCryptoError } from './errors.ts';
import { shaDigestSize, shaHash, type ShaName } from './sha.ts';
import { aesCbcDecrypt, aesCbcEncrypt, hmac, randomBytes } from './webcrypto.ts';

const SEGMENT = 4096;

// Block keys (MS-OFFCRYPTO 2.3.4.12-2.3.4.14), appended to the iterated hash to derive each purpose's key.
const BLOCK_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_VALUE = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
const BLOCK_INTEGRITY_KEY = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]);
const BLOCK_INTEGRITY_VALUE = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]);

interface CipherParams {
  saltValue: Uint8Array;
  hashName: ShaName;       // Web Crypto / spec spelling: 'SHA-512'
  saltSize: number;
  blockSize: number;
  keyBits: number;
  hashSize: number;
}

export interface AgileInfo {
  keyData: CipherParams;
  encryptedKey: CipherParams;
  spinCount: number;
  encryptedHmacKey: Uint8Array;
  encryptedHmacValue: Uint8Array;
  encryptedVerifierHashInput: Uint8Array;
  encryptedVerifierHashValue: Uint8Array;
  encryptedKeyValue: Uint8Array;
}

function hashNameFromXml(spelling: string | undefined): ShaName {
  // Descriptors spell it "SHA512"/"SHA-512"; we implement only SHA-512 and SHA-1 for agile.
  const s = (spelling ?? '').toUpperCase().replace('-', '');
  if (s === 'SHA512') return 'SHA-512';
  if (s === 'SHA1') return 'SHA-1';
  throw new OfficeCryptoError('unsupported', `unsupported agile hash algorithm ${JSON.stringify(spelling)} (only SHA-512 and SHA-1)`);
}

function requireAes(cipher: string | undefined): void {
  if ((cipher ?? '').toUpperCase() !== 'AES') {
    throw new OfficeCryptoError('unsupported', `unsupported agile cipher ${JSON.stringify(cipher)} (only AES)`);
  }
}

function findDescendant(el: XmlElement, local: string): XmlElement | undefined {
  for (const c of childElements(el)) {
    if (c.local === local) return c;
    const deep = findDescendant(c, local);
    if (deep) return deep;
  }
  return undefined;
}

function reqAttr(el: XmlElement, name: string): string {
  const v = getAttr(el, name);
  if (v === undefined) throw new OfficeCryptoError('malformed', `agile descriptor is missing the ${name} attribute`);
  return v;
}

function reqInt(el: XmlElement, name: string): number {
  const n = Number(reqAttr(el, name));
  if (!Number.isInteger(n) || n < 0) throw new OfficeCryptoError('malformed', `agile descriptor has a bad ${name}`);
  return n;
}

function parseCipherParams(el: XmlElement): CipherParams {
  requireAes(getAttr(el, 'cipherAlgorithm'));
  const blockSize = reqInt(el, 'blockSize');
  if (blockSize !== 16) throw new OfficeCryptoError('unsupported', `unsupported agile block size ${blockSize}`);
  return {
    saltValue: fromBase64(reqAttr(el, 'saltValue')),
    hashName: hashNameFromXml(getAttr(el, 'hashAlgorithm')),
    saltSize: reqInt(el, 'saltSize'),
    blockSize,
    keyBits: reqInt(el, 'keyBits'),
    hashSize: reqInt(el, 'hashSize'),
  };
}

/** Parse the EncryptionInfo stream (the 8-byte header then the UTF-8 XML descriptor). */
export function parseAgileInfo(encryptionInfo: Uint8Array): AgileInfo {
  let root: XmlElement;
  try {
    root = parseXml(decodeXmlBytes(encryptionInfo.subarray(8)).text).root;
  } catch (e) {
    throw new OfficeCryptoError('malformed', `agile EncryptionInfo XML did not parse: ${e instanceof Error ? e.message : String(e)}`, { cause: e });
  }
  const keyDataEl = findDescendant(root, 'keyData');
  const dataIntegrityEl = findDescendant(root, 'dataIntegrity');
  const encryptedKeyEl = findDescendant(root, 'encryptedKey');
  if (!keyDataEl || !encryptedKeyEl) throw new OfficeCryptoError('malformed', 'agile descriptor is missing keyData or encryptedKey');
  if (!dataIntegrityEl) throw new OfficeCryptoError('malformed', 'agile descriptor is missing dataIntegrity');

  const spinCount = reqInt(encryptedKeyEl, 'spinCount');
  if (spinCount > 10_000_000) throw new OfficeCryptoError('limit', `spinCount ${spinCount} is implausibly large`);

  return {
    keyData: parseCipherParams(keyDataEl),
    encryptedKey: parseCipherParams(encryptedKeyEl),
    spinCount,
    encryptedHmacKey: fromBase64(reqAttr(dataIntegrityEl, 'encryptedHmacKey')),
    encryptedHmacValue: fromBase64(reqAttr(dataIntegrityEl, 'encryptedHmacValue')),
    encryptedVerifierHashInput: fromBase64(reqAttr(encryptedKeyEl, 'encryptedVerifierHashInput')),
    encryptedVerifierHashValue: fromBase64(reqAttr(encryptedKeyEl, 'encryptedVerifierHashValue')),
    encryptedKeyValue: fromBase64(reqAttr(encryptedKeyEl, 'encryptedKeyValue')),
  };
}

/** H0 = H(salt + UTF16LE(pw)); Hn = H(LE32(n) + H(n-1)) for spinCount iterations. The hot loop; uses sha.ts. */
function iteratedHash(hashName: ShaName, salt: Uint8Array, passwordUtf16: Uint8Array, spinCount: number): Uint8Array {
  let h = shaHash(hashName, concatBytes(salt, passwordUtf16));
  const buf = new Uint8Array(4 + h.length); // LE32(i) + previous digest, rewritten each turn
  buf.set(h, 4);
  for (let i = 0; i < spinCount; i++) {
    buf[0] = i & 0xff;
    buf[1] = (i >>> 8) & 0xff;
    buf[2] = (i >>> 16) & 0xff;
    buf[3] = (i >>> 24) & 0xff;
    h = shaHash(hashName, buf);
    buf.set(h, 4);
  }
  return h;
}

/** Final key for a purpose: H(Hfinal + blockKey) truncated to keyBits/8 (MS-OFFCRYPTO 2.3.4.11). */
function deriveKey(hashName: ShaName, hFinal: Uint8Array, blockKey: Uint8Array, keyBits: number): Uint8Array {
  return shaHash(hashName, concatBytes(hFinal, blockKey)).subarray(0, keyBits / 8);
}

/** Segment IV = H(keyData salt + LE32(index)) truncated to the block size (or the salt itself when no block key). */
function segmentIv(hashName: ShaName, salt: Uint8Array, index: number, blockSize: number): Uint8Array {
  return shaHash(hashName, concatBytes(salt, u32le(index))).subarray(0, blockSize);
}

function blockKeyIv(info: CipherParams, blockKey: Uint8Array): Uint8Array {
  return shaHash(info.hashName, concatBytes(info.saltValue, blockKey)).subarray(0, info.blockSize);
}

/** Verify the password against the verifier, returning the derived package (secret) key, or throwing wrong-password. */
export async function deriveSecretKey(info: AgileInfo, passwordUtf16: Uint8Array): Promise<Uint8Array> {
  const ek = info.encryptedKey;
  const hFinal = iteratedHash(ek.hashName, ek.saltValue, passwordUtf16, info.spinCount);

  const keyInput = deriveKey(ek.hashName, hFinal, BLOCK_VERIFIER_INPUT, ek.keyBits);
  const keyValue = deriveKey(ek.hashName, hFinal, BLOCK_VERIFIER_VALUE, ek.keyBits);
  const verifierInput = await aesCbcDecrypt(keyInput, ek.saltValue, info.encryptedVerifierHashInput);
  const expected = await aesCbcDecrypt(keyValue, ek.saltValue, info.encryptedVerifierHashValue);
  const actual = shaHash(ek.hashName, verifierInput);
  if (!timingSafeEqual(actual, expected.subarray(0, actual.length))) {
    throw new OfficeCryptoError('wrong-password', 'The password is incorrect.');
  }

  const keyForValue = deriveKey(ek.hashName, hFinal, BLOCK_KEY_VALUE, ek.keyBits);
  return aesCbcDecrypt(keyForValue, ek.saltValue, info.encryptedKeyValue);
}

/** Check the data-integrity HMAC over the EncryptedPackage stream; throws on mismatch (tamper/corruption). */
export async function verifyIntegrity(info: AgileInfo, secretKey: Uint8Array, encryptedPackage: Uint8Array): Promise<void> {
  const kd = info.keyData;
  const hmacKey = await aesCbcDecrypt(secretKey, blockKeyIv(kd, BLOCK_INTEGRITY_KEY), info.encryptedHmacKey);
  const hmacValue = await aesCbcDecrypt(secretKey, blockKeyIv(kd, BLOCK_INTEGRITY_VALUE), info.encryptedHmacValue);
  const actual = await hmac(kd.hashName, hmacKey.subarray(0, kd.hashSize), encryptedPackage);
  if (!timingSafeEqual(actual, hmacValue.subarray(0, kd.hashSize))) {
    throw new OfficeCryptoError('integrity', 'The data-integrity check failed: the encrypted package was tampered with or is corrupt.');
  }
}

/** Decrypt the EncryptedPackage stream to the original ZIP bytes. */
export async function decryptPayload(info: AgileInfo, secretKey: Uint8Array, encryptedPackage: Uint8Array): Promise<Uint8Array> {
  const kd = info.keyData;
  if (encryptedPackage.length < 8) throw new OfficeCryptoError('malformed', 'EncryptedPackage stream is too short');
  const totalSize = readU64le(encryptedPackage, 0);
  const body = encryptedPackage.subarray(8);
  if (totalSize > body.length) throw new OfficeCryptoError('malformed', `EncryptedPackage declares ${totalSize} bytes but holds ${body.length}`);
  const out = new Uint8Array(totalSize);
  let written = 0;
  for (let i = 0, off = 0; off < body.length; i++, off += SEGMENT) {
    const segment = body.subarray(off, Math.min(off + SEGMENT, body.length));
    const dec = await aesCbcDecrypt(secretKey, segmentIv(kd.hashName, kd.saltValue, i, kd.blockSize), padToBlock(segment, kd.blockSize));
    const take = Math.min(dec.length, totalSize - written);
    out.set(dec.subarray(0, take), written);
    written += take;
    if (written >= totalSize) break;
  }
  return out;
}

export interface EncryptOptions {
  spinCount?: number;   // lower it in tests; defaults to the PowerPoint value
}
const DEFAULT_SPIN_COUNT = 100_000;

/** The result of encrypting: the two streams a compound file must hold. */
export interface EncryptedStreams {
  encryptionInfo: Uint8Array;
  encryptedPackage: Uint8Array;
}

/** Encrypt ZIP bytes with a fresh random key and salts: AES-256-CBC, SHA-512, data-integrity HMAC. */
export async function encryptAgile(zipBytes: Uint8Array, passwordUtf16: Uint8Array, options: EncryptOptions = {}): Promise<EncryptedStreams> {
  const spinCount = options.spinCount ?? DEFAULT_SPIN_COUNT;
  if (!Number.isInteger(spinCount) || spinCount < 0 || spinCount > 10_000_000) {
    throw new OfficeCryptoError('limit', `spinCount ${spinCount} is out of range`);
  }
  const hashName: ShaName = 'SHA-512';
  const keyBits = 256;
  const blockSize = 16;
  const saltSize = 16;
  const hashSize = shaDigestSize(hashName); // 64

  const keyDataSalt = randomBytes(saltSize);
  const encKeySalt = randomBytes(saltSize);
  const secretKey = randomBytes(keyBits / 8); // the package key, 32 random bytes

  const kd: CipherParams = { saltValue: keyDataSalt, hashName, saltSize, blockSize, keyBits, hashSize };
  const ek: CipherParams = { saltValue: encKeySalt, hashName, saltSize, blockSize, keyBits, hashSize };

  // Encrypt the payload in 4096-byte segments, each CBC-encrypted under its own IV and padded to the block size.
  // The output length is the 8-byte size prefix plus each segment rounded up to a multiple of the block size.
  const totalSize = zipBytes.length;
  const segCount = Math.ceil(totalSize / SEGMENT);
  const bodyLen = segCount === 0 ? 0 : (segCount - 1) * SEGMENT + Math.ceil((totalSize - (segCount - 1) * SEGMENT) / blockSize) * blockSize;
  const encryptedPackage = new Uint8Array(8 + bodyLen);
  encryptedPackage.set(u64le(totalSize), 0);
  let writeAt = 8;
  for (let i = 0; i < segCount; i++) {
    const seg = zipBytes.subarray(i * SEGMENT, (i + 1) * SEGMENT);
    const enc = await aesCbcEncrypt(secretKey, segmentIv(hashName, keyDataSalt, i, blockSize), padToBlock(seg, blockSize));
    encryptedPackage.set(enc, writeAt);
    writeAt += enc.length;
  }

  // Verifier: a random value, its AES-CBC form, and the hash of it, all keyed off the password.
  const hFinal = iteratedHash(hashName, encKeySalt, passwordUtf16, spinCount);
  const keyInput = deriveKey(hashName, hFinal, BLOCK_VERIFIER_INPUT, keyBits);
  const keyValue = deriveKey(hashName, hFinal, BLOCK_VERIFIER_VALUE, keyBits);
  const keyForValue = deriveKey(hashName, hFinal, BLOCK_KEY_VALUE, keyBits);

  const verifierInput = randomBytes(saltSize);
  const encryptedVerifierHashInput = await aesCbcEncrypt(keyInput, encKeySalt, verifierInput);
  const encryptedVerifierHashValue = await aesCbcEncrypt(keyValue, encKeySalt, padToBlock(shaHash(hashName, verifierInput), blockSize));
  const encryptedKeyValue = await aesCbcEncrypt(keyForValue, encKeySalt, secretKey);

  // Data integrity: a random HMAC key, encrypted, and the HMAC of the encrypted package, encrypted.
  const hmacKey = randomBytes(hashSize);
  const encryptedHmacKey = await aesCbcEncrypt(secretKey, blockKeyIv(kd, BLOCK_INTEGRITY_KEY), hmacKey);
  const hmacValue = await hmac(hashName, hmacKey, encryptedPackage);
  const encryptedHmacValue = await aesCbcEncrypt(secretKey, blockKeyIv(kd, BLOCK_INTEGRITY_VALUE), padToBlock(hmacValue, blockSize));

  const info: AgileInfo = {
    keyData: kd, encryptedKey: ek, spinCount,
    encryptedHmacKey, encryptedHmacValue,
    encryptedVerifierHashInput, encryptedVerifierHashValue, encryptedKeyValue,
  };
  return { encryptionInfo: buildEncryptionInfo(info), encryptedPackage };
}

/** Serialize an AgileInfo to an EncryptionInfo stream, matching the layout PowerPoint writes. */
export function buildEncryptionInfo(info: AgileInfo): Uint8Array {
  const kd = info.keyData;
  const ek = info.encryptedKey;
  const hashSpelling = kd.hashName === 'SHA-512' ? 'SHA512' : 'SHA1';
  const b64 = toBase64;
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
    `<keyData saltSize="${kd.saltSize}" blockSize="${kd.blockSize}" keyBits="${kd.keyBits}" hashSize="${kd.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${hashSpelling}" saltValue="${b64(kd.saltValue)}"/>` +
    `<dataIntegrity encryptedHmacKey="${b64(info.encryptedHmacKey)}" encryptedHmacValue="${b64(info.encryptedHmacValue)}"/>` +
    '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
    `<p:encryptedKey spinCount="${info.spinCount}" saltSize="${ek.saltSize}" blockSize="${ek.blockSize}" keyBits="${ek.keyBits}" hashSize="${ek.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${hashSpelling}" saltValue="${b64(ek.saltValue)}" encryptedVerifierHashInput="${b64(info.encryptedVerifierHashInput)}" encryptedVerifierHashValue="${b64(info.encryptedVerifierHashValue)}" encryptedKeyValue="${b64(info.encryptedKeyValue)}"/>` +
    '</keyEncryptor></keyEncryptors></encryption>';
  const body = new TextEncoder().encode(xml);
  const header = new Uint8Array([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]); // version 4.4, reserved 0x40
  return concatBytes(header, body);
}
