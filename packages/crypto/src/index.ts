// packages/crypto/src/index.ts
// Opening and saving password-protected Open XML packages (MS-OFFCRYPTO). Agile (PowerPoint 2013+) and the older
// standard scheme are decrypted; encryption always writes agile, as PowerPoint does. No Node, no DOM: Web Crypto
// for AES/HMAC, a plain-TS hash for the key-derivation spin loop, and the `cfb` package for the compound file.

import {
  assertSafeContainer, cfbHasStream, DEFAULT_CONTAINER_LIMITS, isCompoundFile, readContainer, writeContainer,
  type ContainerLimits,
} from './container.ts';
import {
  decryptPayload, deriveSecretKey, encryptAgile, parseAgileInfo, verifyIntegrity, type EncryptOptions,
} from './agile.ts';
import { DATA_SPACE_MAP, PRIMARY, STRONG_ENCRYPTION_DATA_SPACE, VERSION } from './dataspaces.ts';
import { decryptStandardPayload, parseStandardInfo, verifyStandardPassword } from './standard.ts';
import { OfficeCryptoError } from './errors.ts';
import { utf16le } from './bytes.ts';

export { OfficeCryptoError, type OfficeCryptoErrorCode } from './errors.ts';
export { isCompoundFile, type ContainerLimits, DEFAULT_CONTAINER_LIMITS } from './container.ts';

/** How the EncryptionInfo stream declares its scheme (MS-OFFCRYPTO 2.3.4.1). */
function detectScheme(encryptionInfo: Uint8Array): 'agile' | 'standard' {
  if (encryptionInfo.length < 4) throw new OfficeCryptoError('malformed', 'EncryptionInfo stream is too short');
  const major = encryptionInfo[0]! | (encryptionInfo[1]! << 8);
  const minor = encryptionInfo[2]! | (encryptionInfo[3]! << 8);
  if (major === 4 && minor === 4) return 'agile';
  if ((major === 2 || major === 3 || major === 4) && minor === 2) return 'standard';
  if ((major === 3 || major === 4) && minor === 3) {
    throw new OfficeCryptoError('unsupported', 'extensible encryption (EncryptionInfo 3.3/4.3) is not supported');
  }
  throw new OfficeCryptoError('unsupported', `unsupported EncryptionInfo version ${major}.${minor}`);
}

/** True when the bytes are a compound file that carries an EncryptedPackage stream (a password-protected package). */
export function isEncryptedPackage(bytes: Uint8Array): boolean {
  return isCompoundFile(bytes) && cfbHasStream(bytes, 'EncryptedPackage');
}

export interface DecryptOptions {
  limits?: Partial<ContainerLimits>;
  /** Verify the agile data-integrity HMAC (default true). Off only for diagnostics; standard has no HMAC. */
  checkIntegrity?: boolean;
}

/**
 * Decrypt a password-protected package to its ZIP (.pptx) bytes. The password is verified; for agile the
 * data-integrity HMAC is checked too. Throws OfficeCryptoError with a `code`: 'not-encrypted', 'malformed',
 * 'unsupported', 'wrong-password', 'integrity' or 'limit'.
 */
export async function decryptPackage(bytes: Uint8Array, password: string, options: DecryptOptions = {}): Promise<Uint8Array> {
  if (!isCompoundFile(bytes)) throw new OfficeCryptoError('not-encrypted', 'Not an encrypted Office package: not a compound file.');
  const limits: ContainerLimits = { ...DEFAULT_CONTAINER_LIMITS, ...options.limits };
  const container = readContainer(bytes, limits);

  const encryptionInfo = container.stream('EncryptionInfo');
  const encryptedPackage = container.stream('EncryptedPackage');
  if (!encryptionInfo || !encryptedPackage) {
    throw new OfficeCryptoError('not-encrypted', 'Not an encrypted Office package: no EncryptionInfo/EncryptedPackage streams.');
  }
  const password16 = utf16le(password);

  if (detectScheme(encryptionInfo) === 'agile') {
    const info = parseAgileInfo(encryptionInfo);
    const secretKey = await deriveSecretKey(info, password16);     // throws wrong-password
    if (options.checkIntegrity !== false) await verifyIntegrity(info, secretKey, encryptedPackage);
    return decryptPayload(info, secretKey, encryptedPackage);
  }
  const info = parseStandardInfo(encryptionInfo);
  const key = await verifyStandardPassword(info, password16);      // throws wrong-password
  return decryptStandardPayload(key, encryptedPackage);
}

/**
 * Encrypt ZIP (.pptx) bytes into a password-protected compound file: agile encryption with fresh random salts,
 * AES-256-CBC, SHA-512, spinCount 100000 (lower it with options.spinCount in tests), a data-integrity HMAC, and
 * the DataSpaces structures PowerPoint expects.
 */
export async function encryptPackage(zipBytes: Uint8Array, password: string, options: EncryptOptions = {}): Promise<Uint8Array> {
  const { encryptionInfo, encryptedPackage } = await encryptAgile(zipBytes, utf16le(password), options);
  const container = writeContainer([
    { path: '/EncryptedPackage', content: encryptedPackage },
    { path: '/\u0006DataSpaces/Version', content: VERSION },
    { path: '/\u0006DataSpaces/DataSpaceMap', content: DATA_SPACE_MAP },
    { path: '/\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', content: STRONG_ENCRYPTION_DATA_SPACE },
    { path: '/\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', content: PRIMARY },
    { path: '/EncryptionInfo', content: encryptionInfo },
  ]);
  // Defence in depth: the bytes we just wrote must pass the same guard we apply to untrusted input.
  assertSafeContainer(container, DEFAULT_CONTAINER_LIMITS);
  return container;
}

// Lower-level pieces, exported for tests and for callers that already hold the streams.
export { parseAgileInfo, encryptAgile, type EncryptOptions } from './agile.ts';
export { parseStandardInfo } from './standard.ts';
export { readContainer, writeContainer, assertSafeContainer } from './container.ts';
