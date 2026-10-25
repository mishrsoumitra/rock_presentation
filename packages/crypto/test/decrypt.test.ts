// Decrypting real encrypted packages: the msoffcrypto-encrypted fixture (byte-for-byte against its plaintext) and
// files encrypted by Microsoft Office / Apache POI. Plus password and integrity failure paths.
import * as CFB from 'cfb';
import { describe, expect, test } from 'vitest';
import { decryptPackage, isEncryptedPackage, OfficeCryptoError, parseAgileInfo, readContainer } from '../src/index.ts';
import { ENCRYPTED_FIXTURES, POWERPOINT_AGILE_UNKNOWN_PW, readFixture } from './fixtures.ts';

const isZip = (b: Uint8Array) => b.length > 4 && b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04;

describe('isEncryptedPackage', () => {
  test('true for every encrypted fixture, false for a plain ZIP and for junk', () => {
    for (const f of ENCRYPTED_FIXTURES) expect(isEncryptedPackage(readFixture(f.file)), f.file).toBe(true);
    expect(isEncryptedPackage(readFixture(POWERPOINT_AGILE_UNKNOWN_PW))).toBe(true);
    expect(isEncryptedPackage(readFixture('plain/default.pptx'))).toBe(false);
    expect(isEncryptedPackage(new Uint8Array([1, 2, 3, 4]))).toBe(false);
  });
});

describe('decryptPackage on fixtures from other implementations', () => {
  for (const f of ENCRYPTED_FIXTURES) {
    test(`${f.scheme}: ${f.file}`, async () => {
      const out = await decryptPackage(readFixture(f.file), f.password);
      expect(isZip(out), 'decrypts to a ZIP').toBe(true);
      if (f.plain) expect(Buffer.from(out).equals(Buffer.from(readFixture(f.plain))), 'byte-for-byte vs original').toBe(true);
    });
  }

  test('the msoffcrypto fixture decrypts to its plaintext byte for byte', async () => {
    const out = await decryptPackage(readFixture('msoffcrypto/default-agile.pptx'), 'pptx-crypto-test-2026');
    expect(Buffer.from(out).equals(Buffer.from(readFixture('plain/default.pptx')))).toBe(true);
  });
});

describe('password and integrity failures', () => {
  test('wrong password throws code "wrong-password" (agile and standard)', async () => {
    for (const f of ENCRYPTED_FIXTURES) {
      await expect(decryptPackage(readFixture(f.file), 'definitely-not-it'), f.file).rejects.toMatchObject({
        name: 'OfficeCryptoError', code: 'wrong-password',
      });
    }
  });

  test('wrong password on a genuine PowerPoint deck fails cleanly, not with a crash', async () => {
    await expect(decryptPackage(readFixture(POWERPOINT_AGILE_UNKNOWN_PW), 'password')).rejects.toBeInstanceOf(OfficeCryptoError);
  });

  test('a tampered EncryptedPackage fails the agile integrity HMAC', async () => {
    // Rebuild the compound file with one byte of the EncryptedPackage flipped, keeping the verifier intact.
    const original = readFixture('msoffcrypto/default-agile.pptx');
    const container = readContainer(original);
    const pkg = container.stream('EncryptedPackage')!.slice();
    const last = pkg.length - 1;
    pkg[last] = pkg[last]! ^ 0xff; // flip a byte in the last segment
    const tampered = rebuildWithPackage(original, pkg);
    await expect(decryptPackage(tampered, 'pptx-crypto-test-2026')).rejects.toMatchObject({ code: 'integrity' });
    // With the integrity check off, the wrong bytes still come through (proving the HMAC was what caught it).
    const out = await decryptPackage(tampered, 'pptx-crypto-test-2026', { checkIntegrity: false });
    expect(out.length).toBeGreaterThan(0);
  });

  test('a corrupted verifier reads as a wrong password', async () => {
    const info = parseAgileInfo(readContainer(readFixture('msoffcrypto/default-agile.pptx')).stream('EncryptionInfo')!);
    expect(info.encryptedKey.keyBits).toBe(256);
    expect(info.encryptedKey.hashName).toBe('SHA-512');
    expect(info.spinCount).toBe(100000);
  });
});

// Reassemble a compound file identical to `source` but with a replaced EncryptedPackage stream.
function rebuildWithPackage(source: Uint8Array, pkg: Uint8Array): Uint8Array {
  const cfb = CFB.read(source, { type: 'buffer' });
  CFB.utils.cfb_add(cfb, 'EncryptedPackage', pkg);
  const out = CFB.write(cfb, { type: 'buffer' }) as Uint8Array;
  return new Uint8Array(out);
}
