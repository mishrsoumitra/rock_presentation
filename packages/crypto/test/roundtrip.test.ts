// Our encrypt -> our decrypt, and the shape of what we write (DataSpaces streams, descriptor, versions).
import { describe, expect, test } from 'vitest';
import * as CFB from 'cfb';
import {
  decryptPackage, encryptPackage, isEncryptedPackage, parseAgileInfo, readContainer,
} from '../src/index.ts';
import { DATA_SPACE_MAP, PRIMARY, STRONG_ENCRYPTION_DATA_SPACE, VERSION } from '../src/dataspaces.ts';
import { readFixture } from './fixtures.ts';

const sameBytes = (a: Uint8Array, b: Uint8Array) => a.length === b.length && Buffer.from(a).equals(Buffer.from(b));

describe('encrypt -> decrypt round trip', () => {
  test('a real deck round-trips (low spinCount for speed)', async () => {
    const zip = readFixture('plain/default.pptx');
    const enc = await encryptPackage(zip, 'correct horse battery staple', { spinCount: 1000 });
    expect(isEncryptedPackage(enc)).toBe(true);
    expect(sameBytes(await decryptPackage(enc, 'correct horse battery staple'), zip)).toBe(true);
  });

  test('payload sizes around the 4096-byte segment boundary and empty input', async () => {
    for (const n of [0, 1, 15, 16, 17, 4095, 4096, 4097, 8192, 10000]) {
      const data = new Uint8Array(n).map((_, i) => (i * 37 + 11) & 0xff);
      const enc = await encryptPackage(data, 'pw', { spinCount: 1 });
      expect(sameBytes(await decryptPackage(enc, 'pw'), data), `size ${n}`).toBe(true);
    }
  });

  test('a wrong password on our own output is rejected', async () => {
    const enc = await encryptPackage(readFixture('plain/default.pptx'), 'right', { spinCount: 1 });
    await expect(decryptPackage(enc, 'wrong')).rejects.toMatchObject({ code: 'wrong-password' });
  });

  test('fresh salts make two encryptions of the same input differ', async () => {
    const zip = readFixture('plain/default.pptx');
    const a = await encryptPackage(zip, 'pw', { spinCount: 1 });
    const b = await encryptPackage(zip, 'pw', { spinCount: 1 });
    expect(sameBytes(a, b)).toBe(false);
  });
});

describe('what encryptPackage writes', () => {
  test('container holds the expected streams with the fixed DataSpaces bytes', async () => {
    const enc = await encryptPackage(readFixture('plain/default.pptx'), 'pw', { spinCount: 1 });
    const cfb = CFB.read(enc, { type: 'buffer' });
    const paths = cfb.FullPaths.filter(p => !p.endsWith('/'));
    for (const p of ['EncryptedPackage', 'EncryptionInfo', '\u0006DataSpaces/Version', '\u0006DataSpaces/DataSpaceMap',
      '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace',
      '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary']) {
      expect(paths.some(x => x.endsWith(p)), `has ${p}`).toBe(true);
    }
    const get = (name: string) => new Uint8Array(CFB.find(cfb, name)!.content as ArrayLike<number>);
    expect(sameBytes(get('Version'), VERSION)).toBe(true);
    expect(sameBytes(get('DataSpaceMap'), DATA_SPACE_MAP)).toBe(true);
    expect(sameBytes(get('StrongEncryptionDataSpace'), STRONG_ENCRYPTION_DATA_SPACE)).toBe(true);
    expect(sameBytes(get('\u0006Primary'), PRIMARY)).toBe(true);
  });

  test('the EncryptionInfo we write is agile AES-256 / SHA-512 and re-parses', async () => {
    const enc = await encryptPackage(readFixture('plain/default.pptx'), 'pw', { spinCount: 1234 });
    const info = parseAgileInfo(readContainer(enc).stream('EncryptionInfo')!);
    expect(info.encryptedKey.keyBits).toBe(256);
    expect(info.keyData.keyBits).toBe(256);
    expect(info.encryptedKey.hashName).toBe('SHA-512');
    expect(info.spinCount).toBe(1234);
    expect(info.keyData.hashSize).toBe(64);
  });
});
