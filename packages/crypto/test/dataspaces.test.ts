// The embedded \x06DataSpaces streams must equal the ones a real PowerPoint-encrypted deck carries, byte for byte.
import { describe, expect, test } from 'vitest';
import * as CFB from 'cfb';
import { DATA_SPACE_MAP, PRIMARY, STRONG_ENCRYPTION_DATA_SPACE, VERSION } from '../src/dataspaces.ts';
import { readFixture, POWERPOINT_AGILE_UNKNOWN_PW } from './fixtures.ts';

const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && Buffer.from(a).equals(Buffer.from(b));

describe('embedded DataSpaces constants', () => {
  test('match the streams in a genuine PowerPoint-encrypted file', () => {
    const cfb = CFB.read(readFixture(POWERPOINT_AGILE_UNKNOWN_PW), { type: 'buffer' });
    const get = (name: string) => new Uint8Array(CFB.find(cfb, name)!.content as ArrayLike<number>);
    expect(same(get('Version'), VERSION)).toBe(true);
    expect(same(get('DataSpaceMap'), DATA_SPACE_MAP)).toBe(true);
    expect(same(get('StrongEncryptionDataSpace'), STRONG_ENCRYPTION_DATA_SPACE)).toBe(true);
    expect(same(get('\u0006Primary'), PRIMARY)).toBe(true);
  });

  test('expected lengths (MS-OFFCRYPTO 2.1)', () => {
    expect(VERSION.length).toBe(76);
    expect(DATA_SPACE_MAP.length).toBe(112);
    expect(STRONG_ENCRYPTION_DATA_SPACE.length).toBe(64);
    expect(PRIMARY.length).toBe(200);
  });
});
