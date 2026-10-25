// SHA-1 and SHA-512 (sha.ts) against Node's crypto, which is the same primitive as crypto.subtle.digest. The spin
// loop depends on these being bit-exact, so this is the test the brief calls for.
import { createHash, webcrypto } from 'node:crypto';
import { describe, expect, test } from 'vitest';
import { sha1, sha512, shaHash } from '../src/sha.ts';

const hex = (u: Uint8Array) => Buffer.from(u).toString('hex');
const ref = (algo: 'sha1' | 'sha512', d: Uint8Array) => createHash(algo).update(Buffer.from(d)).digest('hex');

describe('pure-TS hashes match the platform hash', () => {
  test('known vectors', () => {
    expect(hex(sha1(new Uint8Array(0)))).toBe('da39a3ee5e6b4b0d3255bfef95601890afd80709');
    expect(hex(sha1(new TextEncoder().encode('abc')))).toBe('a9993e364706816aba3e25717850c26c9cd0d89d');
    expect(hex(sha512(new Uint8Array(0)))).toBe(
      'cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e',
    );
    expect(hex(sha512(new TextEncoder().encode('abc')))).toBe(
      'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
    );
  });

  test('every length from 0 to 400 bytes, both algorithms', () => {
    for (let len = 0; len <= 400; len++) {
      const d = new Uint8Array(len);
      for (let i = 0; i < len; i++) d[i] = (i * 131 + len * 17) & 0xff;
      expect(hex(sha1(d)), `sha1 len ${len}`).toBe(ref('sha1', d));
      expect(hex(sha512(d)), `sha512 len ${len}`).toBe(ref('sha512', d));
    }
  });

  test('inputs that straddle the block boundary (SHA-1 64, SHA-512 128)', () => {
    for (const len of [55, 56, 57, 63, 64, 65, 111, 112, 113, 127, 128, 129, 239, 240, 241]) {
      const d = new Uint8Array(len).map((_, i) => (i * 7 + 3) & 0xff);
      expect(hex(sha1(d))).toBe(ref('sha1', d));
      expect(hex(sha512(d))).toBe(ref('sha512', d));
    }
  });

  test('shaHash dispatch matches crypto.subtle.digest', async () => {
    const d = new TextEncoder().encode('The quick brown fox jumps over the lazy dog');
    expect(hex(shaHash('SHA-512', d))).toBe(hex(new Uint8Array(await webcrypto.subtle.digest('SHA-512', d))));
    expect(hex(shaHash('SHA-1', d))).toBe(hex(new Uint8Array(await webcrypto.subtle.digest('SHA-1', d))));
  });
});
