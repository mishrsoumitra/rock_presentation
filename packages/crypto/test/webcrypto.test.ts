// The AES primitives built on Web Crypto: unpadded CBC and the ECB-via-CBC trick, checked against Node's crypto.
import { createCipheriv, randomBytes } from 'node:crypto';
import { describe, expect, test } from 'vitest';
import { aesCbcDecrypt, aesCbcEncrypt, aesEcbDecrypt, hmac } from '../src/webcrypto.ts';

const hex = (u: Uint8Array) => Buffer.from(u).toString('hex');
const u8 = (b: Buffer) => new Uint8Array(b);

describe('AES-CBC with no padding', () => {
  test('encrypt matches Node and decrypt is its inverse, for 128- and 256-bit keys', async () => {
    for (let t = 0; t < 40; t++) {
      const key = u8(randomBytes(t % 2 ? 32 : 16));
      const iv = u8(randomBytes(16));
      const data = u8(randomBytes(16 * (1 + (t % 6))));
      const c = createCipheriv(key.length === 32 ? 'aes-256-cbc' : 'aes-128-cbc', Buffer.from(key), Buffer.from(iv));
      c.setAutoPadding(false);
      const refCt = Buffer.concat([c.update(Buffer.from(data)), c.final()]);
      const ct = await aesCbcEncrypt(key, iv, data);
      expect(hex(ct)).toBe(refCt.toString('hex'));
      expect(hex(await aesCbcDecrypt(key, iv, ct))).toBe(hex(data));
    }
  });

  test('rejects a non-block-multiple length', async () => {
    await expect(aesCbcEncrypt(new Uint8Array(16), new Uint8Array(16), new Uint8Array(17))).rejects.toMatchObject({ code: 'malformed' });
  });
});

describe('AES-ECB decryption (via CBC with a zero IV)', () => {
  test('matches Node ECB for 128- and 256-bit keys', async () => {
    for (let t = 0; t < 40; t++) {
      const key = u8(randomBytes(t % 2 ? 32 : 16));
      const data = u8(randomBytes(16 * (1 + (t % 6))));
      const ec = createCipheriv(key.length === 32 ? 'aes-256-ecb' : 'aes-128-ecb', Buffer.from(key), null);
      ec.setAutoPadding(false);
      const ct = Buffer.concat([ec.update(Buffer.from(data)), ec.final()]);
      expect(hex(await aesEcbDecrypt(key, u8(ct)))).toBe(hex(data));
    }
  });
});

describe('HMAC', () => {
  test('HMAC-SHA512 and HMAC-SHA1 match Node', async () => {
    const { createHmac } = await import('node:crypto');
    const key = u8(randomBytes(64));
    const msg = u8(randomBytes(9000));
    expect(hex(await hmac('SHA-512', key, msg))).toBe(createHmac('sha512', Buffer.from(key)).update(Buffer.from(msg)).digest('hex'));
    expect(hex(await hmac('SHA-1', key, msg))).toBe(createHmac('sha1', Buffer.from(key)).update(Buffer.from(msg)).digest('hex'));
  });
});
