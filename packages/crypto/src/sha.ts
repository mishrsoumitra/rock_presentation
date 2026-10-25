// packages/crypto/src/sha.ts
// Synchronous SHA-512 and SHA-1 in plain TypeScript. MS-OFFCRYPTO key derivation runs the hash 50,000-100,000
// times over tiny inputs; done through the async `crypto.subtle.digest` that is one awaited microtask per
// iteration, which is far slower than the hashing itself. These are used only for that spin loop and are
// fuzz-tested against `subtle.digest` (see test/sha.test.ts). AES and HMAC still go through Web Crypto.

export type ShaName = 'SHA-1' | 'SHA-512';

export function shaDigestSize(name: ShaName): number {
  return name === 'SHA-512' ? 64 : 20;
}

export function shaHash(name: ShaName, data: Uint8Array): Uint8Array {
  return name === 'SHA-512' ? sha512(data) : sha1(data);
}

// --- SHA-1 (FIPS 180-4) ---------------------------------------------------------------------------------------

export function sha1(data: Uint8Array): Uint8Array {
  const bitLen = data.length * 8;
  // message + 0x80 + zero pad to 56 mod 64 + 8-byte big-endian length
  const withPad = ((data.length + 8) >> 6 << 6) + 64;
  const msg = new Uint8Array(withPad);
  msg.set(data);
  msg[data.length] = 0x80;
  const dv = new DataView(msg.buffer);
  // 64-bit length, but a part never reaches 2^32 bytes, so only the low word is non-zero.
  dv.setUint32(withPad - 4, bitLen >>> 0, false);
  dv.setUint32(withPad - 8, Math.floor(bitLen / 0x100000000), false);

  let h0 = 0x67452301, h1 = 0xefcdab89, h2 = 0x98badcfe, h3 = 0x10325476, h4 = 0xc3d2e1f0;
  const w = new Int32Array(80);

  for (let off = 0; off < withPad; off += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getInt32(off + i * 4, false);
    for (let i = 16; i < 80; i++) {
      const x = w[i - 3]! ^ w[i - 8]! ^ w[i - 14]! ^ w[i - 16]!;
      w[i] = (x << 1) | (x >>> 31);
    }
    let a = h0, b = h1, c = h2, d = h3, e = h4;
    for (let i = 0; i < 80; i++) {
      let f: number, k: number;
      if (i < 20) { f = (b & c) | (~b & d); k = 0x5a827999; }
      else if (i < 40) { f = b ^ c ^ d; k = 0x6ed9eba1; }
      else if (i < 60) { f = (b & c) | (b & d) | (c & d); k = 0x8f1bbcdc; }
      else { f = b ^ c ^ d; k = 0xca62c1d6; }
      const t = (((a << 5) | (a >>> 27)) + f + e + k + w[i]!) | 0;
      e = d; d = c; c = (b << 30) | (b >>> 2); b = a; a = t;
    }
    h0 = (h0 + a) | 0; h1 = (h1 + b) | 0; h2 = (h2 + c) | 0; h3 = (h3 + d) | 0; h4 = (h4 + e) | 0;
  }

  const out = new Uint8Array(20);
  const odv = new DataView(out.buffer);
  odv.setUint32(0, h0 >>> 0, false);
  odv.setUint32(4, h1 >>> 0, false);
  odv.setUint32(8, h2 >>> 0, false);
  odv.setUint32(12, h3 >>> 0, false);
  odv.setUint32(16, h4 >>> 0, false);
  return out;
}

// --- SHA-512 (FIPS 180-4), 64-bit words as (hi, lo) uint32 pairs ----------------------------------------------

// Round constants K[t], 80 64-bit values, stored as interleaved hi, lo.
const K512 = new Uint32Array([
  0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc,
  0x3956c25b, 0xf348b538, 0x59f111f1, 0xb605d019, 0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118,
  0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe, 0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2,
  0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235, 0xc19bf174, 0xcf692694,
  0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3, 0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65,
  0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5,
  0x983e5152, 0xee66dfab, 0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4,
  0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725, 0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70,
  0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed, 0x53380d13, 0x9d95b3df,
  0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b,
  0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001, 0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30,
  0xd192e819, 0xd6ef5218, 0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8,
  0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53, 0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8,
  0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb, 0x5b9cca4f, 0x7763e373, 0x682e6ff3, 0xd6b2b8a3,
  0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec,
  0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b,
  0xca273ece, 0xea26619c, 0xd186b8c7, 0x21c0c207, 0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178,
  0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6, 0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b,
  0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc, 0x431d67c4, 0x9c100d4c,
  0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a, 0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817,
]);

export function sha512(data: Uint8Array): Uint8Array {
  const bitLen = data.length * 8;
  // message + 0x80 + zero pad to 112 mod 128 + 16-byte big-endian length
  const withPad = ((data.length + 16) >> 7 << 7) + 128;
  const msg = new Uint8Array(withPad);
  msg.set(data);
  msg[data.length] = 0x80;
  const dv = new DataView(msg.buffer);
  dv.setUint32(withPad - 4, bitLen >>> 0, false);
  dv.setUint32(withPad - 8, Math.floor(bitLen / 0x100000000) >>> 0, false);

  // Initial hash values H0..H7, as hi/lo pairs.
  let h0h = 0x6a09e667, h0l = 0xf3bcc908, h1h = 0xbb67ae85, h1l = 0x84caa73b;
  let h2h = 0x3c6ef372, h2l = 0xfe94f82b, h3h = 0xa54ff53a, h3l = 0x5f1d36f1;
  let h4h = 0x510e527f, h4l = 0xade682d1, h5h = 0x9b05688c, h5l = 0x2b3e6c1f;
  let h6h = 0x1f83d9ab, h6l = 0xfb41bd6b, h7h = 0x5be0cd19, h7l = 0x137e2179;

  const wh = new Uint32Array(80);
  const wl = new Uint32Array(80);

  for (let off = 0; off < withPad; off += 128) {
    for (let i = 0; i < 16; i++) {
      wh[i] = dv.getUint32(off + i * 8, false);
      wl[i] = dv.getUint32(off + i * 8 + 4, false);
    }
    for (let i = 16; i < 80; i++) {
      // s0 = rotr(w[i-15],1) ^ rotr(w[i-15],8) ^ shr(w[i-15],7)
      const x15h = wh[i - 15]!, x15l = wl[i - 15]!;
      const s0h = ((x15h >>> 1) | (x15l << 31)) ^ ((x15h >>> 8) | (x15l << 24)) ^ (x15h >>> 7);
      const s0l = ((x15l >>> 1) | (x15h << 31)) ^ ((x15l >>> 8) | (x15h << 24)) ^ ((x15l >>> 7) | (x15h << 25));
      // s1 = rotr(w[i-2],19) ^ rotr(w[i-2],61) ^ shr(w[i-2],6)
      const x2h = wh[i - 2]!, x2l = wl[i - 2]!;
      const s1h = ((x2h >>> 19) | (x2l << 13)) ^ ((x2l >>> 29) | (x2h << 3)) ^ (x2h >>> 6);
      const s1l = ((x2l >>> 19) | (x2h << 13)) ^ ((x2h >>> 29) | (x2l << 3)) ^ ((x2l >>> 6) | (x2h << 26));
      // w[i] = w[i-16] + s0 + w[i-7] + s1
      const lo = (wl[i - 16]! >>> 0) + (s0l >>> 0) + (wl[i - 7]! >>> 0) + (s1l >>> 0);
      const hi = (wh[i - 16]! >>> 0) + (s0h >>> 0) + (wh[i - 7]! >>> 0) + (s1h >>> 0) + Math.floor(lo / 0x100000000);
      wl[i] = lo >>> 0;
      wh[i] = hi >>> 0;
    }

    let ah = h0h, al = h0l, bh = h1h, bl = h1l, ch = h2h, cl = h2l, dh = h3h, dl = h3l;
    let eh = h4h, el = h4l, fh = h5h, fl = h5l, gh = h6h, gl = h6l, hh = h7h, hl = h7l;

    for (let i = 0; i < 80; i++) {
      // S1 = rotr(e,14) ^ rotr(e,18) ^ rotr(e,41)
      const s1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const s1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      // ch = (e & f) ^ (~e & g)
      const chh = (eh & fh) ^ (~eh & gh);
      const chl = (el & fl) ^ (~el & gl);
      // T1 = h + S1 + ch + K[i] + w[i]
      let lo = (hl >>> 0) + (s1l >>> 0) + (chl >>> 0) + (K512[i * 2 + 1]! >>> 0) + (wl[i]! >>> 0);
      let hi = (hh >>> 0) + (s1h >>> 0) + (chh >>> 0) + (K512[i * 2]! >>> 0) + (wh[i]! >>> 0) + Math.floor(lo / 0x100000000);
      const t1l = lo >>> 0;
      const t1h = hi >>> 0;
      // S0 = rotr(a,28) ^ rotr(a,34) ^ rotr(a,39)
      const s0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const s0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      // maj = (a & b) ^ (a & c) ^ (b & c)
      const majh = (ah & bh) ^ (ah & ch) ^ (bh & ch);
      const majl = (al & bl) ^ (al & cl) ^ (bl & cl);
      // T2 = S0 + maj
      lo = (s0l >>> 0) + (majl >>> 0);
      hi = (s0h >>> 0) + (majh >>> 0) + Math.floor(lo / 0x100000000);
      const t2l = lo >>> 0;
      const t2h = hi >>> 0;

      hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
      // e = d + T1
      lo = (dl >>> 0) + (t1l >>> 0);
      hi = (dh >>> 0) + (t1h >>> 0) + Math.floor(lo / 0x100000000);
      eh = hi >>> 0; el = lo >>> 0;
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      // a = T1 + T2
      lo = (t1l >>> 0) + (t2l >>> 0);
      hi = (t1h >>> 0) + (t2h >>> 0) + Math.floor(lo / 0x100000000);
      ah = hi >>> 0; al = lo >>> 0;
    }

    const add = (xh: number, xl: number, yh: number, yl: number): [number, number] => {
      const lo = (xl >>> 0) + (yl >>> 0);
      const hi = (xh >>> 0) + (yh >>> 0) + Math.floor(lo / 0x100000000);
      return [hi >>> 0, lo >>> 0];
    };
    [h0h, h0l] = add(h0h, h0l, ah, al); [h1h, h1l] = add(h1h, h1l, bh, bl);
    [h2h, h2l] = add(h2h, h2l, ch, cl); [h3h, h3l] = add(h3h, h3l, dh, dl);
    [h4h, h4l] = add(h4h, h4l, eh, el); [h5h, h5l] = add(h5h, h5l, fh, fl);
    [h6h, h6l] = add(h6h, h6l, gh, gl); [h7h, h7l] = add(h7h, h7l, hh, hl);
  }

  const out = new Uint8Array(64);
  const odv = new DataView(out.buffer);
  const hs = [h0h, h0l, h1h, h1l, h2h, h2l, h3h, h3l, h4h, h4l, h5h, h5l, h6h, h6l, h7h, h7l];
  for (let i = 0; i < 16; i++) odv.setUint32(i * 4, hs[i]! >>> 0, false);
  return out;
}
