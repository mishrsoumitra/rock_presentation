// Mutation fuzzing: seeded random edits of real and hand-made metafiles must parse and play, or fail with a
// MetafileError, never anything else, and quickly. Seeded, so a failure reproduces from its case number.
import { expect, test } from 'vitest';
import { drawMetafile, MetafileError, parseMetafile } from '../src/index.ts';
import { corpusMetafiles } from './corpus.ts';
import { MockCanvas, RecordingContext } from './mock.ts';
import { bitmaps, plusOnly, textAndClip, wmfWithEmf } from './samples.ts';

function rng(seed: number): () => number {             // mulberry32
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// values that break sizes, counts and offsets
const NASTY = [0, 1, 2, 3, 4, 7, 8, 0xff, 0x100, 0x7fff, 0x8000, 0xffff, 0x10000, 0x7fffffff, 0x80000000, 0xfffffffe, 0xffffffff, 0x3fffffff, 16_384, 16_385];

export function mutate(src: Uint8Array, rnd: () => number): Uint8Array {
  let b: Uint8Array = src.slice();
  const edits = 1 + Math.floor(rnd() * 4);
  for (let k = 0; k < edits; k++) {
    const at = Math.floor(rnd() * b.length);
    const r = rnd();
    if (r < 0.3) {                                                   // flip a few bytes
      for (let i = 0; i < 1 + Math.floor(rnd() * 4); i++) b[Math.min(b.length - 1, at + i)] = Math.floor(rnd() * 256);
    } else if (r < 0.6) {                                            // a nasty 32-bit (or 16-bit) value
      const v = NASTY[Math.floor(rnd() * NASTY.length)]!;
      const p = at & ~(rnd() < 0.5 ? 3 : 1);
      for (let i = 0; i < 4 && p + i < b.length; i++) b[p + i] = (v >>> (i * 8)) & 0xff;
    } else if (r < 0.72) {                                           // truncate
      b = b.subarray(0, at);
    } else if (r < 0.82) {                                           // delete a run
      const n = 1 + Math.floor(rnd() * 16);
      b = Uint8Array.from([...b.subarray(0, at), ...b.subarray(at + n)]);
    } else if (r < 0.92) {                                           // insert random bytes
      const n = 1 + Math.floor(rnd() * 8);
      b = Uint8Array.from([...b.subarray(0, at), ...Array.from({ length: n }, () => Math.floor(rnd() * 256)), ...b.subarray(at)]);
    } else {                                                         // copy a slice over another place
      const from = Math.floor(rnd() * b.length), n = 1 + Math.floor(rnd() * 64);
      b.set(b.slice(from, from + n).subarray(0, Math.max(0, b.length - at)), at);
    }
    if (b.length === 0) break;
  }
  return b;
}

test('mutated metafiles parse and play, or fail with MetafileError only; every case is quick', () => {
  const seeds = [
    plusOnly(), wmfWithEmf(), bitmaps(), textAndClip(),
    ...corpusMetafiles().map(p => p.bytes).filter(b => b.length <= 120_000),
  ];
  const rnd = rng(0x5eed);
  let parsed = 0, played = 0, refused = 0, cases = 0, slowest = 0;
  const outcomes = new Map<string, number>();
  const deadline = performance.now() + 20_000;
  for (let round = 0; cases < 6000 && performance.now() < deadline; round++) {
    for (const seed of seeds) {
      const input = mutate(seed, rnd);
      const started = performance.now();
      cases++;
      try {
        const mf = parseMetafile(input);
        parsed++;
        const ctx = new RecordingContext(300, 300);
        drawMetafile(ctx, mf, { x: 0, y: 0, w: 300, h: 300 }, { createCanvas: (w, h) => new MockCanvas(w, h), maxWork: 2_000_000, snap: true });
        played++;
        expect(ctx.depth).toBe(0);
      } catch (e) {
        if (!(e instanceof MetafileError)) {
          throw new Error(`case ${cases} (round ${round}, seed ${seeds.indexOf(seed)}): ${e instanceof Error ? e.stack : String(e)}`);
        }
        refused++;
        outcomes.set(e.code, (outcomes.get(e.code) ?? 0) + 1);
      }
      slowest = Math.max(slowest, performance.now() - started);
      if (cases >= 6000) break;
    }
  }
  expect(cases).toBeGreaterThan(2000);
  expect(played).toBeGreaterThan(cases / 4);
  expect(refused).toBeGreaterThan(50);
  expect(parsed).toBeGreaterThanOrEqual(played);
  expect(slowest).toBeLessThan(2000);
}, 60_000);

test('hostile sizes are refused before anything big is allocated', () => {
  // a 4-byte DIB claiming 10000 x 6000 pixels
  const header = new Uint8Array(40);
  const dv = new DataView(header.buffer);
  dv.setUint32(0, 40, true); dv.setInt32(4, 10_000, true); dv.setInt32(8, 6_000, true); dv.setUint16(12, 1, true); dv.setUint16(14, 32, true);
  const rec = new Uint8Array(100 + 40 + 4);
  const r = new DataView(rec.buffer);
  r.setUint32(0, 76, true); r.setUint32(4, rec.length, true);
  r.setInt32(32, 10, true); r.setInt32(36, 10, true); r.setUint32(40, 0xcc0020, true);
  r.setFloat32(52, 1, true); r.setFloat32(64, 1, true);
  r.setUint32(84, 100, true); r.setUint32(88, 40, true); r.setUint32(92, 140, true); r.setUint32(96, 4, true);
  rec.set(header, 100);
  const head = new Uint8Array(108);
  const h = new DataView(head.buffer);
  h.setUint32(0, 1, true); h.setUint32(4, 108, true); h.setInt32(32, 2540, true); h.setInt32(36, 2540, true);
  h.setUint32(40, 0x464d4520, true); h.setInt32(72, 1000, true); h.setInt32(76, 1000, true); h.setInt32(80, 254, true); h.setInt32(84, 254, true);
  const bytes = Uint8Array.from([...head, ...rec]);
  const ctx = new RecordingContext();
  const started = performance.now();
  drawMetafile(ctx, parseMetafile(bytes), { x: 0, y: 0, w: 100, h: 100 }, { createCanvas: (w, hh) => new MockCanvas(w, hh) });
  expect(ctx.of('drawImage')).toHaveLength(0);
  expect(performance.now() - started).toBeLessThan(500);
});

test('a million tiny records hit the record limit, not the memory', () => {
  // EMR_SAVEDC is 8 bytes: 1,000,001 of them after the header
  const n = 1_000_001;
  const bytes = new Uint8Array(108 + n * 8);
  const h = new DataView(bytes.buffer);
  h.setUint32(0, 1, true); h.setUint32(4, 108, true); h.setInt32(32, 2540, true); h.setInt32(36, 2540, true);
  h.setUint32(40, 0x464d4520, true); h.setInt32(72, 1000, true); h.setInt32(76, 1000, true); h.setInt32(80, 254, true); h.setInt32(84, 254, true);
  for (let i = 0; i < n; i++) { h.setUint32(108 + i * 8, 33, true); h.setUint32(112 + i * 8, 8, true); }
  expect(() => parseMetafile(bytes)).toThrow(expect.objectContaining({ code: 'limit' }));
});
