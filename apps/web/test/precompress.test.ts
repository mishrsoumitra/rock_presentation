// apps/web/test/precompress.test.ts  -- the build's compressed copies (apps/web/precompress.ts)
import { randomBytes } from 'node:crypto';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { brotliDecompressSync, gunzipSync } from 'node:zlib';
import { expect, test } from 'vitest';
import { compressible, precompressDir } from '../precompress.ts';

test('which files get copies: compressible types over 1 KB, never maps or the copies themselves', () => {
  expect(compressible('index-abc.js', 2000)).toBe(true);
  expect(compressible('Carlito-Regular.ttf', 600_000)).toBe(true);
  expect(compressible('harfbuzz-abc.wasm', 441_359)).toBe(true);
  expect(compressible('index.html', 740)).toBe(false);
  expect(compressible('index-abc.js.map', 5_000_000)).toBe(false);
  expect(compressible('index-abc.js.br', 2000)).toBe(false);
  expect(compressible('index-abc.js.gz', 2000)).toBe(false);
  expect(compressible('photo.png', 20_000)).toBe(false);
});

test('writes a Brotli and a gzip copy next to each file worth them, spread over worker threads', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'pptx-precompress-'));
  try {
    mkdirSync(join(dir, 'assets'));
    mkdirSync(join(dir, 'fonts'));
    const script = `console.log(${JSON.stringify('x'.repeat(4000))});`;
    const style = 'body { margin: 0 }\n'.repeat(100);
    writeFileSync(join(dir, 'assets/index-abc.js'), script);
    writeFileSync(join(dir, 'assets/index-abc.js.map'), '{}'.repeat(2000));
    writeFileSync(join(dir, 'assets/index-def.css'), style);
    writeFileSync(join(dir, 'fonts/noise.ttf'), randomBytes(8192));        // a copy would not save a tenth
    writeFileSync(join(dir, 'fonts/.hidden.js'), 'x'.repeat(2000));         // dotfiles are never served
    writeFileSync(join(dir, 'index.html'), '<!doctype html><title>x</title>');
    expect(await precompressDir(dir, 3)).toBe(4);
    expect(brotliDecompressSync(readFileSync(join(dir, 'assets/index-abc.js.br'))).toString()).toBe(script);
    expect(gunzipSync(readFileSync(join(dir, 'assets/index-abc.js.gz'))).toString()).toBe(script);
    expect(brotliDecompressSync(readFileSync(join(dir, 'assets/index-def.css.br'))).toString()).toBe(style);
    expect(gunzipSync(readFileSync(join(dir, 'assets/index-def.css.gz'))).toString()).toBe(style);
    for (const left of ['assets/index-abc.js.map.br', 'fonts/noise.ttf.br', 'fonts/noise.ttf.gz', 'fonts/.hidden.js.br', 'index.html.br', 'assets/index-abc.js.br.br']) {
      expect(existsSync(join(dir, left)), left).toBe(false);
    }
    // a second run makes the same copies again, and nothing more
    expect(await precompressDir(dir, 1)).toBe(4);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
