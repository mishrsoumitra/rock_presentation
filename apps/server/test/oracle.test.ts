// The CI oracle routes: absent unless enabled; compare (O3), roundtrip, the validator CLI (O1) and the PowerPoint
// bridge, the last two against stand-ins (a script that answers like tools/validator, a local HTTP stub).
import { writeFileSync } from 'node:fs';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { join } from 'node:path';
import type { FastifyInstance } from 'fastify';
import pngjs from 'pngjs';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { cleanup, deck, form, readRepoFile, tempDir, testServer, upload } from './helpers.ts';

const ROUTES = ['validate', 'compare', 'roundtrip', 'open', 'export', 'text-metrics', 'resave'];

/** A width x height PNG, white, with a black square of `square` pixels at the top left. */
function png(width: number, height: number, square = 0, options: { interlace?: boolean } = {}): Buffer {
  const img = new pngjs.PNG({ width, height });
  img.data.fill(255);
  for (let y = 0; y < square; y++) for (let x = 0; x < square; x++) img.data.fill(0, (y * width + x) * 4, (y * width + x) * 4 + 3);
  const out = pngjs.PNG.sync.write(img);
  if (options.interlace) out[28] = 1;     // the IHDR interlace flag (the server refuses it before decoding)
  return out;
}

async function compare(app: FastifyInstance, expected: Uint8Array, actual: Uint8Array, query = '') {
  const body = await form({ expected: { data: expected, filename: 'a.png', type: 'image/png' }, actual: { data: actual, filename: 'b.png', type: 'image/png' } });
  return app.inject({ method: 'POST', url: `/api/oracle/compare${query}`, ...body });
}

afterAll(() => cleanup());

test('without ENABLE_ORACLES every oracle route is 404', async () => {
  const app = await testServer();
  try {
    for (const name of ROUTES) {
      const res = await upload(app, deck(1), { url: `/api/oracle/${name}` });
      expect(res.statusCode, name).toBe(404);
      expect(res.json().error).toBe('not-found');
    }
  } finally {
    await app.close();
  }
});

describe('with ENABLE_ORACLES', () => {
  let app: FastifyInstance;
  beforeAll(async () => {
    app = await testServer({ oracles: { enabled: true } });
  });
  afterAll(async () => {
    await app.close();
  });

  test('compare: identical images match; a 10 x 10 square is 100 pixels; the diff image decodes', async () => {
    const same = await compare(app, png(200, 100), png(200, 100));
    expect(same.statusCode, same.body).toBe(200);
    expect(same.json()).toMatchObject({ width: 200, height: 100, mismatched: 0, ratio: 0, threshold: 0.1 });

    const res = await compare(app, png(200, 100), png(200, 100, 10), '?threshold=0.05');
    expect(res.json()).toMatchObject({ mismatched: 100, ratio: 100 / 20_000, threshold: 0.05 });
    const diff = pngjs.PNG.sync.read(Buffer.from(res.json().diffPng, 'base64'));
    expect([diff.width, diff.height]).toEqual([200, 100]);
    expect((await compare(app, png(20, 20), png(20, 20, 5), '?diff=false')).json().diffPng).toBeUndefined();
  });

  test('compare refuses images it cannot compare', async () => {
    const sizes = await compare(app, png(20, 10), png(10, 20));
    expect([sizes.statusCode, sizes.json().error]).toEqual([422, 'size-mismatch']);
    const notPng = await compare(app, png(10, 10), new TextEncoder().encode('not an image at all, honestly'));
    expect([notPng.statusCode, notPng.json().error]).toEqual([422, 'not-a-png']);
    const interlaced = await compare(app, png(10, 10), png(10, 10, 0, { interlace: true }));
    expect([interlaced.statusCode, interlaced.json().error]).toEqual([422, 'unsupported-png']);
    const huge = png(10, 10);
    huge.writeUInt32BE(100_000, 16);              // IHDR claims 100,000 x 10: refused before pngjs allocates
    huge.writeUInt32BE(100_000, 20);
    const big = await compare(app, png(10, 10), huge);
    expect([big.statusCode, big.json().error]).toEqual([413, 'limit']);
    const one = await form({ expected: { data: png(10, 10) } });
    expect((await app.inject({ method: 'POST', url: '/api/oracle/compare', ...one })).json().error).toBe('missing-file');
  });

  test('roundtrip: an OpcPackage re-save keeps every part', async () => {
    const res = await upload(app, readRepoFile('packages/crypto/test/fixtures/plain/default.pptx'), { url: '/api/oracle/roundtrip' });
    expect(res.statusCode, res.body).toBe(200);
    expect(res.json()).toMatchObject({ lossless: true, reserializedLossless: true, differences: [] });
    expect(res.json().xmlParts).toBeGreaterThan(10);
    const encrypted = await upload(app, readRepoFile('packages/crypto/test/fixtures/msoffcrypto/default-agile.pptx'), { url: '/api/oracle/roundtrip' });
    expect([encrypted.statusCode, encrypted.json().error]).toEqual([409, 'encrypted']);
  });

  test('validate is 501 without PPTX_VALIDATOR; the PowerPoint routes are 501 without POWERPOINT_BRIDGE_URL', async () => {
    const res = await upload(app, deck(1), { url: '/api/oracle/validate' });
    expect([res.statusCode, res.json().error]).toEqual([501, 'not-configured']);
    expect(res.json().message).toContain('PPTX_VALIDATOR');
    for (const name of ['open', 'export', 'text-metrics', 'resave']) {
      const r = await upload(app, deck(1), { url: `/api/oracle/${name}` });
      expect([r.statusCode, r.json().error], name).toEqual([501, 'not-configured']);
      expect(r.json().message).toContain('POWERPOINT_BRIDGE_URL');
    }
  });
});

describe('the validator CLI and the PowerPoint bridge, against stand-ins', () => {
  let app: FastifyInstance;
  let bridge: Server;
  const seen: { url: string; type: string; bytes: number }[] = [];
  beforeAll(async () => {
    // answers as tools/validator does: a JSON array of errors and exit code 1 (2 and {error} for a file it cannot open)
    const dir = tempDir();
    const script = join(dir, 'validator.mjs');
    writeFileSync(script, `
      import { readFileSync } from 'node:fs';
      const [file, format] = process.argv.slice(2);
      const bytes = readFileSync(file);
      if (bytes[0] !== 0x50) { console.log(JSON.stringify({ error: 'not a package' })); process.exit(2); }
      console.log(JSON.stringify([{ Id: 'Sem_Stand_In', Description: 'checked ' + bytes.length + ' bytes for ' + (format ?? 'Microsoft365'), Part: '/ppt/presentation.xml', Type: 'Semantic' }]));
      process.exit(1);
    `);
    bridge = createServer((req, res) => {
      const chunks: Buffer[] = [];
      req.on('data', (c: Buffer) => chunks.push(c));
      req.on('end', () => {
        seen.push({ url: req.url ?? '', type: req.headers['content-type'] ?? '', bytes: Buffer.concat(chunks).length });
        if (req.url === '/oracle/export') {
          res.writeHead(200, { 'content-type': 'image/png' }).end(png(4, 4));
        } else {
          res.writeHead(req.url === '/oracle/open' ? 200 : 422, { 'content-type': 'application/json' }).end(JSON.stringify({ opened: req.url === '/oracle/open' }));
        }
      });
    });
    await new Promise<void>(resolve => bridge.listen(0, '127.0.0.1', resolve));
    const port = (bridge.address() as AddressInfo).port;
    app = await testServer({ oracles: { enabled: true, validator: [process.execPath, script], bridgeUrl: `http://127.0.0.1:${port}/` } });
  });
  afterAll(async () => {
    await app.close();
    await new Promise(resolve => bridge.close(resolve));
  });

  test('validate runs the program on a private copy of the deck and reports its errors', async () => {
    const bytes = deck(2);
    const res = await upload(app, bytes, { url: '/api/oracle/validate?format=Office2019' });
    expect(res.statusCode, res.body).toBe(200);
    expect(res.json()).toMatchObject({ valid: false, exitCode: 1, errorCount: 1 });
    expect(res.json().errors[0]).toMatchObject({ Id: 'Sem_Stand_In', Description: `checked ${bytes.length} bytes for Office2019` });
    // hostile input never reaches the validator: our own parser refuses it first
    expect((await upload(app, new TextEncoder().encode('junk'), { url: '/api/oracle/validate' })).json().error).toBe('not-a-package');
    expect((await upload(app, deck(1), { url: '/api/oracle/validate?format=x;rm' })).statusCode).toBe(400);
  });

  test('the PowerPoint routes are forwarded to the bridge as they came, and its answer relayed', async () => {
    const bytes = deck(1);
    const open = await upload(app, bytes, { url: '/api/oracle/open' });
    expect(open.statusCode).toBe(200);
    expect(open.json()).toEqual({ opened: true });
    expect(seen.at(-1)).toMatchObject({ url: '/oracle/open' });
    expect(seen.at(-1)!.type).toContain('multipart/form-data');
    expect(seen.at(-1)!.bytes).toBeGreaterThan(bytes.length);

    const exported = await upload(app, bytes, { url: '/api/oracle/export' });
    expect(exported.statusCode).toBe(200);
    expect(exported.headers['content-type']).toBe('image/png');
    expect(exported.headers['content-disposition']).toContain('attachment');

    const refused = await upload(app, bytes, { url: '/api/oracle/resave' });
    expect(refused.statusCode).toBe(422);
    expect(seen.map(s => s.url)).toEqual(['/oracle/open', '/oracle/export', '/oracle/resave']);
  });
});
