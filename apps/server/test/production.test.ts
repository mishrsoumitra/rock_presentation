// What a deployment adds: API tokens, proxies, and the editor served next to the API.
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { brotliDecompressSync, gunzipSync } from 'node:zlib';
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { configFromEnv } from '../src/config.ts';
import { cleanup, deck, tempDir, testServer, upload } from './helpers.ts';

const TOKEN = 'a'.repeat(40);
const CSP = "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; object-src 'none'";

afterAll(cleanup);

describe('settings from the environment', () => {
  test('a HOST on the network needs API_TOKENS, or ALLOW_ANONYMOUS_API=1', () => {
    expect(() => configFromEnv({ HOST: '0.0.0.0' })).toThrow(/API_TOKENS/);
    expect(configFromEnv({ HOST: '0.0.0.0', ALLOW_ANONYMOUS_API: '1' }).apiTokens).toEqual([]);
    expect(configFromEnv({ HOST: '0.0.0.0', API_TOKENS: `${TOKEN}, ${'b'.repeat(32)}` }).apiTokens).toEqual([TOKEN, 'b'.repeat(32)]);
    for (const host of ['127.0.0.1', 'localhost', '::1', '127.0.0.2']) expect(configFromEnv({ HOST: host }).host).toBe(host);
    expect(configFromEnv({}).host).toBe('127.0.0.1');
  });

  test('tokens must be long, without spaces', () => {
    expect(() => configFromEnv({ API_TOKENS: 'short' })).toThrow(/at least 32/);
    expect(() => configFromEnv({ API_TOKENS: `${TOKEN} x` })).toThrow(/without spaces/);
  });

  test('TRUST_PROXY is a hop count, true, or addresses', () => {
    expect(configFromEnv({}).trustProxy).toBe(false);
    expect(configFromEnv({ TRUST_PROXY: '2' }).trustProxy).toBe(2);
    expect(configFromEnv({ TRUST_PROXY: 'true' }).trustProxy).toBe(true);
    expect(configFromEnv({ TRUST_PROXY: '10.0.0.0/8, 127.0.0.1' }).trustProxy).toEqual(['10.0.0.0/8', '127.0.0.1']);
  });
});

describe('API tokens', () => {
  let app: FastifyInstance;
  beforeAll(async () => { app = await testServer({ apiTokens: [TOKEN] }); });
  afterAll(async () => { await app.close(); });

  test('every API route but the health check wants one', async () => {
    expect((await app.inject({ url: '/api/health' })).statusCode).toBe(200);
    const anonymous = await upload(app, deck(1));
    expect([anonymous.statusCode, anonymous.json().error, anonymous.headers['www-authenticate']]).toEqual([401, 'unauthorized', 'Bearer']);
    const wrong = await upload(app, deck(1), { headers: { authorization: `Bearer ${'b'.repeat(40)}` } });
    expect(wrong.statusCode).toBe(401);
    const right = await upload(app, deck(1), { headers: { authorization: `Bearer ${TOKEN}` } });
    expect(right.statusCode).toBe(201);
    const id = right.json().id as string;
    expect((await app.inject({ url: `/api/decks/${id}` })).statusCode).toBe(401);
    expect((await app.inject({ url: `/api/decks/${id}`, headers: { authorization: `bearer ${TOKEN}` } })).statusCode).toBe(200);
    expect((await app.inject({ url: `/api/decks/${id}/file`, headers: { authorization: TOKEN } })).statusCode).toBe(401);
  });
});

describe('the editor, served next to the API', () => {
  let dir: string;
  let app: FastifyInstance;
  const script = `console.log(${JSON.stringify('x'.repeat(4000))});`;
  beforeAll(async () => {
    dir = tempDir('pptx-web-');
    mkdirSync(join(dir, 'assets'));
    mkdirSync(join(dir, 'fonts/carlito'), { recursive: true });
    writeFileSync(join(dir, 'index.html'), `<!doctype html><html><head>\n<meta http-equiv="Content-Security-Policy" content="${CSP}"><script type="module" src="/assets/index-abc123.js"></script></head></html>`);
    writeFileSync(join(dir, 'assets/index-abc123.js'), script);
    writeFileSync(join(dir, 'assets/index-abc123.js.map'), '{}');
    writeFileSync(join(dir, 'fonts/carlito/Carlito-Regular.ttf'), new Uint8Array(64));
    writeFileSync(join(dir, '.env'), 'SECRET=1');
    app = await testServer({ webDir: dir, apiTokens: [TOKEN] });
  });
  afterAll(async () => { await app.close(); });

  test('index.html, with its policy as a header too, and no token needed', async () => {
    const res = await app.inject({ url: '/' });
    expect(res.statusCode).toBe(200);
    expect(res.headers['content-type']).toBe('text/html; charset=utf-8');
    expect(res.headers['content-security-policy']).toBe(`${CSP}; frame-ancestors 'none'`);
    expect(res.headers['cache-control']).toBe('no-cache');
    expect(res.headers['x-frame-options']).toBe('DENY');
    expect((await app.inject({ method: 'HEAD', url: '/' })).statusCode).toBe(200);
  });

  test('hashed assets: cached for good, compressed, and 304 when unchanged', async () => {
    const plain = await app.inject({ url: '/assets/index-abc123.js' });
    expect([plain.statusCode, plain.headers['cache-control'], plain.body]).toEqual([200, 'public, max-age=31536000, immutable', script]);
    // a script may start a worker, which takes its policy from this response
    expect(plain.headers['content-security-policy']).toBe(`${CSP}; frame-ancestors 'none'`);
    const br = await app.inject({ url: '/assets/index-abc123.js', headers: { 'accept-encoding': 'gzip, br' } });
    expect(br.headers['content-encoding']).toBe('br');
    expect(brotliDecompressSync(br.rawPayload).toString()).toBe(script);
    const gz = await app.inject({ url: '/assets/index-abc123.js', headers: { 'accept-encoding': 'gzip' } });
    expect(gunzipSync(gz.rawPayload).toString()).toBe(script);
    const again = await app.inject({ url: '/assets/index-abc123.js', headers: { 'if-none-match': plain.headers.etag as string } });
    expect(again.statusCode).toBe(304);
    expect((await app.inject({ url: '/fonts/carlito/Carlito-Regular.ttf' })).headers['content-type']).toBe('font/ttf');
  });

  test('nothing else: no source maps, dotfiles, paths out of the build, or API paths', async () => {
    for (const url of ['/assets/index-abc123.js.map', '/.env', '/../package.json', '/%2e%2e/package.json', '/assets/../../package.json', '/nope']) {
      expect((await app.inject({ url })).statusCode, url).toBe(404);
    }
    const api = await app.inject({ url: '/api/nope', headers: { authorization: `Bearer ${TOKEN}` } });
    expect([api.statusCode, api.json().error, api.json().message]).toEqual([404, 'not-found', 'No such route.']);
  });

  test('without a build, only the API is served; a build without its policy is refused', async () => {
    const bare = await testServer({ webDir: tempDir('pptx-no-web-') });
    expect((await bare.inject({ url: '/' })).statusCode).toBe(404);
    await bare.close();
    const unsafe = tempDir('pptx-unsafe-web-');
    writeFileSync(join(unsafe, 'index.html'), '<!doctype html><title>x</title>');
    await expect(testServer({ webDir: unsafe })).rejects.toThrow(/Content-Security-Policy/);
  });
});
