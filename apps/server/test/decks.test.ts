// The deck API through fastify.inject: upload, metadata, download, versions with If-Match, delete, refusals, headers.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { corpusName, opcFiles } from '../../../tools/corpus/list.ts';
import { configFromEnv } from '../src/config.ts';
import { describeError, validateDeck } from '../src/tasks.ts';
import { cleanup, compoundFile, corpusFile, deck, form, readRepoFile, tempDir, testServer, upload, zip } from './helpers.ts';

const PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

let app: FastifyInstance;
let dataDir: string;
beforeAll(async () => {
  dataDir = tempDir();
  app = await testServer({ dataDir });
});
afterAll(async () => {
  await app.close();
  cleanup();
});

const blobs = () => readdirSync(join(dataDir, 'blobs'), { recursive: true }).filter(n => /^[0-9a-f]{2}\/[0-9a-f]{64}$/.test(String(n)));

describe('decks', () => {
  test('an upload is checked, stored, described and downloaded byte for byte', async () => {
    const bytes = deck(3, { note: 'upload' });
    const res = await upload(app, bytes, { filename: 'Quarterly review.pptx' });
    expect(res.statusCode).toBe(201);
    const meta = res.json();
    expect(meta).toMatchObject({
      name: 'Quarterly review.pptx', version: 1, size: bytes.length, slideCount: 3, slideSize: { cx: 9144000, cy: 6858000 },
      encrypted: false, macros: false, contentType: PPTX,
    });
    expect(meta.id).toMatch(/^[0-9a-f]{32}$/);
    expect(res.headers.location).toBe(`/api/decks/${meta.id}`);
    expect(res.headers.etag).toBe(meta.etag);
    expect(meta.versions).toHaveLength(1);

    const got = await app.inject({ url: `/api/decks/${meta.id}` });
    expect(got.statusCode).toBe(200);
    expect(got.json()).toEqual(meta);
    expect(got.headers.etag).toBe(meta.etag);
    expect((await app.inject({ url: `/api/decks/${meta.id}`, headers: { 'if-none-match': meta.etag } })).statusCode).toBe(304);

    const file = await app.inject({ url: `/api/decks/${meta.id}/file` });
    expect(file.statusCode).toBe(200);
    expect(file.headers['content-type']).toBe(PPTX);
    expect(file.headers['content-disposition']).toBe(`attachment; filename="Quarterly review.pptx"; filename*=UTF-8''Quarterly%20review.pptx`);
    expect(Buffer.compare(file.rawPayload, bytes)).toBe(0);
  });

  test('a real python-pptx template from the repository is accepted', async () => {
    const res = await upload(app, readRepoFile('packages/crypto/test/fixtures/plain/default.pptx'), { fields: { name: 'Default' } });
    expect(res.statusCode).toBe(201);
    expect(res.json()).toMatchObject({ name: 'Default', encrypted: false, contentType: PPTX });
  });

  test('names are cleaned: no path, no control or bidi characters; downloads get the right extension', async () => {
    const res = await upload(app, deck(1), { filename: 'C:\\Users\\me\\evil\u202Etxt.pptx' });
    const meta = res.json();
    expect(meta.name).toBe('eviltxt.pptx');
    const macro = deck(1, { mainType: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml', extra: [{ name: 'ppt/vbaProject.bin', data: 'not run' }] });
    const m = (await upload(app, macro, { filename: 'tools.pptx' })).json();
    expect(m).toMatchObject({ macros: true, contentType: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12' });
    const file = await app.inject({ url: `/api/decks/${m.id}/file` });
    expect(file.headers['content-type']).toBe('application/vnd.ms-powerpoint.presentation.macroEnabled.12');
    expect(file.headers['content-disposition']).toContain('filename="tools.pptm"');
  });

  test('a new version needs If-Match: 428 without, 412 when stale, 200 when current', async () => {
    const v1 = deck(1, { note: 'v1' });
    const meta = (await upload(app, v1)).json();
    const v2 = deck(2, { note: 'v2' });

    const missing = await upload(app, v2, { method: 'PUT', url: `/api/decks/${meta.id}` });
    expect(missing.statusCode).toBe(428);
    expect(missing.json().error).toBe('precondition-required');

    const ok = await upload(app, v2, { method: 'PUT', url: `/api/decks/${meta.id}`, headers: { 'if-match': meta.etag } });
    expect(ok.statusCode).toBe(200);
    const next = ok.json();
    expect(next).toMatchObject({ id: meta.id, version: 2, slideCount: 2, size: v2.length, name: meta.name });
    expect(next.etag).not.toBe(meta.etag);
    expect(ok.headers.etag).toBe(next.etag);
    expect(next.versions.map((v: { version: number }) => v.version)).toEqual([1, 2]);

    const stale = await upload(app, deck(1, { note: 'v3' }), { method: 'PUT', url: `/api/decks/${meta.id}`, headers: { 'if-match': meta.etag } });
    expect(stale.statusCode).toBe(412);
    expect(stale.json().error).toBe('precondition-failed');
    expect(stale.headers.etag).toBe(next.etag);
    expect((await upload(app, v2, { method: 'PUT', url: `/api/decks/${meta.id}`, headers: { 'if-match': `W/${next.etag}` } })).statusCode).toBe(412);

    expect(Buffer.compare((await app.inject({ url: `/api/decks/${meta.id}/file?version=1` })).rawPayload, v1)).toBe(0);
    expect(Buffer.compare((await app.inject({ url: `/api/decks/${meta.id}/file` })).rawPayload, v2)).toBe(0);
    expect((await app.inject({ url: `/api/decks/${meta.id}/file?version=3` })).statusCode).toBe(404);
  });

  test('two saves racing on the same ETag: exactly one wins', async () => {
    const meta = (await upload(app, deck(1, { note: 'race' }))).json();
    const put = (note: string) => upload(app, deck(1, { note }), { method: 'PUT', url: `/api/decks/${meta.id}`, headers: { 'if-match': meta.etag } });
    const results = await Promise.all([put('a'), put('b'), put('c')]);
    expect(results.map(r => r.statusCode).sort()).toEqual([200, 412, 412]);
    expect((await app.inject({ url: `/api/decks/${meta.id}` })).json().version).toBe(2);
  });

  test('delete: 412 when stale, 204, then 404; a blob shared with another deck stays until both are gone', async () => {
    const bytes = deck(2, { note: 'shared' });
    const a = (await upload(app, bytes)).json();
    const b = (await upload(app, bytes)).json();
    expect(a.versions[0].sha256).toBe(b.versions[0].sha256);
    const before = blobs().length;

    expect((await app.inject({ method: 'DELETE', url: `/api/decks/${a.id}`, headers: { 'if-match': '"1-0000"' } })).statusCode).toBe(412);
    expect((await app.inject({ method: 'DELETE', url: `/api/decks/${a.id}`, headers: { 'if-match': a.etag } })).statusCode).toBe(204);
    expect((await app.inject({ url: `/api/decks/${a.id}` })).statusCode).toBe(404);
    expect((await app.inject({ method: 'DELETE', url: `/api/decks/${a.id}` })).statusCode).toBe(404);
    expect(Buffer.compare((await app.inject({ url: `/api/decks/${b.id}/file` })).rawPayload, bytes)).toBe(0);
    expect(blobs().length).toBe(before);

    expect((await app.inject({ method: 'DELETE', url: `/api/decks/${b.id}` })).statusCode).toBe(204);
    expect(blobs().length).toBe(before - 1);
  });

  test('ids are matched by pattern and never reach the filesystem', async () => {
    for (const url of ['/api/decks/..%2F..%2Fetc%2Fpasswd', '/api/decks/ABCDEF0123456789abcdef0123456789', '/api/decks/1234/file']) {
      const res = await app.inject({ url });
      expect(res.statusCode, url).toBe(400);
      expect(res.json().error).toBe('invalid-request');
    }
    expect((await app.inject({ url: `/api/decks/${'0'.repeat(32)}` })).statusCode).toBe(404);
    expect((await app.inject({ url: '/api/nothing-here' })).json()).toEqual({ statusCode: 404, error: 'not-found', message: 'No such route.' });
  });
});

describe('refusals', () => {
  test('files that are not PowerPoint Open XML: 415 with a reason', async () => {
    const text = await upload(app, new TextEncoder().encode('just some text'));
    expect(text.statusCode).toBe(415);
    expect(text.json().error).toBe('not-a-package');

    const word = zip([
      {
        name: '[Content_Types].xml',
        data: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          + '</Types>',
      },
      {
        name: '_rels/.rels',
        data: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
          + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
      },
      { name: 'word/document.xml', data: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>' },
    ]);
    const doc = await upload(app, word);
    expect(doc.statusCode).toBe(415);
    expect(doc.json().error).toBe('not-a-presentation');

    const plainZip = await upload(app, zip([{ name: 'a.txt', data: 'hello' }]));
    expect(plainZip.statusCode).toBe(422);
    expect(plainZip.json().error).toBe('invalid-package');

    const json = await app.inject({ method: 'POST', url: '/api/decks', payload: { file: 'x' } });
    expect(json.statusCode).toBe(415);
    const octets = await app.inject({ method: 'POST', url: '/api/decks', payload: Buffer.from('x'), headers: { 'content-type': 'application/octet-stream' } });
    expect(octets.statusCode).toBe(415);
  });

  test('slides that are not well-formed, nest too deep, or carry a DOCTYPE are refused: 422 invalid-part', async () => {
    const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
    const cases = {
      unclosed: `<p:sld ${P}><p:cSld>`,
      deep: `<p:sld ${P}>${'<p:x>'.repeat(10_000)}${'</p:x>'.repeat(10_000)}</p:sld>`,
      doctype: `<?xml version="1.0"?><!DOCTYPE p:sld [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;">]><p:sld ${P}>&lol2;</p:sld>`,
    };
    for (const [name, slideXml] of Object.entries(cases)) {
      const res = await upload(app, deck(2, { slideXml }));
      expect([res.statusCode, res.json().error], name).toEqual([422, 'invalid-part']);
      expect(res.json().message).toContain('/ppt/slides/slide1.xml');
    }
  });

  test('a legacy .ppt is refused with the reason (synthetic compound file)', async () => {
    const res = await upload(app, compoundFile(['Root Entry', 'Current User', 'PowerPoint Document']), { filename: 'old.ppt' });
    expect(res.statusCode).toBe(415);
    expect(res.json()).toMatchObject({ error: 'legacy-ppt' });
    expect(res.json().message).toContain('.ppt');
  });

  const legacy = corpusFile('poi/test-data~slideshow~bug52297.ppt');
  test.skipIf(!legacy)('a legacy .ppt from the corpus (Apache POI bug52297.ppt) is refused as legacy-ppt', async () => {
    const res = await upload(app, legacy!, { filename: 'bug52297.ppt' });
    expect(res.statusCode).toBe(415);
    expect(res.json()).toMatchObject({ error: 'legacy-ppt' });
  });

  const expectedRefusals = Object.entries(JSON.parse(new TextDecoder().decode(readRepoFile('corpus/expected-refusals.json'))) as Record<string, { code: string }>);
  // refused by the server's check, which parses every slide and wants a presentation, though OpcPackage.open alone succeeds
  const serverOnly: Record<string, string> = {
    'libreoffice/sd~qa~unit~data~pptx~fail~ofz35597-1.pptx': 'invalid-part',   // an OSS-Fuzz slide that is not well-formed
    'poi/test-data~slideshow~testPPT.thmx': 'not-a-presentation',             // an Office theme
  };
  test.skipIf(opcFiles.length === 0)('the corpus: every package is accepted but the reviewed refusals, each with its code', () => {
    const { limits, container } = app.services.config;
    const expected = { ...Object.fromEntries(expectedRefusals.map(([f, { code }]) => [f, code])), ...serverOnly };
    const wrong: string[] = [];
    for (const file of opcFiles) {
      const name = corpusName(file);
      let code = 'accepted';
      try {
        if (validateDeck({ bytes: new Uint8Array(readFileSync(file)), limits, container }).encrypted) code = 'encrypted';
      } catch (e) {
        code = describeError(e).code;
      }
      if (code !== (expected[name] ?? 'accepted')) wrong.push(`${name}: ${code}`);
    }
    expect(wrong).toEqual([]);
  });

  test.skipIf(!corpusFile(expectedRefusals[0]![0]))('every file on corpus/expected-refusals.json is refused with its code', async () => {
    for (const [file, { code }] of expectedRefusals) {
      const res = await upload(app, corpusFile(file)!, { filename: 'x.pptx' });
      // a password-protected package is the one exception: the server stores it as uploaded (see the encryption test)
      if (code === 'encrypted') expect(res.statusCode, file).toBe(201);
      else expect(res.json().error, file).toBe(code);
    }
  });

  test('an empty file, a missing file and unexpected fields are refused', async () => {
    expect((await upload(app, new Uint8Array(0))).json().error).toBe('empty-file');
    const noFile = await form({}, { name: 'x' });
    expect((await app.inject({ method: 'POST', url: '/api/decks', ...noFile })).json().error).toBe('missing-file');
    const extra = await form({ file: { data: deck(1) } }, { colour: 'blue' });
    expect((await app.inject({ method: 'POST', url: '/api/decks', ...extra })).statusCode).toBe(400);
    const two = await form({ file: { data: deck(1) }, other: { data: deck(1) } });
    const tooMany = await app.inject({ method: 'POST', url: '/api/decks', ...two });
    expect([tooMany.statusCode, tooMany.json().error]).toEqual([413, 'too-many-files']);
    const wrongName = await form({ upload: { data: deck(1) } });
    expect((await app.inject({ method: 'POST', url: '/api/decks', ...wrongName })).json().error).toBe('unexpected-file');
    const long = await form({ file: { data: deck(1) } }, { name: 'x'.repeat(5000) });
    expect((await app.inject({ method: 'POST', url: '/api/decks', ...long })).statusCode).toBe(413);
  });

  test('an upload over MAX_UPLOAD_MB is refused with 413 and nothing is stored', async () => {
    const dir = tempDir();
    const small = await testServer({ dataDir: dir, maxUploadBytes: 64 * 1024 });
    try {
      const big = deck(1, { extra: [{ name: 'ppt/media/image1.png', data: new Uint8Array(200 * 1024).fill(7) }] });
      const res = await upload(small, big);
      expect(res.statusCode).toBe(413);
      expect(res.json().error).toBe('too-large');
      expect(readdirSync(join(dir, 'decks'))).toEqual([]);
      expect((await upload(small, deck(1))).statusCode).toBe(201);
    } finally {
      await small.close();
    }
  });
});

describe('password-protected decks', () => {
  test('are stored as uploaded with encrypted=true, and the server refuses to draw them (409)', async () => {
    const bytes = readRepoFile('packages/crypto/test/fixtures/msoffcrypto/default-agile.pptx');
    const res = await upload(app, bytes, { filename: 'secret.pptx' });
    expect(res.statusCode).toBe(201);
    const meta = res.json();
    expect(meta).toMatchObject({ encrypted: true, slideCount: null, slideSize: null, size: bytes.length, contentType: PPTX });
    const file = await app.inject({ url: `/api/decks/${meta.id}/file` });
    expect(Buffer.compare(file.rawPayload, bytes)).toBe(0);
    for (const url of [`/api/decks/${meta.id}/slides/1.png`, `/api/decks/${meta.id}/export.pdf`]) {
      const r = await app.inject({ url });
      expect(r.statusCode).toBe(409);
      expect(r.json()).toMatchObject({ error: 'encrypted' });
      expect(r.json().message).toContain('password');
    }
  });

  test('a real PowerPoint-encrypted file is accepted the same way', async () => {
    const res = await upload(app, readRepoFile('packages/crypto/test/fixtures/real/powerpoint-agile.pptx'));
    expect(res.statusCode).toBe(201);
    expect(res.json().encrypted).toBe(true);
  });

  test('a compound file with a looping FAT and an EncryptedPackage entry is refused as malformed', async () => {
    const res = await upload(app, compoundFile(['Root Entry', 'EncryptedPackage', 'EncryptionInfo'], { fatCycle: true }));
    expect(res.statusCode).toBe(422);
    expect(res.json().error).toBe('malformed');
  });
});

describe('security', () => {
  test('security headers on every answer, success or error; no CORS', async () => {
    const meta = (await upload(app, deck(1))).json();
    const answers = [
      await app.inject({ url: '/api/health' }),
      await app.inject({ url: `/api/decks/${meta.id}/file`, headers: { origin: 'https://evil.example' } }),
      await app.inject({ url: '/api/decks/nope' }),
      await app.inject({ url: '/missing' }),
      await app.inject({ method: 'OPTIONS', url: '/api/decks', headers: { origin: 'https://evil.example', 'access-control-request-method': 'POST' } }),
    ];
    for (const r of answers) {
      expect(r.headers['content-security-policy']).toBe("default-src 'none';base-uri 'none';form-action 'none';frame-ancestors 'none';sandbox");
      expect(r.headers['x-content-type-options']).toBe('nosniff');
      expect(r.headers['x-frame-options']).toBe('DENY');
      expect(r.headers['referrer-policy']).toBe('no-referrer');
      expect(r.headers['cross-origin-resource-policy']).toBe('same-origin');
      expect(r.headers['access-control-allow-origin']).toBeUndefined();
    }
  });

  test('errors are JSON without stack traces', async () => {
    const res = await app.inject({ url: '/api/decks/zz' });
    expect(res.headers['content-type']).toContain('application/json');
    const body = res.json();
    expect(Object.keys(body).sort()).toEqual(['error', 'message', 'statusCode']);
    expect(JSON.stringify(body)).not.toMatch(/at .*\.ts:\d+/);
  });

  test('under pressure (here: any heap use at all) every request gets 503 with Retry-After', async () => {
    const pressed = await testServer({ pressure: { maxHeapUsedBytes: 1, sampleIntervalMs: 10 } });
    try {
      await new Promise(r => setTimeout(r, 100));
      const r = await pressed.inject({ url: '/api/health' });
      expect(r.statusCode).toBe(503);
      expect(r.json()).toMatchObject({ statusCode: 503, error: 'unavailable' });
      expect(r.headers['retry-after']).toBe('10');
      expect(r.headers['x-content-type-options']).toBe('nosniff');
    } finally {
      await pressed.close();
    }
  });

  test('requests past the rate limit get 429', async () => {
    const limited = await testServer({ rateLimit: { max: 3, upload: 1, windowMs: 60_000 } });
    try {
      const codes = [];
      for (let i = 0; i < 4; i++) codes.push((await limited.inject({ url: '/api/health' })).statusCode);
      expect(codes).toEqual([200, 200, 200, 429]);
      const r = await limited.inject({ url: '/api/health' });
      expect(r.json()).toMatchObject({ statusCode: 429, error: 'rate-limited' });
      expect(r.headers['retry-after']).toBeDefined();
    } finally {
      await limited.close();
    }
  });
});

describe('configuration from the environment', () => {
  test('defaults, values, and refusal of values that do not parse', () => {
    const c = configFromEnv({ DATA_DIR: '/tmp/x', PORT: '8080', MAX_UPLOAD_MB: '20', RENDER_PAGES: '0', ENABLE_ORACLES: '1', PPTX_VALIDATOR: '["dotnet","v.dll"]' });
    expect(c).toMatchObject({ host: '127.0.0.1', port: 8080, dataDir: '/tmp/x', maxUploadBytes: 20 << 20, render: { pages: 0 }, oracles: { enabled: true, validator: ['dotnet', 'v.dll'] } });
    expect(c.container.maxBytes).toBe(20 << 20);
    expect(configFromEnv({}).oracles.enabled).toBe(false);
    expect(() => configFromEnv({ PORT: 'eighty' })).toThrow(/PORT/);
    expect(() => configFromEnv({ MAX_UPLOAD_MB: '0' })).toThrow(/MAX_UPLOAD_MB/);
    expect(() => configFromEnv({ POWERPOINT_BRIDGE_URL: 'file:///etc/passwd' })).toThrow(/http/);
    expect(configFromEnv({ RATE_LIMIT: '0' }).rateLimit).toBe(false);
  });
});
