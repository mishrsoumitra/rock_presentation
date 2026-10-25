// apps/server/src/oracle.ts  -- the CI oracle routes; they exist only with ENABLE_ORACLES=1 (404 otherwise)
//
//   POST /api/oracle/validate        O1: the Open XML SDK validator (tools/validator, .NET) on a deck; PPTX_VALIDATOR
//   POST /api/oracle/compare         O3: pixelmatch of two PNGs: mismatched pixels, their ratio, and a diff image
//   POST /api/oracle/roundtrip       open and re-save with OpcPackage: did every part's canonical XML survive
//   POST /api/oracle/open            PowerPoint itself, on the Windows runner (did it open the file), and its slide
//   POST /api/oracle/export          PNGs, every text box's lines and bounds, and its own re-save: forwarded to
//   POST /api/oracle/text-metrics    POWERPOINT_BRIDGE_URL as they came, or 501 without one
//   POST /api/oracle/resave
import { execFile } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { FastifyInstance, FastifyRequest } from 'fastify';
import type { Static } from 'typebox';
import { HttpError, refusal } from './errors.ts';
import { cleanName, contentDisposition, readUpload } from './http.ts';
import { uploadRate, workerHttpError, type Services } from './routes.ts';
import { CompareQuery, CompareView, ERRORS, NoFields, NoQuery, RoundtripView, UploadFields, ValidatorQuery, ValidatorView } from './schemas.ts';
import { PACKAGE_TYPES, PPTX_TYPE, type DeckCheck } from './tasks.ts';

const EXTENSIONS = new Map(Object.values(PACKAGE_TYPES).map(t => [t.type, t.ext]));
const MAX_BRIDGE_RESPONSE = 512 * (1 << 20);
const RELAYED_TYPES = ['application/json', 'image/png', 'application/zip', 'application/pdf', 'text/plain', ...EXTENSIONS.keys()];

interface Run { code: number; stdout: string; stderr: string; }

/** Run a program without a shell, under a deadline, with its output capped. Non-zero exits are results, not errors. */
function run(file: string, args: string[], timeoutMs: number): Promise<Run> {
  return new Promise((resolve, reject) => {
    execFile(file, args, { timeout: timeoutMs, killSignal: 'SIGKILL', maxBuffer: 64 * (1 << 20), windowsHide: true }, (error, stdout, stderr) => {
      const exit = (error as { code?: unknown } | null)?.code;
      if (error && typeof exit !== 'number') reject(error);
      else resolve({ code: typeof exit === 'number' ? exit : 0, stdout: String(stdout), stderr: String(stderr) });
    });
  });
}

/** A response body, refused once it passes `max` bytes. */
async function readCapped(res: Response, max: number): Promise<Buffer> {
  if (!res.body) return Buffer.alloc(0);
  const reader = res.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) return Buffer.concat(chunks);
    total += value.length;
    if (total > max) {
      await reader.cancel();
      throw new HttpError(502, 'bridge-failed', `The PowerPoint bridge answered with more than ${max} bytes.`);
    }
    chunks.push(value);
  }
}

export function oracleRoutes(app: FastifyInstance, s: Services): void {
  const { oracles } = s.config;
  const rate = uploadRate(s.config);

  /** A deck for the validator, checked by our own parser first: hostile input never reaches the .NET one. */
  const checkedDeck = async (request: FastifyRequest) => {
    const file = (await readUpload(request, ['file'], NoFields)).files.get('file')!;
    let check: DeckCheck;
    try {
      check = await s.pool.run('validate', { bytes: file.bytes, limits: s.config.limits, container: s.config.container });
    } catch (e) {
      throw workerHttpError(e, request);
    }
    if (check.encrypted) throw refusal('encrypted', 'This presentation is password-protected; the server cannot open it.');
    return { bytes: file.bytes, ext: EXTENSIONS.get(check.contentType) ?? 'pptx' };
  };

  app.post<{ Querystring: Static<typeof ValidatorQuery> }>('/api/oracle/validate', {
    config: rate, schema: { querystring: ValidatorQuery, response: { 200: ValidatorView, ...ERRORS } },
  }, async request => {
    const argv = oracles.validator;
    if (!argv) {
      throw new HttpError(501, 'not-configured', 'The Open XML SDK validator is not set up on this server: '
        + 'build tools/validator (it needs the .NET SDK) and set PPTX_VALIDATOR to the program.');
    }
    const deck = await checkedDeck(request);
    const dir = await mkdtemp(join(s.config.dataDir, 'tmp', 'validate-'));
    try {
      const file = join(dir, `deck.${deck.ext}`);
      await writeFile(file, deck.bytes, { mode: 0o600 });
      let r: Run;
      try {
        r = await run(argv[0]!, [...argv.slice(1), file, ...(request.query.format ? [request.query.format] : [])], oracles.validatorTimeoutMs);
      } catch (e) {
        request.log.error({ err: e }, 'validator did not run');
        throw new HttpError(502, 'validator-failed', 'The validator could not be run or did not finish in time.');
      }
      // tools/validator/Program.cs: exit 0 valid, 1 validation errors, 2 the package could not be opened; JSON on stdout
      let out: unknown;
      try {
        out = JSON.parse(r.stdout);
      } catch {
        throw new HttpError(502, 'validator-failed', `The validator's output is not JSON (exit code ${r.code}).`);
      }
      if (!Array.isArray(out)) {
        const error = (out as { error?: unknown } | null)?.error;
        const why = typeof error === 'string' ? error : 'the package could not be opened';
        return { valid: false, exitCode: r.code, errorCount: null, errors: [], error: why };
      }
      return { valid: r.code === 0 && out.length === 0, exitCode: r.code, errorCount: out.length, errors: out as Record<string, unknown>[] };
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  app.post<{ Querystring: Static<typeof CompareQuery> }>('/api/oracle/compare', {
    config: rate, schema: { querystring: CompareQuery, response: { 200: CompareView, ...ERRORS } },
  }, async request => {
    const upload = await readUpload(request, ['expected', 'actual'], NoFields, { files: 2, parts: 2, fields: 0, fileSize: oracles.maxPngBytes });
    const threshold = request.query.threshold ?? 0.1;
    try {
      const r = await s.pool.run('compare', {
        expected: upload.files.get('expected')!.bytes, actual: upload.files.get('actual')!.bytes,
        threshold, includeAA: request.query.includeAA ?? false, diff: request.query.diff ?? true, maxPixels: s.config.render.maxPixels,
      });
      return {
        width: r.width, height: r.height, mismatched: r.mismatched, ratio: r.ratio, threshold,
        ...(r.diff ? { diffPng: Buffer.from(r.diff).toString('base64') } : {}),
      };
    } catch (e) {
      throw workerHttpError(e, request);
    }
  });

  app.post('/api/oracle/roundtrip', {
    config: rate, schema: { querystring: NoQuery, response: { 200: RoundtripView, ...ERRORS } },
  }, async request => {
    const file = (await readUpload(request, ['file'], NoFields)).files.get('file')!;
    try {
      return await s.pool.run('roundtrip', { bytes: file.bytes, limits: s.config.limits });
    } catch (e) {
      throw workerHttpError(e, request);
    }
  });

  // The deck goes to PowerPoint unparsed: what PowerPoint makes of a file we would refuse is part of the question.
  for (const name of ['open', 'export', 'text-metrics', 'resave']) {
    app.post(`/api/oracle/${name}`, { config: rate, schema: { querystring: NoQuery, response: ERRORS } }, async (request, reply) => {
      if (!oracles.bridgeUrl) {
        throw new HttpError(501, 'not-configured', `POST /oracle/${name} asks PowerPoint itself, which runs only on the `
          + 'Windows CI runner: set POWERPOINT_BRIDGE_URL to its bridge.');
      }
      const upload = await readUpload(request, ['file'], UploadFields);
      const file = upload.files.get('file')!;
      const form = new FormData();
      const filename = cleanName(upload.fields.name ?? file.filename) ?? 'deck.pptx';
      form.set('file', new Blob([file.bytes as BlobPart], { type: file.mimetype || PPTX_TYPE }), filename);
      let res: Response;
      try {
        res = await fetch(new URL(`oracle/${name}`, oracles.bridgeUrl), {
          method: 'POST', body: form, redirect: 'error', signal: AbortSignal.timeout(oracles.bridgeTimeoutMs),
        });
      } catch (e) {
        request.log.error({ err: e }, 'PowerPoint bridge unreachable');
        throw new HttpError(502, 'bridge-failed', 'The PowerPoint bridge did not answer.');
      }
      const body = await readCapped(res, MAX_BRIDGE_RESPONSE);
      const type = (res.headers.get('content-type') ?? '').split(';')[0]!.trim().toLowerCase();
      reply.code(res.status >= 200 && res.status <= 599 ? res.status : 502).type(RELAYED_TYPES.includes(type) ? type : 'application/octet-stream');
      if (type !== 'application/json' && type !== 'text/plain') {
        reply.header('content-disposition', contentDisposition('attachment', `${name}-result`));
      }
      return reply.send(body);
    });
  }
}
