// apps/server/src/routes.ts  -- the deck API: upload, metadata, download, new versions, delete, slide PNGs, PDF, health
import type { FastifyContextConfig, FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import type {} from '@fastify/rate-limit';
import type { Static } from 'typebox';
import type { ServerConfig } from './config.ts';
import { busy, HttpError, notFound, refusal } from './errors.ts';
import { cleanName, contentDisposition, downloadName, ifMatchSatisfied, ifNoneMatchHit, readUpload } from './http.ts';
import { QueueFullError, type JobQueue } from './jobs.ts';
import { RenderError, type Renderer } from './render.ts';
import {
  DeckParams, DeckView, ERRORS, HealthView, IfMatchHeaders, IfNoneMatchHeaders, NoQuery, SlideParams, SlideQuery, UploadFields,
  VersionQuery, type DeckViewT,
} from './schemas.ts';
import { currentVersion, type DeckMeta, type DeckStore, type DeckVersion } from './storage.ts';
import { PACKAGE_TYPES, type DeckCheck } from './tasks.ts';
import { TaskFailure, WorkerError, type WorkerPool } from './workers.ts';

export interface Services {
  config: ServerConfig;
  store: DeckStore;
  pool: WorkerPool;
  renderer: Renderer | undefined;   // undefined when RENDER_PAGES is 0
  renders: JobQueue;
}

export const etagOf = (v: DeckVersion) => `"${v.version}-${v.sha256}"`;

/** Route config for routes that take a file: the tighter per-IP budget. */
export function uploadRate(config: ServerConfig): FastifyContextConfig {
  return config.rateLimit ? { rateLimit: { max: config.rateLimit.upload, timeWindow: config.rateLimit.windowMs } } : {};
}

const EXTENSIONS = new Map(Object.values(PACKAGE_TYPES).map(t => [t.type, t.ext]));
const started = Date.now();

function view(meta: DeckMeta): DeckViewT {
  const v = currentVersion(meta);
  return {
    id: meta.id, name: meta.name, etag: etagOf(v), version: v.version, size: v.size, slideCount: v.slideCount,
    slideSize: v.slideSize, encrypted: v.encrypted, macros: v.macros, contentType: v.contentType,
    createdAt: meta.createdAt, updatedAt: meta.updatedAt,
    versions: meta.versions.map(x => ({
      version: x.version, sha256: x.sha256, size: x.size, createdAt: x.createdAt, encrypted: x.encrypted,
      slideCount: x.slideCount, contentType: x.contentType, macros: x.macros,
    })),
  };
}

/** The HTTP answer for a worker job that failed: the task's own refusal, or what happened to the thread. */
export function workerHttpError(e: unknown, request: FastifyRequest): unknown {
  if (e instanceof TaskFailure) {
    if (e.code !== 'unexpected') return refusal(e.code, e.message);
    request.log.warn({ reason: e.message }, 'parser failed on an upload');
    return new HttpError(422, 'unreadable', 'The file could not be read as a presentation.');
  }
  if (!(e instanceof WorkerError)) return e;
  switch (e.code) {
    case 'busy': return busy('checking other files');
    case 'timeout': case 'out-of-memory':
      return new HttpError(422, 'too-complex', `The file could not be checked within the server's limits: ${e.message}`);
    case 'closed': return new HttpError(503, 'unavailable', e.message);
    default:
      request.log.error({ reason: e.message }, 'parser worker crashed');
      return new HttpError(500, 'worker-crashed', 'The file checker failed unexpectedly.');
  }
}

function renderHttpError(e: unknown): unknown {
  if (e instanceof QueueFullError) return busy('drawing other slides');
  if (!(e instanceof RenderError)) return e;
  switch (e.code) {
    case 'timeout': return new HttpError(504, 'render-timeout', e.message);
    case 'unavailable': return new HttpError(503, 'render-unavailable', e.message);
    case 'no-slides': return new HttpError(422, 'no-slides', e.message);
    default: return new HttpError(422, 'render-failed', `The slide could not be drawn: ${e.message}`);
  }
}

interface ReceivedDeck { bytes: Buffer; name: string | undefined; fileName: string | undefined; check: DeckCheck; }

/** Read one deck upload (`file`, and optionally a `name` field) and have a worker check it. */
async function receiveDeck(s: Services, request: FastifyRequest): Promise<ReceivedDeck> {
  const upload = await readUpload(request, ['file'], UploadFields);
  const file = upload.files.get('file')!;
  let check: DeckCheck;
  try {
    check = await s.pool.run('validate', { bytes: file.bytes, limits: s.config.limits, container: s.config.container });
  } catch (e) {
    throw workerHttpError(e, request);
  }
  return { bytes: file.bytes, name: cleanName(upload.fields.name), fileName: cleanName(file.filename), check };
}

const versionInfo = (c: DeckCheck) => ({
  encrypted: c.encrypted, slideCount: c.slideCount, slideSize: c.slideSize, contentType: c.contentType, macros: c.macros,
});

async function deck(s: Services, id: string): Promise<DeckMeta> {
  const meta = await s.store.get(id);
  if (!meta) throw notFound();
  return meta;
}

function version(meta: DeckMeta, n: number | undefined): DeckVersion {
  const v = n === undefined ? currentVersion(meta) : meta.versions[n - 1];
  if (!v) throw notFound('version');
  return v;
}

/** Encrypted decks are stored as uploaded and only a client holding the password can open them. */
function drawable(v: DeckVersion): void {
  if (v.encrypted) {
    throw new HttpError(409, 'encrypted', 'This presentation is password-protected. The server never receives passwords, '
      + 'so it cannot draw or convert it; open it in the editor instead.');
  }
}

/** A precondition for changing a deck: it exists and, when `required` or given, If-Match names its current version. */
function ifMatch(header: string | undefined, required: boolean) {
  return (meta: DeckMeta | undefined): DeckMeta => {
    if (!meta) throw notFound();
    const etag = etagOf(currentVersion(meta));
    if (header === undefined) {
      if (required) throw new HttpError(428, 'precondition-required', 'Send If-Match with the ETag of the version you changed.');
      return meta;
    }
    if (!ifMatchSatisfied(header, etag)) {
      throw new HttpError(412, 'precondition-failed', 'The deck has a newer version than the one you changed: load it again first.', { etag });
    }
    return meta;
  };
}

/** Send a cached or freshly drawn image or PDF, honouring If-None-Match. */
function sendDrawing(request: FastifyRequest, reply: FastifyReply, bytes: Uint8Array, type: string, disposition: string, etag: string) {
  reply.header('etag', etag).header('cache-control', 'private, no-cache');
  if (ifNoneMatchHit(request.headers['if-none-match'], etag)) return reply.code(304).send();
  return reply.type(type).header('content-disposition', disposition).send(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength));
}

export function deckRoutes(app: FastifyInstance, s: Services): void {
  const rate = uploadRate(s.config);
  const blob = (v: DeckVersion) => async () => {
    const bytes = await s.store.blobs.get(v.sha256);
    if (!bytes) throw new RenderError('render-failed', 'the deck\'s file is missing');
    return bytes;
  };
  const renderer = (): Renderer => {
    if (!s.renderer) throw new HttpError(503, 'render-unavailable', 'Rendering is turned off on this server (RENDER_PAGES=0).');
    if (!s.renderer.built) throw new HttpError(503, 'render-unavailable', 'The render page is not built on this server.');
    return s.renderer;
  };

  app.get('/api/health', { schema: { querystring: NoQuery, response: { 200: HealthView, ...ERRORS } } }, async () => ({
    status: 'ok',
    uptimeSeconds: Math.round((Date.now() - started) / 1000),
    workers: { ...s.pool.stats() },
    render: { enabled: s.renderer !== undefined, built: s.renderer?.built ?? false, ...s.renderer?.stats(), ...s.renders.stats() },
  }));

  app.post('/api/decks', { config: rate, schema: { querystring: NoQuery, response: { 201: DeckView, ...ERRORS } } }, async (request, reply) => {
    const up = await receiveDeck(s, request);
    const meta = await s.store.create(up.name ?? up.fileName ?? 'Untitled.pptx', up.bytes, versionInfo(up.check));
    const v = currentVersion(meta);
    request.log.info({ deck: meta.id, size: v.size, slides: v.slideCount, encrypted: v.encrypted, warnings: up.check.diagnostics.length }, 'deck created');
    return reply.code(201).header('location', `/api/decks/${meta.id}`).header('etag', etagOf(v)).send(view(meta));
  });

  app.get<{ Params: Static<typeof DeckParams> }>('/api/decks/:id', {
    schema: { params: DeckParams, querystring: NoQuery, headers: IfNoneMatchHeaders, response: { 200: DeckView, ...ERRORS } },
  }, async (request, reply) => {
    const meta = await deck(s, request.params.id);
    const etag = etagOf(currentVersion(meta));
    reply.header('etag', etag).header('cache-control', 'private, no-cache');
    if (ifNoneMatchHit(request.headers['if-none-match'], etag)) return reply.code(304).send();
    return view(meta);
  });

  app.get<{ Params: Static<typeof DeckParams>; Querystring: Static<typeof VersionQuery> }>('/api/decks/:id/file', {
    schema: { params: DeckParams, querystring: VersionQuery, response: ERRORS },
  }, async (request, reply) => {
    const meta = await deck(s, request.params.id);
    const v = version(meta, request.query.version);
    const bytes = await s.store.blobs.get(v.sha256);
    if (!bytes) throw notFound('file');
    return reply
      .type(v.contentType)
      .header('content-disposition', contentDisposition('attachment', downloadName(meta.name, EXTENSIONS.get(v.contentType) ?? 'pptx')))
      .header('etag', etagOf(v))
      .header('cache-control', 'private, no-cache')
      .send(bytes);
  });

  app.put<{ Params: Static<typeof DeckParams>; Headers: Static<typeof IfMatchHeaders> }>('/api/decks/:id', {
    config: rate, schema: { params: DeckParams, querystring: NoQuery, headers: IfMatchHeaders, response: { 200: DeckView, ...ERRORS } },
  }, async (request, reply) => {
    const precondition = ifMatch(request.headers['if-match'], true);
    precondition(await s.store.get(request.params.id));       // refuse before reading the upload; checked again under the lock
    const up = await receiveDeck(s, request);
    const meta = await s.store.addVersion(request.params.id, up.bytes, versionInfo(up.check), precondition, up.name);
    const v = currentVersion(meta);
    request.log.info({ deck: meta.id, version: v.version, size: v.size, warnings: up.check.diagnostics.length }, 'deck version saved');
    return reply.header('etag', etagOf(v)).send(view(meta));
  });

  app.delete<{ Params: Static<typeof DeckParams>; Headers: Static<typeof IfMatchHeaders> }>('/api/decks/:id', {
    schema: { params: DeckParams, querystring: NoQuery, headers: IfMatchHeaders, response: ERRORS },
  }, async (request, reply) => {
    await s.store.delete(request.params.id, ifMatch(request.headers['if-match'], false));
    request.log.info({ deck: request.params.id }, 'deck deleted');
    return reply.code(204).send();
  });

  app.get<{ Params: Static<typeof SlideParams>; Querystring: Static<typeof SlideQuery> }>('/api/decks/:id/slides/:n.png', {
    schema: { params: SlideParams, querystring: SlideQuery, headers: IfNoneMatchHeaders, response: ERRORS },
  }, async (request, reply) => {
    const { n } = request.params;
    const meta = await deck(s, request.params.id);
    const v = version(meta, request.query.version);
    drawable(v);
    if (v.slideCount !== null && n > v.slideCount) throw notFound('slide');
    const width = request.query.width ?? s.config.render.defaultWidth;
    if (width > s.config.render.maxWidth) throw new HttpError(400, 'invalid-request', `width must be at most ${s.config.render.maxWidth}`);
    const name = `slide-${n}-w${width}.png`;
    let png = await s.store.cache.get(v.sha256, name);
    if (!png) {
      const r = renderer();
      try {
        png = await s.renders.run(`png:${v.sha256}:${n}:${width}`, async () => {
          const out = await r.renderSlide({ key: v.sha256, load: blob(v) }, n - 1, width);
          await s.store.cache.put(v.sha256, name, out);
          return out;
        });
      } catch (e) {
        throw renderHttpError(e);
      }
    }
    return sendDrawing(request, reply, png, 'image/png', contentDisposition('inline', `slide${n}.png`), `"${v.sha256}-${n}-${width}"`);
  });

  app.get<{ Params: Static<typeof DeckParams>; Querystring: Static<typeof VersionQuery> }>('/api/decks/:id/export.pdf', {
    schema: { params: DeckParams, querystring: VersionQuery, headers: IfNoneMatchHeaders, response: ERRORS },
  }, async (request, reply) => {
    const meta = await deck(s, request.params.id);
    const v = version(meta, request.query.version);
    drawable(v);
    if (v.slideCount === 0) throw new HttpError(422, 'no-slides', 'The presentation has no slides.');
    const width = s.config.render.pdfWidth;
    const name = `deck-w${width}.pdf`;
    let pdf = await s.store.cache.get(v.sha256, name);
    if (!pdf) {
      const r = renderer();
      try {
        pdf = await s.renders.run(`pdf:${v.sha256}:${width}`, async () => {
          const out = await r.exportPdf({ key: v.sha256, load: blob(v) }, width);
          if (out.errors.length) {
            request.log.warn({ deck: meta.id, failed: out.errors.length, first: out.errors[0] }, 'slides left blank in a PDF');
          }
          await s.store.cache.put(v.sha256, name, out.pdf);
          return out.pdf;
        });
      } catch (e) {
        throw renderHttpError(e);
      }
    }
    const disposition = contentDisposition('attachment', downloadName(meta.name, 'pdf'));
    return sendDrawing(request, reply, pdf, 'application/pdf', disposition, `"${v.sha256}-pdf-${width}"`);
  });
}
