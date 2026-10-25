// apps/server/src/app.ts  -- buildServer(config): the Fastify app with its security defaults, storage, workers and renderer
//
// Security defaults, outermost first: helmet (a CSP that allows nothing, since the API serves JSON and files and
// never HTML; no framing, no sniffing, no referrer), under-pressure (503 when the event loop lags or the heap
// fills), a per-IP rate limit with a tighter budget for uploads, and multipart limits (one file of at most
// MAX_UPLOAD_MB, a few short fields). With API_TOKENS set, every /api route but the health check wants one of
// them as a bearer token, checked before an upload's body is read. Every route has TypeBox schemas; ids are matched
// by pattern and only ever name records, never paths. Errors are JSON without stack traces. There is no CORS:
// browsers on other origins get nothing. Logs carry methods, URLs, ids and sizes, never file bytes, tokens or
// passwords (which never reach the server). The built editor, when there is one, is served at / (src/web.ts).
import { timingSafeEqual } from 'node:crypto';
import Fastify, { type FastifyInstance } from 'fastify';
import helmet from '@fastify/helmet';
import multipart from '@fastify/multipart';
import rateLimit from '@fastify/rate-limit';
import underPressure from '@fastify/under-pressure';
import { resolveConfig, type ConfigInput } from './config.ts';
import { HttpError, sendError } from './errors.ts';
import { LocalJobQueue } from './jobs.ts';
import { oracleRoutes } from './oracle.ts';
import { Renderer } from './render.ts';
import { deckRoutes, type Services } from './routes.ts';
import { DeckStore } from './storage.ts';
import { loadWeb, webRoutes } from './web.ts';
import { WorkerPool } from './workers.ts';

declare module 'fastify' {
  interface FastifyInstance {
    services: Services;
  }
}

export async function buildServer(input: ConfigInput = {}): Promise<FastifyInstance> {
  const config = resolveConfig(input);
  const store = await DeckStore.openDirectory(config.dataDir);
  const web = await loadWeb(config.webDir);
  const pool = new WorkerPool(config.workers);
  const renderer = config.render.pages > 0 ? new Renderer(config.render, config.limits) : undefined;
  const renders = new LocalJobQueue(config.render.pages, config.render.queue);
  const services: Services = { config, store, pool, renderer, renders };

  const app = Fastify({
    logger: config.logLevel === 'silent' ? false : { level: config.logLevel },
    bodyLimit: 64 * 1024,                 // no route takes a JSON body; files come as multipart, capped below
    requestTimeout: config.requestTimeoutMs,
    requestIdHeader: false,               // request ids are the server's own, never a client's
    // a hop count trusts that many proxies nearest the server
    trustProxy: typeof config.trustProxy === 'number' ? ((_address: string, hop: number) => hop < (config.trustProxy as number)) : config.trustProxy,
    ajv: { customOptions: { removeAdditional: true, coerceTypes: 'array', useDefaults: true } },
  });
  app.decorate('services', services);
  app.setErrorHandler(sendError);
  app.setNotFoundHandler((_request, reply) => reply.code(404).send({ statusCode: 404, error: 'not-found', message: 'No such route.' }));
  app.addHook('onClose', async () => {
    await Promise.allSettled([pool.close(), renderer?.close()]);
  });

  await app.register(helmet, {
    contentSecurityPolicy: {
      useDefaults: false,
      directives: { defaultSrc: ["'none'"], baseUri: ["'none'"], formAction: ["'none'"], frameAncestors: ["'none'"], sandbox: [] },
    },
    crossOriginEmbedderPolicy: { policy: 'require-corp' },
    crossOriginOpenerPolicy: { policy: 'same-origin' },
    crossOriginResourcePolicy: { policy: 'same-origin' },
    referrerPolicy: { policy: 'no-referrer' },
    strictTransportSecurity: { maxAge: 31_536_000, includeSubDomains: true },
    xFrameOptions: { action: 'deny' },
  });
  await app.register(underPressure, {
    maxEventLoopDelay: config.pressure.maxEventLoopDelayMs,
    maxHeapUsedBytes: config.pressure.maxHeapUsedBytes,
    maxRssBytes: config.pressure.maxRssBytes,
    sampleInterval: config.pressure.sampleIntervalMs,
    retryAfter: 10,
  });
  if (config.rateLimit) await app.register(rateLimit, { global: true, max: config.rateLimit.max, timeWindow: config.rateLimit.windowMs });
  if (config.apiTokens.length) {
    // compared in constant time; only a token's length (32 characters or more, and not secret) can show
    const allowed = config.apiTokens.map(token => Buffer.from(token));
    app.addHook('onRequest', async request => {
      const route = request.routeOptions.url;
      if (!route?.startsWith('/api/') || route === '/api/health') return;
      const given = /^Bearer +(\S+)$/i.exec(request.headers.authorization ?? '')?.[1];
      const bytes = given === undefined ? undefined : Buffer.from(given);
      if (!bytes || !allowed.some(a => a.length === bytes.length && timingSafeEqual(a, bytes))) {
        throw new HttpError(401, 'unauthorized', 'This API needs a token: Authorization: Bearer <token>.', { 'www-authenticate': 'Bearer' });
      }
    });
  }
  await app.register(multipart, {
    throwFileSizeLimit: true,
    limits: { fileSize: config.maxUploadBytes, files: 1, fields: 2, fieldSize: 1024, fieldNameSize: 64, parts: 3, headerPairs: 32 },
  });

  deckRoutes(app, services);
  if (config.oracles.enabled) oracleRoutes(app, services);
  if (web) webRoutes(app, web);
  return app;
}
