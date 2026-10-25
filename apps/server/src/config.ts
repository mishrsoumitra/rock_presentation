// apps/server/src/config.ts  -- server settings: the defaults, and reading them from environment variables
import { availableParallelism } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { getHeapStatistics } from 'node:v8';
import { DEFAULT_PACKAGE_LIMITS, type PackageLimits } from '@pptx/opc';
import { DEFAULT_CONTAINER_LIMITS, type ContainerLimits } from '@pptx/crypto';

const HERE = dirname(fileURLToPath(import.meta.url));
const MB = 1 << 20;

/** Worker threads that open untrusted decks (src/workers.ts). */
export interface WorkerSettings {
  size: number;           // threads
  queue: number;          // jobs that may wait for a thread; beyond that requests get 503
  timeoutMs: number;      // wall-clock deadline per job; the thread is terminated on overrun
  heapMb: number;         // resourceLimits.maxOldGenerationSizeMb
  youngMb: number;        // resourceLimits.maxYoungGenerationSizeMb
  stackMb: number;        // resourceLimits.stackSizeMb
  testTasks: boolean;     // accept the test-only tasks (spin, alloc, crash); never read from the environment
}

/** Slide images and PDFs, drawn by the product's renderer in headless Chromium (src/render.ts). */
export interface RenderSettings {
  pages: number;          // Chromium pages drawing at once; 0 turns rendering off
  queue: number;          // render jobs that may wait for a page
  timeoutMs: number;      // per slide image
  pdfTimeoutMs: number;   // per PDF
  defaultWidth: number;   // PNG width when ?width= is not given
  maxWidth: number;
  pdfWidth: number;       // pixel width of each slide image in a PDF
  maxPixels: number;      // the largest image drawn, whatever the slide's aspect ratio
  heapMb: number;         // V8 heap limit of Chromium's renderer processes
  sandbox: boolean;       // Chromium's own process sandbox (Playwright turns it off by default); RENDER_SANDBOX=0 where it cannot start
  pageDir: string;        // the built render page (build:render-page)
  fontsDir: string;       // assets/fonts
}

export interface OracleSettings {
  enabled: boolean;               // ENABLE_ORACLES=1; the /api/oracle routes do not exist otherwise
  validator?: string[];           // PPTX_VALIDATOR: argv of the Open XML SDK validator (tools/validator)
  validatorTimeoutMs: number;
  bridgeUrl?: string;             // POWERPOINT_BRIDGE_URL: the Windows runner's PowerPoint bridge
  bridgeTimeoutMs: number;
  maxPngBytes: number;            // each image given to /api/oracle/compare
}

export interface ServerConfig {
  host: string;
  port: number;
  dataDir: string;
  maxUploadBytes: number;
  /** ZIP and XML limits for opening uploads, checked before anything is inflated. Tighter than the library's. */
  limits: PackageLimits;
  /** Compound-file limits for password-protected packages, which are only checked, never decrypted. */
  container: ContainerLimits;
  workers: WorkerSettings;
  render: RenderSettings;
  oracles: OracleSettings;
  /** Requests per IP per window; `upload` is the tighter budget of routes that take a file. False: off. */
  rateLimit: { max: number; upload: number; windowMs: number } | false;
  /** Above any of these the server answers 503 (@fastify/under-pressure); 0 turns a check off. Sampled every sampleIntervalMs. */
  pressure: { maxEventLoopDelayMs: number; maxHeapUsedBytes: number; maxRssBytes: number; sampleIntervalMs: number };
  requestTimeoutMs: number;
  logLevel: string;
  /** Bearer tokens that may use the API (GET /api/health needs none). Empty: the API is open. */
  apiTokens: string[];
  /** Fastify's trustProxy: the proxies whose X-Forwarded-For names the client, for rate limits and logs. */
  trustProxy: boolean | number | string[];
  /** The built editor (apps/web/dist), served at /; when it has no index.html, only the API is served. */
  webDir: string;
}

export type ConfigInput = Partial<Omit<ServerConfig, 'workers' | 'render' | 'oracles' | 'pressure' | 'limits' | 'container'>> & {
  workers?: Partial<WorkerSettings>;
  render?: Partial<RenderSettings>;
  oracles?: Partial<OracleSettings>;
  pressure?: Partial<ServerConfig['pressure']>;
  limits?: Partial<PackageLimits>;
  container?: Partial<ContainerLimits>;
};

/** The object without its undefined values, so that spreading it never replaces a default with undefined. */
function defined<T extends object>(o: T | undefined): Partial<T> {
  return Object.fromEntries(Object.entries(o ?? {}).filter(([, v]) => v !== undefined)) as Partial<T>;
}

/** The settings with every gap filled from the defaults. */
export function resolveConfig(input: ConfigInput = {}): ServerConfig {
  const maxUploadBytes = input.maxUploadBytes ?? 100 * MB;
  const inflated = 512 * MB;
  return {
    host: input.host ?? '127.0.0.1',
    port: input.port ?? 3000,
    dataDir: resolve(input.dataDir ?? 'data'),
    maxUploadBytes,
    limits: { ...DEFAULT_PACKAGE_LIMITS, maxEntryBytes: inflated, maxTotalBytes: inflated, ...defined(input.limits) },
    container: { ...DEFAULT_CONTAINER_LIMITS, maxBytes: maxUploadBytes, ...defined(input.container) },
    // each thread may hold its heap plus the inflated entries (MAX_INFLATED_MB) of the deck it checks
    workers: {
      size: Math.max(1, Math.min(2, availableParallelism() - 1)),
      queue: 32, timeoutMs: 20_000, heapMb: 512, youngMb: 32, stackMb: 4, testTasks: false,
      ...defined(input.workers),
    },
    render: {
      pages: 2, queue: 64, timeoutMs: 30_000, pdfTimeoutMs: 180_000,
      defaultWidth: 1280, maxWidth: 3840, pdfWidth: 1920, maxPixels: 3840 * 2160 * 2, heapMb: 1024, sandbox: true,
      pageDir: join(HERE, '../dist/render-page'),
      fontsDir: join(HERE, '../../../assets/fonts'),
      ...defined(input.render),
    },
    oracles: { enabled: false, validatorTimeoutMs: 300_000, bridgeTimeoutMs: 300_000, maxPngBytes: 64 * MB, ...defined(input.oracles) },
    rateLimit: input.rateLimit ?? { max: 600, upload: 60, windowMs: 60_000 },
    pressure: {
      maxEventLoopDelayMs: 1000,
      maxHeapUsedBytes: Math.floor(getHeapStatistics().heap_size_limit * 0.9),
      maxRssBytes: 0,
      sampleIntervalMs: 1000,
      ...defined(input.pressure),
    },
    requestTimeoutMs: input.requestTimeoutMs ?? 300_000,
    logLevel: input.logLevel ?? 'info',
    apiTokens: input.apiTokens ?? [],
    trustProxy: input.trustProxy ?? false,
    webDir: resolve(input.webDir ?? join(HERE, '../../web/dist')),
  };
}

const isLoopback = (host: string) => host === 'localhost' || host === '::1' || /^127\.\d+\.\d+\.\d+$/.test(host);
const yes = (raw: string | undefined) => raw?.trim() === '1' || raw?.trim().toLowerCase() === 'true';

/** API_TOKENS: comma-separated, each at least 32 characters (openssl rand -hex 32 makes one). */
function tokens(raw: string | undefined): string[] {
  const list = (raw ?? '').split(',').map(t => t.trim()).filter(Boolean);
  if (list.some(t => t.length < 32 || /\s/.test(t))) throw new Error('API_TOKENS must be comma-separated tokens of at least 32 characters, without spaces');
  return list;
}

/** TRUST_PROXY: a number of proxy hops, "true" for any, or comma-separated proxy addresses and CIDR ranges. */
function trustProxy(raw: string | undefined): ServerConfig['trustProxy'] | undefined {
  const v = raw?.trim();
  if (!v) return undefined;
  if (/^\d+$/.test(v)) return Number(v);
  if (v === 'true' || v === 'false') return v === 'true';
  return v.split(',').map(a => a.trim()).filter(Boolean);
}

function int(env: NodeJS.ProcessEnv, name: string, min: number, max: number): number | undefined {
  const raw = env[name]?.trim();
  if (raw === undefined || raw === '') return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n) || n < min || n > max) throw new Error(`${name} must be an integer from ${min} to ${max}, not "${raw}"`);
  return n;
}

/** PPTX_VALIDATOR is a path to the validator executable, or a JSON array of argv (e.g. ["dotnet", "validator.dll"]). */
function argv(raw: string | undefined): string[] | undefined {
  const v = raw?.trim();
  if (!v) return undefined;
  if (!v.startsWith('[')) return [v];
  const parsed: unknown = JSON.parse(v);
  if (!Array.isArray(parsed) || parsed.length === 0 || !parsed.every(a => typeof a === 'string' && a.length > 0)) {
    throw new Error('PPTX_VALIDATOR must be a path or a JSON array of strings');
  }
  return parsed as string[];
}

function httpUrl(raw: string | undefined, name: string): string | undefined {
  const v = raw?.trim();
  if (!v) return undefined;
  const url = new URL(v);
  if (url.protocol !== 'http:' && url.protocol !== 'https:') throw new Error(`${name} must be an http(s) URL`);
  if (!url.pathname.endsWith('/')) url.pathname += '/';
  return url.href;
}

/**
 * Settings from the environment: PORT, HOST (default 127.0.0.1), DATA_DIR, MAX_UPLOAD_MB, RENDER_PAGES,
 * ENABLE_ORACLES, PPTX_VALIDATOR, POWERPOINT_BRIDGE_URL, plus WORKERS, WORKER_HEAP_MB, PARSE_TIMEOUT_MS,
 * RENDER_TIMEOUT_MS, RENDER_SANDBOX (0 to run Chromium without its sandbox), MAX_INFLATED_MB, RATE_LIMIT
 * (requests per minute per IP, 0 for none), LOG_LEVEL, API_TOKENS, ALLOW_ANONYMOUS_API, TRUST_PROXY and WEB_DIR.
 * A value that does not parse stops the server at startup rather than running with a surprise, and so does a HOST
 * that reaches the network without API_TOKENS, unless ALLOW_ANONYMOUS_API=1 says something in front authenticates.
 */
export function configFromEnv(env: NodeJS.ProcessEnv): ServerConfig {
  const mb = (v: number | undefined) => (v === undefined ? undefined : v * MB);
  const inflated = mb(int(env, 'MAX_INFLATED_MB', 16, 2048));
  const rate = int(env, 'RATE_LIMIT', 0, 1_000_000);
  const host = env.HOST?.trim() || undefined;
  const apiTokens = tokens(env.API_TOKENS);
  if (host && !isLoopback(host) && !apiTokens.length && !yes(env.ALLOW_ANONYMOUS_API)) {
    throw new Error(`HOST=${host} would open the API to the network without API_TOKENS. Set API_TOKENS, `
      + 'or ALLOW_ANONYMOUS_API=1 when a proxy in front of the server authenticates.');
  }
  return resolveConfig({
    host,
    port: int(env, 'PORT', 0, 65535),
    dataDir: env.DATA_DIR?.trim() || undefined,
    maxUploadBytes: mb(int(env, 'MAX_UPLOAD_MB', 1, 2048)),
    limits: { maxEntryBytes: inflated, maxTotalBytes: inflated },
    workers: {
      size: int(env, 'WORKERS', 1, 64),
      heapMb: int(env, 'WORKER_HEAP_MB', 32, 16384),
      timeoutMs: int(env, 'PARSE_TIMEOUT_MS', 100, 600_000),
    },
    render: {
      pages: int(env, 'RENDER_PAGES', 0, 32),
      timeoutMs: int(env, 'RENDER_TIMEOUT_MS', 1000, 600_000),
      sandbox: env.RENDER_SANDBOX === '0' || env.RENDER_SANDBOX === 'false' ? false : undefined,
    },
    oracles: {
      enabled: env.ENABLE_ORACLES === '1' || env.ENABLE_ORACLES === 'true',
      validator: argv(env.PPTX_VALIDATOR),
      bridgeUrl: httpUrl(env.POWERPOINT_BRIDGE_URL, 'POWERPOINT_BRIDGE_URL'),
    },
    rateLimit: rate === undefined ? undefined
      : rate === 0 ? false
      : { max: rate, upload: Math.max(1, Math.ceil(rate / 10)), windowMs: 60_000 },
    logLevel: env.LOG_LEVEL?.trim() || undefined,
    apiTokens,
    trustProxy: trustProxy(env.TRUST_PROXY),
    webDir: env.WEB_DIR?.trim() || undefined,
  });
}
