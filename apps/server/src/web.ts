// apps/server/src/web.ts  -- the editor (apps/web's build), served from memory next to the API
//
// The editor runs entirely in the browser, so serving it is serving files: index.html, the hashed bundles under
// /assets/ and the fonts. They are read once at startup, so a request can only ever name a file that was there
// then (no paths reach the file system). The build leaves a Brotli and a gzip copy next to every file worth
// compressing (apps/web/precompress.ts); they are read here and sent to the clients that take them. A file the
// build left no copy of is compressed in the background once the server is up, so a start never waits for it.
// index.html carries its content security policy in a meta tag; the same policy goes in a header on every file,
// with frame-ancestors added (which a meta tag cannot carry): workers get their policy from their script's
// response, not from the page. Source maps and the compressed copies are not served as files of their own.
import { createHash } from 'node:crypto';
import { readdir, readFile } from 'node:fs/promises';
import { extname, join, relative, sep } from 'node:path';
import { promisify } from 'node:util';
import { brotliCompress, constants, gzip } from 'node:zlib';
import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import { HttpError, notFound } from './errors.ts';

const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json', '.wasm': 'application/wasm', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.ico': 'image/x-icon', '.txt': 'text/plain; charset=utf-8', '.ttf': 'font/ttf', '.otf': 'font/otf', '.woff2': 'font/woff2',
};
const COMPRESSIBLE = new Set(['.html', '.js', '.css', '.json', '.wasm', '.svg', '.txt', '.ttf', '.otf', '.ico']);
const MIN_BYTES = 1024;
const WORTH_IT = 0.9;

const brotli = promisify(brotliCompress);
const gzipped = promisify(gzip);

interface WebFile {
  type: string;
  etag: string;
  cache: string;
  body: Buffer;
  br?: Buffer;
  gzip?: Buffer;
}

export interface WebFiles {
  files: Map<string, WebFile>;
  /** index.html's policy, with frame-ancestors added. */
  csp: string;
  /** Settles once every file the build left no compressed copy of has been compressed here, in the background. */
  compressed: Promise<void>;
}

/** The build's compressed copy at `path`, or undefined when the build left none. */
async function copyAt(path: string): Promise<Buffer | undefined> {
  try {
    return await readFile(path);
  } catch {
    return undefined;
  }
}

/** Brotli (quality 11) and gzip for each file, one file at a time and off the event loop; a copy that saves less than a tenth is dropped. */
async function compressLater(files: WebFile[]): Promise<void> {
  for (const file of files) {
    try {
      const [br, gz] = await Promise.all([
        brotli(file.body, { params: { [constants.BROTLI_PARAM_QUALITY]: 11, [constants.BROTLI_PARAM_SIZE_HINT]: file.body.length } }),
        gzipped(file.body, { level: 9 }),
      ]);
      if (br.length < file.body.length * WORTH_IT) file.br = br;
      if (gz.length < file.body.length * WORTH_IT) file.gzip = gz;
    } catch {
      // the file is served as it is
    }
  }
}

/** The built editor in `dir`, or undefined when there is none (no index.html): the server then serves only the API. */
export async function loadWeb(dir: string): Promise<WebFiles | undefined> {
  let index: Buffer;
  try {
    index = await readFile(join(dir, 'index.html'));
  } catch {
    return undefined;
  }
  const meta = /<meta\s+http-equiv="Content-Security-Policy"\s+content="([^"]+)"/i.exec(index.toString('utf8'));
  if (!meta) throw new Error(`${join(dir, 'index.html')} has no Content-Security-Policy meta tag: was it built with vite build?`);
  const files = new Map<string, WebFile>();
  const later: WebFile[] = [];
  for (const entry of await readdir(dir, { recursive: true, withFileTypes: true })) {
    if (!entry.isFile()) continue;
    const path = join(entry.parentPath, entry.name);
    const name = relative(dir, path).split(sep).join('/');
    if (name.split('/').some(s => s.startsWith('.')) || /\.(map|br|gz)$/.test(name)) continue;
    const ext = extname(name).toLowerCase();
    const body = await readFile(path);
    const file: WebFile = {
      type: TYPES[ext] ?? 'application/octet-stream',
      etag: `"${createHash('sha256').update(body).digest('base64url').slice(0, 22)}"`,
      // Vite names everything under assets/ by its content; the rest (index.html, fonts) keeps its name
      cache: name.startsWith('assets/') ? 'public, max-age=31536000, immutable' : name === 'index.html' ? 'no-cache' : 'public, max-age=86400',
      body,
    };
    if (COMPRESSIBLE.has(ext) && body.length > MIN_BYTES) {
      file.br = await copyAt(`${path}.br`);
      file.gzip = await copyAt(`${path}.gz`);
      if (!file.br && !file.gzip) later.push(file);
    }
    files.set(name, file);
  }
  return { files, csp: `${meta[1]!.replace(/;\s*$/, '')}; frame-ancestors 'none'`, compressed: compressLater(later) };
}

/** GET (and HEAD) for every path outside /api/: a file of the build, or 404. */
export function webRoutes(app: FastifyInstance, web: WebFiles): void {
  const send = (request: FastifyRequest, reply: FastifyReply, name: string) => {
    const file = web.files.get(name);
    if (!file) throw notFound('file');
    reply.header('content-type', file.type).header('cache-control', file.cache).header('etag', file.etag).header('vary', 'accept-encoding');
    // the editor's own policy replaces the API's (which allows nothing) on every file: a worker takes its policy
    // from its own script's response, so worker bundles need it as much as the page does
    reply.header('content-security-policy', web.csp);
    if (request.headers['if-none-match'] === file.etag) return reply.code(304).send();
    const accept = String(request.headers['accept-encoding'] ?? '');
    if (file.br && /\bbr\b/.test(accept)) return reply.header('content-encoding', 'br').send(file.br);
    if (file.gzip && /\bgzip\b/.test(accept)) return reply.header('content-encoding', 'gzip').send(file.gzip);
    return reply.send(file.body);
  };
  app.get('/', (request, reply) => send(request, reply, 'index.html'));
  app.get('/*', (request, reply) => {
    const name = (request.params as { '*': string })['*'];
    // unknown API routes stay the API's JSON 404
    if (name === 'api' || name.startsWith('api/')) throw new HttpError(404, 'not-found', 'No such route.');
    return send(request, reply, name);
  });
}
