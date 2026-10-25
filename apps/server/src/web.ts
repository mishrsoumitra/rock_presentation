// apps/server/src/web.ts  -- the editor (apps/web's build), served from memory next to the API
//
// The editor runs entirely in the browser, so serving it is serving files: index.html, the hashed bundles under
// /assets/ and the fonts. They are read once at startup, so a request can only ever name a file that was there
// then (no paths reach the file system), and compressed once (Brotli and gzip) for the clients that take it.
// index.html carries its content security policy in a meta tag; the same policy goes in a header too, with
// frame-ancestors added, which a meta tag cannot carry. Source maps are not served.
import { createHash } from 'node:crypto';
import { readdir, readFile } from 'node:fs/promises';
import { extname, join, relative, sep } from 'node:path';
import { brotliCompressSync, constants, gzipSync } from 'node:zlib';
import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import { HttpError, notFound } from './errors.ts';

const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json', '.wasm': 'application/wasm', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.ico': 'image/x-icon', '.txt': 'text/plain; charset=utf-8', '.ttf': 'font/ttf', '.otf': 'font/otf', '.woff2': 'font/woff2',
};
const COMPRESSIBLE = new Set(['.html', '.js', '.css', '.json', '.wasm', '.svg', '.txt', '.ttf', '.otf', '.ico']);

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
  for (const entry of await readdir(dir, { recursive: true, withFileTypes: true })) {
    if (!entry.isFile()) continue;
    const path = join(entry.parentPath, entry.name);
    const name = relative(dir, path).split(sep).join('/');
    if (name.split('/').some(s => s.startsWith('.')) || name.endsWith('.map')) continue;
    const ext = extname(name).toLowerCase();
    const body = await readFile(path);
    const file: WebFile = {
      type: TYPES[ext] ?? 'application/octet-stream',
      etag: `"${createHash('sha256').update(body).digest('base64url').slice(0, 22)}"`,
      // Vite names everything under assets/ by its content; the rest (index.html, fonts) keeps its name
      cache: name.startsWith('assets/') ? 'public, max-age=31536000, immutable' : name === 'index.html' ? 'no-cache' : 'public, max-age=86400',
      body,
    };
    if (COMPRESSIBLE.has(ext) && body.length > 1024) {
      const br = brotliCompressSync(body, { params: { [constants.BROTLI_PARAM_QUALITY]: 11, [constants.BROTLI_PARAM_SIZE_HINT]: body.length } });
      const gz = gzipSync(body, { level: 9 });
      if (br.length < body.length * 0.9) file.br = br;
      if (gz.length < body.length * 0.9) file.gzip = gz;
    }
    files.set(name, file);
  }
  return { files, csp: `${meta[1]!.replace(/;\s*$/, '')}; frame-ancestors 'none'` };
}

/** GET (and HEAD) for every path outside /api/: a file of the build, or 404. */
export function webRoutes(app: FastifyInstance, web: WebFiles): void {
  const send = (request: FastifyRequest, reply: FastifyReply, name: string) => {
    const file = web.files.get(name);
    if (!file) throw notFound('file');
    reply.header('content-type', file.type).header('cache-control', file.cache).header('etag', file.etag).header('vary', 'accept-encoding');
    // the editor's own policy replaces the API's (which allows nothing)
    if (name === 'index.html') reply.header('content-security-policy', web.csp);
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

