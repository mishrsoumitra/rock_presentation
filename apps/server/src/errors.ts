// apps/server/src/errors.ts  -- errors that carry their HTTP answer, and the one JSON shape every error is sent in
import type { FastifyError, FastifyReply, FastifyRequest } from 'fastify';

/** An error with the status, a stable machine-readable code and a message safe to show the client. */
export class HttpError extends Error {
  readonly statusCode: number;
  readonly code: string;
  readonly headers: Record<string, string>;
  constructor(statusCode: number, code: string, message: string, headers: Record<string, string> = {}) {
    super(message);
    this.name = 'HttpError';
    this.statusCode = statusCode;
    this.code = code;
    this.headers = headers;
  }
}

export const notFound = (what = 'deck') => new HttpError(404, 'not-found', `No such ${what}.`);
export const busy = (what: string) => new HttpError(503, 'busy', `The server is busy ${what}; try again shortly.`, { 'retry-after': '5' });

/**
 * The answer for a file the worker refused, by the refusal's code (OpcError, OfficeCryptoError, or our own): files
 * that are not PowerPoint Open XML at all are 415, files over a size limit 413, a password-protected file where the
 * server would have to open it 409, broken packages 422.
 */
export function refusal(code: string, message: string): HttpError {
  switch (code) {
    case 'not-a-package': case 'legacy-ppt': case 'unsupported-container': case 'not-a-presentation':
      return new HttpError(415, code, message);
    case 'limit':
      return new HttpError(413, code, message);
    case 'encrypted':
      return new HttpError(409, code, message);
    default:
      return new HttpError(422, code, message);
  }
}

/** The body of every error response. Never a stack trace; the message is cut short. */
export interface ErrorBody { statusCode: number; error: string; message: string; }

// Fastify's and its plugins' own error codes, as the stable codes this API answers with
const FASTIFY_CODES: Record<string, string> = {
  FST_ERR_VALIDATION: 'invalid-request',
  FST_ERR_CTP_INVALID_MEDIA_TYPE: 'unsupported-media-type',
  FST_ERR_CTP_BODY_TOO_LARGE: 'too-large',
  FST_ERR_CTP_EMPTY_JSON_BODY: 'invalid-request',
  FST_ERR_CTP_INVALID_JSON_BODY: 'invalid-request',
  FST_REQ_FILE_TOO_LARGE: 'too-large',
  FST_FILES_LIMIT: 'too-many-files',
  FST_FIELDS_LIMIT: 'too-many-fields',
  FST_PARTS_LIMIT: 'too-many-parts',
  FST_INVALID_MULTIPART_CONTENT_TYPE: 'unsupported-media-type',
  FST_PROTO_VIOLATION: 'invalid-request',
  FST_INVALID_JSON_FIELD_ERROR: 'invalid-request',
};

/** The Fastify error handler: HttpErrors as they are, client errors with their own message, anything else as 500. */
export function sendError(error: FastifyError | HttpError | Error, request: FastifyRequest, reply: FastifyReply): void {
  let status = (error as { statusCode?: number }).statusCode ?? 500;
  if (!Number.isInteger(status) || status < 400 || status > 599) status = 500;
  let code: string;
  let message: string;
  if (error instanceof HttpError) {
    code = error.code;
    message = error.message;
    for (const [k, v] of Object.entries(error.headers)) reply.header(k, v);
  } else if (status < 500) {
    const fst = (error as FastifyError).code;
    code = (fst && FASTIFY_CODES[fst]) ?? (status === 429 ? 'rate-limited' : 'bad-request');
    if (fst === 'FST_INVALID_MULTIPART_CONTENT_TYPE') status = 415;
    message = error.message;
  } else {
    code = status === 503 ? 'unavailable' : 'internal';
    message = status === 503 ? 'The server cannot take this request right now.' : 'Something went wrong on the server.';
  }
  if (status >= 500) request.log.error({ err: error }, 'request failed');
  else request.log.info({ code, status }, 'request refused');
  const body: ErrorBody = { statusCode: status, error: code, message: message.slice(0, 500) };
  reply.code(status).type('application/json; charset=utf-8').send(body);
}
