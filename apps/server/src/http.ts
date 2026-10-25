// apps/server/src/http.ts  -- request and response helpers: multipart uploads, file names, ETags, Content-Disposition
import type { FastifyRequest } from 'fastify';
import type {} from '@fastify/multipart';
import Value from 'typebox/value';
import type { TSchema } from 'typebox';
import { HttpError } from './errors.ts';

export interface UploadedFile { bytes: Buffer; filename: string | undefined; mimetype: string; }
export interface Upload { files: Map<string, UploadedFile>; fields: Record<string, string>; }
export type UploadLimits = NonNullable<Parameters<FastifyRequest['parts']>[0]>['limits'];

/**
 * Read a multipart/form-data upload: exactly the file parts named in `files`, each once, and small text fields
 * that must satisfy `fieldsSchema`. Sizes and counts are capped by the multipart plugin (busboy stops reading past
 * them and the request fails with 413); `limits` tightens or loosens them for one route.
 */
export async function readUpload(request: FastifyRequest, files: string[], fieldsSchema: TSchema, limits?: UploadLimits): Promise<Upload> {
  if (!request.isMultipart()) throw new HttpError(415, 'unsupported-media-type', 'Send the file as multipart/form-data.');
  const out = new Map<string, UploadedFile>();
  const fields: Record<string, string> = {};
  for await (const part of request.parts(limits ? { limits } : {})) {
    if (part.type === 'file') {
      if (!files.includes(part.fieldname) || out.has(part.fieldname)) {
        part.file.resume();
        throw new HttpError(400, 'unexpected-file', `The form has an unexpected file field "${part.fieldname.slice(0, 64)}".`);
      }
      out.set(part.fieldname, { bytes: await part.toBuffer(), filename: part.filename || undefined, mimetype: part.mimetype });
    } else {
      if (part.fieldnameTruncated || part.valueTruncated) throw new HttpError(413, 'too-large', 'A form field is too long.');
      fields[part.fieldname] = String(part.value);
    }
  }
  if (!Value.Check(fieldsSchema, fields)) throw new HttpError(400, 'invalid-request', 'The form has unexpected or invalid fields.');
  for (const name of files) {
    const f = out.get(name);
    if (!f) throw new HttpError(400, 'missing-file', `The form has no "${name}" file.`);
    if (f.bytes.length === 0) throw new HttpError(400, 'empty-file', `The "${name}" file is empty.`);
  }
  return { files: out, fields };
}

/**
 * A display name from an upload: the last path segment only, without control and format characters (bidirectional
 * overrides could disguise an extension), at most 200 characters. Undefined when nothing is left.
 */
export function cleanName(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const base = raw.normalize('NFC').replace(/^.*[\\/]/, '').replace(/[\p{Cc}\p{Cf}]/gu, '').trim();
  const chars = [...base];
  const name = chars.length > 200 ? chars.slice(0, 200).join('').trim() : base;
  return name || undefined;
}

/** The download's file name: the deck's name with the extension of what is sent. */
export function downloadName(name: string, ext: string): string {
  const base = name.replace(/\.[A-Za-z0-9]{1,5}$/, '').trim() || 'presentation';
  return `${base}.${ext}`;
}

/** RFC 6266 Content-Disposition, with an ASCII fallback name and the full name in RFC 8187 encoding. */
export function contentDisposition(kind: 'attachment' | 'inline', filename: string): string {
  const fallback = filename.replace(/[^\x20-\x7e]/g, '_').replace(/["\\]/g, '_');
  const encoded = encodeURIComponent(filename).replace(/['()*]/g, c => `%${c.charCodeAt(0).toString(16).toUpperCase()}`);
  return `${kind}; filename="${fallback}"; filename*=UTF-8''${encoded}`;
}

/** RFC 9110 If-Match: "*" or a list of entity tags compared strongly, so a weak tag never matches. */
export function ifMatchSatisfied(header: string, etag: string): boolean {
  if (header.trim() === '*') return true;
  return header.split(',').some(t => t.trim() === etag);
}

/** RFC 9110 If-None-Match, compared weakly: true when the client already holds this representation. */
export function ifNoneMatchHit(header: string | undefined, etag: string): boolean {
  if (!header) return false;
  if (header.trim() === '*') return true;
  return header.split(',').some(t => t.trim().replace(/^W\//, '') === etag);
}
