// packages/metafile/src/errors.ts  -- the package's one error type, and the limits that stop a whole picture
export type MetafileErrorCode =
  | 'malformed'     // the bytes break the format's structure: a bad header, a record that does not fit
  | 'unsupported'   // a well-formed file this package does not draw
  | 'limit';        // a count, size or work limit was exceeded

/** A metafile that cannot be read or drawn. Metafiles come from untrusted files, so callers catch this. */
export class MetafileError extends Error {
  readonly code: MetafileErrorCode;
  constructor(code: MetafileErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'MetafileError';
    this.code = code;
  }
}

const STOP = Symbol('stop');

/**
 * A limit that stops the whole picture (records, work, nesting, table sizes), as opposed to one record that is
 * too big or broken, which is skipped like GDI skips a call that fails.
 */
export const stop = (message: string) => new MetafileError('limit', message, { cause: STOP });
export const stops = (e: unknown): boolean => e instanceof MetafileError && e.cause === STOP;
