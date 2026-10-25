// packages/geometry/src/errors.ts

/** A geometry that cannot be read or evaluated. Custom geometry comes from untrusted files, so callers catch this. */
export class GeometryError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'GeometryError';
  }
}
