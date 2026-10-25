// apps/server/src/schemas.ts  -- TypeBox schemas: every route's params, query and headers, and its JSON responses
//
// Fastify validates requests against these with Ajv (unknown query parameters are dropped) and serializes JSON
// responses through them, so a field that is not declared here never reaches a client.
import Type, { type Static } from 'typebox';

const DeckId = Type.String({ pattern: '^[0-9a-f]{32}$', description: 'a deck id: 128 random bits, hex' });
const VersionNumber = Type.Integer({ minimum: 1, maximum: 1_000_000 });

export const NoQuery = Type.Object({});
export const DeckParams = Type.Object({ id: DeckId });
export const SlideParams = Type.Object({ id: DeckId, n: Type.Integer({ minimum: 1, maximum: 100_000 }) });
export const VersionQuery = Type.Object({ version: Type.Optional(VersionNumber) });
export const SlideQuery = Type.Object({ version: Type.Optional(VersionNumber), width: Type.Optional(Type.Integer({ minimum: 16, maximum: 3840 })) });
export const IfMatchHeaders = Type.Object({ 'if-match': Type.Optional(Type.String({ maxLength: 2048 })) });
export const IfNoneMatchHeaders = Type.Object({ 'if-none-match': Type.Optional(Type.String({ maxLength: 2048 })) });

/** Text fields an upload may carry besides the file. */
export const UploadFields = Type.Object({ name: Type.Optional(Type.String({ minLength: 1, maxLength: 1024 })) }, { additionalProperties: false });
export const NoFields = Type.Object({}, { additionalProperties: false });

export const CompareQuery = Type.Object({
  threshold: Type.Optional(Type.Number({ minimum: 0, maximum: 1 })),
  includeAA: Type.Optional(Type.Boolean()),
  diff: Type.Optional(Type.Boolean()),
});
export const ValidatorQuery = Type.Object({
  format: Type.Optional(Type.String({ pattern: '^[A-Za-z0-9]{1,32}$', description: 'FileFormatVersions, e.g. Microsoft365' })),
});

const SlideSize = Type.Union([Type.Object({ cx: Type.Number(), cy: Type.Number() }), Type.Null()]);

export const VersionView = Type.Object({
  version: Type.Integer(),
  sha256: Type.String(),
  size: Type.Integer(),
  createdAt: Type.String(),
  encrypted: Type.Boolean(),
  slideCount: Type.Union([Type.Integer(), Type.Null()]),
  contentType: Type.String(),
  macros: Type.Boolean(),
});

export const DeckView = Type.Object({
  id: Type.String(),
  name: Type.String(),
  etag: Type.String({ description: 'the current version\'s entity tag, for If-Match' }),
  version: Type.Integer(),
  size: Type.Integer(),
  slideCount: Type.Union([Type.Integer(), Type.Null()]),
  slideSize: SlideSize,
  encrypted: Type.Boolean(),
  macros: Type.Boolean(),
  contentType: Type.String(),
  createdAt: Type.String(),
  updatedAt: Type.String(),
  versions: Type.Array(VersionView),
});
export type DeckViewT = Static<typeof DeckView>;

export const ErrorView = Type.Object({ statusCode: Type.Integer(), error: Type.String(), message: Type.String() });

const Counts = Type.Record(Type.String(), Type.Union([Type.Number(), Type.Boolean()]));
export const HealthView = Type.Object({ status: Type.String(), uptimeSeconds: Type.Number(), workers: Counts, render: Counts });

export const CompareView = Type.Object({
  width: Type.Integer(),
  height: Type.Integer(),
  mismatched: Type.Integer(),
  ratio: Type.Number(),
  threshold: Type.Number(),
  diffPng: Type.Optional(Type.String({ description: 'base64' })),
});

export const RoundtripView = Type.Object({
  lossless: Type.Boolean(),
  reserializedLossless: Type.Boolean(),
  parts: Type.Integer(),
  xmlParts: Type.Integer(),
  unparseableXmlParts: Type.Integer(),
  differences: Type.Array(Type.String()),
  diagnostics: Type.Array(Type.String()),
});

export const ValidatorView = Type.Object({
  valid: Type.Boolean(),
  exitCode: Type.Integer(),
  errorCount: Type.Union([Type.Integer(), Type.Null()]),
  errors: Type.Array(Type.Record(Type.String(), Type.Unknown())),
  error: Type.Optional(Type.String()),
});

/** Error answers every route may give, by status class. */
export const ERRORS = { '4xx': ErrorView, '5xx': ErrorView };
