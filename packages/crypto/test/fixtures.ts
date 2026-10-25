// packages/crypto/test/fixtures.ts
// Test fixtures and their documented passwords. Tests run in Node and read these from disk. Every encrypted
// fixture was produced by a different implementation than ours, so decrypting it is a real cross-check:
//   real/*            encrypted by Microsoft Office / Apache POI (and the SDK) -- see provenance below
//   msoffcrypto/*     encrypted by the Python msoffcrypto-tool (the reference); plain/* holds the originals
// These are third-party test files, treated as untrusted data (we only read their bytes).
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
export const FIXTURES = join(HERE, 'fixtures');

export function readFixture(rel: string): Uint8Array {
  return new Uint8Array(readFileSync(join(FIXTURES, rel)));
}

export interface EncryptedFixture {
  file: string;
  password: string;
  scheme: 'agile' | 'standard';
  /** The original plaintext ZIP, when we hold it, for a byte-for-byte decrypt check. */
  plain?: string;
  source: string;
}

export const ENCRYPTED_FIXTURES: EncryptedFixture[] = [
  // Encrypted by the Python msoffcrypto-tool (reference), original committed alongside.
  {
    file: 'msoffcrypto/default-agile.pptx', password: 'pptx-crypto-test-2026', scheme: 'agile',
    plain: 'plain/default.pptx',
    source: 'python-pptx default template, encrypted by msoffcrypto-tool 6.0.0',
  },
  // Encrypted by Microsoft Office / Apache POI, independent of both msoffcrypto and us.
  {
    file: 'real/excel-standard.xlsx', password: 'VelvetSweatshop', scheme: 'standard',
    source: 'Apache POI test-data/poifs/protect.xlsx (Excel standard encryption, AES-128/SHA-1); POI DEFAULT_PASSWORD',
  },
  {
    file: 'real/word-agile.docx', password: 'VelvetSweatshop', scheme: 'agile',
    source: 'Apache POI test-data/poifs/protected_agile.docx (Word agile encryption); POI DEFAULT_PASSWORD',
  },
  {
    file: 'real/excel-agile-aes256.xlsx', password: 'Test001!!', scheme: 'agile',
    source: 'Apache POI test-data/poifs/60320-protected.xlsx (agile AES-256/SHA-512); POI TestDecryptor',
  },
];

// A genuine PowerPoint-encrypted .pptx whose password is not documented anywhere in the Open XML SDK repo (the
// SDK uses it only to test encryption *detection*, never decryption). Kept for detection, descriptor-parse and
// wrong-password tests on real PowerPoint output.
export const POWERPOINT_AGILE_UNKNOWN_PW = 'real/powerpoint-agile.pptx';
