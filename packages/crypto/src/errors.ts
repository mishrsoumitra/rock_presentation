// packages/crypto/src/errors.ts -- typed errors, in the spirit of OpcError in packages/opc/src/errors.ts.

export type OfficeCryptoErrorCode =
  | 'not-encrypted'        // asked to decrypt something that is not an encrypted compound file
  | 'malformed'            // the compound file or an EncryptionInfo structure cannot be read
  | 'unsupported'          // an encryption scheme we do not implement (RC4, XOR, extensible, non-AES/SHA agile)
  | 'wrong-password'       // the password failed the verifier
  | 'integrity'            // the data-integrity HMAC did not match: tampering or corruption
  | 'limit';               // a size/structure limit was exceeded before or during parsing

export class OfficeCryptoError extends Error {
  readonly code: OfficeCryptoErrorCode;
  constructor(code: OfficeCryptoErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'OfficeCryptoError';
    this.code = code;
  }
}
