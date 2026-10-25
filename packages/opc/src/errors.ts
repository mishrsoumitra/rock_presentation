// packages/opc/src/errors.ts
export type OpcErrorCode =
  | 'not-a-package'          // neither a ZIP nor a compound file
  | 'encrypted'              // a password-protected package inside a compound file (MS-OFFCRYPTO)
  | 'legacy-ppt'             // the binary PowerPoint 97-2003 format (MS-PPT), out of scope
  | 'unsupported-container'  // some other compound file
  | 'invalid-zip'            // the ZIP structure cannot be read
  | 'limit'                  // a size or count limit was exceeded before inflation
  | 'unsafe-name'            // a ZIP entry name that cannot be a part name
  | 'invalid-package'        // readable ZIP, but not an OPC package
  | 'invalid-part'           // an XML part that does not parse
  | 'missing-part'
  | 'part-exists'
  | 'missing-relationship'
  | 'dangling-relationship'  // a relationship whose target part is not in the package
  | 'referenced-part';       // removing a part something still points at

export class OpcError extends Error {
  readonly code: OpcErrorCode;
  constructor(code: OpcErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'OpcError';
    this.code = code;
  }
}
