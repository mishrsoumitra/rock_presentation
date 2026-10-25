// packages/crypto/src/dataspaces.ts
// The \x06DataSpaces structures Office writes for a strong-encrypted package (MS-OFFCRYPTO 2.1 and 2.3.4.x):
// Version, DataSpaceMap, DataSpaceInfo/StrongEncryptionDataSpace, TransformInfo/StrongEncryptionTransform/\x06Primary.
// These streams are format-fixed: the bytes below were taken from a real PowerPoint-encrypted deck and verified
// byte-for-byte identical across the Open XML SDK sample, two Apache POI samples and msoffcrypto-tool's output
// (see test/dataspaces.test.ts). They are embedded verbatim so our encryptor lays out what PowerPoint expects.

function hex(s: string): Uint8Array {
  const out = new Uint8Array(s.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = parseInt(s.slice(i * 2, i * 2 + 2), 16);
  return out;
}

// "Microsoft.Container.DataSpaces" version record (MS-OFFCRYPTO 2.1.5).
export const VERSION = hex('3c0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e004400610074006100530070006100630065007300010000000100000001000000');

// DataSpaceMap: maps the "EncryptedPackage" stream to the "StrongEncryptionDataSpace" (MS-OFFCRYPTO 2.1.6).
export const DATA_SPACE_MAP = hex('08000000010000006800000001000000000000002000000045006e0063007200790070007400650064005000610063006b00610067006500320000005300740072006f006e00670045006e006300720079007000740069006f006e004400610074006100530070006100630065000000');

// StrongEncryptionDataSpace: lists the single transform "StrongEncryptionTransform" (MS-OFFCRYPTO 2.1.7/2.1.8).
export const STRONG_ENCRYPTION_DATA_SPACE = hex('0800000001000000320000005300740072006f006e00670045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d000000');

// \x06Primary: the StrongEncryptionTransform transform info and the AES128 feature GUID (MS-OFFCRYPTO 2.1.9/2.2.6).
export const PRIMARY = hex('58000000010000004c0000007b00460046003900410033004600300033002d0035003600450046002d0034003600310033002d0042004400440035002d003500410034003100430031004400300037003200340036007d004e0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e0045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d00000001000000010000000100000000000000000000000000000004000000');
