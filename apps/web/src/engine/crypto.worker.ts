// apps/web/src/engine/crypto.worker.ts  -- decrypting off the main thread: the password's spin count is up to the file
import { decryptPackage, OfficeCryptoError } from '@pptx/crypto';

self.onmessage = async (e: MessageEvent<{ bytes: Uint8Array; password: string }>) => {
  try {
    const bytes = await decryptPackage(e.data.bytes, e.data.password);
    (self as unknown as Worker).postMessage({ ok: true, bytes }, [bytes.buffer]);
  } catch (err) {
    (self as unknown as Worker).postMessage({
      ok: false, error: err instanceof Error ? err.message : String(err), code: err instanceof OfficeCryptoError ? err.code : undefined,
    });
  }
};
