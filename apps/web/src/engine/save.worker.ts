// apps/web/src/engine/save.worker.ts  -- zipping (and encrypting) a deck off the main thread
import { encryptPackage } from '@pptx/crypto';
import { writeZip } from '@pptx/opc';

export interface SaveRequest {
  entries: { name: string; data: Uint8Array; store: boolean }[];
  /** Encrypt with this password (the deck was opened with one, or the user set one). Never leaves the browser. */
  password?: string;
}

self.onmessage = async (e: MessageEvent<SaveRequest>) => {
  try {
    let bytes = writeZip(e.data.entries);
    if (e.data.password) bytes = await encryptPackage(bytes, e.data.password);
    (self as unknown as Worker).postMessage({ ok: true, bytes }, [bytes.buffer]);
  } catch (err) {
    (self as unknown as Worker).postMessage({ ok: false, error: err instanceof Error ? err.message : String(err) });
  }
};
