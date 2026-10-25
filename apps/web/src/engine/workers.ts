// apps/web/src/engine/workers.ts  -- the page's side of the decryption and metafile workers
import { OfficeCryptoError, type OfficeCryptoErrorCode } from '@pptx/crypto';

/** Decrypt a protected deck in a worker; rejects with an OfficeCryptoError (wrong-password, ...) as the library does. */
export function decryptInWorker(bytes: Uint8Array, password: string): Promise<Uint8Array> {
  const worker = new Worker(new URL('./crypto.worker.ts', import.meta.url), { type: 'module' });
  return new Promise<Uint8Array>((resolve, reject) => {
    worker.onmessage = (e: MessageEvent<{ ok: boolean; bytes?: Uint8Array; error?: string; code?: OfficeCryptoErrorCode }>) => {
      if (e.data.ok && e.data.bytes) resolve(e.data.bytes);
      else reject(e.data.code ? new OfficeCryptoError(e.data.code, e.data.error ?? 'decryption failed') : new Error(e.data.error ?? 'decryption failed'));
    };
    worker.onerror = e => reject(new Error(e.message));
    const copy = bytes.slice();
    worker.postMessage({ bytes: copy, password }, [copy.buffer]);
  }).finally(() => worker.terminate());
}

const DEADLINE = 10_000;
let metafiles: Worker | undefined;
let next = 1;
const waiting = new Map<number, { resolve: (b: ImageBitmap | undefined) => void; timer: ReturnType<typeof setTimeout> }>();

function metafileWorker(): Worker {
  if (metafiles) return metafiles;
  const w = new Worker(new URL('./metafile.worker.ts', import.meta.url), { type: 'module' });
  w.onmessage = (e: MessageEvent<{ id: number; bitmap?: ImageBitmap }>) => {
    const job = waiting.get(e.data.id);
    if (!job) return;
    clearTimeout(job.timer);
    waiting.delete(e.data.id);
    job.resolve(e.data.bitmap);
  };
  metafiles = w;
  return w;
}

/** Play an EMF or WMF in the worker; undefined when it cannot be drawn or overruns its deadline (then the worker is replaced). */
export function metafileInWorker(bytes: Uint8Array): Promise<ImageBitmap | undefined> {
  const id = next++;
  return new Promise(resolve => {
    const timer = setTimeout(() => {
      // a picture that takes this long is treated as hostile: everything queued behind it is retried on a new worker
      metafiles?.terminate();
      metafiles = undefined;
      waiting.delete(id);
      resolve(undefined);
      for (const [other, job] of waiting) {
        clearTimeout(job.timer);
        waiting.delete(other);
        job.resolve(undefined);
      }
    }, DEADLINE);
    waiting.set(id, { resolve, timer });
    const copy = bytes.slice();
    metafileWorker().postMessage({ id, bytes: copy }, [copy.buffer]);
  });
}
