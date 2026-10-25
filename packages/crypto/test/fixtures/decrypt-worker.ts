// Decrypts inside a worker thread so a test can put a deadline on hostile input: a synchronous hang in the main
// thread could never be timed out. Mirrors packages/opc's open-worker.ts.
import { parentPort, workerData } from 'node:worker_threads';
import { decryptPackage, OfficeCryptoError } from '../../src/index.ts';

const { bytes, password } = workerData as { bytes: Uint8Array; password: string };
decryptPackage(bytes, password).then(
  out => parentPort!.postMessage({ ok: true, length: out.length }),
  (e: unknown) => parentPort!.postMessage({
    ok: false,
    code: e instanceof OfficeCryptoError ? e.code : `crash:${e instanceof Error ? e.name : typeof e}`,
    message: e instanceof Error ? e.message : String(e),
  }),
);
