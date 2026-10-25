// Opens a package inside a worker thread, so a test can put a time limit on hostile input:
// a synchronous hang in the main thread could never be timed out.
import { parentPort, workerData } from 'node:worker_threads';
import { OpcError, OpcPackage } from '../../src/index.ts';

const { bytes, parse } = workerData as { bytes: Uint8Array; parse?: string };
try {
  const pkg = OpcPackage.open(bytes);
  if (parse) pkg.xml(parse);
  parentPort!.postMessage({ ok: true, parts: pkg.partNames().length, diagnostics: pkg.diagnostics });
} catch (e) {
  parentPort!.postMessage({
    ok: false,
    code: e instanceof OpcError ? e.code : `crash:${e instanceof Error ? e.name : typeof e}`,
    message: e instanceof Error ? e.message : String(e),
  });
}
