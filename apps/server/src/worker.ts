// apps/server/src/worker.ts  -- a worker thread of the pool in src/workers.ts: runs one task at a time, answers by id
import { parentPort, workerData } from 'node:worker_threads';
import { describeError, runTask } from './tasks.ts';

const { testTasks } = workerData as { testTasks: boolean };
const port = parentPort!;

port.on('message', (msg: { id: number; task: string; input: unknown }) => {
  try {
    const { result, transfer } = runTask(msg.task, msg.input, testTasks);
    port.postMessage({ id: msg.id, ok: true, result }, transfer);
  } catch (e) {
    port.postMessage({ id: msg.id, ok: false, error: describeError(e) });
  }
});
// modules are loaded: from here on the pool's deadline is the job's own
port.postMessage({ ready: true });
