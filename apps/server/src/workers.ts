// apps/server/src/workers.ts  -- a bounded pool of worker threads for everything that parses untrusted bytes
//
// Each thread runs under V8 resource limits (old and young generation, stack) and each job under a wall-clock
// deadline: a thread that runs out of heap is stopped by V8, one that overruns is terminated here. Either way only
// that job fails and the thread is replaced by a fresh one. Jobs beyond the pool's size wait in a bounded queue;
// when that is full, run() rejects at once with 'busy', which the routes answer with 503.
import { Worker } from 'node:worker_threads';
import type { WorkerSettings } from './config.ts';
import type { TaskName, Tasks } from './tasks.ts';

export type WorkerFailure = 'busy' | 'timeout' | 'out-of-memory' | 'crashed' | 'closed';

/** A job that got no answer from its task: the pool was full or closing, or the thread overran, ran out of memory or died. */
export class WorkerError extends Error {
  readonly code: WorkerFailure;
  constructor(code: WorkerFailure, message: string) {
    super(message);
    this.name = 'WorkerError';
    this.code = code;
  }
}

/** A refusal the task reported itself: the code of the OpcError, OfficeCryptoError or TaskError it threw. */
export class TaskFailure extends Error {
  readonly code: string;
  constructor(code: string, message: string) {
    super(message);
    this.name = 'TaskFailure';
    this.code = code;
  }
}

interface Job {
  id: number;
  task: TaskName;
  input: unknown;
  transfer: ArrayBuffer[];
  timeoutMs: number;
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
}

interface Slot { worker: Worker; ready: boolean; job?: Job; timer?: NodeJS.Timeout; }

type Reply = { ready: true } | { id: number; ok: true; result: unknown } | { id: number; ok: false; error: { code: string; message: string } };

const WORKER_URL = new URL('./worker.ts', import.meta.url);
/** Extra time for a job that lands on a thread still loading its modules. */
const STARTUP_GRACE_MS = 15_000;

export interface PoolStats { size: number; threads: number; busy: number; queued: number; timeouts: number; outOfMemory: number; crashes: number; }
export interface RunOptions { transfer?: ArrayBuffer[]; timeoutMs?: number; }

export class WorkerPool {
  private readonly settings: WorkerSettings;
  private readonly slots = new Set<Slot>();
  private readonly idle: Slot[] = [];
  private readonly queue: Job[] = [];
  private nextId = 1;
  private closed = false;
  private readonly counts = { timeouts: 0, outOfMemory: 0, crashes: 0 };

  constructor(settings: WorkerSettings) {
    this.settings = settings;
  }

  /**
   * Run a task on a thread. `transfer` lists buffers of the input to move rather than copy (the caller loses them).
   * Rejects with TaskFailure when the task refused the input, WorkerError otherwise.
   */
  run<K extends TaskName>(task: K, input: Tasks[K]['input'], options: RunOptions = {}): Promise<Tasks[K]['output']> {
    if (this.closed) return Promise.reject(new WorkerError('closed', 'The server is shutting down.'));
    const free = this.idle.length + this.settings.size - this.slots.size;
    if (free <= 0 && this.queue.length >= this.settings.queue) {
      return Promise.reject(new WorkerError('busy', 'Every worker is busy and the queue is full.'));
    }
    return new Promise((resolve, reject) => {
      this.queue.push({
        id: this.nextId++, task, input, transfer: options.transfer ?? [], timeoutMs: options.timeoutMs ?? this.settings.timeoutMs,
        resolve: resolve as (value: unknown) => void, reject,
      });
      this.pump();
    });
  }

  stats(): PoolStats {
    let busy = 0;
    for (const s of this.slots) if (s.job) busy++;
    return { size: this.settings.size, threads: this.slots.size, busy, queued: this.queue.length, ...this.counts };
  }

  /** Fail queued and running jobs and stop every thread. */
  async close(): Promise<void> {
    this.closed = true;
    for (const job of this.queue.splice(0)) job.reject(new WorkerError('closed', 'The server is shutting down.'));
    const workers = [...this.slots].map(s => s.worker);
    for (const s of [...this.slots]) this.retire(s, new WorkerError('closed', 'The server is shutting down.'));
    await Promise.all(workers.map(w => w.terminate()));
  }

  private pump(): void {
    while (this.queue.length > 0 && !this.closed) {
      const slot = this.idle.pop() ?? (this.slots.size < this.settings.size ? this.spawn() : undefined);
      if (!slot) return;
      this.start(slot, this.queue.shift()!);
    }
  }

  private spawn(): Slot {
    const worker = new Worker(WORKER_URL, {
      name: 'pptx-parser',
      workerData: { testTasks: this.settings.testTasks },
      env: {},     // the parser needs none of the server's environment
      resourceLimits: {
        maxOldGenerationSizeMb: this.settings.heapMb,
        maxYoungGenerationSizeMb: this.settings.youngMb,
        stackSizeMb: this.settings.stackMb,
      },
    });
    worker.unref();
    const slot: Slot = { worker, ready: false };
    worker.on('message', (m: Reply) => this.onMessage(slot, m));
    worker.on('error', (e: Error & { code?: string }) => {
      const oom = e.code === 'ERR_WORKER_OUT_OF_MEMORY';
      this.counts[oom ? 'outOfMemory' : 'crashes']++;
      this.retire(slot, oom
        ? new WorkerError('out-of-memory', 'The job ran out of memory.')
        : new WorkerError('crashed', `The worker failed: ${e.message}`));
    });
    worker.on('exit', code => {
      if (slot.job) this.counts.crashes++;
      this.retire(slot, new WorkerError('crashed', `The worker stopped (exit code ${code}).`));
    });
    this.slots.add(slot);
    return slot;
  }

  /** The job's deadline; on a thread still loading, it starts over (without the grace) once the thread is ready. */
  private arm(slot: Slot, job: Job): void {
    clearTimeout(slot.timer);
    slot.timer = setTimeout(() => {
      this.counts.timeouts++;
      this.retire(slot, new WorkerError('timeout', `The job took longer than ${job.timeoutMs} ms.`));
      void slot.worker.terminate();
    }, job.timeoutMs + (slot.ready ? 0 : STARTUP_GRACE_MS));
  }

  private start(slot: Slot, job: Job): void {
    slot.job = job;
    this.arm(slot, job);
    try {
      slot.worker.postMessage({ id: job.id, task: job.task, input: job.input }, job.transfer);
    } catch (e) {
      // the input could not be cloned; the thread is fine
      clearTimeout(slot.timer);
      slot.job = undefined;
      this.idle.push(slot);
      job.reject(e instanceof Error ? e : new Error(String(e)));
    }
  }

  private onMessage(slot: Slot, m: Reply): void {
    if ('ready' in m) {
      slot.ready = true;
      if (slot.job) this.arm(slot, slot.job);
      return;
    }
    const job = slot.job;
    if (!job || m.id !== job.id) return;
    clearTimeout(slot.timer);
    slot.job = undefined;
    if (m.ok) job.resolve(m.result);
    else job.reject(new TaskFailure(m.error.code, m.error.message));
    this.idle.push(slot);
    this.pump();
  }

  /** Take a thread out of the pool and fail its job. Runs once per thread; the next job gets a fresh thread. */
  private retire(slot: Slot, failure: WorkerError): void {
    if (!this.slots.delete(slot)) return;
    const at = this.idle.indexOf(slot);
    if (at >= 0) this.idle.splice(at, 1);
    clearTimeout(slot.timer);
    const job = slot.job;
    slot.job = undefined;
    job?.reject(failure);
    this.pump();
  }
}
