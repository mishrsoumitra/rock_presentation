// apps/server/src/jobs.ts  -- the render job queue: bounded concurrency, a bounded wait list, one run per key
//
// The plan's production design puts renders on a shared queue (BullMQ or pg-boss) so that any instance can draw
// any deck. JobQueue is that seam; LocalJobQueue runs the jobs in this process.

/** Thrown when the wait list is full; the caller answers 503. */
export class QueueFullError extends Error {
  constructor() {
    super('the render queue is full');
    this.name = 'QueueFullError';
  }
}

export interface JobQueue {
  /**
   * Run `job` when a slot is free. Callers asking for a key that is already queued or running share that run
   * and its result (a slide requested twice is drawn once). Rejects with QueueFullError when the wait list is full.
   */
  run<T>(key: string, job: () => Promise<T>): Promise<T>;
  stats(): { running: number; waiting: number };
}

export class LocalJobQueue implements JobQueue {
  private readonly concurrency: number;
  private readonly maxWaiting: number;
  private running = 0;
  private readonly waiting: (() => void)[] = [];
  private readonly inflight = new Map<string, Promise<unknown>>();

  constructor(concurrency: number, maxWaiting: number) {
    this.concurrency = Math.max(1, concurrency);
    this.maxWaiting = maxWaiting;
  }

  run<T>(key: string, job: () => Promise<T>): Promise<T> {
    const shared = this.inflight.get(key);
    if (shared) return shared as Promise<T>;
    if (this.running >= this.concurrency && this.waiting.length >= this.maxWaiting) return Promise.reject(new QueueFullError());
    const result = this.slot().then(async () => {
      try {
        return await job();
      } finally {
        this.running--;
        this.waiting.shift()?.();
      }
    });
    this.inflight.set(key, result);
    void result.then(() => this.inflight.delete(key), () => this.inflight.delete(key));
    return result;
  }

  /** Resolves once this caller may run; the slot is taken on its behalf. */
  private slot(): Promise<void> {
    if (this.running < this.concurrency) {
      this.running++;
      return Promise.resolve();
    }
    return new Promise(resolve => this.waiting.push(() => {
      this.running++;
      resolve();
    }));
  }

  stats(): { running: number; waiting: number } {
    return { running: this.running, waiting: this.waiting.length };
  }
}
