// apps/server/src/main.ts  -- start the server from environment variables; stop it cleanly on SIGINT or SIGTERM
//
//   DATA_DIR=/var/lib/pptx PORT=3000 node apps/server/src/main.ts
//
// See configFromEnv in src/config.ts for every variable. Shutting down stops taking requests, lets those in flight
// finish, then closes the worker threads and Chromium; a second signal, or 20 s, forces the exit.
import { buildServer } from './app.ts';
import { configFromEnv } from './config.ts';

const config = configFromEnv(process.env);
const app = await buildServer(config);

let stopping = false;
async function stop(signal: string): Promise<void> {
  if (stopping) process.exit(1);
  stopping = true;
  app.log.info({ signal }, 'shutting down');
  setTimeout(() => {
    app.log.error('shutdown took too long; exiting');
    process.exit(1);
  }, 20_000).unref();
  try {
    await app.close();
    app.log.info('stopped');
    process.exitCode = 0;
  } catch (e) {
    app.log.error({ err: e }, 'shutdown failed');
    process.exitCode = 1;
  }
}
process.on('SIGINT', signal => void stop(signal));
process.on('SIGTERM', signal => void stop(signal));

await app.listen({ host: config.host, port: config.port });
const { dataDir, workers, render, oracles } = config;
app.log.info({ dataDir, workers: workers.size, renderPages: render.pages, oracles: oracles.enabled }, 'ready');
