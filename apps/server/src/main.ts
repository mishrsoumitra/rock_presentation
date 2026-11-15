// apps/server/src/main.ts  -- start the server from environment variables; stop it cleanly on SIGINT or SIGTERM
//
//   DATA_DIR=/var/lib/pptx PORT=3000 node apps/server/src/main.ts
//
// See configFromEnv in src/config.ts for every variable. Shutting down stops taking requests, lets those in flight
// finish, then closes the worker threads and Chromium; a second signal a second or more later, or 20 s, forces the
// exit. (Within a second it is the same Ctrl-C: the terminal sends it to every process in the group, and npm start
// then forwards the one it got to the server as well.)
import { buildServer } from './app.ts';
import { configFromEnv } from './config.ts';

const config = configFromEnv(process.env);
const app = await buildServer(config);

let stopping: number | undefined;   // when the first signal came
async function stop(signal: string): Promise<void> {
  if (stopping !== undefined) {
    if (Date.now() - stopping >= 1000) process.exit(1);
    return;
  }
  stopping = Date.now();
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
const { dataDir, workers, render, oracles, apiTokens, webDir } = config;
app.log.info({
  dataDir, workers: workers.size, renderPages: render.pages, oracles: oracles.enabled, apiTokens: apiTokens.length,
  editor: app.hasRoute({ method: 'GET', url: '/' }) ? webDir : false,
}, 'ready');
