// tools/start/start.ts  -- npm start: install what is missing, build what is stale, then run the server in production mode
//
//   npm start                     the editor at http://127.0.0.1:3000/, the API under /api
//
// Only Node is needed (24.11 or later; .node-version names the one used here). The script fetches pnpm through
// corepack (or npx, where Node ships without corepack), installs the dependencies from the lockfile, downloads
// Playwright's headless Chromium, builds the editor and the server's render page at the same time, and then runs
// apps/server/src/main.ts in this process: NODE_ENV=production, the settings of .env when there is one (see
// .env.example), and Node's compile cache on, so the server's modules are compiled once and reused.
//
// A step runs only when it has to. Its stamp under node_modules/.cache/rock-start holds a hash of everything it
// depends on (the lockfile and manifests; the sources, fonts and configs of the builds), so a start after the first
// takes about as long as the server itself. A missing output (node_modules, a build, the browser) brings the step
// back, and so does deleting node_modules.
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createRequire, enableCompileCache } from 'node:module';
import { homedir } from 'node:os';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

export const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const CACHE = join(ROOT, 'node_modules/.cache/rock-start');
const STAMP_FORMAT = '1';   // part of every hash: bump it when what a stamp covers changes

interface Manifest { engines?: { node?: string }; packageManager?: string }
const manifest = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as Manifest;

// ------------------------------------------------------------------------------------------------ stamps

/** Whether `version` (x.y.z) satisfies `range`, which is ">=x.y.z" (the form package.json's engines use). */
export function satisfies(version: string, range: string): boolean {
  const m = /^>=\s*(\d+)\.(\d+)\.(\d+)$/.exec(range.trim());
  if (!m) throw new Error(`unsupported version range: ${range}`);
  const want = [Number(m[1]), Number(m[2]), Number(m[3])];
  const have = version.split('.').map(Number);
  for (let i = 0; i < 3; i++) {
    const h = have[i] ?? 0, w = want[i]!;
    if (h !== w) return h > w;
  }
  return true;
}

/** Every file under `dir` (absolute paths), dotfiles left out; nothing when there is no such directory. */
export function filesUnder(dir: string): string[] {
  if (!existsSync(dir)) return [];
  return readdirSync(dir, { recursive: true, withFileTypes: true })
    .filter(e => e.isFile() && !e.name.startsWith('.'))
    .map(e => join(e.parentPath, e.name));
}

/** A hash of these files' paths (relative to `root`) and contents, in a fixed order; a file that is not there counts as empty. */
export function hashFiles(paths: string[], root = ROOT): string {
  const h = createHash('sha256').update(STAMP_FORMAT);
  for (const p of [...new Set(paths)].sort()) {
    h.update(relative(root, p).split(sep).join('/')).update('\0');
    if (existsSync(p)) h.update(readFileSync(p));
    h.update('\0');
  }
  return h.digest('hex');
}

const stampPath = (name: string) => join(CACHE, `${name}.sha256`);
const readStamp = (name: string) => (existsSync(stampPath(name)) ? readFileSync(stampPath(name), 'utf8') : undefined);
function writeStamp(name: string, hash: string): void {
  mkdirSync(CACHE, { recursive: true });
  writeFileSync(stampPath(name), hash);
}

// ------------------------------------------------------------------------------------------------ running things

const started = performance.now();
const seconds = (since: number) => `${((performance.now() - since) / 1000).toFixed(1)} s`;
const ok = (what: string) => console.log(`  ✓ ${what}`);
const busy = (what: string) => console.log(`  … ${what}`);

/** A step that could not be done; its message says it all, so no stack trace is shown. */
class StepError extends Error {}

/** Run a command to completion with its output on this terminal; rejects when it fails (code ENOENT when it is not there). */
function run(command: string, args: string[], options: { cwd?: string; env?: NodeJS.ProcessEnv } = {}): Promise<void> {
  return new Promise((done, reject) => {
    const child = spawn(command, args, {
      cwd: options.cwd ?? ROOT, env: { ...process.env, ...options.env }, stdio: 'inherit', shell: process.platform === 'win32',
    });
    child.on('error', reject);
    child.on('exit', (code, signal) => (code === 0 ? done() : reject(new StepError(`${command} ${args.join(' ')} failed (${signal ?? `exit code ${code}`})`))));
  });
}

/** pnpm at the version package.json names: through corepack, or through npx where Node ships without corepack. */
async function pnpm(args: string[]): Promise<void> {
  const env = { COREPACK_ENABLE_DOWNLOAD_PROMPT: '0' };
  try {
    await run('corepack', ['pnpm', ...args], { env });
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== 'ENOENT') throw e;
    const version = manifest.packageManager?.replace(/^pnpm@/, '').split('+')[0] || 'latest';
    await run('npx', ['--yes', `pnpm@${version}`, ...args], { env });
  }
}

/** A step with a stamp: skipped when its inputs hash as they did when it last ran and its outputs are all there. */
async function step(label: string, name: string, inputs: string[], outputs: string[], work: () => Promise<unknown>): Promise<void> {
  const hash = hashFiles(inputs);
  const missing = outputs.filter(o => !existsSync(join(ROOT, o)));
  if (!missing.length && readStamp(name) === hash) return ok(`${label}: up to date`);
  busy(label);
  const since = performance.now();
  await work();
  const left = outputs.filter(o => !existsSync(join(ROOT, o)));
  if (left.length) throw new StepError(`${label}: not produced: ${left.join(', ')}`);
  writeStamp(name, hash);
  ok(`${label} (${seconds(since)})`);
}

/** Vite's command-line entry, as the app at `dir` resolves it. */
const viteBin = (dir: string) => join(dirname(createRequire(join(ROOT, dir, 'package.json')).resolve('vite/package.json')), 'bin/vite.js');

const subdirs = (dir: string) => readdirSync(join(ROOT, dir)).filter(n => !n.startsWith('.'));

// ------------------------------------------------------------------------------------------------ the steps

const INSTALL_OUTPUTS = [
  'node_modules/.modules.yaml', 'apps/web/node_modules/vite/package.json', 'apps/server/node_modules/vite/package.json',
  'apps/server/node_modules/playwright/package.json',
];
const installInputs = () => [
  ...['package.json', 'pnpm-lock.yaml', 'pnpm-workspace.yaml', '.npmrc'].map(p => join(ROOT, p)),
  ...['apps', 'packages', 'tools'].flatMap(d => subdirs(d).map(n => join(ROOT, d, n, 'package.json'))),
];

const BUILD_OUTPUTS = ['apps/web/dist/index.html', 'apps/server/dist/render-page/index.html'];
const buildInputs = () => [
  ...['pnpm-lock.yaml', 'tsconfig.base.json', 'apps/web/index.html', 'apps/web/vite.config.ts', 'apps/web/precompress.ts', 'apps/web/tsconfig.json']
    .map(p => join(ROOT, p)),
  ...['apps/web/src', 'assets', 'apps/server/render-page'].flatMap(d => filesUnder(join(ROOT, d))),
  ...subdirs('packages').flatMap(p => filesUnder(join(ROOT, 'packages', p, 'src'))),
];

interface Shell { dir: string; version: string; cli: string }

/** Where Playwright keeps its browsers, by the rules it applies itself: PLAYWRIGHT_BROWSERS_PATH, or the platform's cache directory. */
export function playwrightBrowsersDir(coreDir: string, env: NodeJS.ProcessEnv = process.env, platform = process.platform, home = homedir()): string {
  const set = env.PLAYWRIGHT_BROWSERS_PATH?.trim();
  if (set === '0') return join(coreDir, '.local-browsers');
  if (set) return resolve(env.INIT_CWD || process.cwd(), set);
  const cache = platform === 'linux' ? env.XDG_CACHE_HOME || join(home, '.cache')
    : platform === 'darwin' ? join(home, 'Library', 'Caches')
    : platform === 'win32' ? env.LOCALAPPDATA || join(home, 'AppData', 'Local')
    : undefined;
  if (!cache) throw new Error(`Playwright has no browsers for ${platform}`);
  return join(cache, 'ms-playwright');
}

/** The headless Chromium the server draws with: where Playwright (as apps/server resolves it) installs it, and the CLI that does. */
export function chromiumShell(env: NodeJS.ProcessEnv = process.env): Shell {
  const server = createRequire(join(ROOT, 'apps/server/package.json'));
  const playwright = dirname(server.resolve('playwright/package.json'));
  const core = dirname(createRequire(join(playwright, 'package.json')).resolve('playwright-core/package.json'));
  const { browsers } = JSON.parse(readFileSync(join(core, 'browsers.json'), 'utf8')) as { browsers: { name: string; revision: string; browserVersion?: string }[] };
  const shell = browsers.find(b => b.name === 'chromium-headless-shell');
  if (!shell) throw new Error(`${core}/browsers.json names no chromium-headless-shell`);
  // Playwright's browser directories are the name with its dashes made underscores, then the revision
  return { dir: join(playwrightBrowsersDir(core, env), `chromium_headless_shell-${shell.revision}`), version: shell.browserVersion ?? shell.revision, cli: join(playwright, 'cli.js') };
}

const INSTALLED = 'INSTALLATION_COMPLETE';   // Playwright's own marker of a finished download

async function installChromium(): Promise<void> {
  const shell = chromiumShell();
  if (existsSync(join(shell.dir, INSTALLED))) return ok(`Chromium headless shell ${shell.version}: installed`);
  busy(`Chromium headless shell ${shell.version}`);
  const since = performance.now();
  // --with-deps adds the system libraries Chromium needs on Linux, which takes root (the container image does this itself)
  const withDeps = process.platform === 'linux' && process.getuid?.() === 0;
  await run(process.execPath, [shell.cli, 'install', ...(withDeps ? ['--with-deps'] : []), '--only-shell', 'chromium']);
  if (!existsSync(join(shell.dir, INSTALLED))) throw new StepError(`Chromium headless shell: nothing at ${shell.dir}`);
  ok(`Chromium headless shell ${shell.version} (${seconds(since)})`);
  if (process.platform === 'linux' && !withDeps) console.log('    (if Chromium does not start, it lacks system libraries: sudo npx playwright install-deps chromium)');
}

export async function main(): Promise<void> {
  process.chdir(ROOT);
  console.log('rock_presentation');
  const wanted = manifest.engines?.node ?? '>=0.0.0';
  if (!satisfies(process.versions.node, wanted)) {
    throw new StepError(`Node ${wanted} is needed (.node-version names ${readFileSync(join(ROOT, '.node-version'), 'utf8').trim()}); this is v${process.versions.node}`);
  }
  ok(`Node v${process.versions.node}`);

  await step('dependencies', 'install', installInputs(), INSTALL_OUTPUTS, () => pnpm(['install', '--frozen-lockfile', '--config.confirmModulesPurge=false']));
  await installChromium();
  await step('editor and render page', 'build', buildInputs(), BUILD_OUTPUTS, () => Promise.all([
    run(process.execPath, [viteBin('apps/web'), 'build', '--logLevel', 'warn'], { cwd: join(ROOT, 'apps/web') }),
    run(process.execPath, [viteBin('apps/server'), 'build', '--config', 'render-page/vite.config.ts'], { cwd: join(ROOT, 'apps/server') }),
  ]));

  process.env.NODE_ENV ??= 'production';
  const dotenv = join(ROOT, '.env');
  if (existsSync(dotenv)) process.loadEnvFile(dotenv);     // what is already in the environment wins
  enableCompileCache(join(CACHE, 'compile-cache'));
  const host = process.env.HOST?.trim() || '127.0.0.1';
  const port = process.env.PORT?.trim() || '3000';
  const data = resolve(process.env.DATA_DIR?.trim() || 'data');
  console.log(`  ▶ http://${host}:${port}/  (the editor; the API under /api; decks in ${data}; NODE_ENV=${process.env.NODE_ENV}; ${seconds(started)} to here)`);
  await import(pathToFileURL(join(ROOT, 'apps/server/src/main.ts')).href);
}

if (import.meta.main) {
  try {
    await main();
  } catch (e) {
    console.error(e instanceof StepError ? `  ✗ ${e.message}` : e);
    process.exit(1);
  }
}
