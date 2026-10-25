// tools/render/metafiles.ts  -- draw EMF and WMF pictures to PNG in headless Chromium, with @pptx/metafile
//
//   node tools/render/metafiles.ts [--out out/metafile/render] [--plus auto|emf|emf+] [--background css|checker]
//                                  [--match text] [--list file-of-paths] [deck, folder or .emf/.wmf file]...
//
// With no inputs, every metafile part of every corpus deck. Writes <out>/<NNN~deck-part>/slide1.png (the layout
// contact.ts tiles) and a summary.json with each picture's record counts, errors and timing.
import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { readZip } from '@pptx/opc';
import { chromium } from 'playwright';
import { createServer, type ViteDevServer } from 'vite';
import { corpusFiles } from '../corpus/list.ts';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = join(HERE, '../..');

interface Result { png?: string; width?: number; height?: number; error?: string; info?: string; ms: number; }
interface Picture { name: string; source: string; bytes: Uint8Array; }

const DECK = new Set(['.pptx', '.pptm', '.potx', '.potm', '.ppsx', '.ppsm']);
const USAGE = 'usage: metafiles.ts [--out dir] [--plus auto|emf|emf+] [--background css|checker] [--match text] [--list file] [input]...';

const isMetafile = (b: Uint8Array) =>
  (b.length >= 44 && b[0] === 1 && b[1] === 0 && b[2] === 0 && b[3] === 0 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46)
  || (b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a)
  || ((b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0);

/** Metafile parts of a deck: media named .emf/.wmf, and anything else that is one by content. */
function fromDeck(file: string): Picture[] {
  let entries;
  try {
    entries = readZip(new Uint8Array(readFileSync(file)));
  } catch {
    return [];
  }
  const deck = basename(file).replace(/\.[^.]+$/, '').replace(/^.*~/, '');
  return entries
    .filter(e => /\.(emf|wmf)$/i.test(e.name) || (/^ppt\/(media|embeddings)\//.test(e.name) && isMetafile(e.data)))
    .map(e => ({ name: `${deck}-${basename(e.name)}`, source: `${file}!${e.name}`, bytes: e.data }));
}

function collect(paths: string[]): Picture[] {
  const out: Picture[] = [];
  for (const p of paths) {
    if (statSync(p).isDirectory()) { out.push(...collect(readdirSync(p).sort().map(n => join(p, n)))); continue; }
    const ext = extname(p).toLowerCase();
    if (DECK.has(ext)) out.push(...fromDeck(p));
    else if (ext === '.emf' || ext === '.wmf') out.push({ name: basename(p), source: p, bytes: new Uint8Array(readFileSync(p)) });
  }
  return out;
}

async function main(): Promise<number> {
  const args = process.argv.slice(2);
  const opt = (name: string, dflt: string) => {
    const i = args.indexOf(name);
    if (i < 0) return dflt;
    const v = args[i + 1] ?? dflt;
    args.splice(i, 2);
    return v;
  };
  const outDir = resolve(opt('--out', join(ROOT, 'out/metafile/render')));
  const plus = opt('--plus', 'auto');
  const background = opt('--background', '#fff');
  const match = opt('--match', '');
  const list = opt('--list', '');
  if (!['auto', 'emf', 'emf+'].includes(plus) || args.some(a => a.startsWith('--'))) {
    console.error(USAGE);
    return 2;
  }
  if (list) args.push(...readFileSync(list, 'utf8').split('\n').map(l => l.trim()).filter(Boolean));
  const pictures = (args.length ? collect(args) : corpusFiles.flatMap(fromDeck)).filter(p => !match || p.name.includes(match) || p.source.includes(match));

  const server: ViteDevServer = await createServer({
    root: join(HERE, 'metafile-harness'),
    publicDir: join(ROOT, 'assets'),
    logLevel: 'error',
    server: { port: 0, strictPort: false, fs: { allow: [ROOT] } },
  });
  await server.listen();
  const url = server.resolvedUrls?.local[0];
  if (!url) throw new Error('vite did not report a URL');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const pageErrors: string[] = [];
  page.on('pageerror', e => pageErrors.push(String(e)));
  const summary: { name: string; source: string; info?: string; error?: string; width?: number; height?: number; ms: number }[] = [];
  try {
    await page.goto(url);
    await page.waitForFunction(() => (globalThis as unknown as { harnessReady?: boolean }).harnessReady === true, null, { timeout: 60_000 })
      .catch(() => { throw new Error(`harness did not start: ${pageErrors.join('\n')}`); });
    for (const [i, pic] of pictures.entries()) {
      const r: Result = await page.evaluate(
        ([b64, o]) => (globalThis as unknown as { renderMetafile: (d: string, o: object) => Promise<Result> }).renderMetafile(b64, o),
        [Buffer.from(pic.bytes).toString('base64'), { plus, background }] as const,
      );
      const dir = join(outDir, `${String(i).padStart(3, '0')}~${pic.name}`);
      mkdirSync(dir, { recursive: true });
      if (r.png) writeFileSync(join(dir, 'slide1.png'), Buffer.from(r.png, 'base64'));
      summary.push({ name: pic.name, source: pic.source, info: r.info, error: r.error?.slice(0, 2000), width: r.width, height: r.height, ms: Math.round(r.ms) });
      console.log(`${r.error ? 'ERR ' : 'ok  '}  ${pic.name}: ${r.width ?? '-'}x${r.height ?? '-'} ${Math.round(r.ms)} ms  ${r.info ?? ''}${r.error ? `  ${r.error.split('\n')[0]}` : ''}`);
    }
  } finally {
    await browser.close();
    await server.close();
  }
  mkdirSync(outDir, { recursive: true });
  writeFileSync(join(outDir, 'summary.json'), JSON.stringify(summary, null, 2));
  const failed = summary.filter(s => s.error).length;
  console.log(`${summary.length} pictures, ${failed} not drawn -> ${outDir}`);
  return failed ? 1 : 0;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) process.exitCode = await main();
