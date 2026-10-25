// tools/render/contact.ts  -- contact sheets of rendered slides, for looking at many renders at once
//
//   node tools/render/contact.ts <render dir> <out prefix> [--per 40] [--match substring]
//
// Tiles every slideN.png under the render directory (as written by render.ts) into labelled sheets.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { chromium } from 'playwright';

const [dir, prefix, ...rest] = process.argv.slice(2);
if (!dir || !prefix) throw new Error('usage: contact.ts <render dir> <out prefix> [--per N] [--match text]');
const per = Number(rest[rest.indexOf('--per') + 1] ?? 40) || 40;
const matchAt = rest.indexOf('--match');
const match = matchAt >= 0 ? rest[matchAt + 1] ?? '' : '';

const items: { label: string; file: string }[] = [];
for (const deck of readdirSync(dir).sort()) {
  const d = join(dir, deck);
  if (!statSync(d).isDirectory() || (match && !deck.includes(match))) continue;
  const slides = readdirSync(d).filter(f => /^slide\d+\.png$/.test(f)).sort((a, b) => Number(a.slice(5, -4)) - Number(b.slice(5, -4)));
  for (const s of slides) items.push({ label: `${deck.replace(/^.*~/, '')} #${s.slice(5, -4)}`, file: join(d, s) });
}
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
for (let sheet = 0; sheet * per < items.length; sheet++) {
  const chunk = items.slice(sheet * per, (sheet + 1) * per);
  const cells = chunk.map(it => `<figure><img src="data:image/png;base64,${readFileSync(it.file).toString('base64')}"><figcaption>${it.label.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</figcaption></figure>`).join('');
  await page.setContent(`<style>body{margin:8px;font:11px sans-serif;display:grid;grid-template-columns:repeat(8,1fr);gap:6px;background:#ddd;align-items:start}
    figure{margin:0;background:#fff;padding:3px}img{width:100%;display:block;border:1px solid #999}figcaption{overflow:hidden;white-space:nowrap;text-overflow:ellipsis}</style>${cells}`);
  await page.screenshot({ path: `${prefix}-${sheet + 1}.png`, fullPage: true });
}
await browser.close();
console.log(`${items.length} slides on ${Math.ceil(items.length / per)} sheets: ${prefix}-*.png`);
