// tools/geometry/gallery.ts  -- `node tools/geometry/gallery.ts`: every preset shape and text warp, drawn for a look
//
// Writes out/geometry-gallery.html (one labelled SVG per geometry, evaluated at its default adjust values) and
// renders it with Playwright's Chromium to out/geometry-gallery.png, plus one PNG per section
// (out/geometry-gallery-<n>.png) small enough to inspect. Each cell shows the paths (fill by path fill mode, the
// outline where the path is stroked), the text rectangle (dashed red), the adjust handles (yellow diamonds) and the
// connection sites (blue crosses).
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import {
  evaluateGeometry, PRESET_NAMES, presetGeometry, TEXT_WARP_NAMES, textWarpGeometry,
  type GeometryPath, type PathCommand, type ShapeGeometry,
} from '@pptx/geometry';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../..');
const OUT = join(ROOT, 'out');
const W = 1600000, H = 1100000;                 // EMU; not square, so ss-based and w/h-based guides differ
const PAD = 300000;                             // room for handles; callout tails may overflow into the next cell
const FILLS: Record<GeometryPath['fill'], string> = {
  norm: '#9dc3e6', lighten: '#cfe2f3', lightenLess: '#b6d4ee', darken: '#5b8fc0', darkenLess: '#7aa9d6', none: 'none',
};

const n = (v: number) => Math.round(v).toString();
function d(commands: PathCommand[]): string {
  return commands.map(c => {
    switch (c.op) {
      case 'M': case 'L': return `${c.op}${n(c.x)} ${n(c.y)}`;
      case 'Q': return `Q${n(c.x1)} ${n(c.y1)} ${n(c.x)} ${n(c.y)}`;
      case 'C': return `C${n(c.x1)} ${n(c.y1)} ${n(c.x2)} ${n(c.y2)} ${n(c.x)} ${n(c.y)}`;
      case 'Z': return 'Z';
    }
  }).join('');
}

function cell(name: string, geom: ShapeGeometry, fill: boolean): string {
  let svg: string;
  try {
    const e = evaluateGeometry(geom, W, H);
    const parts = [`<rect x="0" y="0" width="${W}" height="${H}" fill="none" stroke="#ddd" stroke-width="1" vector-effect="non-scaling-stroke"/>`];
    // Fills first, then outlines: chartPlus, chartStar and chartX list their lines before the filled square, and
    // they show in PowerPoint, so a renderer cannot simply paint each path in turn.
    if (fill) for (const p of e.paths) if (p.fill !== 'none') parts.push(`<path d="${d(p.commands)}" fill="${FILLS[p.fill]}" fill-rule="evenodd"/>`);
    for (const p of e.paths) {
      if (p.stroke) parts.push(`<path d="${d(p.commands)}" fill="none" stroke="#1f3864" stroke-width="1.5" vector-effect="non-scaling-stroke"/>`);
    }
    const t = e.textRect;
    parts.push(`<rect x="${n(t.x)}" y="${n(t.y)}" width="${n(t.w)}" height="${n(t.h)}" fill="none" stroke="#e03030" stroke-dasharray="3 2" stroke-width="1" vector-effect="non-scaling-stroke"/>`);
    const r = 45000;
    for (const s of e.connectionSites) {
      parts.push(`<path d="M${n(s.x - r)} ${n(s.y - r)}L${n(s.x + r)} ${n(s.y + r)}M${n(s.x - r)} ${n(s.y + r)}L${n(s.x + r)} ${n(s.y - r)}" stroke="#2060ff" stroke-width="1.2" vector-effect="non-scaling-stroke"/>`);
    }
    const q = 60000;
    for (const h of e.handles) {
      parts.push(`<path d="M${n(h.x)} ${n(h.y - q)}L${n(h.x + q)} ${n(h.y)}L${n(h.x)} ${n(h.y + q)}L${n(h.x - q)} ${n(h.y)}Z" fill="#ffd400" stroke="#8a6d00" stroke-width="1" vector-effect="non-scaling-stroke"/>`);
    }
    svg = `<svg viewBox="${-PAD} ${-PAD} ${W + 2 * PAD} ${H + 2 * PAD}" width="200" height="160" overflow="visible">${parts.join('')}</svg>`;
  } catch (err) {
    svg = `<div class="error">${(err as Error).message}</div>`;
  }
  return `<figure>${svg}<figcaption>${name}</figcaption></figure>`;
}

const PAGE = 48;
const sections: { title: string; cells: string[] }[] = [];
for (let i = 0; i < PRESET_NAMES.length; i += PAGE) {
  const names = PRESET_NAMES.slice(i, i + PAGE);
  sections.push({ title: `Preset shapes ${i + 1}-${i + names.length}`, cells: names.map(nm => cell(nm, presetGeometry(nm)!, true)) });
}
// A text warp's paths are the curves text is fitted between, never filled.
sections.push({ title: 'Text warps', cells: TEXT_WARP_NAMES.map(nm => cell(nm, textWarpGeometry(nm)!, false)) });

const html = `<!doctype html>
<html><head><meta charset="utf-8"><title>Geometry gallery</title><style>
  body { font: 12px system-ui, sans-serif; margin: 16px; background: #fff; color: #222; }
  section { display: grid; grid-template-columns: repeat(8, 210px); gap: 4px; margin-bottom: 24px; width: max-content; }
  h2 { grid-column: 1 / -1; font-size: 15px; margin: 0 0 4px; }
  figure { margin: 0; text-align: center; }
  figcaption { font-family: ui-monospace, monospace; font-size: 11px; }
  .error { width: 200px; height: 160px; color: #c00; overflow: hidden; }
</style></head><body>
<p>${PRESET_NAMES.length} preset shapes and ${TEXT_WARP_NAMES.length} text warps at ${W} x ${H} EMU, default adjust values.
Dashed red: text rectangle. Yellow: adjust handles. Blue: connection sites. Fills (evenodd) are drawn before outlines.</p>
${sections.map(s => `<section><h2>${s.title}</h2>${s.cells.join('\n')}</section>`).join('\n')}
</body></html>
`;

mkdirSync(OUT, { recursive: true });
writeFileSync(join(OUT, 'geometry-gallery.html'), html);
const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 1720, height: 1000 } });
  await page.setContent(html);
  await page.screenshot({ path: join(OUT, 'geometry-gallery.png'), fullPage: true });
  const blocks = page.locator('section');
  for (let i = 0; i < (await blocks.count()); i++) await blocks.nth(i).screenshot({ path: join(OUT, `geometry-gallery-${i + 1}.png`) });
} finally {
  await browser.close();
}
console.log(`wrote ${join(OUT, 'geometry-gallery.html')} and ${sections.length + 1} PNGs`);
