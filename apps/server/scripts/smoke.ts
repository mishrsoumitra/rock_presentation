// apps/server/scripts/smoke.ts  -- check a running server end to end: the editor, the API, a slide image and a PDF
//
//   node apps/server/scripts/smoke.ts http://127.0.0.1:3000 [api-token]
//
// For the container image, in CI and after a deploy. It makes its own two-slide deck with the editor's code,
// uploads it, draws a slide, exports the PDF and deletes the deck again. Exits 1 at the first thing that is wrong.
import { PNG } from 'pngjs';
import { OpcPackage } from '@pptx/opc';
import { addSlide, Editor, insertTextBox } from '../../../packages/editor-core/src/index.ts';
import { blankPresentation } from '../../web/src/engine/blank.ts';

const base = new URL(process.argv[2] ?? 'http://127.0.0.1:3000');
const token = process.argv[3] ?? process.env.API_TOKEN;
const auth: Record<string, string> = token ? { authorization: `Bearer ${token}` } : {};
const at = (path: string) => new URL(path, base);

function check(ok: unknown, what: string): asserts ok {
  if (!ok) {
    console.error(`FAIL ${what}`);
    process.exit(1);
  }
  console.log(`ok   ${what}`);
}

/** Two slides, the first with a line of text. */
function deck(): Uint8Array {
  const ed = new Editor(OpcPackage.open(blankPresentation()));
  const master = ed.doc.masterParts()[0]!;
  const layout = ed.doc.layoutsOf(master).at(-1)!;
  while (ed.doc.slideParts().length < 2) ed.run(addSlide(ed, layout, ed.doc.slideParts().length).command);
  const slide = ed.doc.slideParts()[0]!;
  ed.run(insertTextBox(ed, slide, { x: 914400, y: 914400, w: 6 * 914400, h: 914400 }, 'Smoke test').command);
  return ed.save();
}

const started = Date.now();
let health: Response | undefined;
while (Date.now() - started < 90_000) {
  health = await fetch(at('/api/health')).catch(() => undefined);
  if (health?.ok) break;
  await new Promise(r => setTimeout(r, 1000));
}
check(health?.ok, `the server answers ${at('/api/health')}`);
const status = await health.json() as { render: { enabled: boolean; built: boolean } };
check(status.render.enabled && status.render.built, 'slide rendering is on, with its page built');

const editor = await fetch(at('/'));
check(editor.ok && (editor.headers.get('content-security-policy') ?? '').includes("frame-ancestors 'none'"), 'the editor is served, with its policy');

const upload = (headers: Record<string, string>) => {
  const form = new FormData();
  form.set('file', new Blob([deck() as BlobPart]), 'smoke.pptx');
  return fetch(at('/api/decks'), { method: 'POST', body: form, headers });
};
if (token) check((await upload({})).status === 401, 'an upload without the token is refused');
const created = await upload(auth);
check(created.status === 201, `the deck uploads (${created.status})`);
const { id, slideCount } = await created.json() as { id: string; slideCount: number };
check(slideCount === 2, `the server counts its two slides (${slideCount})`);

const png = await fetch(at(`/api/decks/${id}/slides/1.png?width=640`), { headers: auth });
const image = png.ok ? PNG.sync.read(Buffer.from(await png.arrayBuffer())) : undefined;
check(image?.width === 640, `slide 1 is drawn as a 640-pixel PNG (${png.status})`);
const colours = new Set<number>();
for (let i = 0; i < image.data.length; i += 4) colours.add(image.data.readUInt32BE(i));
check(colours.size > 2, 'and the text is on it');

const pdf = await fetch(at(`/api/decks/${id}/export.pdf`), { headers: auth });
const bytes = Buffer.from(await pdf.arrayBuffer());
check(pdf.ok && bytes.subarray(0, 5).toString() === '%PDF-', `the PDF exports (${pdf.status})`);
check((bytes.toString('latin1').match(/\/Type\s*\/Page(?![a-zA-Z])/g) ?? []).length === 2, 'with a page per slide');

check((await fetch(at(`/api/decks/${id}`), { method: 'DELETE', headers: auth })).status === 204, 'the deck deletes');
console.log(`smoke test passed in ${Math.round((Date.now() - started) / 1000)} s`);
