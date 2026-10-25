// tools/geometry/gen-presets.ts  -- generate packages/geometry/src/presets.generated.ts and textwarps.generated.ts
//
//   node tools/geometry/gen-presets.ts [dir]      (default: corpus/.cache/ecma-376/geometries)
//
// Reads presetShapeDefinitions.xml and presetTextWarpDefinitions.xml, the ECMA-376 Part 1 Annex D files from
// OfficeOpenXML-DrawingMLGeometries.zip, with our own XML parser and the same reader a:custGeom goes through, and
// writes them out as plain strings in the compact form of packages/geometry/src/compact.ts. The files are untrusted
// data: they are parsed and nothing in them is executed; names must be plain identifiers and every other value
// passes the compact encoder's checks, and each geometry must decode back to exactly what was read.
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { isDeepStrictEqual } from 'node:util';
import { childElements, decodeXmlBytes, parseXml } from '@pptx/xml';
// Relative, not '@pptx/geometry': the package's index imports the files this script writes.
import { decodeGeometry, encodeGeometry } from '../../packages/geometry/src/compact.ts';
import { parseCustomGeometry } from '../../packages/geometry/src/parse.ts';
import type { ShapeGeometry } from '../../packages/geometry/src/types.ts';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../..');
export const DEFAULT_DIR = join(ROOT, 'corpus/.cache/ecma-376/geometries');
export const OUTPUTS = {
  presets: join(ROOT, 'packages/geometry/src/presets.generated.ts'),
  textWarps: join(ROOT, 'packages/geometry/src/textwarps.generated.ts'),
};

/** Every shape element of one Annex D file, by name. A name given twice must have the same definition both times. */
export function readDefinitions(file: string): Map<string, ShapeGeometry> {
  const doc = parseXml(decodeXmlBytes(new Uint8Array(readFileSync(file))).text);
  const out = new Map<string, ShapeGeometry>();
  for (const el of childElements(doc.root)) {
    if (!/^[a-z][A-Za-z0-9]*$/.test(el.local)) throw new Error(`${file}: unexpected element <${el.local}>`);
    const geom = parseCustomGeometry(el);
    const known = out.get(el.local);
    if (known && !isDeepStrictEqual(known, geom)) throw new Error(`${file}: two different definitions of ${el.local}`);
    out.set(el.local, geom);
  }
  return out;
}

// Annex D lacks upArrow: its file has upDownArrow twice instead. This is downArrow upside down, in the clockwise
// order the other arrows use (test/presets.test.ts checks it against downArrow at many sizes and adjust values).
const UP_ARROW = `<upArrow xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
  <avLst><gd name="adj1" fmla="val 50000"/><gd name="adj2" fmla="val 50000"/></avLst>
  <gdLst>
    <gd name="maxAdj2" fmla="*/ 100000 h ss"/>
    <gd name="a1" fmla="pin 0 adj1 100000"/>
    <gd name="a2" fmla="pin 0 adj2 maxAdj2"/>
    <gd name="dy2" fmla="*/ ss a2 100000"/>
    <gd name="y2" fmla="+- t dy2 0"/>
    <gd name="dx1" fmla="*/ w a1 200000"/>
    <gd name="x1" fmla="+- hc 0 dx1"/>
    <gd name="x2" fmla="+- hc dx1 0"/>
    <gd name="dy1" fmla="*/ x1 dy2 wd2"/>
    <gd name="y1" fmla="+- y2 0 dy1"/>
  </gdLst>
  <ahLst>
    <ahXY gdRefX="adj1" minX="0" maxX="100000"><pos x="x1" y="b"/></ahXY>
    <ahXY gdRefY="adj2" minY="0" maxY="maxAdj2"><pos x="l" y="y2"/></ahXY>
  </ahLst>
  <cxnLst>
    <cxn ang="3cd4"><pos x="hc" y="t"/></cxn>
    <cxn ang="cd2"><pos x="l" y="y2"/></cxn>
    <cxn ang="cd4"><pos x="hc" y="b"/></cxn>
    <cxn ang="0"><pos x="r" y="y2"/></cxn>
  </cxnLst>
  <rect l="x1" t="y1" r="x2" b="b"/>
  <pathLst>
    <path>
      <moveTo><pt x="l" y="y2"/></moveTo>
      <lnTo><pt x="hc" y="t"/></lnTo>
      <lnTo><pt x="r" y="y2"/></lnTo>
      <lnTo><pt x="x2" y="y2"/></lnTo>
      <lnTo><pt x="x2" y="b"/></lnTo>
      <lnTo><pt x="x1" y="b"/></lnTo>
      <lnTo><pt x="x1" y="y2"/></lnTo>
      <close/>
    </path>
  </pathLst>
</upArrow>`;

const PRESET_NOTES = [
  'Corrections to that file:',
  '- upArrow is missing; the file has upDownArrow twice instead. upArrow is added as downArrow upside down.',
  '- curvedLeftArrow gives its connection site on the left edge the angle "cd3", which is no guide. It is cd2',
  '  here (pointing left), mirroring the site on the right edge of curvedRightArrow, which points right (0).',
  '- circularArrow, leftCircularArrow and leftRightCircularArrow define xB as "+- xH 0 dxB 0" and yB as',
  '  "+- yH 0 dyB 0", one argument more than "+-" takes. The stray trailing 0 is dropped: xB = xH - dxB.',
  '- pie has <rect l="il" t="ir" r="it" b="ib">, which puts its text in a sliver outside the shape; t and r are',
  '  swapped back: l="il" t="it" r="ir" b="ib", the rectangle inscribed in the ellipse as in chord and others.',
  '- squareTabs has a connection site at (dx, x1), below the shape unless it is square; it is (dx, y1), the',
  '  inner corner of the bottom-left tab, matching the sites at (dx, dx), (x1, dx) and (x1, y1).',
];

/** The preset shapes, with the corrections listed in PRESET_NOTES. Each applies only where the error is found. */
export function presetShapes(dir = DEFAULT_DIR): Map<string, ShapeGeometry> {
  const shapes = readDefinitions(join(dir, 'presetShapeDefinitions.xml'));
  const edit = (name: string, change: (g: ShapeGeometry) => ShapeGeometry) => {
    const g = shapes.get(name);
    if (g) shapes.set(name, change(g));
  };
  if (!shapes.has('upArrow')) shapes.set('upArrow', parseCustomGeometry(parseXml(UP_ARROW).root));
  edit('curvedLeftArrow', g => ({ ...g, cxnLst: g.cxnLst.map(c => (c.ang === 'cd3' ? { ...c, ang: 'cd2' } : c)) }));
  for (const name of ['circularArrow', 'leftCircularArrow', 'leftRightCircularArrow']) {
    edit(name, g => ({ ...g, gdLst: g.gdLst.map(gd => ({ ...gd, fmla: gd.fmla.replace(/^(\+- \S+ \S+ \S+) 0$/, '$1') })) }));
  }
  edit('pie', g => (g.rect?.t === 'ir' && g.rect.r === 'it' ? { ...g, rect: { ...g.rect, t: 'it', r: 'ir' } } : g));
  edit('squareTabs', g => ({
    ...g,
    cxnLst: g.cxnLst.map(c => (c.pos.x === 'dx' && c.pos.y === 'x1' ? { ...c, pos: { x: 'dx', y: 'y1' } } : c)),
  }));
  return shapes;
}

export const textWarps = (dir = DEFAULT_DIR) => readDefinitions(join(dir, 'presetTextWarpDefinitions.xml'));

function tsModule(file: string, constName: string, description: string[], defs: Map<string, ShapeGeometry>): string {
  const lines = [
    `// ${file}  -- GENERATED by tools/geometry/gen-presets.ts. Do not edit.`,
    '//',
    ...description.map(l => `// ${l}`),
    '// Regenerate with `node tools/geometry/gen-presets.ts [directory holding the Annex D files]`.',
    "import type { CompactGeometry } from './compact.ts';",
    '',
    `export const ${constName}: Readonly<Record<string, CompactGeometry>> = {`,
  ];
  for (const name of [...defs.keys()].sort()) {
    const geom = defs.get(name)!;
    const compact = encodeGeometry(geom);
    if (!isDeepStrictEqual(decodeGeometry(compact), geom)) throw new Error(`${name} does not survive the compact form`);
    lines.push(`  ${name}: {`);
    for (const key of ['av', 'gd', 'ah', 'cxn', 'rect'] as const) {
      const value = compact[key];
      if (value !== undefined) lines.push(`    ${key}: ${JSON.stringify(value)},`);
    }
    if (compact.path.length === 1) lines.push(`    path: [${JSON.stringify(compact.path[0])}],`);
    else lines.push('    path: [', ...compact.path.map(p => `      ${JSON.stringify(p)},`), '    ],');
    lines.push('  },');
  }
  lines.push('};', '');
  return lines.join('\n');
}

/** The two generated modules, as text, and the definitions they were made from. */
export function generate(dir = DEFAULT_DIR) {
  const shapes = presetShapes(dir);
  const warps = textWarps(dir);
  return {
    shapes,
    warps,
    presets: tsModule('packages/geometry/src/presets.generated.ts', 'PRESETS', [
      `The ${shapes.size} preset shape geometries (ST_ShapeType) of ECMA-376 Part 1, 5th edition, Annex D:`,
      'presetShapeDefinitions.xml from OfficeOpenXML-DrawingMLGeometries.zip, in the compact form of compact.ts.',
      ...PRESET_NOTES,
    ], shapes),
    textWarps: tsModule('packages/geometry/src/textwarps.generated.ts', 'TEXT_WARPS', [
      `The ${warps.size} preset text warps (ST_TextShapeType, all but textNoShape) of ECMA-376 Part 1, 5th edition,`,
      'Annex D: presetTextWarpDefinitions.xml from OfficeOpenXML-DrawingMLGeometries.zip, in the compact form of',
      'compact.ts.',
    ], warps),
  };
}

function main(argv: string[]): void {
  const dir = resolve(argv[0] ?? DEFAULT_DIR);
  const out = generate(dir);
  writeFileSync(OUTPUTS.presets, out.presets);
  writeFileSync(OUTPUTS.textWarps, out.textWarps);
  console.log(`${out.shapes.size} preset shapes -> ${OUTPUTS.presets} (${out.presets.length} bytes)`);
  console.log(`${out.warps.size} text warps -> ${OUTPUTS.textWarps} (${out.textWarps.length} bytes)`);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main(process.argv.slice(2));
