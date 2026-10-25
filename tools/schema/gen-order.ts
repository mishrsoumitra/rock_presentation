// tools/schema/gen-order.ts  -- generate packages/model/src/schema.generated.ts from the ECMA-376 Transitional XSDs.
//
//   node tools/schema/gen-order.ts [schema dir]      (default: corpus/.cache/ecma-376/transitional)
//
// For every complex type in the DrawingML and PresentationML schemas it records the allowed child elements in schema
// order: children of one xsd:choice share a rank, so they replace each other; a child under maxOccurs > 1 is
// repeatable. It also records each child's type, so the model can find the type of any element by walking down from
// the part's root element. Strict documents use the same element names under other URIs; the model maps both
// flavours to the same logical prefixes, so one table serves both.
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { childElements, getAttr, parseXml, XMLNS_NS, type XmlElement } from '@pptx/xml';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../..');
const dir = process.argv[2] ?? join(ROOT, 'corpus/.cache/ecma-376/transitional');
const FILES = ['dml-main.xsd', 'pml.xsd', 'dml-chart.xsd', 'dml-chartDrawing.xsd', 'dml-diagram.xsd', 'dml-picture.xsd'];
const XS = 'http://www.w3.org/2001/XMLSchema';

const KEY: Record<string, string> = {
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/presentationml/2006/main': 'p',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing': 'cdr',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
};

interface Schema { root: XmlElement; key: string; }
const schemas: Schema[] = FILES.map(f => {
  const root = parseXml(readFileSync(join(dir, f), 'utf8')).root;
  const key = KEY[getAttr(root, 'targetNamespace') ?? ''];
  if (!key) throw new Error(`${f}: unknown target namespace`);
  return { root, key };
});

const xs = (el: XmlElement, local: string) => el.ns === XS && el.local === local;
const xsChildren = (el: XmlElement) => childElements(el).filter(c => c.ns === XS);

/** Resolve a QName in an attribute value (type="a:CT_X") through the schema's declarations, to 'a:CT_X'. */
function qref(value: string, at: XmlElement): string | null {
  const colon = value.indexOf(':');
  const prefix = colon < 0 ? '' : value.slice(0, colon);
  const local = colon < 0 ? value : value.slice(colon + 1);
  for (let e: XmlElement | null = at; e; e = e.parent) {
    for (const a of e.attrs) {
      if (a.ns !== XMLNS_NS || (a.prefix === null ? '' : a.local) !== prefix) continue;
      const key = KEY[a.value];
      return key ? `${key}:${local}` : null;      // xsd:, shared simple types and the like: not elements we order
    }
  }
  return null;
}

const types = new Map<string, { el: XmlElement; key: string }>();
const groups = new Map<string, { el: XmlElement; key: string }>();
const globals = new Map<string, string | null>();
for (const s of schemas) {
  for (const c of xsChildren(s.root)) {
    const name = getAttr(c, 'name');
    if (!name) continue;
    if (xs(c, 'complexType')) types.set(`${s.key}:${name}`, { el: c, key: s.key });
    else if (xs(c, 'group')) groups.set(`${s.key}:${name}`, { el: c, key: s.key });
    else if (xs(c, 'element')) globals.set(`${s.key}:${name}`, getAttr(c, 'type') ? qref(getAttr(c, 'type')!, c) : null);
  }
}

interface Slot { name: string; type: string | null; rank: number; repeat: boolean; }

const many = (el: XmlElement) => {
  const max = getAttr(el, 'maxOccurs');
  return max === 'unbounded' || (max !== undefined && Number(max) > 1);
};

/** Append the element slots of one particle to `out`; returns the next free rank. */
function particle(el: XmlElement, key: string, rank: number, repeat: boolean, out: Slot[], seen: Set<string>): number {
  const rep = repeat || many(el);
  if (xs(el, 'element')) {
    const ref = getAttr(el, 'ref');
    const name = ref ? qref(ref, el) : `${key}:${getAttr(el, 'name')}`;
    if (!name) return rank + 1;
    const type = ref ? globals.get(name) ?? null : getAttr(el, 'type') ? qref(getAttr(el, 'type')!, el) : null;
    if (!out.some(s => s.name === name)) out.push({ name, type, rank, repeat: rep });
    return rank + 1;
  }
  if (xs(el, 'sequence')) {
    let r = rank;
    for (const c of xsChildren(el)) r = particle(c, key, r, rep, out, seen);
    return r;
  }
  if (xs(el, 'choice') || xs(el, 'all')) {
    // Alternatives of a choice replace each other, so they all start at one rank; xsd:all is unordered, the same
    // thing here. An alternative may be a whole sequence (c:CT_DLbl: delete, or the shared label settings), so what
    // follows the choice ranks after the longest alternative, not merely after its first rank.
    let end = rank + 1;
    for (const c of xsChildren(el)) end = Math.max(end, particle(c, key, rank, rep || xs(el, 'all'), out, seen));
    return end;
  }
  if (xs(el, 'group')) {
    const ref = getAttr(el, 'ref');
    const name = ref ? qref(ref, el) : null;
    const g = name ? groups.get(name) : undefined;
    if (!g || seen.has(name!)) return rank + 1;
    seen.add(name!);
    let r = rank;
    for (const c of xsChildren(g.el)) r = particle(c, g.key, r, rep, out, seen);
    seen.delete(name!);
    return r;
  }
  return rank;                                    // annotations, attributes, xsd:any
}

function slotsOf(typeName: string, seen = new Set<string>()): Slot[] {
  const t = types.get(typeName);
  if (!t || seen.has(typeName)) return [];
  seen.add(typeName);
  const out: Slot[] = [];
  let rank = 0;
  for (const c of xsChildren(t.el)) {
    if (xs(c, 'complexContent') || xs(c, 'simpleContent')) {
      for (const d of xsChildren(c)) {
        if (!xs(d, 'extension') && !xs(d, 'restriction')) continue;
        const base = getAttr(d, 'base') ? qref(getAttr(d, 'base')!, d) : null;
        if (base && xs(d, 'extension')) {
          for (const s of slotsOf(base, seen)) out.push({ ...s });
          rank = out.reduce((m, s) => Math.max(m, s.rank + 1), 0);
        }
        for (const e of xsChildren(d)) rank = particle(e, t.key, rank, false, out, new Set());
      }
    } else {
      rank = particle(c, t.key, rank, false, out, new Set());
    }
  }
  seen.delete(typeName);
  return out;
}

// Compact encoding, parsed lazily by the model: slots separated by ';', each "name=type@rank" with '*' when repeatable.
const lines: string[] = [];
for (const name of [...types.keys()].sort()) {
  const slots = slotsOf(name);
  if (!slots.length) continue;
  lines.push(`  '${name}': '${slots.map(s => `${s.name}=${s.type ?? ''}@${s.rank}${s.repeat ? '*' : ''}`).join(';')}',`);
}
const globalLines = [...globals].filter(([, t]) => t).sort(([a], [b]) => (a < b ? -1 : 1)).map(([n, t]) => `  '${n}': '${t}',`);

const out = `// GENERATED by tools/schema/gen-order.ts from the ECMA-376 5th edition Transitional schemas
// (${FILES.join(', ')}). Do not edit; run \`node tools/schema/gen-order.ts\` to regenerate.
//
// TYPES: complex type -> its child elements in schema order, as "name=type@rank" separated by ';'. Elements of one
// xsd:choice share a rank; '*' marks a repeatable element. GLOBALS: root elements of parts -> their types.

export const TYPES: Record<string, string> = {
${lines.join('\n')}
};

export const GLOBALS: Record<string, string> = {
${globalLines.join('\n')}
};
`;
const target = join(ROOT, 'packages/model/src/schema.generated.ts');
writeFileSync(target, out);
console.log(`${lines.length} complex types, ${globalLines.length} global elements -> ${target} (${out.length} bytes)`);
