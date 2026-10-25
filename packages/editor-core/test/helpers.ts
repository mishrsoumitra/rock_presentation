// Test helpers: open decks, fingerprint a package, and the phase 1 oracle for edited packages.
import { normalizeName, OpcPackage, relsNameFor, resolveTarget, type Part } from '@pptx/opc';
import { checkOrder } from '@pptx/model';
import { canonical, decodeXmlBytes, parseXml, type XmlElement } from '@pptx/xml';
import { Editor, relRefs, shapeIds, type Command } from '../src/index.ts';

export function openEditor(bytes: Uint8Array): Editor {
  let t = 0;
  return new Editor(OpcPackage.open(bytes), { now: () => (t += 10) });
}

export interface PartState { contentType: string; bytes: Uint8Array; xml?: string; }

/** Every part: its content type, its bytes, and its canonical XML when it is parsed. */
export function state(ed: Editor): Map<string, PartState> {
  const out = new Map<string, PartState>();
  for (const name of ed.pkg.partNames()) {
    const p = ed.pkg.getPart(name)!;
    out.set(name, { contentType: p.contentType, bytes: p.bytes, ...(p.xml ? { xml: canonical(p.xml) } : {}) });
  }
  return out;
}

/** A part's canonical XML: from its tree when parsed, else from its bytes. */
function xmlOf(s: PartState): string | undefined {
  if (s.xml !== undefined) return s.xml;
  try {
    return canonical(parseXml(decodeXmlBytes(s.bytes).text));
  } catch {
    return undefined;
  }
}

/** Whether two package states hold the same parts with the same content; parsing a part is not a change. */
export function diffStates(a: Map<string, PartState>, b: Map<string, PartState>): string[] {
  const out: string[] = [];
  if ([...a.keys()].join('\n') !== [...b.keys()].join('\n')) out.push(`parts differ: ${[...a.keys()].filter(k => !b.has(k)).join(' ')} | ${[...b.keys()].filter(k => !a.has(k)).join(' ')}`);
  for (const [k, x] of a) {
    const y = b.get(k);
    if (!y) continue;
    if (x.contentType !== y.contentType) out.push(`${k}: content type`);
    else if (x.xml === undefined && y.xml === undefined) { if (x.bytes !== y.bytes) out.push(`${k}: bytes`); }
    else if (xmlOf(x) !== xmlOf(y)) out.push(`${k}: xml`);
  }
  return out;
}

/** Problems a package has, as text, for the parts given: schema order and relationship references. */
function problemsOf(pkg: OpcPackage, names: string[]): Set<string> {
  const out = new Set<string>();
  for (const name of names) {
    const p = pkg.getPart(name);
    if (!p || !p.xml || name.endsWith('.rels')) continue;
    for (const x of checkOrder(p.xml.root)) out.add(`${name}: ${x}`);
    const rels = pkg.relationships(name);
    const known = new Set(rels.map(r => r.id));
    for (const ref of relRefs(p.xml.root)) if (!known.has(ref.value)) out.add(`${name}: ${ref.el.local}/@r:${ref.local} ${ref.value} has no relationship`);
    for (const r of rels) {
      if (!r.external && !pkg.getPart(resolveTarget(normalizeName(name), r.target))) out.add(`${name}: relationship ${r.id} to missing ${r.target}`);
    }
    // every animation target must be a shape of the slide
    const timing = p.xml.root.children.find(c => c.kind === 'element' && c.local === 'timing');
    if (timing && timing.kind === 'element') {
      const ids = shapeIds(p.xml.root);
      const walk = (e: XmlElement) => {
        if (e.local === 'spTgt') {
          const spid = Number(e.attrs.find(a => a.local === 'spid')?.value);
          if (!ids.has(spid)) out.add(`${name}: an animation targets shape ${spid}, which is not on the slide`);
        }
        for (const c of e.children) if (c.kind === 'element') walk(c);
      };
      walk(timing);
    }
  }
  return out;
}

/**
 * The phase 1 oracle for an edited package: it saves, opens again with the same part names and content types,
 * every parsed part reads back to the same canonical XML, and edited parts gained no schema-order or relationship
 * problems they did not have before (`baseline`).
 */
export function oracle(ed: Editor, baseline: Set<string> | ((parts: string[]) => Set<string>)): string[] {
  const saved = ed.pkg.save();
  const back = OpcPackage.open(saved);
  const problems: string[] = [];
  const names = ed.pkg.partNames();
  const backNames = new Set(back.partNames());
  for (const n of names) {
    const p = ed.pkg.getPart(n)!;
    if (!backNames.has(n)) { problems.push(`${n} lost on save`); continue; }
    if (back.getPart(n)!.contentType !== p.contentType) problems.push(`${n} content type changed on save`);
    if (p.xml && canonical(back.xml(n)) !== canonical(p.xml)) problems.push(`${n} reads back differently`);
  }
  if (backNames.size !== names.length) problems.push('extra parts after save');
  const dirty = names.filter(n => ed.pkg.getPart(n)!.dirty);
  const found = [...problemsOf(ed.pkg, dirty)];
  if (!found.length) return problems;
  // a copied part (an imported master) may carry its source's problems: compare what, not where
  const known = new Set([...(typeof baseline === 'function' ? baseline(dirty) : baseline)].map(withoutPart));
  for (const x of found) if (!known.has(withoutPart(x))) problems.push(x);
  return problems;
}

const withoutPart = (problem: string) => problem.slice(problem.indexOf(': ') + 2);

/** Parse every XML part, so fingerprints compare trees rather than buffers whichever parts a command reads. */
export function parseAll(ed: Editor): void {
  for (const n of ed.pkg.partNames()) if (n.endsWith('.xml') || n.endsWith('.rels')) touchParse(ed.pkg.getPart(n)!, ed.pkg);
}

/** Problems already present in the parts a command may touch, so the oracle reports only new ones. */
export function baselineOf(ed: Editor): Set<string> {
  return problemsOf(ed.pkg, ed.pkg.partNames().filter(n => ed.pkg.getPart(n)!.xml));
}

/**
 * Run a command through its whole life: run, check the oracle, undo (the package is exactly as before), redo
 * (exactly as after), and undo again, leaving the package as it was. Returns every problem found.
 */
export function cycle(ed: Editor, make: () => Command, options: { parse?: boolean } = {}): string[] {
  if (options.parse ?? true) parseAll(ed);
  const before = state(ed);
  const cmd = make();
  ed.run(cmd);
  ed.log.seal();
  const after = state(ed);
  // the problems the edited parts had before the command are the baseline: read them once undone
  const problems = oracle(ed, parts => {
    ed.undo();
    const b = problemsOf(ed.pkg, parts.filter(p => ed.pkg.getPart(p)).map(p => (ed.pkg.xml(p), p)));
    ed.redo();
    return b;
  });
  ed.undo();
  problems.push(...diffStates(before, state(ed)).map(d => `undo ${cmd.label}: ${d}`));
  ed.redo();
  problems.push(...diffStates(after, state(ed)).map(d => `redo ${cmd.label}: ${d}`));
  ed.undo();
  problems.push(...diffStates(before, state(ed)).map(d => `undo again ${cmd.label}: ${d}`));
  return problems;
}

/** Parse the parts a test compares, so fingerprints compare trees rather than buffers. */
function touchParse(p: Part, pkg: OpcPackage): void {
  if (p.xml || p.bytes.length > 4 << 20) return;
  try {
    pkg.xml(p.name);
  } catch {
    // unreadable parts stay bytes
  }
}

export { relsNameFor };
