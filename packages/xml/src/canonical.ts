// packages/xml/src/canonical.ts  (test helper)
import { XML_NS, type XmlDocument, type XmlElement, type XmlMisc } from './tree.ts';

type Scope = Map<string | null, string>;   // prefix (null = default) -> namespace URI

/**
 * A comparison form for tests: names resolved to namespace URIs, attributes sorted, comments dropped,
 * adjacent text and CDATA merged, and whitespace-only text dropped where an element also has element
 * children. Two parts that differ only in prefixes, attribute order, CDATA versus escaped text or
 * indentation compare equal, the same way PowerPoint treats them. Processing instructions are kept, in and
 * around the root, because the application they are addressed to reads them. Namespaces are resolved from the
 * xmlns attributes in the tree, not from the `ns` fields, so a tree whose prefixes and declarations
 * disagree compares unequal to what a reader would see.
 */
export function canonical(doc: XmlDocument): string {
  const out: string[] = [];
  const pis = (misc: XmlMisc[]) => {
    for (const m of misc) if (m.kind === 'pi') out.push('<?', JSON.stringify(m.target), JSON.stringify(m.data), '?>');
  };
  pis(doc.prolog);
  canonElement(doc.root, new Map([['xml', XML_NS]]), out);
  pis(doc.epilog);
  return out.join('');
}

function canonElement(el: XmlElement, parentScope: Scope, out: string[]): void {
  let scope = parentScope;
  for (const a of el.attrs) {
    if (a.prefix === 'xmlns' || (a.prefix === null && a.local === 'xmlns')) {
      if (scope === parentScope) scope = new Map(parentScope);
      scope.set(a.prefix === null ? null : a.local, a.value);
    }
  }
  const resolve = (prefix: string | null, isAttr: boolean) =>
    prefix === null ? (isAttr ? '' : scope.get(null) ?? '') : scope.get(prefix) ?? `?${prefix}`;

  out.push('<{', resolve(el.prefix, false), '}', el.local);
  const attrs = el.attrs
    .filter(a => !(a.prefix === 'xmlns' || (a.prefix === null && a.local === 'xmlns')))
    .map(a => ({ key: `{${resolve(a.prefix, true)}}${a.local}`, value: a.value }))
    .sort((x, y) => (x.key < y.key ? -1 : x.key > y.key ? 1 : 0));
  for (const a of attrs) out.push(' ', a.key, '=', JSON.stringify(a.value));
  out.push('>');

  const hasElementChild = el.children.some(c => c.kind === 'element');
  let run: string | null = null;
  const flush = () => {
    // indentation between elements is never meaningful in OOXML
    if (run !== null && !(hasElementChild && /^[ \t\r\n]*$/.test(run))) out.push('(', JSON.stringify(run), ')');
    run = null;
  };
  for (const c of el.children) {
    if (c.kind === 'element') {
      flush();
      canonElement(c, scope, out);
    } else if (c.kind === 'text' || c.kind === 'cdata') {
      run = (run ?? '') + c.value;
    } else if (c.kind === 'pi') {
      flush();
      out.push('<?', JSON.stringify(c.target), JSON.stringify(c.data), '?>');
    }
    // comments are ignored and do not split a text run
  }
  flush();
  out.push('</>');
}

/** Where two canonical forms first differ, with context; null when they are equal. */
export function canonicalDiff(expected: string, actual: string, context = 120): string | null {
  if (expected === actual) return null;
  let i = 0;
  while (i < expected.length && i < actual.length && expected[i] === actual[i]) i++;
  const from = Math.max(0, i - context);
  return `first difference at ${i}\n  expected: …${expected.slice(from, i + context)}\n  actual:   …${actual.slice(from, i + context)}`;
}
