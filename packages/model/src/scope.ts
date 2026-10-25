// packages/model/src/scope.ts
import { XMLNS_NS, type XmlElement } from '@pptx/xml';

/** The namespace URI bound to `prefix` at `el` ('' for the default namespace), or undefined. Used for prefixes that
 *  appear in attribute values, such as mc:Ignorable and mc:Choice Requires, which the parser cannot resolve. */
export function lookupNamespace(el: XmlElement | null, prefix: string): string | undefined {
  for (let e = el; e; e = e.parent) {
    for (const a of e.attrs) {
      if (a.ns !== XMLNS_NS) continue;
      if ((a.prefix === null ? '' : a.local) === prefix) return a.value === '' ? undefined : a.value;
    }
  }
  return undefined;
}
