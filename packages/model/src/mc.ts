// packages/model/src/mc.ts  -- Markup Compatibility (ECMA-376 Part 3)
import type { XmlElement } from '@pptx/xml';
import { lookupNamespace } from './scope.ts';
import { nsKey, type NsKey } from './ns.ts';
import { attr, is } from './dom.ts';

/**
 * The element children to read for `el`, with each mc:AlternateContent replaced by the branch a consumer that
 * understands `understood` must read: the first mc:Choice whose Requires prefixes all resolve to understood
 * namespaces, else the mc:Fallback, else nothing. The tree is not changed; saving keeps every branch.
 */
export function effectiveChildren(el: XmlElement, understood: ReadonlySet<NsKey>): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (c.kind !== 'element') continue;
    if (!is(c, 'mc:AlternateContent')) {
      out.push(c);
      continue;
    }
    const branch = selectBranch(c, understood);
    if (branch) out.push(...effectiveChildren(branch, understood));
  }
  return out;
}

/** The branch of one mc:AlternateContent that applies, or undefined when there is none. */
export function selectBranch(alt: XmlElement, understood: ReadonlySet<NsKey>): XmlElement | undefined {
  for (const c of alt.children) {
    if (c.kind !== 'element') continue;
    if (is(c, 'mc:Choice')) {
      const requires = (attr(c, 'Requires') ?? '').split(/\s+/).filter(Boolean);
      const ok = requires.length > 0 && requires.every(prefix => {
        const uri = lookupNamespace(c, prefix);
        const key = uri === undefined ? undefined : nsKey(uri);
        return key !== undefined && understood.has(key);
      });
      if (ok) return c;
    } else if (is(c, 'mc:Fallback')) {
      return c;
    }
  }
  return undefined;
}
