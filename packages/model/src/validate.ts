// packages/model/src/validate.ts  -- schema child-order check, a local stand-in for part of the Open XML validator
import type { XmlElement } from '@pptx/xml';
import { lname } from './dom.ts';
import { schemaType } from './order.ts';
import { TYPES } from './schema.generated.ts';

interface Slot { rank: number; repeat: boolean; }
const cache = new Map<string, Map<string, Slot>>();

function slots(type: string): Map<string, Slot> | undefined {
  let m = cache.get(type);
  if (m) return m;
  const src = TYPES[type];
  if (src === undefined) return undefined;
  m = new Map();
  for (const entry of src.split(';')) {
    const at = entry.lastIndexOf('@');
    const repeat = entry.endsWith('*');
    const name = entry.slice(0, entry.indexOf('='));
    if (!m.has(name)) m.set(name, { rank: Number(entry.slice(at + 1, repeat ? -1 : undefined)), repeat });
  }
  cache.set(type, m);
  return m;
}

/**
 * Every place under `root` where children the schema knows are out of order, or where a slot that holds one element
 * holds several. Children the schema does not know (Office extensions) are not judged; mc:AlternateContent is
 * judged inside each branch. An empty list means the order is valid for everything the ECMA schemas cover.
 */
export function checkOrder(root: XmlElement): string[] {
  const problems: string[] = [];
  const walk = (el: XmlElement, where: string) => {
    const type = schemaType(el);
    const s = type ? slots(type) : undefined;
    let last = -1;
    const used = new Set<number>();
    for (const c of el.children) {
      if (c.kind !== 'element') continue;
      const slot = s?.get(lname(c));
      if (s && slot) {
        if (slot.rank < last) problems.push(`${where}/${lname(c)}: out of schema order in ${type}`);
        else if (!slot.repeat && used.has(slot.rank)) problems.push(`${where}/${lname(c)}: second element in a single slot of ${type}`);
        last = Math.max(last, slot.rank);
        used.add(slot.rank);
      }
      walk(c, `${where}/${lname(c)}`);
    }
  };
  walk(root, lname(root));
  return problems;
}
