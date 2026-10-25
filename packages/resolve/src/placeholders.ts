// packages/resolve/src/placeholders.ts  -- which layout and master placeholder a placeholder inherits from
import type { Placeholder, ShapeView } from '@pptx/model';

const TITLES = new Set(['title', 'ctrTitle']);
const FOOTERS = new Set(['dt', 'ftr', 'sldNum', 'hdr']);

/** Placeholders of a slide-like part, flattened out of groups. */
export function placeholdersOf(shapes: ShapeView[]): ShapeView[] {
  const out: ShapeView[] = [];
  const walk = (list: ShapeView[]) => {
    for (const s of list) {
      if (s.placeholder) out.push(s);
      if (s.kind === 'grpSp') walk(s.children);
    }
  };
  walk(shapes);
  return out;
}

/**
 * The layout placeholder a slide placeholder inherits from. Titles match titles; a placeholder that states an index
 * matches that index (its type may differ: a slide "body" with idx 1 inherits from a layout "obj" with idx 1);
 * otherwise the same type, then the body family (body, obj, subTitle).
 */
export function matchLayoutPlaceholder(ph: Placeholder, layout: ShapeView[]): ShapeView | undefined {
  const cands = placeholdersOf(layout);
  if (TITLES.has(ph.type)) {
    return cands.find(s => s.placeholder!.type === ph.type) ?? cands.find(s => TITLES.has(s.placeholder!.type));
  }
  if (ph.hasIdx) {
    const byIdx = cands.filter(s => s.placeholder!.idx === ph.idx && s.placeholder!.hasIdx);
    const hit = byIdx.find(s => s.placeholder!.type === ph.type) ?? byIdx[0];
    if (hit) return hit;
  }
  const byType = cands.find(s => s.placeholder!.type === ph.type);
  if (byType) return byType;
  if (FOOTERS.has(ph.type)) return undefined;
  return cands.find(s => ['body', 'obj'].includes(s.placeholder!.type) && (!ph.hasIdx || !s.placeholder!.hasIdx || s.placeholder!.idx === ph.idx))
    ?? cands.find(s => ['body', 'obj', 'subTitle'].includes(s.placeholder!.type));
}

/** The master placeholder a layout (or slide) placeholder inherits from: masters hold one per role. */
export function matchMasterPlaceholder(ph: Placeholder, master: ShapeView[]): ShapeView | undefined {
  const cands = placeholdersOf(master);
  if (TITLES.has(ph.type)) return cands.find(s => TITLES.has(s.placeholder!.type));
  if (FOOTERS.has(ph.type)) return cands.find(s => s.placeholder!.type === ph.type);
  if (ph.type === 'sldImg') return cands.find(s => s.placeholder!.type === 'sldImg');
  return cands.find(s => s.placeholder!.type === 'body') ?? cands.find(s => s.placeholder!.type === 'obj');
}

/** The master text style a shape's text falls back to: title, body, or other (non-placeholders and footers). */
export function textStyleBucket(ph: Placeholder | null): 'title' | 'body' | 'other' {
  if (!ph) return 'other';
  if (TITLES.has(ph.type)) return 'title';
  if (FOOTERS.has(ph.type)) return 'other';
  return 'body';
}
