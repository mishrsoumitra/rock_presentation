// packages/editor-core/src/snapshot.ts  -- the state of the elements an edit touches, so undo can put it back
import type { XmlAttr, XmlElement, XmlNode } from '@pptx/xml';

interface ElementState {
  readonly el: XmlElement;
  readonly attrs: readonly XmlAttr[];
  readonly children: readonly XmlNode[];
  readonly selfClosing: boolean;
}

/**
 * Every element under some roots (the roots included) with its attributes, its child list and its self-closing
 * flag. Element objects are kept, not copied: restoring puts the very same elements back, so views and commands
 * that hold elements stay valid across undo and redo. Text nodes are never changed in place, so sharing them is
 * safe; attributes are, so they are copied both ways.
 */
export type Snapshot = readonly ElementState[];

export function capture(roots: readonly XmlElement[]): Snapshot {
  const out: ElementState[] = [];
  const seen = new Set<XmlElement>();
  const stack = [...roots];
  while (stack.length) {
    const el = stack.pop()!;
    if (seen.has(el)) continue;
    seen.add(el);
    out.push({ el, attrs: el.attrs.map(a => ({ ...a })), children: el.children.slice(), selfClosing: el.selfClosing });
    for (const c of el.children) if (c.kind === 'element') stack.push(c);
  }
  return out;
}

/** Put every captured element back as it was, parent pointers included. */
export function restore(snapshot: Snapshot): void {
  for (const s of snapshot) {
    s.el.attrs = s.attrs.map(a => ({ ...a }));
    s.el.children = s.children.slice();
    s.el.selfClosing = s.selfClosing;
    for (const c of s.children) if (c.kind === 'element') c.parent = s.el;
  }
}
