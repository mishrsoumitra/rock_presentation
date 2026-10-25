// packages/editor-core/src/save.ts  -- what changes on save: document statistics and the main part's type
import { normalizeName, resolveTarget } from '@pptx/opc';
import { CONTENT_TYPE, relKind, setText, textOf } from '@pptx/model';
import type { Editor } from './editor.ts';

/**
 * Bring derived data up to date before writing: docProps/app.xml's slide, hidden slide and notes counts, and
 * (when a template or show is saved as a presentation) the main part's content type. Not undoable, and nothing
 * to undo: it only restates the deck.
 */
export function prepareForSave(ed: Editor, options: { asPresentation?: boolean } = {}): void {
  const pkg = ed.pkg;
  if (options.asPresentation && ed.mainKind !== 'presentation') {
    pkg.setContentType(ed.doc.mainPart, ed.mainKind === 'macro' ? CONTENT_TYPE.macroPresentation : CONTENT_TYPE.presentation);
  }
  const rel = pkg.relationships('/').find(r => relKind(r.type) === 'extended-properties' && !r.external);
  const app = rel && pkg.getPart(resolveTarget('/', rel.target))?.name;
  if (!app) return;
  let root;
  try {
    root = pkg.xml(app).root;
  } catch {
    return;                                         // an unreadable app.xml is written back as it was
  }
  const slides = ed.doc.slideParts();
  const counts: Record<string, number> = {
    Slides: slides.length,
    HiddenSlides: slides.filter(s => ed.doc.slide(s).hidden).length,
    Notes: slides.filter(s => ed.doc.notesOf(s)).length,
  };
  let changed = false;
  for (const [name, n] of Object.entries(counts)) {
    const el = root.children.find(c => c.kind === 'element' && c.local === name);
    if (!el || el.kind !== 'element' || textOf(el) === String(n)) continue;
    setText(el, String(n));
    changed = true;
  }
  if (changed) pkg.markDirty(normalizeName(app));
}
