// packages/editor-core/src/editor.ts  -- an open deck being edited: package, document, command log, versions
import { relsNameFor, type OpcPackage } from '@pptx/opc';
import { child, CONTENT_TYPE, PresentationDocument, ShapeView, shapeKind, type SlideView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import { CommandLog, edit, type Command, type CommandLogOptions, type EditSpec } from './history.ts';
import { prepareForSave } from './save.ts';

export type ChangeListener = (parts: readonly string[]) => void;

/**
 * The editing session. The package's trees are edited in place by commands; every command bumps a version counter
 * per part it touched, which is what the UI subscribes to (a drag on slide 7 re-renders slide 7 and its thumbnail
 * only) and what drops the document's caches for those parts.
 */
export class Editor {
  readonly pkg: OpcPackage;
  readonly doc: PresentationDocument;
  readonly log: CommandLog;
  private readonly versions = new Map<string, number>();
  private readonly listeners = new Set<ChangeListener>();
  /** Bumped on every change anywhere. */
  version = 0;

  constructor(pkg: OpcPackage, options: CommandLogOptions = {}) {
    this.pkg = pkg;
    this.doc = new PresentationDocument(pkg);
    this.log = new CommandLog(pkg, parts => this.changed(parts), options);
  }

  /** The command for an in-place edit of this package (see `edit`). */
  edit(spec: EditSpec): Command { return edit(this.pkg, spec); }

  run(cmd: Command): void { this.log.run(cmd); }
  undo(): Command | undefined { return this.log.undo(); }
  redo(): Command | undefined { return this.log.redo(); }

  versionOf(part: string): number { return this.versions.get(part) ?? 0; }

  subscribe(fn: ChangeListener): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  /** Report parts changed outside a command (loading fonts, decoding pictures): re-render, nothing to undo. */
  changed(parts: readonly string[]): void {
    for (const p of parts) {
      this.versions.set(p, (this.versions.get(p) ?? 0) + 1);
      this.doc.invalidate(p);
    }
    this.version++;
    for (const fn of this.listeners) fn(parts);
  }

  get strict(): boolean { return this.doc.strict; }
  slide(part: string): SlideView { return this.doc.slide(part); }
  spTree(part: string): XmlElement {
    const tree = this.doc.slide(part).spTree;
    if (!tree) throw new Error(`${part} has no shape tree`);
    return tree;
  }

  /** A shape of a slide by id, searching groups; the first one when a broken file repeats an id. */
  shape(part: string, id: number): ShapeView | undefined {
    const find = (shapes: ShapeView[]): ShapeView | undefined => {
      for (const s of shapes) {
        if (s.id === id) return s;
        const inner = s.kind === 'grpSp' ? find(s.children) : undefined;
        if (inner) return inner;
      }
      return undefined;
    };
    return find(this.doc.slide(part).shapes);
  }

  /** The view of a shape element of a part. */
  view(part: string, el: XmlElement): ShapeView {
    if (!shapeKind(el)) throw new Error(`${el.local} is not a shape`);
    return new ShapeView(el, this.doc.owner(part));
  }

  /** The relationships part of a part, when it has one. */
  relsOf(part: string): string | undefined { return this.pkg.getPart(relsNameFor(part))?.name; }

  /** The deck as .pptx bytes (a template or show opened here is written as a presentation if `asPresentation`). */
  save(options: { asPresentation?: boolean } = {}): Uint8Array {
    prepareForSave(this, options);
    return this.pkg.save();
  }

  /** Whether the main part is a template (.potx) or a show (.ppsx) rather than a presentation. */
  get mainKind(): 'presentation' | 'template' | 'slideshow' | 'macro' {
    const ct = this.pkg.getPart(this.doc.mainPart)?.contentType ?? '';
    if (ct === CONTENT_TYPE.template || ct === CONTENT_TYPE.macroTemplate) return 'template';
    if (ct === CONTENT_TYPE.slideshow || ct === CONTENT_TYPE.macroSlideshow) return 'slideshow';
    if (ct === CONTENT_TYPE.macroPresentation) return 'macro';
    return 'presentation';
  }

  /** p:presentation's root element. */
  get presentationRoot(): XmlElement { return this.doc.root(this.doc.mainPart); }

  /** Child of p:presentation by name, for commands that scope it. */
  presentationChild(name: string): XmlElement | undefined { return child(this.presentationRoot, name); }
}
