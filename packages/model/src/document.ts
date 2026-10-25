// packages/model/src/document.ts  -- a presentation package seen as slides, layouts, masters and themes
import { normalizeName, resolveTarget, type OpcPackage } from '@pptx/opc';
import type { XmlElement } from '@pptx/xml';
import { attrNs, child, children } from './dom.ts';
import { isStrictUri, relKind } from './ns.ts';
import { parseTheme, type ThemeDef } from './theme.ts';
import { PresentationView, SlideView, type PartOwner } from './views.ts';

export class PresentationDocument {
  readonly pkg: OpcPackage;
  readonly mainPart: string;
  /** Strict documents get Strict namespaces and relationship types in everything the editor writes. */
  readonly strict: boolean;
  private readonly themes = new Map<string, ThemeDef>();

  constructor(pkg: OpcPackage) {
    this.pkg = pkg;
    const main = pkg.relationships('/').find(r => relKind(r.type) === 'officeDocument' && !r.external);
    if (!main) throw new Error('not a presentation: the package has no officeDocument relationship');
    const name = pkg.getPart(resolveTarget('/', main.target))?.name;
    if (!name) throw new Error(`not a presentation: ${main.target} is missing`);
    this.mainPart = name;
    this.strict = isStrictUri(this.root(name).ns);
  }

  /** The handle views use to mark a part dirty after writing to it. */
  owner(part: string): PartOwner {
    return { part, touch: () => this.pkg.markDirty(part) };
  }

  root(part: string): XmlElement {
    return this.pkg.xml(part).root;
  }

  get presentation(): PresentationView {
    return new PresentationView(this.root(this.mainPart), this.owner(this.mainPart));
  }

  /** The part an internal relationship points at; undefined when it is missing, external or dangling. */
  target(source: string, rId: string): string | undefined {
    const rel = this.pkg.relationships(source).find(r => r.id === rId);
    if (!rel || rel.external) return undefined;
    return this.pkg.getPart(resolveTarget(normalizeName(source), rel.target))?.name;
  }

  /** The parts that relationships of one kind ('slideLayout', 'image', ...) point at, in relationship order. */
  targets(source: string, kind: string): string[] {
    const out: string[] = [];
    for (const r of this.pkg.relationships(source)) {
      if (r.external || relKind(r.type) !== kind) continue;
      const name = this.pkg.getPart(resolveTarget(normalizeName(source), r.target))?.name;
      if (name) out.push(name);
    }
    return out;
  }

  /** Slide parts in presentation order (p:sldIdLst). */
  slideParts(): string[] {
    return this.presentation.slideIds.map(e => this.target(this.mainPart, e.rId)).filter((p): p is string => p !== undefined);
  }

  masterParts(): string[] {
    return this.presentation.masterIds.map(e => this.target(this.mainPart, e.rId)).filter((p): p is string => p !== undefined);
  }

  /** A master's layouts in the order of its p:sldLayoutIdLst. */
  layoutsOf(master: string): string[] {
    const ids = children(child(this.root(master), 'p:sldLayoutIdLst'), 'p:sldLayoutId');
    return ids.map(e => this.target(master, attrNs(e, 'r', 'id') ?? '')).filter((p): p is string => p !== undefined);
  }

  layoutOf(slide: string): string | undefined { return this.targets(slide, 'slideLayout')[0]; }
  masterOf(layout: string): string | undefined { return this.targets(layout, 'slideMaster')[0]; }
  themeOf(master: string): string | undefined { return this.targets(master, 'theme')[0]; }
  notesOf(slide: string): string | undefined { return this.targets(slide, 'notesSlide')[0]; }
  notesMaster(): string | undefined {
    const rId = this.presentation.notesMasterRId;
    return rId ? this.target(this.mainPart, rId) : undefined;
  }
  tableStylesPart(): string | undefined { return this.targets(this.mainPart, 'tableStyles')[0]; }

  slide(part: string): SlideView {
    return new SlideView(this.root(part), this.owner(part));
  }

  /** The parsed theme of a theme part, cached until `invalidate(part)`. */
  theme(part: string): ThemeDef {
    let t = this.themes.get(part);
    if (!t) {
      t = parseTheme(this.root(part));
      this.themes.set(part, t);
    }
    return t;
  }

  /** Drop what is cached for a part after it was edited. */
  invalidate(part: string): void {
    this.themes.delete(part);
  }
}
