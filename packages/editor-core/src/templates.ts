// packages/editor-core/src/templates.ts  -- importing a template's masters into the open deck
import { partKey, type OpcPackage } from '@pptx/opc';
import { attr, child, children, isStrictUri, PresentationDocument, relKind, relType, setAttr } from '@pptx/model';
import { encodeXmlText, serializeXml } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { copyPart, hashMedia, mediaPart } from './parts.ts';
import { convertFlavour, element, insertAt } from './xml.ts';

/** Masters and layouts share one id space, starting at 2147483648 (ECMA-376 Part 1, 19.2.1.25 and 19.2.1.31). */
const FIRST_MASTER_ID = 2147483648;

function usedMasterLayoutIds(ed: Editor): number[] {
  const ids = ed.doc.presentation.masterIds.map(m => m.id);
  for (const m of ed.doc.masterParts()) {
    for (const l of children(child(ed.doc.root(m), 'p:sldLayoutIdLst'), 'p:sldLayoutId')) ids.push(Number(attr(l, 'id') ?? 0));
  }
  return ids;
}

/** Hash every media part of a template so importing it can reuse pictures the deck already has. */
export async function prepareTemplate(ed: Editor, from: OpcPackage): Promise<Map<string, string>> {
  const hashes = new Map<string, string>();
  for (const n of from.partNames()) if (n.startsWith('/ppt/media/')) hashes.set(n, await hashMedia(ed, from.getPart(n)!.bytes));
  return hashes;
}

/**
 * Import masters from a template (another deck or a .potx): each master is copied with its layouts, its theme and
 * every part they use, under fresh part names; pictures identical to ones in the deck are shared. The presentation
 * gets a slideMaster relationship and a p:sldMasterId, and every master and layout id is made unique in the deck.
 * A template of the other flavour (Strict or Transitional) is converted. Returns the new master part names.
 */
export function importMasters(ed: Editor, from: OpcPackage, hashes: ReadonlyMap<string, string>, masters?: string[]): { command: Command; masters: string[] } {
  const src = new PresentationDocument(from);
  const list = masters ?? src.masterParts();
  const pres = ed.presentationRoot;
  const main = ed.doc.mainPart;
  const created: string[] = [];
  const command = ed.edit({
    label: 'Import Slide Masters', parts: [main], structural: true, scope: () => [pres],
    apply: () => {
      created.length = 0;
      const copied = new Map<string, string>();
      const before = new Set(ed.pkg.partNames());
      const mapShared = (kind: string, target: string): string | null | undefined => {
        if (kind === 'image' || kind === 'media' || kind === 'video' || kind === 'audio') {
          const p = from.getPart(target);
          const hash = hashes.get(target);
          if (!p || !hash) return undefined;
          return mediaPart(ed.pkg, p.bytes, target.slice(target.lastIndexOf('.') + 1).toLowerCase(), p.contentType, hash, kind === 'image' ? 'image' : 'media');
        }
        if (kind === 'slideLayout' || kind === 'slideMaster' || kind === 'theme') return copied.get(partKey(target));
        return null;                                    // links to the template's slides do not come along
      };
      for (const m of list) created.push(copyPart(ed.pkg, from, m, copied, mapShared));
      if (src.strict !== ed.strict) convertParts(ed, [...ed.pkg.partNames()].filter(n => !before.has(n)));
      let next = Math.max(FIRST_MASTER_ID - 1, ...usedMasterLayoutIds(ed)) + 1;
      for (const m of created) {
        // layout ids inside the copied master are renumbered too, so none collides with the deck's
        for (const l of children(child(ed.doc.root(m), 'p:sldLayoutIdLst'), 'p:sldLayoutId')) setAttr(l, 'id', next++);
        ed.pkg.markDirty(m);
        const rId = ed.pkg.addRelationship(main, relType('slideMaster', ed.strict), m);
        const lst = child(pres, 'p:sldMasterIdLst')!;
        insertAt(lst, lst.children.length, element(lst, `<p:sldMasterId id="${next++}" r:id="${rId}"/>`, ed.strict));
      }
    },
  });
  return { command, masters: created };
}

/** Rewrite copied parts in the deck's flavour: element and attribute namespaces, and relationship types. */
function convertParts(ed: Editor, names: string[]): void {
  for (const n of names) {
    const p = ed.pkg.getPart(n)!;
    if (!/xml$/.test(p.contentType) && !n.endsWith('.rels')) continue;
    const doc = ed.pkg.xml(n);
    if (n.endsWith('.rels')) {
      for (const c of doc.root.children) {
        if (c.kind !== 'element') continue;
        const t = c.attrs.find(a => a.local === 'Type');
        if (t && (isStrictUri(t.value) !== ed.strict) && /officeDocument/.test(t.value)) t.value = relType(relKind(t.value), ed.strict);
      }
    } else convertFlavour(doc.root, ed.strict);
    ed.pkg.markDirty(n);
  }
}

/** The bytes of a part as the editor would write it now (for tests and debugging). */
export function partText(pkg: OpcPackage, name: string): Uint8Array {
  const p = pkg.getPart(name)!;
  return p.dirty && p.xml ? encodeXmlText(serializeXml(p.xml), p.encoding) : p.bytes;
}
