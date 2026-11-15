// packages/editor-core/src/notes.ts  -- speaker notes: reading a slide's notes and writing them
import { encodeXmlText, type XmlElement } from '@pptx/xml';
import { CONTENT_TYPE, insertInOrder, relType, uriFor } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { copyPart, freeName } from './parts.ts';
import { paragraphsXml } from './shapes.ts';
import { bodyPlainText, replaceText } from './text.ts';
import { element, insertAt, nextShapeId, unitsIn } from './xml.ts';

/** The text body of a notes page: its body placeholder, where PowerPoint keeps what the speaker reads. */
function notesBody(ed: Editor, notes: string): XmlElement | undefined {
  return ed.doc.slide(notes).shapes.find(s => s.placeholder?.type === 'body')?.txBody;
}

/** A slide's speaker notes as plain text: paragraphs and line breaks both read as new lines. Empty without any. */
export function notesText(ed: Editor, slide: string): string {
  const notes = ed.doc.notesOf(slide);
  const body = notes ? notesBody(ed, notes) : undefined;
  return body ? bodyPlainText(body).replace(/\u000b/g, '\n') : '';
}

const TREE_HEAD = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
  + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const decls = (strict: boolean) => `xmlns:a="${uriFor('a', strict)}" xmlns:r="${uriFor('r', strict)}" xmlns:p="${uriFor('p', strict)}"`;

const bodyShapeXml = (id: number, text: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Notes Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>`
  + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml(text)}</p:txBody></p:sp>`;

/** A notes page as PowerPoint makes one: the slide's picture above, the notes text below, both placed by the master. */
function notesXml(strict: boolean, text: string): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<p:notes ${decls(strict)}><p:cSld><p:spTree>${TREE_HEAD}`
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
    + `${bodyShapeXml(3, text)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

/** The notes master a deck without one gets: a portrait page, the slide picture in its top half, the text below. */
function notesMasterXml(strict: boolean, size: { cx: number; cy: number }): string {
  const w = size.cx, h = size.cy;
  const r = Math.round;
  const level = (n: number) => `<a:lvl${n}pPr marL="${(n - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
    + '<a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>'
    + `</a:lvl${n}pPr>`;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<p:notesMaster ${decls(strict)}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${TREE_HEAD}`
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr>'
    + `<p:spPr><a:xfrm><a:off x="${r(w * 0.1)}" y="${r(h * 0.125)}"/><a:ext cx="${r(w * 0.8)}" cy="${r(h * 0.3375)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>`
    + '<a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr>'
    + `<p:spPr><a:xfrm><a:off x="${r(w * 0.1)}" y="${r(h * 0.48125)}"/><a:ext cx="${r(w * 0.8)}" cy="${r(h * 0.39375)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
    + '<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/>'
    + '<a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>'
    + '</p:spTree></p:cSld>'
    + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    + `<p:notesStyle>${[1, 2, 3, 4, 5, 6, 7, 8, 9].map(level).join('')}</p:notesStyle></p:notesMaster>`;
}

/**
 * The deck's notes master, made when it has none: a new master with its own copy of the first slide master's theme
 * (a theme part belongs to one master), listed in p:notesMasterIdLst. Runs inside a structural edit.
 */
export function ensureNotesMaster(ed: Editor): string {
  const existing = ed.doc.notesMaster();
  if (existing) return existing;
  const pkg = ed.pkg;
  const main = ed.doc.mainPart;
  const master = freeName(pkg, '/ppt/notesMasters/notesMaster#.xml');
  pkg.addPart(master, CONTENT_TYPE.notesMaster, encodeXmlText(notesMasterXml(ed.strict, ed.doc.presentation.notesSize)));
  const slideMaster = ed.doc.masterParts()[0];
  const theme = slideMaster ? ed.doc.themeOf(slideMaster) : undefined;
  if (theme) pkg.addRelationship(master, relType('theme', ed.strict), copyPart(pkg, pkg, theme, new Map()));
  const rId = pkg.addRelationship(main, relType('notesMaster', ed.strict), master);
  const pres = ed.presentationRoot;
  insertInOrder(pres, element(pres, `<p:notesMasterIdLst><p:notesMasterId r:id="${rId}"/></p:notesMasterIdLst>`, ed.strict));
  return master;
}

/**
 * Set a slide's speaker notes. A slide that has a notes page keeps it, and the text keeps the formatting of its
 * first run; a slide without one gets a page (and the deck a notes master, if it had none). Typing in the notes
 * pane passes one `coalesceKey`, so a run of keystrokes is one undo step.
 */
export function setNotes(ed: Editor, slide: string, text: string, options: { coalesceKey?: string } = {}): Command {
  const existing = ed.doc.notesOf(slide);
  const body = existing ? notesBody(ed, existing) : undefined;
  if (existing && body) return replaceText(ed, existing, body, text, { label: 'Notes', ...options });
  if (existing) {
    // a notes page that lost its text placeholder gets one back
    const root = ed.doc.root(existing);
    const tree = ed.doc.slide(existing).spTree;
    if (!tree) throw new Error(`${existing} has no shape tree`);
    return ed.edit({
      label: 'Notes', parts: [existing], scope: () => [tree],
      apply: () => {
        const units = unitsIn(tree);
        const last = units[units.length - 1];
        insertAt(tree, last ? tree.children.indexOf(last) + 1 : tree.children.length, element(tree, bodyShapeXml(nextShapeId(root), text), ed.strict));
      },
    });
  }
  const main = ed.doc.mainPart;
  const notes = freeName(ed.pkg, '/ppt/notesSlides/notesSlide#.xml');
  return ed.edit({
    label: 'Notes', parts: [main, slide, notes], structural: true, scope: () => [ed.presentationRoot],
    apply: () => {
      const master = ensureNotesMaster(ed);
      ed.pkg.addPart(notes, CONTENT_TYPE.notesSlide, encodeXmlText(notesXml(ed.strict, text)));
      ed.pkg.addRelationship(notes, relType('notesMaster', ed.strict), master);
      ed.pkg.addRelationship(notes, relType('slide', ed.strict), slide);
      ed.pkg.addRelationship(slide, relType('notesSlide', ed.strict), notes);
    },
  });
}
