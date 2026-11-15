// packages/editor-core/src/comments.ts  -- comments on slides: reading, adding, changing and deleting them
import { encodeXmlText, type XmlElement } from '@pptx/xml';
import { attr, child, children, int, relType, setAttr, setText, textOf, uriFor } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { collectOrphans, freeName } from './parts.ts';
import { element, esc, insertAt } from './xml.ts';

const AUTHORS_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml';
const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';

/** A comment on a slide. A comment is named by its author and its number among that author's comments. */
export interface SlideComment {
  authorId: number;
  /** The number of this comment among its author's (unique with `authorId` in the deck). */
  index: number;
  author: string;
  initials: string;
  /** When it was written, as the file has it (ISO 8601). */
  date: string;
  text: string;
}

/** The part that lists who has written comments, when the deck has one. */
function authorsPart(ed: Editor): string | undefined { return ed.doc.targets(ed.doc.mainPart, 'commentAuthors')[0]; }
/** A slide's comments part, when it has one. */
function commentsPart(ed: Editor, slide: string): string | undefined { return ed.doc.targets(slide, 'comments')[0]; }

/**
 * A slide's comments, in the order they were written. These are the comments every version of PowerPoint reads;
 * the threaded comments newer versions add beside them are kept in the file but not listed here.
 */
export function slideComments(ed: Editor, slide: string): SlideComment[] {
  const part = commentsPart(ed, slide);
  if (!part) return [];
  const authors = new Map<number, { name: string; initials: string }>();
  const ap = authorsPart(ed);
  for (const a of ap ? children(ed.doc.root(ap), 'p:cmAuthor') : []) authors.set(int(attr(a, 'id'), 0), { name: attr(a, 'name') ?? '', initials: attr(a, 'initials') ?? '' });
  return children(ed.doc.root(part), 'p:cm').map(cm => {
    const authorId = int(attr(cm, 'authorId'), 0);
    const who = authors.get(authorId);
    return { authorId, index: int(attr(cm, 'idx'), 0), author: who?.name ?? 'Unknown', initials: who?.initials ?? '', date: attr(cm, 'dt') ?? '', text: textOf(child(cm, 'p:text')) };
  });
}

/** How many comments each slide has, for the slide panel. */
export function commentCount(ed: Editor, slide: string): number {
  const part = commentsPart(ed, slide);
  return part ? children(ed.doc.root(part), 'p:cm').length : 0;
}

const decls = (ed: Editor) => `xmlns:a="${uriFor('a', ed.strict)}" xmlns:r="${uriFor('r', ed.strict)}" xmlns:p="${uriFor('p', ed.strict)}"`;
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** Initials for a name: the first letters of its words, as PowerPoint makes them ("Asha Rao" gives "AR"). */
export function initialsOf(name: string): string {
  const letters = name.trim().split(/\s+/).filter(Boolean).map(w => [...w][0]!.toUpperCase());
  return letters.slice(0, 3).join('') || '?';
}

/**
 * Add a comment to a slide. The first comment of a deck makes its list of comment authors, and the first of a
 * slide its comments part. An author is found by name, or added; each author's comments are numbered from 1.
 * `at` is where its marker sits on the slide, in EMU; `when` is the time written into the file (now, by default).
 */
export function addComment(ed: Editor, slide: string, text: string, author: { name: string; initials?: string }, at: { x: number; y: number } = { x: 127000, y: 127000 }, when: Date = new Date()): Command {
  const main = ed.doc.mainPart;
  const authorsName = authorsPart(ed) ?? '/ppt/commentAuthors.xml';
  const commentsName = commentsPart(ed, slide) ?? freeName(ed.pkg, '/ppt/comments/comment#.xml');
  return ed.edit({
    label: 'New Comment', parts: [main, slide, authorsName, commentsName], structural: true,
    scope: () => [...(authorsPart(ed) ? [ed.doc.root(authorsPart(ed)!)] : []), ...(commentsPart(ed, slide) ? [ed.doc.root(commentsPart(ed, slide)!)] : [])],
    apply: () => {
      const name = author.name.trim() || 'Unknown';
      let ap = authorsPart(ed);
      if (!ap) {
        ed.pkg.addPart(authorsName, AUTHORS_TYPE, encodeXmlText(`${HEAD}<p:cmAuthorLst ${decls(ed)}/>`));
        ed.pkg.addRelationship(main, relType('commentAuthors', ed.strict), authorsName);
        ap = authorsName;
      }
      const list = ed.doc.root(ap);
      let who = children(list, 'p:cmAuthor').find(a => attr(a, 'name') === name);
      if (!who) {
        const used = children(list, 'p:cmAuthor').map(a => int(attr(a, 'id'), 0));
        const id = used.length ? Math.max(...used) + 1 : 0;
        who = element(list, `<p:cmAuthor id="${id}" name="${esc(name)}" initials="${esc(author.initials?.trim() || initialsOf(name))}" lastIdx="0" clrIdx="${id % 10}"/>`, ed.strict);
        insertAt(list, list.children.length, who);
      }
      const index = int(attr(who, 'lastIdx'), 0) + 1;
      setAttr(who, 'lastIdx', index);
      ed.pkg.markDirty(ap);

      let cp = commentsPart(ed, slide);
      if (!cp) {
        ed.pkg.addPart(commentsName, COMMENTS_TYPE, encodeXmlText(`${HEAD}<p:cmLst ${decls(ed)}/>`));
        ed.pkg.addRelationship(slide, relType('comments', ed.strict), commentsName);
        cp = commentsName;
      }
      const root = ed.doc.root(cp);
      // PowerPoint places comment markers in units of 1/576 inch
      const unit = (emu: number) => Math.max(0, Math.round((emu * 576) / 914400));
      const cm = element(root, `<p:cm authorId="${attr(who, 'id')}" dt="${when.toISOString().replace('Z', '')}" idx="${index}"><p:pos x="${unit(at.x)}" y="${unit(at.y)}"/><p:text>${esc(text)}</p:text></p:cm>`, ed.strict);
      insertAt(root, root.children.length, cm);
      ed.pkg.markDirty(cp);
    },
  });
}

function find(ed: Editor, slide: string, authorId: number, index: number): { part: string; root: XmlElement; cm: XmlElement } {
  const part = commentsPart(ed, slide);
  const root = part ? ed.doc.root(part) : undefined;
  const cm = root ? children(root, 'p:cm').find(c => int(attr(c, 'authorId'), -1) === authorId && int(attr(c, 'idx'), -1) === index) : undefined;
  if (!part || !root || !cm) throw new Error('that comment is not on this slide');
  return { part, root, cm };
}

/** Change what a comment says. Its author, number and time stay. */
export function editComment(ed: Editor, slide: string, authorId: number, index: number, text: string): Command {
  const { part, cm } = find(ed, slide, authorId, index);
  return ed.edit({
    label: 'Edit Comment', parts: [part], scope: () => [cm],
    apply: () => {
      const t = child(cm, 'p:text');
      if (t) setText(t, text);
      else insertAt(cm, cm.children.length, element(cm, `<p:text>${esc(text)}</p:text>`, ed.strict));
    },
  });
}

/** Delete a comment. The slide's last comment takes the slide's comments part with it. */
export function deleteComment(ed: Editor, slide: string, authorId: number, index: number): Command {
  const { part, root } = find(ed, slide, authorId, index);
  return ed.edit({
    label: 'Delete Comment', parts: [slide, part], structural: true, scope: () => [root],
    apply: () => {
      const cm = find(ed, slide, authorId, index).cm;
      root.children.splice(root.children.indexOf(cm), 1);
      if (children(root, 'p:cm').length) return;
      const rel = ed.pkg.relationships(slide).find(r => !r.external && r.type.endsWith('/comments'));
      if (rel) ed.pkg.removeRelationship(slide, rel.id);
      collectOrphans(ed.pkg, [part]);
    },
  });
}
