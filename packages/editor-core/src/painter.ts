// packages/editor-core/src/painter.ts  -- the format painter: the look of one shape, picked up and put on others
import { attr, child, children, clone, insertInOrder, is, runsOf, setAttr, ShapeView } from '@pptx/model';
import { serializeElement, type XmlElement } from '@pptx/xml';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { paragraphs } from './text.ts';
import { fragment, twins, unitOf } from './xml.ts';

/**
 * A look picked up from a shape (or from text): what the format painter carries. Everything is kept as XML text,
 * so it can be put on shapes of another slide or after the source has changed. A picture fill is not carried: its
 * picture belongs to the slide it is on.
 */
export interface FormatClip {
  /** Shape formatting: the fill, line and effects of its properties, and its theme style. Absent when picked up from text. */
  shape?: { props: string[]; style?: string };
  /** Character formatting: the run properties of the first text. */
  run?: string;
  /** Paragraph alignment, and how the text sits in its box. */
  align?: string;
  anchor?: string;
}

/** What of a shape's properties is its look (not its place, size or geometry). */
const LOOK = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d'];
const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];
/** Run properties that are not a look: the language, links, and what a spelling check left behind. */
const NOT_LOOK_ATTRS = new Set(['lang', 'altLang', 'dirty', 'err', 'smtClean', 'smtId', 'bmk']);
const NOT_LOOK_CHILDREN = ['a:hlinkClick', 'a:hlinkMouseOver', 'a:extLst'];

/** The run properties that say how the first of the text looks: its first run's, else its first paragraph mark's. */
function firstRunProps(txBody: XmlElement | undefined): XmlElement | undefined {
  for (const p of txBody ? paragraphs(txBody) : []) {
    const r = runsOf(p).find(x => x.kind === 'r');
    const rPr = r?.rPr ?? child(p, 'a:endParaRPr');
    if (rPr) return rPr;
  }
  return undefined;
}

/** A copy of run properties with only what is a look. */
function lookOfRun(rPr: XmlElement): XmlElement {
  const c = clone(rPr);
  c.local = 'rPr';
  c.attrs = c.attrs.filter(a => a.ns !== '' || !NOT_LOOK_ATTRS.has(a.local));
  c.children = c.children.filter(n => n.kind === 'element' && !NOT_LOOK_CHILDREN.some(name => is(n, name)));
  c.selfClosing = c.children.length === 0;
  return c;
}

/** Pick up a shape's look: its fill (not a picture fill), line, effects and theme style, and how its text looks. */
export function copyFormat(shape: ShapeView): FormatClip {
  const clip: FormatClip = {};
  if (shape.kind !== 'graphicFrame' && shape.kind !== 'contentPart') {
    const spPr = shape.spPr;
    const style = child(shape.el, 'p:style');
    clip.shape = {
      props: spPr ? LOOK.map(n => child(spPr, n)).filter((x): x is XmlElement => !!x).map(x => serializeElement(clone(x))) : [],
      ...(style ? { style: serializeElement(clone(style)) } : {}),
    };
  }
  const body = shape.txBody;
  const rPr = firstRunProps(body);
  if (rPr) clip.run = serializeElement(lookOfRun(rPr));
  const p = body ? paragraphs(body)[0] : undefined;
  const algn = attr(child(p, 'a:pPr'), 'algn');
  if (algn) clip.align = algn;
  const anchor = attr(child(body, 'a:bodyPr'), 'anchor');
  if (anchor) clip.anchor = anchor;
  return clip;
}

/** Pick up only how text looks, from run properties (the text selected in a box). */
export function copyTextFormat(rPr: XmlElement | undefined): FormatClip {
  return rPr ? { run: serializeElement(lookOfRun(rPr)) } : {};
}

/** Put a look on run properties: everything that is a look is replaced, the language and any link stay. */
function paintRun(ed: Editor, rPr: XmlElement, look: string): void {
  const src = fragment(rPr, look, ed.strict)[0];
  if (!src) return;
  rPr.attrs = [...rPr.attrs.filter(a => a.ns !== '' || NOT_LOOK_ATTRS.has(a.local)), ...src.attrs.filter(a => a.ns === '')];
  const kept = rPr.children.filter(n => n.kind === 'element' && NOT_LOOK_CHILDREN.some(name => is(n, name)));
  rPr.children = [];
  for (const c of src.children) if (c.kind === 'element') insertInOrder(rPr, clone(c, rPr));
  for (const k of kept) if (k.kind === 'element') insertInOrder(rPr, k);
  rPr.selfClosing = rPr.children.length === 0;
}

/**
 * Put a picked-up look on shapes. A shape takes the fill, line, effects and style (a picture or a table keeps its
 * own: it has none of those to set); every run of its text takes the character formatting, its paragraphs the
 * alignment, and its text box the anchoring. What the clip does not carry is left as it is.
 */
export function pasteFormat(ed: Editor, part: string, shapes: ShapeView[], clip: FormatClip): Command {
  return ed.edit({
    label: 'Format Painter', parts: [part], scope: () => shapes.map(s => unitOf(s.el)),
    apply: () => {
      for (const s of shapes) {
        for (const el of twins(s.el)) {
          const v = new ShapeView(el, s.owner);
          if (clip.shape && (v.kind === 'sp' || v.kind === 'cxnSp')) {
            const spPr = v.spPr ?? (() => { const e = fragment(el, '<p:spPr/>', ed.strict)[0]!; insertInOrder(el, e); return e; })();
            // the look goes; a picture fill the target has stays unless the clip brings another fill
            const bringsFill = clip.shape.props.some(x => FILLS.some(f => x.startsWith(`<${f}`)));
            spPr.children = spPr.children.filter(n => !(n.kind === 'element' && (LOOK.some(name => is(n, name)) || (bringsFill && is(n, 'a:blipFill')))));
            for (const x of clip.shape.props) for (const e of fragment(spPr, x, ed.strict)) insertInOrder(spPr, e);
            spPr.selfClosing = spPr.children.length === 0;
            const old = child(el, 'p:style');
            if (old) el.children.splice(el.children.indexOf(old), 1);
            if (clip.shape.style) for (const e of fragment(el, clip.shape.style, ed.strict)) insertInOrder(el, e);
          }
          const body = v.txBody;
          if (!body) continue;
          if (clip.run) {
            for (const p of paragraphs(body)) {
              for (const r of children(p, 'a:r')) {
                let rPr = child(r, 'a:rPr');
                if (!rPr) { rPr = fragment(r, '<a:rPr lang="en-US"/>', ed.strict)[0]!; insertInOrder(r, rPr); }
                paintRun(ed, rPr, clip.run);
              }
              const mark = child(p, 'a:endParaRPr');
              if (mark) paintRun(ed, mark, clip.run);
            }
          }
          if (clip.align) {
            for (const p of paragraphs(body)) {
              const pPr = child(p, 'a:pPr') ?? (() => { const e = fragment(p, '<a:pPr/>', ed.strict)[0]!; insertInOrder(p, e); return e; })();
              setAttr(pPr, 'algn', clip.align);
            }
          }
          const bodyPr = child(body, 'a:bodyPr');
          if (clip.anchor && bodyPr) setAttr(bodyPr, 'anchor', clip.anchor);
        }
      }
    },
  });
}
