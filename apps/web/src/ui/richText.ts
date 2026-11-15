// apps/web/src/ui/richText.ts  -- text on the clipboard: paragraphs of runs, to and from HTML and plain text

/** A stretch of text with one look. Only what every app agrees on travels: bold, italic and underline. */
export interface Seg { text: string; b?: boolean; i?: boolean; u?: boolean; }
/** A paragraph: its stretches in order. A line break inside a paragraph is "\v" in a stretch's text, as in the deck. */
export type Para = Seg[];

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Plain text for the clipboard: paragraphs on lines of their own; a line break is a new line too. */
export function plainOf(paras: Para[]): string {
  return paras.map(p => p.map(s => s.text).join('').replace(/\u000b/g, '\n')).join('\n');
}

/** HTML for the clipboard: a <p> per paragraph, <b>, <i> and <u> around what has them, <br> for a line break. */
export function paragraphsToHtml(paras: Para[]): string {
  return paras.map(p => `<p>${p.map(s => {
    let h = esc(s.text).replace(/\u000b/g, '<br>');
    if (s.u) h = `<u>${h}</u>`;
    if (s.i) h = `<i>${h}</i>`;
    if (s.b) h = `<b>${h}</b>`;
    return h;
  }).join('')}</p>`).join('');
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
function decode(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (all, name: string) => {
    if (name[0] === '#') {
      const code = name[1] === 'x' || name[1] === 'X' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : all;
    }
    return ENTITIES[name.toLowerCase()] ?? all;
  });
}

/** Elements that start a new paragraph, before and after. */
const BLOCKS = new Set(['p', 'div', 'li', 'ul', 'ol', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'tr', 'table', 'blockquote', 'pre', 'section', 'article', 'header', 'footer', 'dd', 'dt']);
/** Elements whose content is not text to paste. */
const SKIPPED = new Set(['style', 'script', 'head', 'title', 'template', 'noscript']);

interface Look { b: boolean; i: boolean; u: boolean; }

/** What an element's tag and inline style say about the look of the text inside it. */
function lookOf(tag: string, attrs: string, outer: Look): Look {
  const look = { ...outer };
  if (tag === 'b' || tag === 'strong' || /^h[1-6]$/.test(tag)) look.b = true;
  if (tag === 'i' || tag === 'em') look.i = true;
  if (tag === 'u') look.u = true;
  const style = /\bstyle\s*=\s*("([^"]*)"|'([^']*)')/i.exec(attrs);
  const css = (style?.[2] ?? style?.[3] ?? '').toLowerCase();
  const weight = /font-weight\s*:\s*([a-z0-9]+)/.exec(css)?.[1];
  if (weight) look.b = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
  const slant = /font-style\s*:\s*([a-z]+)/.exec(css)?.[1];
  if (slant) look.i = slant === 'italic' || slant === 'oblique';
  const deco = /text-decoration(?:-line)?\s*:\s*([^;]+)/.exec(css)?.[1];
  if (deco) look.u = deco.includes('underline');
  return look;
}

/**
 * Paragraphs of text from HTML on the clipboard, by a small reader of its own: nothing in the HTML is ever run or
 * loaded, since it is never made into a document. Block elements and <br> end lines, runs of white space become one
 * space (as a browser shows them), and bold, italic and underline are kept from tags and inline styles.
 */
export function htmlToParagraphs(html: string): Para[] {
  const paras: Para[] = [[]];
  const stack: { tag: string; look: Look }[] = [];
  let look: Look = { b: false, i: false, u: false };
  let skip = 0;
  let pre = 0;
  const cur = () => paras[paras.length - 1]!;
  const breakPara = () => { if (cur().length) paras.push([]); };
  /** Add text; `literal` text (a tab between cells, a line break) is kept as it is, not collapsed as white space. */
  const add = (raw: string, literal = false) => {
    let text = literal ? raw : decode(raw);
    if (!pre && !literal) {
      text = text.replace(/[ \t\r\n\f]+/g, ' ');
      const last = cur()[cur().length - 1];
      // white space at the start of a line, or after white space, does not show
      if (text.startsWith(' ') && (!last || last.text.endsWith(' '))) text = text.slice(1);
    }
    if (!text) return;
    const last = cur()[cur().length - 1];
    if (last && !!last.b === look.b && !!last.i === look.i && !!last.u === look.u) last.text += text;
    else cur().push({ text, ...(look.b ? { b: true } : {}), ...(look.i ? { i: true } : {}), ...(look.u ? { u: true } : {}) });
  };
  const body = html.replace(/<!--[\s\S]*?-->/g, '').replace(/<!doctype[^>]*>/gi, '');
  const tagRe = /<(\/?)([a-zA-Z][a-zA-Z0-9:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*)>/g;
  let at = 0;
  for (let m = tagRe.exec(body); m; m = tagRe.exec(body)) {
    if (m.index > at && !skip) add(body.slice(at, m.index));
    at = tagRe.lastIndex;
    const closing = m[1] === '/', tag = m[2]!.toLowerCase(), attrs = m[3] ?? '';
    const selfClosed = /\/\s*$/.test(attrs) || ['br', 'img', 'hr', 'meta', 'link', 'input', 'col', 'wbr'].includes(tag);
    if (SKIPPED.has(tag)) { skip = Math.max(0, skip + (closing ? -1 : 1)); continue; }
    if (skip) continue;
    if (tag === 'br') {
      // a line break inside a paragraph, unless nothing has been said on this line yet
      if (cur().length) add('\u000b', true); else paras.push([]);
      continue;
    }
    if (BLOCKS.has(tag)) breakPara();
    if (tag === 'td' || tag === 'th') { if (closing && cur().length) add('\t', true); }
    if (selfClosed) continue;
    if (!closing) {
      stack.push({ tag, look });
      look = lookOf(tag, attrs, look);
      if (tag === 'pre') pre++;
    } else {
      // close back to the matching opener; a stray closing tag is ignored
      const i = stack.map(s => s.tag).lastIndexOf(tag);
      if (i < 0) continue;
      for (const s of stack.slice(i)) if (s.tag === 'pre') pre = Math.max(0, pre - 1);
      look = stack[i]!.look;
      stack.length = i;
    }
  }
  if (at < body.length && !skip) add(body.slice(at));
  // trailing white space on a line does not show, and a last empty paragraph is no paragraph
  for (const p of paras) {
    const last = p[p.length - 1];
    if (last) last.text = last.text.replace(/[ \t\u000b]+$/, '');
    if (last && !last.text) p.pop();
  }
  while (paras.length > 1 && !paras[paras.length - 1]!.length) paras.pop();
  return paras;
}

/** Whether any of the text is bold, italic or underlined (otherwise plain typing does). */
export const hasLook = (paras: Para[]) => paras.some(p => p.some(s => s.b || s.i || s.u));
