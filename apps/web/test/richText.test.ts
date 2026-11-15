// Text on the clipboard: what HTML from other apps becomes, and what the editor writes for them.
import { describe, expect, test } from 'vitest';
import { hasLook, htmlToParagraphs, paragraphsToHtml, plainOf, type Para } from '../src/ui/richText.ts';

/** Paragraphs as text, with [b], [i], [u] before a stretch that has them. */
const show = (paras: Para[]) => paras.map(p => p.map(s => `${s.b ? '[b]' : ''}${s.i ? '[i]' : ''}${s.u ? '[u]' : ''}${s.text}`).join('|'));

describe('HTML to paragraphs', () => {
  test('blocks are paragraphs, <br> a line break inside one, and white space collapses as a browser shows it', () => {
    expect(show(htmlToParagraphs('<p>First   line</p>\n<p>Second<br>third</p>'))).toEqual(['First line', 'Second\u000bthird']);
    expect(show(htmlToParagraphs('<div>one</div><div>two</div>  <div>  three  </div>'))).toEqual(['one', 'two', 'three']);
    expect(show(htmlToParagraphs('just text'))).toEqual(['just text']);
    expect(show(htmlToParagraphs(''))).toEqual(['']);
    expect(show(htmlToParagraphs('<ul><li>a</li><li>b</li></ul>'))).toEqual(['a', 'b']);
  });

  test('bold, italic and underline from tags, nested and closed in any order', () => {
    expect(show(htmlToParagraphs('<p>plain <b>bold <i>both</i></b> <u>under</u> <strong>strong</strong> <em>em</em></p>')))
      .toEqual(['plain |[b]bold |[b][i]both| |[u]under| |[b]strong| |[i]em']);
    expect(show(htmlToParagraphs('<b>a<i>b</b>c</i>d'))).toEqual(['[b]a|[b][i]b|cd']);        // badly nested: closing b ends both
    expect(show(htmlToParagraphs('a</b>b'))).toEqual(['ab']);                                  // a stray closing tag
  });

  test('bold, italic and underline from inline styles, as Google Docs and Word write them', () => {
    const docs = '<meta charset="utf-8"><b style="font-weight:normal;" id="docs-internal-guid-1"><span style="font-weight:700;font-style:italic;">Heavy</span><span style="font-weight:400;text-decoration:underline;"> under</span></b>';
    expect(show(htmlToParagraphs(docs))).toEqual(['[b][i]Heavy|[u] under']);
    expect(show(htmlToParagraphs("<span style='font-weight: bold'>x</span><span style=\"FONT-WEIGHT: 300\">y</span>"))).toEqual(['[b]x|y']);
    expect(show(htmlToParagraphs('<h2>Title</h2><p>body</p>'))).toEqual(['[b]Title', 'body']);
  });

  test('entities are decoded; text that looks like markup after decoding stays text', () => {
    expect(show(htmlToParagraphs('a &amp; b &lt;i&gt; &quot;q&quot; &#233; &#x1F600; &nbsp;x &unknown;'))).toEqual(['a & b <i> "q" é 😀  x &unknown;']);
  });

  test('scripts, styles, comments and the head never become text; attributes with > in them do not end a tag early', () => {
    const html = '<html><head><title>T</title><style>p { color: red }</style></head><body><!-- note --><script>alert("x")</script><p title="a > b">shown</p><img src=x onerror=alert(1)></body></html>';
    expect(show(htmlToParagraphs(html))).toEqual(['shown']);
  });

  test('preformatted text keeps its spaces; table cells are separated by tabs', () => {
    expect(show(htmlToParagraphs('<pre>a   b</pre>'))).toEqual(['a   b']);
    expect(show(htmlToParagraphs('<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>'))).toEqual(['a\tb', 'c\td']);
  });
});

describe('paragraphs to the clipboard', () => {
  const paras: Para[] = [[{ text: 'Hello ' }, { text: 'big', b: true }, { text: ' <world> & more', i: true, u: true }], [{ text: 'line\u000bbreak' }], []];

  test('plain text: a line per paragraph, a line break a new line', () => {
    expect(plainOf(paras)).toBe('Hello big <world> & more\nline\nbreak\n');
  });

  test('HTML is escaped, and reading it back gives the same paragraphs', () => {
    const html = paragraphsToHtml(paras);
    expect(html).toBe('<p>Hello <b>big</b><i><u> &lt;world&gt; &amp; more</u></i></p><p>line<br>break</p><p></p>');
    expect(show(htmlToParagraphs(html))).toEqual(['Hello |[b]big|[i][u] <world> & more', 'line\u000bbreak']);
  });

  test('hasLook says whether typing plain text would lose anything', () => {
    expect(hasLook(paras)).toBe(true);
    expect(hasLook([[{ text: 'plain' }]])).toBe(false);
  });
});
