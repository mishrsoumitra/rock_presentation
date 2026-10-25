// packages/charts/src/labels.ts  -- data labels (c:dLbls, c:dLbl): what each says and where it sits
//
// A point's c:dLbl overrides its series' c:dLbls, which overrides the chart group's, property by property. The
// text joins series name, category, value, percentage and bubble size (in that order) with the separator; rich
// text (c:tx/c:rich) replaces it, with fields such as [VALUE] and [CELLRANGE] filled in from the cached data.
import type { XmlElement } from '@pptx/xml';
import { attr, child, coord } from '@pptx/model';
import type { Primitive } from '@pptx/render';
import type { Chart, Group, Layout, Series } from './model.ts';
import { PT } from './axes.ts';
import { pathPrim, rect } from './draw.ts';
import { categoryText } from './legend.ts';
import { formatValue } from './numfmt.ts';
import type { LabelSite } from './series.ts';
import { effectsOf, fillOf, NONE, strokeOf, type Env } from './style.ts';
import { block, isEmpty, placeAt, plainBody, richBody, textPrim, textStyle } from './text.ts';
import { flag, kid, kidsNamed, numOf, valOf } from './xml.ts';

/** What one c:dLbls or c:dLbl element says; absent properties fall through to the next element in the chain. */
interface Spec {
  deleted?: boolean;
  val?: boolean;
  cat?: boolean;
  ser?: boolean;
  pct?: boolean;
  size?: boolean;
  key?: boolean;
  pos?: string;
  numFmt?: { code: string; linked: boolean };
  sep?: string;
  txPr?: XmlElement;
  spPr?: XmlElement;
  rich?: XmlElement;
  layout?: Layout;
}

/** A CT_Boolean child, or undefined when absent. */
const opt = (el: XmlElement, name: string): boolean | undefined => (kid(el, name) ? flag(el, name, true) : undefined);

function specOf(el: XmlElement | undefined): Spec {
  if (!el) return {};
  const out: Spec = {};
  const set = <K extends keyof Spec>(k: K, v: Spec[K] | undefined) => {
    if (v !== undefined) out[k] = v;
  };
  set('deleted', opt(el, 'c:delete'));
  set('val', opt(el, 'c:showVal'));
  set('cat', opt(el, 'c:showCatName'));
  set('ser', opt(el, 'c:showSerName'));
  set('pct', opt(el, 'c:showPercent'));
  set('size', opt(el, 'c:showBubbleSize'));
  set('key', opt(el, 'c:showLegendKey'));
  set('pos', valOf(el, 'c:dLblPos'));
  const nf = kid(el, 'c:numFmt');
  const linked = attr(nf, 'sourceLinked');
  if (nf) out.numFmt = { code: attr(nf, 'formatCode') ?? 'General', linked: linked === '1' || linked === 'true' };
  const sep = kid(el, 'c:separator');
  if (sep) out.sep = sep.children.map(c => (c.kind === 'text' ? c.value : '')).join('');
  set('txPr', kid(el, 'c:txPr'));
  set('spPr', kid(el, 'c:spPr'));
  set('rich', child(kid(el, 'c:tx'), 'c:rich'));
  const m = kid(kid(el, 'c:layout'), 'c:manualLayout');
  if (m) {
    const mode = (n: string) => (valOf(m, n) === 'edge' ? 'edge' as const : 'factor' as const);
    const layout: Layout = {
      target: 'outer', xMode: mode('c:xMode'), yMode: mode('c:yMode'), wMode: 'factor', hMode: 'factor',
    };
    const x = numOf(m, 'c:x');
    const y = numOf(m, 'c:y');
    if (x !== undefined) layout.x = x;
    if (y !== undefined) layout.y = y;
    out.layout = layout;
  }
  return out;
}

/** The specs read so far, per c:dLbls element: its own, and its c:dLbl children's by point index. */
type SpecCache = Map<XmlElement, { all: Spec; points: Map<number, Spec> }>;

/** The label settings of point i: its own c:dLbl, then the series' c:dLbls, then the group's. */
function specFor(g: Group, s: Series, i: number, cache: SpecCache): Spec[] {
  const read = (el: XmlElement | undefined) => {
    if (!el) return undefined;
    let c = cache.get(el);
    if (!c) {
      const points = new Map<number, Spec>();
      for (const d of kidsNamed(el, 'c:dLbl')) {
        const idx = Number.parseInt(valOf(d, 'c:idx') ?? '-1', 10);
        if (idx >= 0) points.set(idx, specOf(d));
      }
      c = { all: specOf(el), points };
      cache.set(el, c);
    }
    return c;
  };
  const ser = read(s.dLbls);
  const grp = read(g.dLbls);
  return [ser?.points.get(i), ser?.all, grp?.points.get(i), grp?.all].filter((x): x is Spec => x !== undefined);
}

/** The nearest value of a property along a chain of specs. */
const first = <K extends keyof Spec>(chain: Spec[], k: K): Spec[K] | undefined =>
  chain.find(c => c[k] !== undefined)?.[k];

/** Where labels go when nothing says: Office's defaults per chart type. */
function defaultPos(g: Group): string {
  switch (g.kind) {
    case 'bar': return g.grouping === 'stacked' || g.grouping === 'percentStacked' ? 'ctr' : 'outEnd';
    case 'pie': case 'ofPie': return 'bestFit';
    case 'doughnut': case 'bubble': case 'area': return 'ctr';
    case 'radar': return 't';
    default: return 'r';
  }
}

/** The format of a point's value: the label's own, else the point's or the series' from the cache. */
function valueFormat(g: Group, s: Series, i: number, nf: Spec['numFmt']): string {
  if (nf && !nf.linked) return nf.code;
  const v = g.kind === 'scatter' || g.kind === 'bubble' ? s.y ?? s.val : s.val;
  return v?.formats[i] ?? v?.formatCode ?? 'General';
}

/** Draw every data label the chart's series ask for. */
export function drawLabels(
  env: Env, chart: Chart, sites: LabelSite[], box: { w: number; h: number }, fontSize: number,
  paths: Primitive[], texts: Primitive[],
): void {
  const cache: SpecCache = new Map();
  const em = fontSize * PT;
  const fmt = (v: number | undefined, code: string) => (v !== undefined ? formatValue(v, code, env.date1904) : '');
  // beyond a few thousand points labels only pile up on each other
  for (const site of sites.slice(0, 5000)) {
    const { g, s, i } = site;
    const chain = specFor(g, s, i, cache);
    if (!chain.length || first(chain, 'deleted')) continue;
    const rich = first(chain, 'rich');
    const [ser, cat, val, pct, size] = (['ser', 'cat', 'val', 'pct', 'size'] as const).map(k => first(chain, k));
    if (!rich && !(ser || cat || val || pct || size)) continue;
    const nf = first(chain, 'numFmt');
    const valueText = fmt(site.value, valueFormat(g, s, i, nf));
    // percentages take the label's format only when it is a percentage format
    const pctText = fmt(site.percent, nf && !nf.linked && nf.code.includes('%') ? nf.code : '0%');
    const sizeText = fmt(site.size, s.size?.formatCode ?? 'General');
    // a scatter chart's "category" is the x value
    const xv = s.x?.pts[i];
    const catText = g.kind === 'scatter' || g.kind === 'bubble'
      ? (xv !== undefined && s.x ? fmt(xv, s.x.formatCode) : String(i + 1))
      : categoryText(env, g.series.find(o => o.cat) ?? s, i);
    const txPrs = chain.map(c => c.txPr).filter((t): t is XmlElement => t !== undefined);
    const d = { size: fontSize, bold: false };
    const style = textStyle(env, [...txPrs, chart.txPr], d);
    let body;
    if (rich) {
      const fields: Record<string, string | undefined> = {
        VALUE: valueText, YVALUE: valueText, SERIESNAME: s.name, CATEGORYNAME: catText, XVALUE: catText,
        PERCENTAGE: pctText,
        // a cell-range label without a cached value shows nothing rather than its placeholder
        BUBBLESIZE: sizeText, CELLRANGE: s.labelRange?.[i] ?? '',
      };
      body = richBody(env, rich, [...txPrs, chart.txPr], d, type => fields[type]).body;
    } else {
      const parts = [ser && s.name, cat && catText, val && valueText, pct && pctText, size && sizeText];
      body = plainBody(style, parts.filter(p => p).join(first(chain, 'sep') ?? ', '));
    }
    if (isEmpty(body)) continue;
    const b = block(env, body);
    const key = first(chain, 'key') ? style.size * PT * 0.6 : 0;
    const keyGap = key ? style.size * PT * 0.25 : 0;
    const w = b.w + key + keyGap;
    const h = b.h;
    let { x, y } = place(site, first(chain, 'pos') ?? defaultPos(g), w, h, em * 0.25);
    const m = first(chain, 'layout');
    if (m) {
      if (m.x !== undefined) x = m.xMode === 'edge' ? m.x * box.w + w / 2 : x + m.x * box.w;
      if (m.y !== undefined) y = m.yMode === 'edge' ? m.y * box.h + h / 2 : y + m.y * box.h;
    }
    // labels stay inside the chart area
    x = Math.max(w / 2, Math.min(box.w - w / 2, x));
    y = Math.max(h / 2, Math.min(box.h - h / 2, y));
    const spPr = first(chain, 'spPr');
    if (spPr) {
      const bodyPr = child(first(chain, 'txPr'), 'a:bodyPr');
      const lIns = coord(attr(bodyPr, 'lIns'), 38100);
      const tIns = coord(attr(bodyPr, 'tIns'), 19050);
      const r = rect({ x: x - w / 2 - lIns, y: y - h / 2 - tIns, w: w + 2 * lIns, h: h + 2 * tIns });
      const bg = pathPrim(r, fillOf(env, spPr, () => NONE), strokeOf(env, spPr, null), effectsOf(env, spPr));
      if (bg) paths.push(bg);
    }
    if (key) {
      const k = pathPrim(rect({ x: x - w / 2, y: y - key / 2, w: key, h: key }), site.style.fill, site.style.line);
      if (k) paths.push(k);
    }
    texts.push(textPrim(b, placeAt(x - w / 2 + key + keyGap, y - h / 2, 0, 0)));
  }
}

/** The centre of a label of size w x h at position `pos` for its site. */
function place(site: LabelSite, pos: string, w: number, h: number, gap: number): { x: number; y: number } {
  const bar = site.bar;
  if (bar) {
    const r = bar.r;
    const centre = { x: r.x + r.w / 2, y: r.y + r.h / 2 };
    if (pos !== 'outEnd' && pos !== 'inEnd' && pos !== 'inBase') return centre;
    // along the bar from its base to its end, which grows in the direction `sign` on the page
    const upright = bar.dir === 'up' || bar.dir === 'down';
    const sign = bar.dir === 'down' || bar.dir === 'right' ? 1 : -1;
    const [lo, hi] = upright ? [r.y, r.y + r.h] : [r.x, r.x + r.w];
    const [base, end] = sign > 0 ? [lo, hi] : [hi, lo];
    const off = gap + (upright ? h : w) / 2;
    const p = pos === 'outEnd' ? end + sign * off : pos === 'inEnd' ? end - sign * off : base + sign * off;
    return upright ? { x: centre.x, y: p } : { x: p, y: centre.y };
  }
  const sl = site.slice;
  if (sl) {
    const c = Math.cos(sl.mid);
    const sn = Math.sin(sl.mid);
    const reach = Math.abs((w / 2) * c) + Math.abs((h / 2) * sn);
    const k = sl.ry / Math.max(1, sl.r);
    const along = (d: number) => ({ x: sl.cx + d * c, y: sl.cy + d * sn * k });
    const outside = () => {
      const p = along(sl.r + gap);
      return { x: p.x + (w / 2) * c, y: p.y + (h / 2) * sn };
    };
    switch (pos) {
      case 'outEnd': return outside();
      case 'inEnd': return along(Math.max(sl.inner, sl.r - gap - reach));
      case 'bestFit': {
        // inside when the label fits between the slice's edges, outside otherwise
        const d = sl.inner ? (sl.inner + sl.r) / 2 : sl.r * 0.62;
        const span = site.percent !== undefined ? site.percent * Math.PI * 2 : Math.PI;
        const room = span >= Math.PI ? Infinity : 2 * d * Math.sin(span / 2);
        return room >= Math.max(w, h) * 0.9 && sl.r - d >= reach * 0.8 ? along(d) : outside();
      }
      default: return along(sl.inner ? (sl.inner + sl.r) / 2 : sl.r / 2);
    }
  }
  const r = site.radius ?? 0;
  switch (pos) {
    case 't': case 'outEnd': case 'inEnd': return { x: site.x, y: site.y - r - gap - h / 2 };
    case 'b': case 'inBase': return { x: site.x, y: site.y + r + gap + h / 2 };
    case 'l': return { x: site.x - r - gap - w / 2, y: site.y };
    case 'ctr': return { x: site.x, y: site.y };
    default: return { x: site.x + r + gap + w / 2, y: site.y };
  }
}
