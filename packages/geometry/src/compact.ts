// packages/geometry/src/compact.ts  -- the text form the generated preset tables are written in
//
// One string per list, tokens separated by single spaces, list items by '; ':
//   av, gd   guides: 'name fmla; name fmla'                            ('adj val 16667; a pin 0 adj 50000')
//   ah       'xy gdRefX minX maxX gdRefY minY maxY x y' or 'polar gdRefR minR maxR gdRefAng minAng maxAng x y',
//            with '-' for an attribute that is absent
//   cxn      'ang x y'
//   rect     'l t r b'
//   path     one string per path: its attributes that differ from the defaults (w=, h=, fill=, stroke=0,
//            extrusionOk=0), then the commands M x y, L x y, A wR hR stAng swAng, Q x1 y1 x y, C x1 y1 x2 y2 x y, Z
import type { AdjPoint, AdjustHandle, ConnectionSite, Guide, PathDef, PathDefCommand, PathFill, ShapeGeometry } from './types.ts';

export interface CompactGeometry {
  readonly av?: string;
  readonly gd?: string;
  readonly ah?: string;
  readonly cxn?: string;
  readonly rect?: string;
  readonly path: readonly string[];
}

const XY = ['gdRefX', 'minX', 'maxX', 'gdRefY', 'minY', 'maxY'] as const;
const POLAR = ['gdRefR', 'minR', 'maxR', 'gdRefAng', 'minAng', 'maxAng'] as const;
const LIST = '; ';

const items = (s: string | undefined) => (s ? s.split(LIST) : []);

function token(s: string): string {
  if (!/^[^\s;=]+$/.test(s) || s === '-') throw new Error(`${JSON.stringify(s)} cannot be written in the compact form`);
  return s;
}

function encodeGuides(guides: readonly Guide[]): string | undefined {
  if (guides.length === 0) return undefined;
  return guides.map(g => {
    if (g.fmla.includes(';') || g.fmla !== g.fmla.trim()) throw new Error(`formula ${JSON.stringify(g.fmla)} cannot be written in the compact form`);
    return `${token(g.name)} ${g.fmla}`;
  }).join(LIST);
}

function decodeGuides(s: string | undefined): Guide[] {
  return items(s).map(item => {
    const space = item.indexOf(' ');
    return { name: item.slice(0, space), fmla: item.slice(space + 1) };
  });
}

const pt = (p: AdjPoint) => `${token(p.x)} ${token(p.y)}`;

function encodeCommand(c: PathDefCommand): string {
  switch (c.op) {
    case 'moveTo': return `M ${pt(c.pt)}`;
    case 'lnTo': return `L ${pt(c.pt)}`;
    case 'arcTo': return `A ${token(c.wR)} ${token(c.hR)} ${token(c.stAng)} ${token(c.swAng)}`;
    case 'quadBezTo': return `Q ${c.pts.map(pt).join(' ')}`;
    case 'cubicBezTo': return `C ${c.pts.map(pt).join(' ')}`;
    case 'close': return 'Z';
  }
}

function encodePath(p: PathDef): string {
  const parts: string[] = [];
  if (p.w) parts.push(`w=${p.w}`);
  if (p.h) parts.push(`h=${p.h}`);
  if (p.fill !== 'norm') parts.push(`fill=${p.fill}`);
  if (!p.stroke) parts.push('stroke=0');
  if (!p.extrusionOk) parts.push('extrusionOk=0');
  return [...parts, ...p.commands.map(encodeCommand)].join(' ');
}

function decodePath(s: string): PathDef {
  const t = s === '' ? [] : s.split(' ');
  let i = 0;
  const attrs = new Map<string, string>();
  for (; i < t.length && t[i]!.includes('='); i++) {
    const [k = '', v = ''] = t[i]!.split('=');
    attrs.set(k, v);
  }
  const next = () => t[i++]!;
  const point = (): AdjPoint => ({ x: next(), y: next() });
  const commands: PathDefCommand[] = [];
  while (i < t.length) {
    const op = next();
    if (op === 'M') commands.push({ op: 'moveTo', pt: point() });
    else if (op === 'L') commands.push({ op: 'lnTo', pt: point() });
    else if (op === 'A') commands.push({ op: 'arcTo', wR: next(), hR: next(), stAng: next(), swAng: next() });
    else if (op === 'Q') commands.push({ op: 'quadBezTo', pts: [point(), point()] });
    else if (op === 'C') commands.push({ op: 'cubicBezTo', pts: [point(), point(), point()] });
    else if (op === 'Z') commands.push({ op: 'close' });
    else throw new Error(`unknown path command ${JSON.stringify(op)}`);
  }
  return {
    w: Number(attrs.get('w') ?? 0),
    h: Number(attrs.get('h') ?? 0),
    fill: (attrs.get('fill') ?? 'norm') as PathFill,
    stroke: attrs.get('stroke') !== '0',
    extrusionOk: attrs.get('extrusionOk') !== '0',
    commands,
  };
}

function encodeHandle(hd: AdjustHandle): string {
  const values = hd.kind === 'xy' ? XY.map(k => hd[k]) : POLAR.map(k => hd[k]);
  return [hd.kind, ...values.map(v => (v === undefined ? '-' : token(v))), pt(hd.pos)].join(' ');
}

function decodeHandle(s: string): AdjustHandle {
  const [kind, ...t] = s.split(' ');
  const attrs: Record<string, string> = {};
  (kind === 'xy' ? XY : POLAR).forEach((k, i) => {
    if (t[i] !== '-') attrs[k] = t[i]!;
  });
  const pos = { x: t[6]!, y: t[7]! };
  return kind === 'xy' ? { kind: 'xy', ...attrs, pos } : { kind: 'polar', ...attrs, pos };
}

export function encodeGeometry(g: ShapeGeometry): CompactGeometry {
  const c: { -readonly [K in keyof CompactGeometry]: CompactGeometry[K] } = { path: g.pathLst.map(encodePath) };
  const av = encodeGuides(g.avLst), gd = encodeGuides(g.gdLst);
  if (av) c.av = av;
  if (gd) c.gd = gd;
  if (g.ahLst.length) c.ah = g.ahLst.map(encodeHandle).join(LIST);
  if (g.cxnLst.length) c.cxn = g.cxnLst.map(s => `${token(s.ang)} ${pt(s.pos)}`).join(LIST);
  if (g.rect) c.rect = [g.rect.l, g.rect.t, g.rect.r, g.rect.b].map(token).join(' ');
  return c;
}

export function decodeGeometry(c: CompactGeometry): ShapeGeometry {
  const rect = c.rect?.split(' ');
  return {
    avLst: decodeGuides(c.av),
    gdLst: decodeGuides(c.gd),
    ahLst: items(c.ah).map(decodeHandle),
    cxnLst: items(c.cxn).map((s): ConnectionSite => {
      const [ang = '', x = '', y = ''] = s.split(' ');
      return { ang, pos: { x, y } };
    }),
    rect: rect ? { l: rect[0]!, t: rect[1]!, r: rect[2]!, b: rect[3]! } : null,
    pathLst: c.path.map(decodePath),
  };
}
