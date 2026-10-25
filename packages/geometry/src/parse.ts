// packages/geometry/src/parse.ts  -- read a:custGeom, and the Annex D preset definitions, which have the same children
import { childElements, getAttr, type XmlElement } from '@pptx/xml';
import { GeometryError } from './errors.ts';
import type { AdjPoint, AdjustHandle, ConnectionSite, Guide, PathDef, PathDefCommand, PathFill, ShapeGeometry } from './types.ts';

// DrawingML main, Transitional and Strict. The Annex D files put each list in the Transitional namespace.
const A_NS = new Set(['http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main']);
const FILLS = new Set<string>(['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess'] satisfies PathFill[]);

const children = (el: XmlElement, local: string) => childElements(el).filter(c => c.local === local && A_NS.has(c.ns));
const child = (el: XmlElement, local: string): XmlElement | undefined => children(el, local)[0];

// Guide names, coordinates and angles are xsd:token based, so surrounding whitespace is not part of the value.
// Formulas are xsd:string and are kept as written.
const optional = (el: XmlElement, name: string) => getAttr(el, name)?.trim();
function required(el: XmlElement, name: string): string {
  const value = optional(el, name);
  if (value === undefined) throw new GeometryError(`<${el.local}> has no ${name} attribute`);
  return value;
}

function bool(el: XmlElement, name: string, otherwise: boolean): boolean {
  const value = optional(el, name);
  if (value === undefined) return otherwise;
  if (value === 'true' || value === '1') return true;
  if (value === 'false' || value === '0') return false;
  throw new GeometryError(`<${el.local}> ${name}="${value}" is not a boolean`);
}

/** w and h of a path: ST_PositiveCoordinate, 0 when absent. */
function extent(el: XmlElement, name: string): number {
  const value = optional(el, name);
  if (value === undefined) return 0;
  if (!/^\+?[0-9]+$/.test(value)) throw new GeometryError(`<path> ${name}="${value}" is not a positive coordinate`);
  return Number(value);
}

function guides(list: XmlElement | undefined): Guide[] {
  if (!list) return [];
  return children(list, 'gd').map(gd => {
    const fmla = getAttr(gd, 'fmla');
    if (fmla === undefined) throw new GeometryError('<gd> has no fmla attribute');
    return { name: required(gd, 'name'), fmla };
  });
}

const point = (el: XmlElement): AdjPoint => ({ x: required(el, 'x'), y: required(el, 'y') });

function pos(el: XmlElement): AdjPoint {
  const p = child(el, 'pos');
  if (!p) throw new GeometryError(`<${el.local}> has no pos`);
  return point(p);
}

function points(el: XmlElement, count: number): AdjPoint[] {
  const pts = children(el, 'pt');
  if (pts.length !== count) throw new GeometryError(`<${el.local}> needs ${count} pt, not ${pts.length}`);
  return pts.map(point);
}

/** The attributes of `names` that are present, as written, in the order given. */
function present<K extends string>(el: XmlElement, names: readonly K[]): Partial<Record<K, string>> {
  const out: Partial<Record<K, string>> = {};
  for (const n of names) {
    const value = optional(el, n);
    if (value !== undefined) out[n] = value;
  }
  return out;
}

function handle(el: XmlElement): AdjustHandle | null {
  if (el.local === 'ahXY') return { kind: 'xy', ...present(el, ['gdRefX', 'minX', 'maxX', 'gdRefY', 'minY', 'maxY']), pos: pos(el) };
  if (el.local === 'ahPolar') return { kind: 'polar', ...present(el, ['gdRefR', 'minR', 'maxR', 'gdRefAng', 'minAng', 'maxAng']), pos: pos(el) };
  return null;
}

function command(el: XmlElement): PathDefCommand | null {
  switch (el.local) {
    case 'moveTo': return { op: 'moveTo', pt: points(el, 1)[0]! };
    case 'lnTo': return { op: 'lnTo', pt: points(el, 1)[0]! };
    case 'arcTo':
      return { op: 'arcTo', wR: required(el, 'wR'), hR: required(el, 'hR'), stAng: required(el, 'stAng'), swAng: required(el, 'swAng') };
    case 'quadBezTo': {
      const [a, b] = points(el, 2);
      return { op: 'quadBezTo', pts: [a!, b!] };
    }
    case 'cubicBezTo': {
      const [a, b, c] = points(el, 3);
      return { op: 'cubicBezTo', pts: [a!, b!, c!] };
    }
    case 'close': return { op: 'close' };
    default: return null;
  }
}

function path(el: XmlElement): PathDef {
  const fill = optional(el, 'fill') ?? 'norm';
  if (!FILLS.has(fill)) throw new GeometryError(`<path> fill="${fill}" is not a path fill mode`);
  const commands: PathDefCommand[] = [];
  for (const c of childElements(el)) {
    const cmd = A_NS.has(c.ns) ? command(c) : null;
    if (cmd) commands.push(cmd);
  }
  return {
    w: extent(el, 'w'),
    h: extent(el, 'h'),
    fill: fill as PathFill,
    stroke: bool(el, 'stroke', true),
    extrusionOk: bool(el, 'extrusionOk', true),   // the schema's default; the prose of 20.1.9.15 says false, the XSD true
    commands,
  };
}

/**
 * The geometry under an a:custGeom element, or under one shape element of the Annex D preset files: avLst, gdLst,
 * ahLst, cxnLst, rect and pathLst children, each optional. Elements in other namespaces are ignored. Missing
 * required attributes throw a GeometryError; guide names are not checked here but when the geometry is evaluated.
 */
export function parseCustomGeometry(custGeom: XmlElement): ShapeGeometry {
  const ahLst = child(custGeom, 'ahLst');
  const cxnLst = child(custGeom, 'cxnLst');
  const rect = child(custGeom, 'rect');
  const pathLst = child(custGeom, 'pathLst');
  return {
    avLst: guides(child(custGeom, 'avLst')),
    gdLst: guides(child(custGeom, 'gdLst')),
    ahLst: ahLst ? childElements(ahLst).filter(c => A_NS.has(c.ns)).map(handle).filter(h => h !== null) : [],
    cxnLst: cxnLst ? children(cxnLst, 'cxn').map((c): ConnectionSite => ({ ang: required(c, 'ang'), pos: pos(c) })) : [],
    rect: rect ? { l: required(rect, 'l'), t: required(rect, 't'), r: required(rect, 'r'), b: required(rect, 'b') } : null,
    pathLst: pathLst ? children(pathLst, 'path').map(path) : [],
  };
}
