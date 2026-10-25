// tools/inspect/inspect.ts  -- the deck inspector (milestone M2): every shape's resolved properties
//
//   node tools/inspect/inspect.ts deck.pptx [--slide N] [--layouts] [--json]
//
// Values are in points and PowerPoint's own terms (Shape.Left, Shape.Width, Font.Name, Font.Size, Fill.ForeColor),
// so the output can be set against what PowerPoint's object model reports for the same file (the property oracle).
import { readFileSync } from 'node:fs';
import { OpcPackage } from '../../packages/opc/src/index.ts';
import { PresentationDocument } from '../../packages/model/src/index.ts';
import { resolveSlide, toHex, type Paint, type ResolvedShape } from '../../packages/resolve/src/index.ts';

const pt = (emu: number) => Math.round((emu / 12700) * 100) / 100;

function paint(p: Paint | undefined): string | Record<string, unknown> {
  if (!p) return 'none';
  switch (p.kind) {
    case 'none': return 'none';
    case 'solid': return toHex(p.color, p.color.a < 1);
    case 'image': return { picture: p.image.part ?? p.image.externalUrl ?? '' };
    case 'pattern': return { pattern: p.preset, fg: toHex(p.fg), bg: toHex(p.bg) };
    case 'linear': case 'path': return { gradient: p.kind, stops: p.stops.map(s => `${toHex(s.color)}@${Math.round(s.pos * 100)}%`) };
    default: return p.kind;
  }
}

export interface ShapeInfo {
  id: number; name: string; layer: string; kind: string; placeholder?: string;
  left: number; top: number; width: number; height: number; rotation: number; flipH: boolean; flipV: boolean;
  geometry?: string; fill: unknown; line?: { weight: number; color: unknown; dash: string };
  graphic?: string;
  text?: { anchor: string; autofit: string; paragraphs: { text: string; level: number; align: string; bullet?: string; font: string; size: number; bold: boolean; italic: boolean; color: unknown }[] };
  children?: ShapeInfo[];
}

export function shapeInfo(s: ResolvedShape): ShapeInfo {
  const x = s.xfrm;
  const info: ShapeInfo = {
    id: s.id, name: s.name, layer: s.layer, kind: s.kind,
    left: pt(x.x), top: pt(x.y), width: pt(x.cx), height: pt(x.cy), rotation: x.rot / 60000, flipH: x.flipH, flipV: x.flipV,
    fill: paint(s.fill),
  };
  if (s.placeholder) info.placeholder = `${s.placeholder.type}${s.placeholder.hasIdx ? `#${s.placeholder.idx}` : ''}`;
  if (s.geometry) info.geometry = s.geometry.kind === 'preset' ? s.geometry.prst : 'custom';
  if (s.line) info.line = { weight: pt(s.line.width), color: paint(s.line.paint), dash: s.line.dash ? s.line.dash.join(' ') : 'solid' };
  if (s.graphic) info.graphic = s.graphic.kind;
  if (s.text) {
    const b = s.text.body;
    info.text = {
      anchor: b.anchor, autofit: b.autofit.kind === 'norm' ? `shrink ${Math.round(b.autofit.fontScale * 100)}%` : b.autofit.kind,
      paragraphs: s.text.paragraphs.map(p => {
        const r = p.runs.find(x => x.kind !== 'break')?.props ?? p.endProps;
        return {
          text: p.runs.map(x => (x.kind === 'break' ? '\v' : x.text)).join(''), level: p.props.level, align: p.props.align,
          ...(p.props.bullet ? { bullet: p.props.bullet.kind === 'char' ? p.props.bullet.char ?? '' : p.props.bullet.kind } : {}),
          font: r.latin, size: r.size, bold: r.bold, italic: r.italic, color: paint(r.fill),
        };
      }),
    };
  }
  if (s.children.length) info.children = s.children.map(shapeInfo);
  return info;
}

function print(info: ShapeInfo, depth: number): void {
  const pad = '  '.repeat(depth);
  const where = `${info.left},${info.top} ${info.width}x${info.height}${info.rotation ? ` rot ${info.rotation}` : ''}${info.flipH ? ' flipH' : ''}${info.flipV ? ' flipV' : ''}`;
  const look = `${info.geometry ?? ''} fill ${JSON.stringify(info.fill)}${info.line ? ` line ${info.line.weight}pt ${JSON.stringify(info.line.color)} ${info.line.dash}` : ''}`;
  console.log(`${pad}[${info.layer}] ${info.kind} #${info.id} "${info.name}"${info.placeholder ? ` ph:${info.placeholder}` : ''}  ${where}  ${look}${info.graphic ? ` (${info.graphic})` : ''}`);
  for (const p of info.text?.paragraphs ?? []) {
    console.log(`${pad}    ${p.bullet !== undefined ? `${p.bullet} ` : ''}lvl ${p.level} ${p.align} | ${p.font} ${p.size}pt${p.bold ? ' bold' : ''}${p.italic ? ' italic' : ''} ${JSON.stringify(p.color)} | ${JSON.stringify(p.text)}`);
  }
  for (const c of info.children ?? []) print(c, depth + 1);
}

function main(): void {
  const args = process.argv.slice(2);
  const file = args.find(a => !a.startsWith('--') && !/^\d+$/.test(a));
  if (!file) {
    console.error('usage: node tools/inspect/inspect.ts deck.pptx [--slide N] [--layouts] [--json]');
    process.exitCode = 2;
    return;
  }
  const at = args.indexOf('--slide');
  const only = at >= 0 ? Number(args[at + 1]) : undefined;
  const doc = new PresentationDocument(OpcPackage.open(new Uint8Array(readFileSync(file))));
  const parts = doc.slideParts().map((p, i) => ({ part: p, label: `slide ${i + 1}` })).filter((_, i) => only === undefined || i + 1 === only);
  if (args.includes('--layouts')) {
    for (const m of doc.masterParts()) {
      parts.push({ part: m, label: 'master' });
      for (const l of doc.layoutsOf(m)) parts.push({ part: l, label: `layout "${doc.slide(l).name}"` });
    }
  }
  const out = parts.map(({ part, label }) => {
    const r = resolveSlide(doc, part);
    return { part, label, size: { width: pt(r.size.cx), height: pt(r.size.cy) }, background: paint(r.background), shapes: r.shapes.map(shapeInfo) };
  });
  if (args.includes('--json')) {
    console.log(JSON.stringify(out, null, 2));
    return;
  }
  for (const s of out) {
    console.log(`== ${s.label} (${s.part})  ${s.size.width}x${s.size.height}pt  background ${JSON.stringify(s.background)}`);
    for (const shape of s.shapes) print(shape, 1);
  }
}

main();
