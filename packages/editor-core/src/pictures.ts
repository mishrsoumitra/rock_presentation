// packages/editor-core/src/pictures.ts  -- inserting, cropping and replacing pictures
import { normalizeName, relativeTarget, resolveTarget } from '@pptx/opc';
import { attrNs, child, insertInOrder, path, relType, removeChildren, setAttrNs, ShapeView } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { collectOrphans, dropUnreferenced, hashMedia, mediaPart, relRefs } from './parts.ts';
import type { Rect } from './shapes.ts';
import { element, esc, insertAt, nextShapeId, unitsIn } from './xml.ts';

export interface ImageData {
  bytes: Uint8Array;
  contentType: string;
  /** File extension without the dot, lower case: png, jpeg, gif, svg... */
  ext: string;
  /** Pixel size and resolution (dots per inch; 96 when the file does not say). */
  width: number;
  height: number;
  dpiX?: number;
  dpiY?: number;
  /** SHA-256 of the bytes, from `prepareImage`. */
  hash: string;
}

/** An SVG picture: PowerPoint stores a PNG rendering for older readers and the SVG beside it. */
export interface SvgImage extends ImageData { svg: { bytes: Uint8Array; hash: string } }

const SVG_EXT = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const ASVG = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

/** Hash the bytes (and the package's media, once) so inserting can reuse an identical media part at once. */
export async function prepareImage(ed: Editor, image: Omit<ImageData, 'hash'>): Promise<ImageData> {
  return { ...image, hash: await hashMedia(ed, image.bytes) };
}

export async function prepareSvg(ed: Editor, png: Omit<ImageData, 'hash'>, svg: Uint8Array): Promise<SvgImage> {
  return { ...png, hash: await hashMedia(ed, png.bytes), svg: { bytes: svg, hash: await hashMedia(ed, svg) } };
}

/** The picture's size on the slide: pixels at its resolution, scaled down to fit the slide if it is larger. */
export function naturalSize(ed: Editor, img: Pick<ImageData, 'width' | 'height' | 'dpiX' | 'dpiY'>): { cx: number; cy: number } {
  let cx = (img.width * 914400) / (img.dpiX || 96);
  let cy = (img.height * 914400) / (img.dpiY || 96);
  const slide = ed.doc.presentation.slideSize;
  const k = Math.min(1, slide.cx / cx, slide.cy / cy);
  cx *= k;
  cy *= k;
  return { cx: Math.round(cx), cy: Math.round(cy) };
}

/** A relationship from `part` to `target` of this kind: an existing one, or a new one. */
function relTo(ed: Editor, part: string, kind: string, target: string): string {
  const type = relType(kind, ed.strict);
  const existing = ed.pkg.relationships(part).find(r => !r.external && r.type.endsWith(`/${kind}`) && resolveTarget(normalizeName(part), r.target) === target);
  return existing?.id ?? ed.pkg.addRelationship(part, type, target);
}

function blipXml(ed: Editor, part: string, img: ImageData | SvgImage): string {
  const media = mediaPart(ed.pkg, img.bytes, img.ext, img.contentType, img.hash);
  const rId = relTo(ed, part, 'image', media);
  if (!('svg' in img)) return `<a:blip r:embed="${rId}"/>`;
  const svgPart = mediaPart(ed.pkg, img.svg.bytes, 'svg', 'image/svg+xml', img.svg.hash);
  const svgId = relTo(ed, part, 'image', svgPart);
  return `<a:blip r:embed="${rId}"><a:extLst><a:ext uri="${SVG_EXT}"><asvg:svgBlip xmlns:asvg="${ASVG}" r:embed="${svgId}"/></a:ext></a:extLst></a:blip>`;
}

/**
 * Insert a picture as PowerPoint does: its media part (an identical one is reused), an image relationship from the
 * slide, and a p:pic that stretches the picture over its frame, centred on the slide unless a box is given.
 */
export function insertPicture(ed: Editor, part: string, img: ImageData | SvgImage, box?: Rect, descr?: string): { command: Command; id: number } {
  const root = ed.doc.root(part);
  const tree = ed.spTree(part);
  const id = nextShapeId(root);
  const size = naturalSize(ed, img);
  const slide = ed.doc.presentation.slideSize;
  const b = box ?? { x: (slide.cx - size.cx) / 2, y: (slide.cy - size.cy) / 2, w: size.cx, h: size.cy };
  const command = ed.edit({
    label: 'Insert Picture', parts: [part], structural: true, scope: () => [tree],
    apply: () => {
      const blip = blipXml(ed, part, img);
      const r = Math.round;
      const pic = element(tree, `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id - 1}"${descr ? ` descr="${esc(descr)}"` : ''}/>`
        + '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        + `<p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`
        + `<p:spPr><a:xfrm><a:off x="${r(b.x)}" y="${r(b.y)}"/><a:ext cx="${r(b.w)}" cy="${r(b.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`, ed.strict);
      const units = unitsIn(tree);
      const last = units[units.length - 1];
      insertAt(tree, last ? tree.children.indexOf(last) + 1 : tree.children.length, pic);
    },
  });
  return { command, id };
}

/** Crop a picture: a:srcRect in thousandths of a percent of the picture's size (negative values pad). */
export function cropPicture(ed: Editor, part: string, pic: ShapeView, crop: { l: number; t: number; r: number; b: number }): Command {
  return ed.edit({
    label: 'Crop', parts: [part], scope: () => [pic.el],
    apply: () => {
      const blipFill = pic.blipFill;
      if (!blipFill) throw new Error('not a picture');
      removeChildren(blipFill, 'a:srcRect');
      const v = (n: number) => (Math.round(n) ? String(Math.round(n)) : undefined);
      const attrs = (['l', 't', 'r', 'b'] as const).map(k => [k, v(crop[k])] as const).filter(([, x]) => x !== undefined).map(([k, x]) => ` ${k}="${x}"`).join('');
      insertInOrder(blipFill, element(blipFill, `<a:srcRect${attrs}/>`, ed.strict));
    },
  });
}

/**
 * Replace a picture's image, keeping its shape, position, crop and effects: the relationship is re-pointed (or a
 * new one made when other elements share it), and the old media part goes if nothing shows it any more. A picture
 * that only links to its image, or whose fill has lost its blip, gets the new image embedded, as PowerPoint does.
 */
export function replacePicture(ed: Editor, part: string, pic: ShapeView, img: ImageData): Command {
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Change Picture', parts: [part], structural: true, scope: () => [root],
    apply: () => {
      const fill = pic.blipFill;
      if (!fill) throw new Error('not a picture');
      let blip = child(fill, 'a:blip');
      if (!blip) {
        blip = element(fill, '<a:blip/>', ed.strict);
        insertInOrder(fill, blip);
      }
      const oldId = attrNs(blip, 'r', 'embed');
      const oldRel = oldId ? ed.pkg.relationships(part).find(r => r.id === oldId) : undefined;
      const oldTarget = oldRel && !oldRel.external ? resolveTarget(normalizeName(part), oldRel.target) : undefined;
      const media = mediaPart(ed.pkg, img.bytes, img.ext, img.contentType, img.hash);
      const shared = oldId !== undefined && relRefs(root).filter(r => r.value === oldId).length > 1;
      if (oldId && oldRel && !shared) retarget(ed, part, oldId, media);
      else setAttrNs(blip, 'r', 'embed', relTo(ed, part, 'image', media));
      // a link to the old image (a linked picture) goes with it
      const linkId = attrNs(blip, 'r', 'link');
      if (linkId) {
        setAttrNs(blip, 'r', 'link', undefined);
        dropUnreferenced(ed.pkg, part, [linkId]);
      }
      // the SVG beside an old picture no longer matches it
      const ext = path(blip, 'a:extLst');
      if (ext) removeChildren(blip, 'a:extLst');
      if (oldTarget && oldTarget !== media) collectOrphans(ed.pkg, [oldTarget]);
    },
  });
}

/** Point an existing relationship at another part. */
export function retarget(ed: Editor, part: string, id: string, target: string): void {
  const rels = ed.relsOf(part);
  if (!rels) throw new Error(`${part} has no relationships`);
  for (const c of ed.pkg.xml(rels).root.children) {
    if (c.kind !== 'element' || c.attrs.find(a => a.local === 'Id')?.value !== id) continue;
    const t = c.attrs.find(a => a.local === 'Target' && a.ns === '');
    if (t) t.value = relativeTarget(normalizeName(part), target);
    c.attrs = c.attrs.filter(a => a.local !== 'TargetMode');
  }
  ed.pkg.markDirty(rels);
}
