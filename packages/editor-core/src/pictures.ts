// packages/editor-core/src/pictures.ts  -- inserting, cropping and replacing pictures
import { normalizeName, relativeTarget, resolveTarget } from '@pptx/opc';
import { attr, attrNs, child, insertInOrder, int, path, relType, removeChildren, setAttrNs, ShapeView, type Xfrm } from '@pptx/model';
import type { Editor } from './editor.ts';
import type { Command } from './history.ts';
import { collectOrphans, dropUnreferenced, hashMedia, mediaPart, relRefs } from './parts.ts';
import { colorXml, type GradientSpec } from './format.ts';
import type { Rect } from './shapes.ts';
import type { ColorSpec } from './text.ts';
import { clone } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
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

/** A picture's crop as the file has it: thousandths of a percent cut off each side (0 when it has none). */
export function pictureCrop(pic: ShapeView): { l: number; t: number; r: number; b: number } {
  const src = child(pic.blipFill, 'a:srcRect');
  return { l: int(attr(src, 'l'), 0), t: int(attr(src, 't'), 0), r: int(attr(src, 'r'), 0), b: int(attr(src, 'b'), 0) };
}

/**
 * Crop a picture and place its frame in one step, for cropping with handles: the frame's edge moves over the
 * picture, so both change together. Drags pass one `coalesceKey`, making the whole drag one undo step.
 */
export function cropAndPlace(ed: Editor, part: string, pic: ShapeView, crop: { l: number; t: number; r: number; b: number }, xfrm: Xfrm, coalesceKey?: string): Command {
  return ed.edit({
    label: 'Crop', parts: [part], scope: () => [pic.el], ...(coalesceKey ? { coalesceKey } : {}),
    apply: () => {
      const blipFill = pic.blipFill;
      if (!blipFill) throw new Error('not a picture');
      removeChildren(blipFill, 'a:srcRect');
      const v = (n: number) => (Math.round(n) ? String(Math.round(n)) : undefined);
      const attrs = (['l', 't', 'r', 'b'] as const).map(k => [k, v(crop[k])] as const).filter(([, x]) => x !== undefined).map(([k, x]) => ` ${k}="${x}"`).join('');
      insertInOrder(blipFill, element(blipFill, `<a:srcRect${attrs}/>`, ed.strict));
      pic.setXfrm({ ...xfrm, x: Math.round(xfrm.x), y: Math.round(xfrm.y), cx: Math.round(xfrm.cx), cy: Math.round(xfrm.cy) });
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

const SHAPE_FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

/**
 * Fill shapes with a picture, stretched over each shape and turning with it. The picture is stored once (an
 * identical one already in the deck is reused). Pictures and graphic frames have no fill and are left alone.
 */
export function setPictureFill(ed: Editor, part: string, shapes: ShapeView[], img: ImageData): Command {
  const root = ed.doc.root(part);
  return ed.edit({
    label: 'Shape Fill', parts: [part], structural: true, scope: () => [root],
    apply: () => {
      const before = new Set(relRefs(root).map(r => r.value));
      const xml = `<a:blipFill dpi="0" rotWithShape="1">${blipXml(ed, part, img)}<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
      for (const s of shapes) {
        if (s.kind === 'graphicFrame' || s.kind === 'pic' || s.kind === 'contentPart') continue;
        const name = s.kind === 'grpSp' ? 'p:grpSpPr' : 'p:spPr';
        const spPr = child(s.el, name) ?? (() => { const e = element(s.el, `<${name}/>`, ed.strict); insertInOrder(s.el, e); return e; })();
        removeChildren(spPr, ...SHAPE_FILLS);
        insertInOrder(spPr, element(spPr, xml, ed.strict));
      }
      // a picture fill this replaced may have been the last use of its picture
      dropUnreferenced(ed.pkg, part, before);
    },
  });
}

/** What a slide's background is filled with: a colour (the theme's or any other), a two-colour gradient, or a picture. */
export type BackgroundSpec = { color: ColorSpec } | { gradient: GradientSpec } | { picture: ImageData };

/** Remove what a part's XML no longer refers to among the relationships it referred to before an edit. */
function dropUnused(ed: Editor, part: string, before: Set<string>): void {
  dropUnreferenced(ed.pkg, part, before);
}

/**
 * Give slides a background of their own (null: back to the layout's). Several slides are one undo step, which is
 * how "apply to all" is done. A picture is stored once however many slides show it.
 */
export function setSlideBackground(ed: Editor, slides: string[], bg: BackgroundSpec | null): Command {
  return ed.edit({
    label: 'Format Background', parts: slides, structural: true, scope: () => slides.map(p => ed.doc.root(p)),
    apply: () => {
      for (const part of slides) {
        const root = ed.doc.root(part);
        const cSld = child(root, 'p:cSld');
        if (!cSld) continue;
        const before = new Set(relRefs(root).map(r => r.value));
        removeChildren(cSld, 'p:bg');
        if (bg) {
          const ang = 'gradient' in bg ? Math.round((((bg.gradient.angle % 360) + 360) % 360) * 60000) : 0;
          const fill = 'color' in bg ? `<a:solidFill>${colorXml(bg.color)}</a:solidFill>`
            : 'gradient' in bg ? `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${colorXml(bg.gradient.from)}</a:gs><a:gs pos="100000">${colorXml(bg.gradient.to)}</a:gs></a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`
              : `<a:blipFill dpi="0" rotWithShape="1">${blipXml(ed, part, bg.picture)}<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
          insertInOrder(cSld, element(cSld, `<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`, ed.strict));
        }
        dropUnused(ed, part, before);
      }
    },
  });
}

/**
 * Give every other slide the background one slide has (PowerPoint's Apply to All): its colour, gradient or picture,
 * or, when it has none of its own, no background of their own either. One undo step.
 */
export function applyBackgroundToAll(ed: Editor, from: string): Command {
  const slides = ed.doc.slideParts().filter(p => p !== from);
  return ed.edit({
    label: 'Apply Background to All', parts: slides, structural: true, scope: () => slides.map(p => ed.doc.root(p)),
    apply: () => {
      const source = child(child(ed.doc.root(from), 'p:cSld'), 'p:bg');
      // the pictures the background shows, by the relationship ids its XML names them with
      const pictures = new Map<string, string>();
      if (source) {
        for (const ref of relRefs(source)) {
          const rel = ed.pkg.relationships(from).find(r => r.id === ref.value && !r.external);
          if (rel) pictures.set(ref.value, resolveTarget(normalizeName(from), rel.target));
        }
      }
      for (const part of slides) {
        const root = ed.doc.root(part);
        const cSld = child(root, 'p:cSld');
        if (!cSld) continue;
        const before = new Set(relRefs(root).map(r => r.value));
        removeChildren(cSld, 'p:bg');
        if (source) {
          const copy = clone(source, cSld);
          const repoint = (e: XmlElement) => {
            for (const a of e.attrs) { const target = pictures.get(a.value); if (target && a.ns.endsWith('/relationships')) a.value = relTo(ed, part, 'image', target); }
            for (const c of e.children) if (c.kind === 'element') repoint(c);
          };
          repoint(copy);
          insertInOrder(cSld, copy);
        }
        dropUnused(ed, part, before);
      }
    },
  });
}

/** A video or audio file to insert: its bytes, type, extension, and hash (from `prepareMedia`). */
export interface MediaData { bytes: Uint8Array; contentType: string; ext: string; hash: string; }

/** Hash a media file (and the package's media, once), so an identical file already in the deck is reused. */
export async function prepareMedia(ed: Editor, media: Omit<MediaData, 'hash'>): Promise<MediaData> {
  return { ...media, hash: await hashMedia(ed, media.bytes) };
}

const MEDIA_EXT = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
const MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media';

/**
 * Insert a video or a sound as PowerPoint stores one: the media file as a part, and a picture (the poster: a frame
 * of the video, or an icon for a sound) that carries two relationships to it, the one every version reads
 * (a:videoFile or a:audioFile) and the one newer versions play from (p14:media). The picture is what is seen on
 * the slide and moved about; the media plays in its place.
 */
export function insertMedia(ed: Editor, part: string, kind: 'video' | 'audio', media: MediaData, poster: ImageData, box?: Rect, name?: string): { command: Command; id: number } {
  const root = ed.doc.root(part);
  const tree = ed.spTree(part);
  const id = nextShapeId(root);
  const size = naturalSize(ed, poster);
  const slide = ed.doc.presentation.slideSize;
  const b = box ?? { x: (slide.cx - size.cx) / 2, y: (slide.cy - size.cy) / 2, w: size.cx, h: size.cy };
  const r = Math.round;
  const command = ed.edit({
    label: kind === 'video' ? 'Insert Video' : 'Insert Audio', parts: [part], structural: true, scope: () => [tree],
    apply: () => {
      const file = mediaPart(ed.pkg, media.bytes, media.ext, media.contentType, media.hash, 'media');
      const link = ed.pkg.addRelationship(part, relType(kind, ed.strict), file);
      const embed = ed.pkg.addRelationship(part, MEDIA_REL, file);
      const xml = `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name ?? (kind === 'video' ? `Video ${id - 1}` : `Audio ${id - 1}`))}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>`
        + `<p:nvPr><a:${kind}File r:link="${link}"/><p:extLst><p:ext uri="${MEDIA_EXT}"><p14:media r:embed="${embed}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>`
        + `<p:blipFill>${blipXml(ed, part, poster)}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`
        + `<p:spPr><a:xfrm><a:off x="${r(b.x)}" y="${r(b.y)}"/><a:ext cx="${r(b.w)}" cy="${r(b.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      const units = unitsIn(tree);
      const last = units[units.length - 1];
      insertAt(tree, last ? tree.children.indexOf(last) + 1 : tree.children.length, element(tree, xml, ed.strict));
    },
  });
  return { command, id };
}

