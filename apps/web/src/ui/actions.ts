// apps/web/src/ui/actions.ts  -- what the ribbon and the keyboard do: commands on the current selection
import {
  addSlide, align, applyLayout, compound, copyShapes, cropPicture, deleteShapes, deleteSlide, distribute, duplicateShapes, duplicateSlide, flip, group, importMasters, replacePicture,
  insertPicture, insertShape, insertTextBox, moveBy, moveSlide, pasteShapes, prepareImage, prepareSvg, prepareTemplate, rotateQuarter,
  setBackground, setFill, setLine, setSlideHidden, setSlideSize, setTransforms, ungroup, zOrder, type Alignment, type ColorSpec, type LineSpec,
  type ScaleMode, type ZOrder,
} from '@pptx/editor-core';
import { OpcPackage } from '@pptx/opc';
import { ShapeView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import { blankPresentation } from '../engine/blank.ts';
import { decodeImage } from '../engine/images.ts';
import { useEditor } from '../state/store.ts';

const st = () => useEditor.getState();

function ctx() {
  const s = st();
  if (!s.session || !s.slide) return undefined;
  const views = s.selection.map(el => new ShapeView(el, s.session!.ed.doc.owner(s.slide!)));
  return { s, ed: s.session.ed, part: s.slide, views };
}

/** The element of a shape just inserted, by id, to select it. */
function inserted(id: number): XmlElement[] {
  const c = ctx();
  const v = c?.ed.shape(c.part, id);
  return v ? [v.el] : [];
}

export function newPresentation(): Promise<void> {
  return st().open(blankPresentation(), 'Presentation1.pptx');
}

export async function openFile(file: File): Promise<void> {
  await st().open(new Uint8Array(await file.arrayBuffer()), file.name);
}

export async function saveAndDownload(): Promise<void> {
  const s = st();
  const bytes = await s.save().catch(e => { s.notify('error', `Saving failed: ${e instanceof Error ? e.message : String(e)}`); return undefined; });
  if (!bytes || !s.session) return;
  const name = s.session.name.replace(/\.(potx|ppsx|pptm|potm|ppsm)$/i, '.pptx');
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = /\.pptx$/i.test(name) ? name : `${name}.pptx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
  s.notify('info', `Saved ${a.download}${s.session.password ? ' (encrypted)' : ''}.`);
}

/** Pick a shape to draw: the next drag on the slide gives its box (a click gives the default size there). */
export function startDrawing(prst: string): void {
  useEditor.setState({ drawing: prst, text: null });
}

/** Insert a preset shape in a box (slide EMU); without one, PowerPoint's default size in the middle of the slide. */
export function insertShapeHere(prst: string, box?: { x: number; y: number; w: number; h: number }): void {
  const c = ctx();
  if (!c) return;
  const size = c.ed.doc.presentation.slideSize;
  const line = /Connector|^line$/.test(prst);
  const w = 1828800, h = line ? 0 : 1371600;
  const r = insertShape(c.ed, c.part, prst, box ?? { x: (size.cx - w) / 2, y: (size.cy - h) / 2, w, h });
  c.s.run(r.command);
  c.s.select(inserted(r.id));
}

export function insertTextBoxHere(): void {
  const c = ctx();
  if (!c) return;
  const size = c.ed.doc.presentation.slideSize;
  const r = insertTextBox(c.ed, c.part, { x: size.cx / 2 - 1524000, y: size.cy / 2 - 184666, w: 3048000, h: 369332 });
  c.s.run(r.command);
  const els = inserted(r.id);
  if (els[0]) useEditor.setState({ selection: els, text: { part: c.part, el: els[0], caret: { p: 0, offset: 0 }, anchor: { p: 0, offset: 0 } } });
}

const EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/tiff': 'tiff', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf' };

/** The content type of a picture file, from its bytes where possible (never trusting the name alone). */
function sniffImage(bytes: Uint8Array, name: string): string | undefined {
  const b = bytes;
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'image/png';
  if (b[0] === 0xff && b[1] === 0xd8) return 'image/jpeg';
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'image/gif';
  if (b[0] === 0x42 && b[1] === 0x4d) return 'image/bmp';
  if (b[0] === 0x52 && b[1] === 0x49 && b[8] === 0x57 && b[9] === 0x45) return 'image/webp';
  if (b[0] === 0x01 && b[1] === 0 && b[2] === 0 && b[3] === 0 && b[40] === 0x20 && b[41] === 0x45) return 'image/x-emf';
  if ((b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) || (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x09 && b[3] === 0x00)) return 'image/x-wmf';
  if (/\.svg$/i.test(name) && /<svg[\s>]/.test(new TextDecoder().decode(b.subarray(0, 4096)))) return 'image/svg+xml';
  return undefined;
}

/** The resolution a PNG or JPEG states, so the picture comes in at the size PowerPoint would give it. */
function dpiOf(bytes: Uint8Array, type: string): number {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (type === 'image/png') {
    for (let i = 8; i + 8 < bytes.length;) {
      const len = dv.getUint32(i);
      const t = String.fromCharCode(...bytes.subarray(i + 4, i + 8));
      if (t === 'pHYs' && bytes[i + 16] === 1) return Math.round(dv.getUint32(i + 8) * 0.0254) || 96;
      if (t === 'IDAT') break;
      i += 12 + len;
    }
  } else if (type === 'image/jpeg' && bytes[6] === 0x4a && bytes[11] === 0 /* JFIF */) {
    const units = bytes[13], x = dv.getUint16(14);
    if (units === 1 && x) return x;
    if (units === 2 && x) return Math.round(x * 2.54);
  }
  return 96;
}

export async function insertPictureFile(file: File): Promise<void> {
  const c = ctx();
  if (!c) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  const type = sniffImage(bytes, file.name);
  if (!type) { c.s.notify('error', `${file.name} is not a picture this editor can insert.`); return; }
  const img = await decodeImage(bytes, type).catch(() => undefined);
  if (!img && type !== 'image/x-emf' && type !== 'image/x-wmf') { c.s.notify('error', `${file.name} could not be read as a picture.`); return; }
  const width = (img as { width?: number } | undefined)?.width ?? 960, height = (img as { height?: number } | undefined)?.height ?? 720;
  const dpi = dpiOf(bytes, type);
  let prepared;
  if (type === 'image/svg+xml' && img) {
    // PowerPoint keeps a PNG of an SVG for readers that cannot draw SVG
    const canvas = new OffscreenCanvas(width, height);
    canvas.getContext('2d')!.drawImage(img, 0, 0);
    const png = new Uint8Array(await (await canvas.convertToBlob({ type: 'image/png' })).arrayBuffer());
    prepared = await prepareSvg(c.ed, { bytes: png, contentType: 'image/png', ext: 'png', width, height }, bytes);
  } else {
    prepared = await prepareImage(c.ed, { bytes, contentType: type, ext: EXT[type]!, width, height, dpiX: dpi, dpiY: dpi });
  }
  const r = insertPicture(c.ed, c.part, prepared, undefined, file.name.replace(/\.[^.]+$/, ''));
  c.s.run(r.command);
  c.s.select(inserted(r.id));
}

export function deleteSelection(): void {
  const c = ctx();
  if (!c || !c.views.length) return;
  c.s.run(deleteShapes(c.ed, c.part, c.views), { selection: [] });
}

export function duplicateSelection(): void {
  const c = ctx();
  if (!c || !c.views.length) return;
  const r = duplicateShapes(c.ed, c.part, c.views);
  c.s.run(r.command);
  c.s.select(r.ids.flatMap(inserted));
}

export function copySelection(cut = false): void {
  const c = ctx();
  if (!c || !c.views.length) return;
  useEditor.setState({ clipboard: copyShapes(c.ed, c.part, c.views) });
  if (cut) deleteSelection();
}

export async function paste(): Promise<void> {
  const c = ctx();
  const clip = st().clipboard;
  if (!c || !clip) return;
  const hashes = clip.from === c.ed.pkg ? undefined : await prepareTemplate(c.ed, clip.from);
  const r = pasteShapes(c.ed, c.part, clip, hashes);
  c.s.run(r.command);
  c.s.select(r.ids.flatMap(inserted));
}

export function groupSelection(): void {
  const c = ctx();
  if (!c || c.views.length < 2) return;
  c.s.run(group(c.ed, c.part, c.views));
  const g = c.views[0]!.el.parent;
  if (g) c.s.select([g]);
}

export function ungroupSelection(): void {
  const c = ctx();
  const g = c?.views.find(v => v.kind === 'grpSp');
  if (!c || !g) return;
  const members = g.children.map(m => m.el);
  c.s.run(ungroup(c.ed, c.part, g));
  c.s.select(members);
}

export function order(op: ZOrder): void {
  const c = ctx();
  if (c?.views.length) c.s.run(zOrder(c.ed, c.part, c.views, op));
}

export function alignSelection(how: Alignment): void {
  const c = ctx();
  if (c?.views.length) c.s.run(align(c.ed, c.part, c.views, how));
}

export function distributeSelection(axis: 'h' | 'v'): void {
  const c = ctx();
  if (c && c.views.length > 2) c.s.run(distribute(c.ed, c.part, c.views, axis));
}

export function rotateSelection(dir: 1 | -1): void {
  const c = ctx();
  if (c?.views.length) c.s.run(rotateQuarter(c.ed, c.part, c.views, dir));
}

export function flipSelection(axis: 'h' | 'v'): void {
  const c = ctx();
  if (c?.views.length) c.s.run(flip(c.ed, c.part, c.views, axis));
}

export function nudge(dx: number, dy: number): void {
  const c = ctx();
  if (c?.views.length) c.s.run(moveBy(c.ed, c.part, c.views, dx, dy, 'Nudge'));
}

export function fillSelection(fill: ColorSpec | 'none' | null): void {
  const c = ctx();
  if (c?.views.length) c.s.run(setFill(c.ed, c.part, c.views, fill));
}

export function outlineSelection(line: LineSpec): void {
  const c = ctx();
  if (c?.views.length) c.s.run(setLine(c.ed, c.part, c.views, line));
}

export function newSlide(layout?: string): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed) return;
  const slides = ed.doc.slideParts();
  const at = s.slide ? slides.indexOf(s.slide) + 1 : slides.length;
  const current = s.slide ?? slides[slides.length - 1];
  const master = ed.doc.masterParts()[0];
  // PowerPoint's New Slide uses the current slide's layout, unless it is the title slide: then Title and Content
  let use = layout;
  if (!use && current) {
    const own = ed.doc.layoutOf(current);
    const kind = own ? ed.doc.slide(own).layoutType : undefined;
    const m = own ? ed.doc.masterOf(own) : master;
    use = kind === 'title' && m ? ed.doc.layoutsOf(m).find(l => ed.doc.slide(l).layoutType === 'obj') ?? own : own;
  }
  use ??= master ? ed.doc.layoutsOf(master)[0] : undefined;
  if (!use) return;
  const r = addSlide(ed, use, at);
  s.run(r.command);
  s.setSlide(r.part);
}

export function deleteCurrentSlide(): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  const slides = ed.doc.slideParts();
  const i = slides.indexOf(s.slide);
  s.run(deleteSlide(ed, s.slide));
  const rest = ed.doc.slideParts();
  s.setSlide(rest[Math.min(i, rest.length - 1)] ?? null);
}

export function duplicateCurrentSlide(): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  const r = duplicateSlide(ed, s.slide);
  s.run(r.command);
  if (r.part) s.setSlide(r.part);
}

export function moveCurrentSlide(to: number, part = st().slide): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !part) return;
  s.run(moveSlide(ed, part, to));
}

export function toggleHidden(): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  s.run(setSlideHidden(ed, s.slide, !ed.doc.slide(s.slide).hidden));
}

export function layoutCurrentSlide(layout: string): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  s.run(applyLayout(ed, s.slide, layout), { selection: [] });
}

export function backgroundCurrentSlide(hex: string | null): void {
  const s = st();
  const ed = s.session?.ed;
  if (ed && s.slide) s.run(setBackground(ed, s.slide, hex ? { hex } : null));
}

export function resizeSlides(cx: number, cy: number, mode: ScaleMode): void {
  const s = st();
  const ed = s.session?.ed;
  if (ed) s.run(setSlideSize(ed, cx, cy, mode), { selection: [] });
}

/** Import the slide masters of another deck or template; their layouts become available for new slides. */
export async function importTemplate(file: File): Promise<void> {
  const s = st();
  const ed = s.session?.ed;
  if (!ed) return;
  try {
    const from = OpcPackage.open(new Uint8Array(await file.arrayBuffer()));
    const hashes = await prepareTemplate(ed, from);
    const r = importMasters(ed, from, hashes);
    s.run(r.command);
    s.notify('info', `Imported ${r.masters.length} slide master${r.masters.length === 1 ? '' : 's'} from ${file.name}.`);
  } catch (e) {
    s.notify('error', `${file.name} could not be imported: ${e instanceof Error ? e.message : String(e)}`);
  }
}

/** Every layout of every master, for the New Slide and Layout galleries. */
export function layouts(): { part: string; name: string; master: string }[] {
  const ed = st().session?.ed;
  if (!ed) return [];
  return ed.doc.masterParts().flatMap(m => ed.doc.layoutsOf(m).map(l => ({ part: l, name: ed.doc.slide(l).name || 'Layout', master: m })));
}

/** The selected picture, if exactly one picture is selected. */
function selectedPicture() {
  const c = ctx();
  const v = c?.views.length === 1 ? c.views[0] : undefined;
  return c && v?.kind === 'pic' ? { c, v } : undefined;
}

/**
 * Crop a picture to an aspect ratio about its centre, as PowerPoint's Crop > Aspect Ratio does (the frame keeps
 * its width; the visible part of the picture is chosen to fill it). Null removes the crop.
 */
export function cropToAspect(ratio: number | null): void {
  const pic = selectedPicture();
  if (!pic) return;
  const { c, v } = pic;
  if (ratio === null) { c.s.run(cropPicture(c.ed, c.part, v, { l: 0, t: 0, r: 0, b: 0 })); return; }
  const x = v.xfrm;
  const scene = c.s.session!.scenes.get(c.part);
  const rs = scene.slide.shapes.find(s => s.source?.view.el === v.el);
  const img = rs?.source.fill.kind === 'image' ? c.s.session!.images.get(rs.source.fill.image.part ?? '') as { width?: number; height?: number } | undefined : undefined;
  if (!x || !img?.width || !img.height) return;
  const pictureAspect = img.width / img.height;
  // the share of the picture's width and height kept, so the kept part has the wanted aspect
  const keepW = Math.min(1, ratio / pictureAspect), keepH = Math.min(1, pictureAspect / ratio);
  const side = (keep: number) => Math.round(((1 - keep) / 2) * 100000);
  // the crop and the frame's new height are one undo step
  const cy = x.cx / ratio;
  c.s.run(compound('Crop', [
    cropPicture(c.ed, c.part, v, { l: side(keepW), r: side(keepW), t: side(keepH), b: side(keepH) }),
    setTransforms(c.ed, c.part, [{ shape: v, xfrm: { ...x, y: x.y + (x.cy - cy) / 2, cy } }], 'Crop'),
  ]));
}

/** Swap the selected picture's image, keeping its frame, crop and effects. */
export async function changePicture(file: File): Promise<void> {
  const pic = selectedPicture();
  if (!pic) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  const type = sniffImage(bytes, file.name);
  if (!type || type === 'image/svg+xml') { pic.c.s.notify('error', `${file.name} cannot replace a picture here.`); return; }
  const img = await decodeImage(bytes, type).catch(() => undefined);
  if (!img) { pic.c.s.notify('error', `${file.name} could not be read as a picture.`); return; }
  const width = (img as { width?: number } | undefined)?.width ?? 960, height = (img as { height?: number } | undefined)?.height ?? 720;
  const prepared = await prepareImage(pic.c.ed, { bytes, contentType: type, ext: EXT[type]!, width, height });
  pic.c.s.run(replacePicture(pic.c.ed, pic.c.part, pic.v, prepared));
}
