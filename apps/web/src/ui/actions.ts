// apps/web/src/ui/actions.ts  -- what the ribbon and the keyboard do: commands on the current selection
import {
  addSlide, align, applyLayout, compound, copyShapes, copySlide, pasteSlide, cropPicture, deleteShapes, deleteSlide, distribute, duplicateShapes, duplicateSlide, flip, group, importMasters, replacePicture,
  insertPicture, insertShape, insertTable, insertTextBox, moveBy, ROW_HEIGHT, moveSlide, pasteShapes, prepareImage, prepareSvg, prepareTemplate, rotateQuarter,
  fillColorOf, setBackground, setFill, setFillTransparency, setGradientFill, setLine, setPictureFill, setShadow, setSlideHidden, setSlideSize, setTransforms, ungroup, zOrder,
  type Alignment, type ColorSpec, type LineSpec,
  type ScaleMode, type ZOrder,
} from '@pptx/editor-core';
import { OpcPackage } from '@pptx/opc';
import { copyFormat, copyTextFormat, paragraphLength, paragraphs, pasteFormat } from '@pptx/editor-core';
import { runsOf } from '@pptx/model';
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

/**
 * Insert a table in the middle of the slide, two thirds of its width, as PowerPoint places one on a slide with no
 * content placeholder to take it.
 */
export function insertTableHere(rows: number, cols: number): void {
  const c = ctx();
  if (!c) return;
  const size = c.ed.doc.presentation.slideSize;
  const w = Math.round((size.cx * 2) / 3);
  const r = insertTable(c.ed, c.part, rows, cols, { x: (size.cx - w) / 2, y: Math.max(0, (size.cy - rows * ROW_HEIGHT) / 2), w, h: rows * ROW_HEIGHT });
  useEditor.setState({ text: null });
  c.s.run(r.command);
  c.s.select(inserted(r.id));
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

/**
 * Pick up a look for the format painter: from the text selected in the box being edited (its character formatting
 * only), or from the first selected shape (its fill, line, effects and text look). False when there is nothing to
 * pick up from.
 */
export function pickUpFormat(): boolean {
  const s = st();
  if (!s.session || !s.slide) return false;
  const t = s.text;
  if (t && (t.anchor.p !== t.caret.p || t.anchor.offset !== t.caret.offset)) {
    // the run the selection starts in
    const body = new ShapeView(t.el, s.session.ed.doc.owner(t.part)).txBody;
    const start = t.anchor.p < t.caret.p || (t.anchor.p === t.caret.p && t.anchor.offset <= t.caret.offset) ? t.anchor : t.caret;
    const p = body ? paragraphs(body)[start.p] : undefined;
    let at = 0;
    let rPr = p ? runsOf(p)[0]?.rPr : undefined;
    for (const r of p ? runsOf(p) : []) {
      const len = r.kind === 'br' ? 1 : r.text.length;
      if (start.offset >= at && start.offset < at + len) { rPr = r.rPr; break; }
      at += len;
    }
    useEditor.setState({ formatClip: copyTextFormat(rPr) });
    return true;
  }
  const el = t?.el ?? s.selection[0];
  if (!el) return false;
  useEditor.setState({ formatClip: copyFormat(new ShapeView(el, s.session.ed.doc.owner(s.slide))) });
  return true;
}

/** Put the picked-up look on shapes (the selected ones by default). */
export function paintFormat(els: XmlElement[] = st().selection): void {
  const s = st();
  if (!s.session || !s.slide || !s.formatClip || !els.length) return;
  s.run(pasteFormat(s.session.ed, s.slide, els.map(el => new ShapeView(el, s.session!.ed.doc.owner(s.slide!))), s.formatClip));
}

/** The Format Painter button: on (for one shape, or until Escape when `sticky`), or off again. */
export function togglePainter(sticky = false): void {
  const s = st();
  if (s.painting && !sticky) { useEditor.setState({ painting: false }); return; }
  if (pickUpFormat()) useEditor.setState({ painting: sticky ? 'sticky' : 'once' });
}

/** Open the selection pane with the keyboard in its alt text box, for the one selected shape. */
export function editAltText(): void {
  useEditor.setState({ pane: 'selection' });
  // the pane (and its box) is in the page after this render
  setTimeout(() => window.dispatchEvent(new CustomEvent('focus-alt-text')), 0);
}

/** Select every shape of the slide itself (not the layout's or master's), as Ctrl+A does. */
export function selectAllShapes(): void {
  const s = st();
  if (!s.session || !s.slide) return;
  const scene = s.session.scenes.get(s.slide);
  s.select(scene.slide.shapes.filter(x => x.layer === 'slide' && x.source).map(x => x.source.view.el));
}

/** Start editing the text of the one selected shape, with the caret at its end (Enter, F2, the menu's Edit Text). */
export function editSelectedText(): boolean {
  const s = st();
  if (!s.session || !s.slide || s.selection.length !== 1) return false;
  const v = new ShapeView(s.selection[0]!, s.session.ed.doc.owner(s.slide));
  const body = v.txBody;
  if (!body) return false;
  const ps = paragraphs(body);
  const end = { p: Math.max(0, ps.length - 1), offset: ps.length ? paragraphLength(ps[ps.length - 1]!) : 0 };
  s.setState({ text: { part: s.slide, el: v.el, caret: end, anchor: end } });
  return true;
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

/**
 * Copy (or cut) from a button or a menu: through the browser's own copy command, so the system clipboard gets the
 * shapes as Ctrl/Cmd+C would put them there; where the browser refuses that, they are kept in memory for pasting
 * here.
 */
export function copySelection(cut = false): void {
  const c = ctx();
  if (!c || !c.views.length) return;
  const before = st().clipId;
  let done = false;
  try {
    done = document.execCommand('copy') && st().clipId !== before;
  } catch {
    done = false;
  }
  if (!done) useEditor.setState({ clipboard: copyShapes(c.ed, c.part, c.views), clipId: null });
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

/** A lighter version of a colour, for the light end of a gradient: a theme colour by its tint, any other by mixing with white. */
function lighter(c: ColorSpec): ColorSpec {
  if (c.kind === 'scheme') return c.lumMod === undefined ? { kind: 'scheme', val: c.val, lumMod: 40000, lumOff: 60000 } : { kind: 'srgb', hex: 'FFFFFF' };
  const n = parseInt(c.hex.replace(/^#/, ''), 16);
  const mix = (v: number) => Math.round(v + (255 - v) * 0.6);
  return { kind: 'srgb', hex: [(n >> 16) & 255, (n >> 8) & 255, n & 255].map(v => mix(v).toString(16).padStart(2, '0')).join('').toUpperCase() };
}

/**
 * A gradient in the selection's own colour: from a lighter version of it to the colour itself, along an angle
 * (90 runs down). A shape with no colour of its own to start from (no fill, a picture) gets the theme's accent.
 */
export function gradientSelection(angle: number): void {
  const c = ctx();
  if (!c?.views.length) return;
  const to = c.views.map(fillColorOf).find(x => x) ?? { kind: 'scheme', val: 'accent1' };
  c.s.run(setGradientFill(c.ed, c.part, c.views, { from: lighter(to), to, angle }));
}

export function transparencySelection(percent: number): void {
  const c = ctx();
  if (c?.views.length) c.s.run(setFillTransparency(c.ed, c.part, c.views, percent));
}

export function shadowSelection(preset: 'bottom-right' | 'bottom' | 'center' | null): void {
  const c = ctx();
  if (c?.views.length) c.s.run(setShadow(c.ed, c.part, c.views, preset));
}

/** Fill the selected shapes with a picture from a file. */
export async function pictureFillSelection(file: File): Promise<void> {
  const c = ctx();
  if (!c?.views.length) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  const type = sniffImage(bytes, file.name);
  if (!type || type === 'image/svg+xml' || type === 'image/x-emf' || type === 'image/x-wmf') { c.s.notify('error', `${file.name} cannot fill a shape here.`); return; }
  const img = await decodeImage(bytes, type).catch(() => undefined);
  if (!img) { c.s.notify('error', `${file.name} could not be read as a picture.`); return; }
  const width = (img as { width?: number }).width ?? 960, height = (img as { height?: number }).height ?? 720;
  c.s.run(setPictureFill(c.ed, c.part, c.views, await prepareImage(c.ed, { bytes, contentType: type, ext: EXT[type]!, width, height })));
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

/** Copy the current slide, to paste here or in another deck opened in this tab. Returns the mark for the clipboard. */
export function copyCurrentSlide(): string | undefined {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return undefined;
  const id = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  useEditor.setState({ slideClip: { clip: copySlide(ed, s.slide), id } });
  return id;
}

/** Cut the current slide: copied, then deleted. The only slide of a deck stays where it is. */
export function cutCurrentSlide(): string | undefined {
  const ed = st().session?.ed;
  if (!ed || ed.doc.slideParts().length < 2) return undefined;
  const id = copyCurrentSlide();
  if (id) deleteCurrentSlide();
  return id;
}

/** Paste the copied slide after the current one, and go to it. */
export function pasteSlideClip(): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slideClip) return;
  const slides = ed.doc.slideParts();
  const r = pasteSlide(ed, s.slideClip.clip, s.slide ? slides.indexOf(s.slide) + 1 : slides.length);
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

/** Crop with handles: on for the one selected picture, or off again. */
export function toggleCrop(): void {
  const s = st();
  if (s.cropping) { useEditor.setState({ cropping: null }); return; }
  const pic = selectedPicture();
  if (pic) useEditor.setState({ cropping: pic.v.el, text: null });
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
