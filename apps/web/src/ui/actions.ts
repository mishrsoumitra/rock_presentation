// apps/web/src/ui/actions.ts  -- what the ribbon and the keyboard do: commands on the current selection
import {
  addSlide, align, applyLayout, compound, copyShapes, copySlide, pasteSlide, cropPicture, deleteShapes, deleteSlide, distribute, duplicateShapes, duplicateSlide, flip, group, importMasters, replacePicture,
  insertPicture, insertShape, insertTable, insertTextBox, moveBy, ROW_HEIGHT, moveSlide, pasteShapes, prepareImage, prepareSvg, prepareTemplate, rotateQuarter,
  addSection, applyBackgroundToAll, chartData, chartPartOf, equationOf, fillColorOf, headerFooter, insertChart, insertEquation, insertMedia, insertWordArt, setEquation, prepareMedia, setChartData, setFill, setHeaderFooter, setHeaderFooterAll, setSlideBackground, setThemeColors, setThemeFonts, themeOfSlide, themeSchemes, setFillTransparency, setGradientFill, setLine, setPictureFill, setShadow, setSlideHidden, setSlideSize, setTransforms, ungroup, zOrder,
  type Alignment, type BackgroundSpec, type ChartData, type ChartKind, type ColorScheme, type HeaderFooter, type WordArtStyle, type ColorSpec, type FontScheme, type LineSpec,
  type ScaleMode, type ZOrder,
} from '@pptx/editor-core';
import { OpcPackage } from '@pptx/opc';
import {
  connectionSites, copyFormat, copyTextFormat, frameBetween, nearestSite, paragraphLength, paragraphs, pasteFormat, setConnectorEnds, type Point,
} from '@pptx/editor-core';
import { runsOf } from '@pptx/model';
import { attr, child, ShapeView } from '@pptx/model';
import type { XmlElement } from '@pptx/xml';
import { blankPresentation } from '../engine/blank.ts';
import { decodeImage } from '../engine/images.ts';
import { autosave, type Autosaved } from '../engine/autosave.ts';
import { versions, type VersionInfo } from '../engine/versions.ts';
import { deckPdf, slidePng, slidesZip } from '../engine/export.ts';
import { audioIcon, mediaTypeOf, videoPoster } from '../engine/media.ts';
import { markSaved, useEditor, type FileHandle } from '../state/store.ts';
import { typeChars } from './textActions.ts';

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

/** The name a saved file gets: the deck's, as a .pptx (a template or a show opened here is saved as a presentation). */
function saveName(name: string): string {
  const n = name.replace(/\.(potx|ppsx|pptm|potm|ppsm)$/i, '.pptx');
  return /\.pptx$/i.test(n) ? n : `${n}.pptx`;
}

const PPTX_TYPE = { description: 'PowerPoint presentation', accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] } };
type Pickers = {
  showSaveFilePicker?: (o: { suggestedName: string; types: unknown[] }) => Promise<FileHandle>;
  showOpenFilePicker?: (o: { types: unknown[]; multiple: false }) => Promise<FileHandle[]>;
};
const pickers = () => window as unknown as Pickers;
const cancelled = (e: unknown) => e instanceof DOMException && e.name === 'AbortError';

async function deckBytes(): Promise<Uint8Array | undefined> {
  const s = st();
  return s.save().catch(e => { s.notify('error', `Saving failed: ${e instanceof Error ? e.message : String(e)}`); return undefined; });
}

/** Download the deck as a file: where the browser cannot write files itself, and for "Save a copy". */
export async function saveAndDownload(): Promise<void> {
  const s = st();
  const bytes = await deckBytes();
  if (!bytes || !s.session) return;
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = saveName(s.session.name);
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
  markSaved();
  await keepVersion(bytes, 'Saved', true);
  s.notify('info', `Saved ${a.download}${s.session.password ? ' (encrypted)' : ''}.`);
}

/**
 * Save (Ctrl/Cmd+S): to the file the deck was opened from or last saved to, with no questions; the first time, the
 * browser asks where. `as` asks where again (Save As). Browsers that cannot write files download a copy instead.
 */
export async function saveFile(as = false): Promise<void> {
  const s = st();
  const session = s.session;
  if (!session) return;
  let handle = as ? undefined : session.handle;
  if (!handle) {
    const pick = pickers().showSaveFilePicker;
    if (!pick) { await saveAndDownload(); return; }
    try {
      handle = await pick({ suggestedName: saveName(session.name), types: [PPTX_TYPE] });
    } catch (e) {
      if (!cancelled(e)) s.notify('error', `Saving failed: ${e instanceof Error ? e.message : String(e)}`);
      return;
    }
  }
  const bytes = await deckBytes();
  if (!bytes) return;
  try {
    const w = await handle.createWritable();
    await w.write(bytes);
    await w.close();
  } catch (e) {
    if (!cancelled(e)) s.notify('error', `${handle.name} could not be written: ${e instanceof Error ? e.message : String(e)}`);
    return;
  }
  session.handle = handle;
  session.name = handle.name;
  markSaved();
  await keepVersion(bytes, 'Saved', true);
  useEditor.setState(x => ({ version: x.version + 1 }));
  s.notify('info', `Saved ${handle.name}${session.password ? ' (encrypted)' : ''}.`);
}

/**
 * Open with the browser's own file picker where it has one, so Save can write back to the same file. Returns false
 * where it has none: the caller then uses an ordinary file input.
 */
export async function openWithPicker(): Promise<boolean> {
  const pick = pickers().showOpenFilePicker;
  if (!pick) return false;
  try {
    const [handle] = await pick({ multiple: false, types: [{ description: 'PowerPoint files', accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx', '.pptm', '.potx', '.ppsx'] } }] });
    if (!handle) return true;
    const file = await handle.getFile();
    await st().open(new Uint8Array(await file.arrayBuffer()), file.name);
    const now = st().session;
    if (now && now.name === file.name) now.handle = handle;
  } catch (e) {
    if (!cancelled(e)) st().notify('error', `That file could not be opened: ${e instanceof Error ? e.message : String(e)}`);
  }
  return true;
}

/** Hand a file made here to the browser to save. */
function download(bytes: Uint8Array, name: string, type: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/**
 * Export from the File tab: the deck as a PDF, the current slide as a PNG, or every slide as PNGs in a ZIP. The
 * deck itself is not changed or saved by this.
 */
export async function exportAs(what: 'pdf' | 'png' | 'pngs'): Promise<void> {
  const s = st();
  const session = s.session;
  if (!session || !s.slide) return;
  const base = session.name.replace(/\.[^.]+$/, '');
  const progress = (done: number, of: number) => s.notify('info', `Exporting slide ${done} of ${of}…`);
  try {
    if (what === 'png') {
      const n = session.ed.doc.slideParts().indexOf(s.slide) + 1;
      download(await slidePng(session, s.slide), `${base} - Slide ${n}.png`, 'image/png');
      s.notify('info', `Exported slide ${n} as a picture.`);
    } else if (what === 'pngs') {
      download(await slidesZip(session, progress), `${base} - slides.zip`, 'application/zip');
      s.notify('info', `Exported ${session.ed.doc.slideParts().length} slides as pictures.`);
    } else {
      download(await deckPdf(session, progress), `${base}.pdf`, 'application/pdf');
      s.notify('info', `Exported ${base}.pdf.`);
    }
  } catch (e) {
    s.notify('error', `The export failed: ${e instanceof Error ? e.message : String(e)}`);
  }
}

/** The deck kept by autosave, if there is one to offer on the welcome screen. */
export const recoverable = (): Promise<Autosaved | undefined> => autosave.get();

// ---------------------------------------------------------------------------------------- version history

/** Keep these bytes as a version of the deck being edited, and tell the Version History pane. */
async function keepVersion(bytes: Uint8Array, label: string, auto: boolean): Promise<number | undefined> {
  const session = st().session;
  if (!session) return undefined;
  const id = await versions.add({ deck: session.name, label, at: Date.now(), bytes, encrypted: !!session.password, slides: session.ed.doc.slideParts().length, auto });
  if (id !== undefined) useEditor.setState(x => ({ versionsTick: x.versionsTick + 1 }));
  return id;
}

/** The versions kept of the deck being edited, newest first. */
export const listVersions = (): Promise<VersionInfo[]> => { const s = st().session; return s ? versions.list(s.name) : Promise.resolve([]); };

/** Save Version: the deck as it is now, under a name (or "Version" with nothing given). It is not a save to the file. */
export async function saveVersion(name?: string): Promise<void> {
  const s = st();
  if (!s.session) return;
  const bytes = await deckBytes();
  if (!bytes) return;
  const id = await keepVersion(bytes, name?.trim() || 'Version', false);
  s.notify(id === undefined ? 'error' : 'info', id === undefined ? 'This version could not be kept: the presentation is too large, or the browser has no room.' : `Kept version "${name?.trim() || 'Version'}".`);
}

/**
 * Restore a version: it becomes the deck in the editor, as unsaved changes (the file is not touched until Save).
 * What was in the editor is kept as a version first, so restoring can itself be gone back on.
 */
export async function restoreVersion(id: number): Promise<void> {
  const s = st();
  const v = await versions.get(id);
  if (!v || !s.session) return;
  const name = s.session.name;
  const before = await deckBytes();
  if (before) await keepVersion(before, 'Before restoring', true);
  await st().open(v.bytes, name);
  if (st().session) {
    useEditor.setState(x => ({ edits: x.edits + 1, pane: 'versions', versionsTick: x.versionsTick + 1 }));
    st().notify('info', `Restored "${v.label}". Save to keep it.`);
  }
}

export async function downloadVersion(id: number): Promise<void> {
  const v = await versions.get(id);
  if (!v) return;
  const url = URL.createObjectURL(new Blob([v.bytes as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
  const a = document.createElement('a');
  a.href = url;
  // the version's own name beside the deck's, in characters a file name can have
  a.download = saveName(v.deck).replace(/\.pptx$/i, '') + ` (${v.label.replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() || 'version'}).pptx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

export async function deleteVersion(id: number): Promise<void> {
  await versions.remove(id);
  useEditor.setState(x => ({ versionsTick: x.versionsTick + 1 }));
}

/** Open the deck autosave kept. It opens with unsaved changes (it was never saved), and stays kept until it is. */
export async function restoreAutosaved(): Promise<void> {
  const kept = await autosave.get();
  if (!kept) return;
  await st().open(kept.bytes, kept.name);
  if (st().session) useEditor.setState(x => ({ edits: x.edits + 1 }));
}

export const discardAutosaved = (): Promise<void> => autosave.clear();

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

/** The shapes of the current slide a connector can attach to: its own top-level shapes that have a frame. */
export function attachable(): ShapeView[] {
  const s = st();
  if (!s.session || !s.slide) return [];
  return s.session.ed.doc.slide(s.slide).shapes.filter(v => connectionSites(v).length > 0);
}

/**
 * Draw a line or connector from one point to another, as it was dragged (so it keeps its direction). An end let go
 * within `reach` of a shape's connection point is attached to it, and then follows that shape.
 */
export function insertConnectorBetween(prst: string, from: Point, to: Point, reach: number): void {
  const c = ctx();
  if (!c) return;
  const shapes = attachable();
  const a = nearestSite(shapes, from, reach), b = nearestSite(shapes, to, reach);
  const start = a?.site ?? from, end = b?.site ?? to;
  const f = frameBetween(start, end);
  const r = insertShape(c.ed, c.part, prst, { x: f.x, y: f.y, w: f.cx, h: f.cy });
  c.s.run(r.command);
  const view = c.ed.shape(c.part, r.id);
  if (view) {
    c.s.run(setConnectorEnds(c.ed, c.part, view, start, end, {
      ...(a ? { start: { id: a.shape.id, index: a.site.index } } : {}), ...(b ? { end: { id: b.shape.id, index: b.site.index } } : {}),
    }));
  }
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

/** The equation of the one selected shape, on one line, when it holds one. */
export function selectedEquation(): string | undefined {
  const c = ctx();
  return c?.views.length === 1 ? equationOf(c.views[0]!) : undefined;
}

/** The Equation dialog's Insert or Apply: change the selected equation, or put a new one in the middle of the slide. */
export function applyEquation(text: string): void {
  const c = ctx();
  if (!c || !text.trim()) return;
  useEditor.setState({ text: null });
  if (c.views.length === 1 && equationOf(c.views[0]!) !== undefined) {
    c.s.run(setEquation(c.ed, c.part, c.views[0]!, text));
    return;
  }
  const size = c.ed.doc.presentation.slideSize;
  // roomy for its text; the box fits itself to the equation when it is drawn
  const w = Math.min(size.cx * 0.9, Math.max(1200000, text.trim().length * 190000)), h = 461665;
  const r = insertEquation(c.ed, c.part, text, { x: (size.cx - w) / 2, y: (size.cy - h) / 2, w, h });
  c.s.run(r.command);
  c.s.select(inserted(r.id));
}

/**
 * Insert WordArt in a look from the gallery: in the middle of the slide, its sample text selected so that typing
 * replaces it, as PowerPoint leaves it.
 */
export function insertWordArtHere(style: WordArtStyle): void {
  const c = ctx();
  if (!c) return;
  const size = c.ed.doc.presentation.slideSize;
  const w = 5486400, h = 923330;
  const text = 'Your text here';
  const r = insertWordArt(c.ed, c.part, style, { x: (size.cx - w) / 2, y: (size.cy - h) / 2, w, h }, text);
  useEditor.setState({ text: null });
  c.s.run(r.command);
  const els = inserted(r.id);
  if (els[0]) useEditor.setState({ selection: els, text: { part: c.part, el: els[0], anchor: { p: 0, offset: 0 }, caret: { p: 0, offset: text.length } } });
}

/** Today's date as PowerPoint's first date format writes it: month, day, year. */
export const todayText = (now = new Date()): string => `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;

/** What the current slide shows at its foot, for the Header & Footer dialog. */
export function footerHere(): HeaderFooter | undefined {
  const c = ctx();
  return c ? headerFooter(c.ed, c.part) : undefined;
}

/** Header & Footer's Apply (this slide) and Apply to All (every slide, title slides left bare if asked). */
export function applyFooter(spec: HeaderFooter, all: boolean, skipTitle: boolean): void {
  const c = ctx();
  if (!c) return;
  useEditor.setState({ text: null });
  const title = (() => { const l = c.ed.doc.layoutOf(c.part); return !!l && attr(c.ed.doc.root(l), 'type') === 'title'; })();
  const none: HeaderFooter = { date: null, footer: null, slideNumber: false };
  c.s.run(all ? setHeaderFooterAll(c.ed, spec, skipTitle) : setHeaderFooter(c.ed, c.part, skipTitle && title ? none : spec));
  c.s.select([]);
}

/**
 * Insert a character from the Symbol menu: at the caret when text is being typed (on the slide, in a table cell,
 * in the notes), and otherwise in a new text box in the middle of the slide, the caret left after it.
 */
export function insertSymbol(char: string): void {
  const s = st();
  if (s.text) { typeChars(char); return; }
  const focus = typeof document === 'undefined' ? null : document.activeElement;
  if (focus instanceof HTMLTextAreaElement || (focus instanceof HTMLInputElement && focus.type === 'text')) {
    document.execCommand('insertText', false, char);                     // as if typed: the field's own undo and change events
    return;
  }
  const c = ctx();
  if (!c) return;
  const size = c.ed.doc.presentation.slideSize;
  const r = insertTextBox(c.ed, c.part, { x: size.cx / 2 - 1524000, y: size.cy / 2 - 184666, w: 3048000, h: 369332 }, char);
  c.s.run(r.command);
  const els = inserted(r.id);
  const at = { p: 0, offset: char.length };
  if (els[0]) useEditor.setState({ selection: els, text: { part: c.part, el: els[0], caret: at, anchor: at } });
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

/**
 * Insert a video or a sound from a file. A video comes in at its own size (fitted to the slide) with a frame of
 * it as its picture; a sound as a small speaker. A file whose type is not one PowerPoint plays is refused.
 */
export async function insertMediaFile(file: File): Promise<void> {
  const c = ctx();
  if (!c) return;
  const what = mediaTypeOf(file.name);
  if (!what) { c.s.notify('error', `${file.name} is not a video or sound this editor can insert.`); return; }
  const bytes = new Uint8Array(await file.arrayBuffer());
  const poster = what.kind === 'video' ? await videoPoster(bytes, what.type) : await audioIcon();
  const media = await prepareMedia(c.ed, { bytes, contentType: what.type, ext: what.ext });
  const image = await prepareImage(c.ed, { bytes: poster.bytes, contentType: 'image/png', ext: 'png', width: poster.width, height: poster.height });
  const size = c.ed.doc.presentation.slideSize;
  const box = what.kind === 'audio' ? { x: (size.cx - 609600) / 2, y: (size.cy - 609600) / 2, w: 609600, h: 609600 } : undefined;
  const r = insertMedia(c.ed, c.part, what.kind, media, image, box, file.name);
  c.s.run(r.command);
  c.s.select(inserted(r.id));
}

/** The media the one selected shape plays, if it plays any that is stored in the deck. */
export function selectedMedia(): { el: XmlElement; kind: 'video' | 'audio'; part: string } | undefined {
  const c = ctx();
  const v = c?.views.length === 1 ? c.views[0] : undefined;
  const m = v?.media;
  const rId = m ? m.embed ?? m.link : undefined;
  const part = c && rId ? c.ed.doc.target(c.part, rId) : undefined;
  return v && m && part && c!.ed.pkg.getPart(part) ? { el: v.el, kind: m.kind, part } : undefined;
}

/** Play the selected video or sound where it is, or stop it. */
export function togglePlaying(): void {
  const s = st();
  if (s.playing) { useEditor.setState({ playing: null }); return; }
  const m = selectedMedia();
  if (m) useEditor.setState({ playing: m.el, text: null });
}

/** Insert a chart of a kind in the middle of the slide, with PowerPoint's sample data, and open its data to change. */
export function insertChartHere(kind: ChartKind): void {
  const c = ctx();
  if (!c) return;
  const size = c.ed.doc.presentation.slideSize;
  const w = Math.round((size.cx * 2) / 3), h = Math.round(size.cy * 0.6);
  const r = insertChart(c.ed, c.part, kind, { x: (size.cx - w) / 2, y: (size.cy - h) / 2, w, h });
  useEditor.setState({ text: null });
  c.s.run(r.command);
  c.s.select(inserted(r.id));
  useEditor.setState({ chartEditing: true });
}

/** The one selected shape, when it is a chart. */
function selectedChart() {
  const c = ctx();
  const v = c?.views.length === 1 ? c.views[0] : undefined;
  return c && v && chartPartOf(c.ed, c.part, v) ? { c, v } : undefined;
}
export const chartSelected = (): boolean => !!selectedChart();

/** The selected chart's data as a table, for the data dialog. */
export function selectedChartData(): ChartData | undefined {
  const x = selectedChart();
  return x ? chartData(x.c.ed, x.c.part, x.v) : undefined;
}

export function applyChartData(data: ChartData): void {
  const x = selectedChart();
  if (x) x.c.s.run(setChartData(x.c.ed, x.c.part, x.v, data));
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

/** Open the comments pane with the keyboard in its box for a new comment. */
export function newComment(): void {
  useEditor.setState({ pane: 'comments' });
  setTimeout(() => window.dispatchEvent(new CustomEvent('focus-new-comment')), 0);
}

/** Select every shape of the slide itself (not the layout's or master's), as Ctrl+A does. */
export function selectAllShapes(): void {
  const s = st();
  if (!s.session || !s.slide) return;
  const scene = s.session.scenes.get(s.slide);
  s.select(scene.slide.shapes.filter(x => x.layer === (scene.slide.own ?? 'slide') && x.source).map(x => x.source.view.el));
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

/** Start a section at the current slide. */
export function addSectionHere(): void {
  const s = st();
  if (!s.session || !s.slide || s.view === 'master') return;
  s.run(addSection(s.session.ed, s.slide));
}

/**
 * Slide Master view: the masters and their layouts are edited as slides are. It opens on the layout of the slide
 * that was showing, and closing it goes back to that slide. While it is open, commands that act on slides (new,
 * delete, duplicate, hide, layout) do nothing.
 */
export function openMasterView(): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || s.view === 'master') return;
  const layout = s.slide ? ed.doc.layoutOf(s.slide) : undefined;
  const first = ed.doc.masterParts()[0];
  const part = layout ?? first;
  if (!part) return;
  ed.log.seal();
  s.session!.thumbs.flush();
  useEditor.setState({ view: 'master', masterReturn: s.slide, slide: part, selection: [], entered: null, text: null, cells: null, cropping: null, pane: s.pane === 'selection' ? 'selection' : 'none' });
}

export function closeMasterView(): void {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || s.view !== 'master') return;
  const slides = ed.doc.slideParts();
  ed.log.seal();
  useEditor.setState({ view: 'normal', masterReturn: null, slide: s.masterReturn && slides.includes(s.masterReturn) ? s.masterReturn : slides[0] ?? null, selection: [], entered: null, text: null, cells: null, cropping: null });
}

/** The masters and, under each, its layouts: what master view lists. */
export function masterPages(): { part: string; kind: 'master' | 'layout'; name: string }[] {
  const ed = st().session?.ed;
  if (!ed) return [];
  return ed.doc.masterParts().flatMap((m, i) => [
    { part: m, kind: 'master' as const, name: ed.doc.slide(m).name || `Slide Master${i ? ` ${i + 1}` : ''}` },
    ...ed.doc.layoutsOf(m).map(l => ({ part: l, kind: 'layout' as const, name: ed.doc.slide(l).name || 'Layout' })),
  ]);
}

export function newSlide(layout?: string): void {
  if (st().view === 'master') return;
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
  if (st().view === 'master') return;
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
  if (st().view === 'master') return;
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  const r = duplicateSlide(ed, s.slide);
  s.run(r.command);
  if (r.part) s.setSlide(r.part);
}

/** Copy the current slide, to paste here or in another deck opened in this tab. Returns the mark for the clipboard. */
export function copyCurrentSlide(): string | undefined {
  if (st().view === 'master') return;
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return undefined;
  const id = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  useEditor.setState({ slideClip: { clip: copySlide(ed, s.slide), id } });
  return id;
}

/** Cut the current slide: copied, then deleted. The only slide of a deck stays where it is. */
export function cutCurrentSlide(): string | undefined {
  if (st().view === 'master') return;
  const ed = st().session?.ed;
  if (!ed || ed.doc.slideParts().length < 2) return undefined;
  const id = copyCurrentSlide();
  if (id) deleteCurrentSlide();
  return id;
}

/** Paste the copied slide after the current one, and go to it. */
export function pasteSlideClip(): void {
  if (st().view === 'master') return;
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slideClip) return;
  const slides = ed.doc.slideParts();
  const r = pasteSlide(ed, s.slideClip.clip, s.slide ? slides.indexOf(s.slide) + 1 : slides.length);
  s.run(r.command);
  if (r.part) s.setSlide(r.part);
}

export function moveCurrentSlide(to: number, part = st().slide): void {
  if (st().view === 'master') return;
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !part) return;
  s.run(moveSlide(ed, part, to));
}

export function toggleHidden(): void {
  if (st().view === 'master') return;
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  s.run(setSlideHidden(ed, s.slide, !ed.doc.slide(s.slide).hidden));
}

export function layoutCurrentSlide(layout: string): void {
  if (st().view === 'master') return;
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  s.run(applyLayout(ed, s.slide, layout), { selection: [] });
}

/** The colour and font schemes of the theme the current slide uses, by name. */
export function currentSchemes(): { colors: string; fonts: string } | undefined {
  const s = st();
  const theme = s.session && s.slide ? themeOfSlide(s.session.ed, s.slide) : undefined;
  return theme ? themeSchemes(s.session!.ed, theme) : undefined;
}

/** Give the current slide's theme another colour scheme: every slide of its master follows. */
export function chooseThemeColors(scheme: ColorScheme): void {
  const s = st();
  const theme = s.session && s.slide ? themeOfSlide(s.session.ed, s.slide) : undefined;
  if (theme) s.run(setThemeColors(s.session!.ed, theme, scheme));
}

/** Give the current slide's theme another pair of fonts. */
export function chooseThemeFonts(pair: FontScheme): void {
  const s = st();
  const theme = s.session && s.slide ? themeOfSlide(s.session.ed, s.slide) : undefined;
  if (theme) s.run(setThemeFonts(s.session!.ed, theme, pair));
}

/** The current slide's background: a colour, a gradient, or (null) the layout's again. */
export function backgroundCurrentSlide(bg: BackgroundSpec | null): void {
  const s = st();
  const ed = s.session?.ed;
  if (ed && s.slide) s.run(setSlideBackground(ed, [s.slide], bg));
}

/** The colour of the current slide's own solid background, when it has one. */
function ownBackgroundColor(): ColorSpec | undefined {
  const s = st();
  if (!s.session || !s.slide) return undefined;
  const fill = child(child(child(child(s.session.ed.doc.root(s.slide), 'p:cSld'), 'p:bg'), 'p:bgPr'), 'a:solidFill');
  const srgb = child(fill, 'a:srgbClr'), scheme = child(fill, 'a:schemeClr');
  if (srgb) return { kind: 'srgb', hex: attr(srgb, 'val') ?? 'FFFFFF' };
  return scheme && !scheme.children.length ? { kind: 'scheme', val: attr(scheme, 'val') ?? 'accent1' } : undefined;
}

/** A gradient background from a lighter tint of the slide's background colour (the theme's accent when it has none) to the colour. */
export function gradientBackground(angle: number): void {
  const to = ownBackgroundColor() ?? { kind: 'scheme', val: 'accent1' };
  backgroundCurrentSlide({ gradient: { from: lighter(to), to, angle } });
}

/** A picture from a file as the current slide's background. */
export async function pictureBackground(file: File): Promise<void> {
  const s = st();
  const ed = s.session?.ed;
  if (!ed || !s.slide) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  const type = sniffImage(bytes, file.name);
  if (!type || type === 'image/svg+xml' || type === 'image/x-emf' || type === 'image/x-wmf') { s.notify('error', `${file.name} cannot be a background here.`); return; }
  const img = await decodeImage(bytes, type).catch(() => undefined);
  if (!img) { s.notify('error', `${file.name} could not be read as a picture.`); return; }
  const width = (img as { width?: number }).width ?? 960, height = (img as { height?: number }).height ?? 720;
  backgroundCurrentSlide({ picture: await prepareImage(ed, { bytes, contentType: type, ext: EXT[type]!, width, height }) });
}

/** Every slide gets the current slide's background. */
export function backgroundToAll(): void {
  const s = st();
  const ed = s.session?.ed;
  if (ed && s.slide) s.run(applyBackgroundToAll(ed, s.slide));
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
