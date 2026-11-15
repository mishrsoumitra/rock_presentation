// apps/web/src/state/store.ts  -- the editor's state: the open deck, the slide on screen, selection, text editing
import { create } from 'zustand';
import { isEncryptedPackage, OfficeCryptoError } from '@pptx/crypto';
import { Editor, paragraphLength, paragraphs, prepareForSave, type CellRef, type Command, type FormatClip, type ShapeClip, type SlideClip, type RunPatch, type TextPos } from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { OpcError, OpcPackage } from '@pptx/opc';
import type { FontRegistry } from '@pptx/text';
import type { XmlElement } from '@pptx/xml';
import { autosave } from '../engine/autosave.ts';
import { loadFonts, registerCssFonts } from '../engine/fonts.ts';
import { ImageCache } from '../engine/images.ts';
import { SceneCache } from '../engine/scene.ts';
import { Thumbnails } from '../engine/thumbs.ts';
import { decryptInWorker } from '../engine/workers.ts';

/** What text editing knows: the shape, the caret and the other end of the selection, in file offsets. */
export interface TextEditing { part: string; el: XmlElement; caret: TextPos; anchor: TextPos; }

/**
 * The cells picked in the selected table: a block from `anchor` to `focus`. `editing` is typing in the anchor cell
 * (in the box laid over it); `seed` is a character typed to start that, added after the cell's text.
 */
export interface Cells { el: XmlElement; anchor: CellRef; focus: CellRef; editing: boolean; seed?: string; }

export type Pane = 'none' | 'animation' | 'selection' | 'comments' | 'spelling';

export interface Session {
  ed: Editor;
  scenes: SceneCache;
  images: ImageCache;
  thumbs: Thumbnails;
  fonts: FontRegistry;
  name: string;
  /** The password the deck was opened with: saving encrypts again with it. It stays in this tab. */
  password: string | null;
  /** The file on this computer the deck was opened from or last saved to, where the browser lets a page write files. */
  handle?: FileHandle;
}

/** The part of the File System Access API's file handle this editor uses. */
export interface FileHandle {
  name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<{ write(data: Uint8Array): Promise<void>; close(): Promise<void> }>;
}

export interface EditorState {
  session: Session | null;
  /** Bumped on every change, so components re-render; per-part versions decide what actually redraws. */
  version: number;
  /** How many edits (commands, undos, redos) the deck has had, and how many it had when it was last saved. */
  edits: number;
  savedEdits: number;
  slide: string | null;
  selection: XmlElement[];
  /** The group being edited inside (double-click on a group), or null at the top of the slide. */
  entered: XmlElement | null;
  text: TextEditing | null;
  /** Table cells picked in the selected table (see `activeCells`, which checks they still hold). */
  cells: Cells | null;
  /** Formatting to apply to the next typed text (Ctrl+B with a caret but no selection). */
  pending: RunPatch | null;
  zoom: number | 'fit';
  /** A preset picked from the shape gallery, waiting to be drawn by dragging on the slide. */
  drawing: string | null;
  mode: 'edit' | 'present';
  /** The slide show is shown with the speaker's view around it: next slide, notes, timer. */
  presenterView: boolean;
  /** Rulers along the slide area, and the guides laid on the slide (EMU from its left and top; this session only). */
  showRuler: boolean;
  showGuides: boolean;
  guides: { x: number[]; y: number[] };
  /** What fills the window while editing: one slide with the panel beside it, or every slide in a grid. */
  view: 'normal' | 'grid' | 'master';
  /** Sections folded shut in the slide panel, by section id. */
  folded: string[];
  /** In master view: the slide that was showing, to go back to. (`slide` then names a layout or a master.) */
  masterReturn: string | null;
  pane: Pane;
  /** The right-click menu, where it was opened (window coordinates) and on what. */
  menu: { x: number; y: number; kind: 'shape' | 'slide' | 'thumb' } | null;
  /** The Header & Footer dialog is open. */
  footerEditing: boolean;
  /** The chart data dialog is open, for the selected chart. */
  chartEditing: boolean;
  /** The hyperlink dialog is open. */
  linking: boolean;
  /** The find and replace panel is open. */
  finding: boolean;
  clipboard: ShapeClip | null;
  /** The video or sound playing in the editor, where its picture is (it is the one selected shape while it plays). */
  playing: XmlElement | null;
  /** The picture being cropped with handles (it is the one selected shape while that lasts). */
  cropping: XmlElement | null;
  /** The look the format painter picked up last (kept for Ctrl+Alt+V after the painter has been used). */
  formatClip: FormatClip | null;
  /** The format painter is on: the next shape clicked takes the look; 'sticky' keeps it on until Escape. */
  painting: false | 'once' | 'sticky';
  /** A copied or cut slide, held apart from the deck, and the mark it carries on the system clipboard. */
  slideClip: { clip: SlideClip; id: string } | null;
  /** The mark the copied shapes carry on the system clipboard, so a paste of that clipboard brings them back whole. */
  clipId: string | null;
  message: { kind: 'info' | 'error'; text: string } | null;
  askPassword: { name: string; bytes: Uint8Array; retry: boolean } | null;

  open(bytes: Uint8Array, name: string, password?: string): Promise<void>;
  run(cmd: Command, after?: Partial<Pick<EditorState, 'selection' | 'slide' | 'text'>>): void;
  undo(): void;
  redo(): void;
  setSlide(part: string | null): void;
  select(els: XmlElement[]): void;
  setState(patch: Partial<EditorState>): void;
  notify(kind: 'info' | 'error', text: string): void;
  save(): Promise<Uint8Array | undefined>;
}

/** Selected elements still in the deck's trees after undo or redo (identity survives both). */
function attached(els: XmlElement[]): XmlElement[] {
  return els.filter(el => {
    let e: XmlElement | null = el;
    while (e?.parent) e = e.parent;
    return e?.local === 'sld' || e?.local === 'sldLayout' || e?.local === 'sldMaster';
  });
}

/**
 * Text editing after undo or redo: it goes on in the same box with a caret, and no selection, at a position the
 * text now has (the selection it had may name text that is gone). It ends when the box itself is gone.
 */
function fitted(s: Session, t: TextEditing | null): TextEditing | null {
  if (!t || !attached([t.el]).length) return null;
  const body = new ShapeView(t.el, s.ed.doc.owner(t.part)).txBody;
  if (!body) return null;
  const ps = paragraphs(body);
  const p = Math.max(0, Math.min(t.caret.p, ps.length - 1));
  const caret = { p, offset: ps[p] ? Math.min(t.caret.offset, paragraphLength(ps[p]!)) : 0 };
  return { ...t, caret, anchor: caret };
}

export const useEditor = create<EditorState>((set, get) => ({
  session: null, version: 0, edits: 0, savedEdits: 0, slide: null, selection: [], entered: null, text: null, cells: null, pending: null,
  zoom: 'fit', drawing: null, mode: 'edit', presenterView: false, view: 'normal', folded: [], masterReturn: null, showRuler: false, showGuides: true, guides: { x: [], y: [] }, pane: 'none', finding: false, linking: false, chartEditing: false, footerEditing: false, menu: null, clipboard: null, clipId: null, slideClip: null, formatClip: null, painting: false, cropping: null, playing: null, message: null, askPassword: null,

  async open(bytes, name, password) {
    performance.mark('open-start');
    let zip = bytes;
    try {
      if (isEncryptedPackage(bytes)) {
        if (!password) {
          set({ askPassword: { name, bytes, retry: false } });
          return;
        }
        zip = await decryptInWorker(bytes, password);
      }
      const [fonts] = await Promise.all([loadFonts(), registerCssFonts()]);
      const pkg = OpcPackage.open(zip);
      const ed = new Editor(pkg);
      get().session?.thumbs.dispose();
      const images = new ImageCache(() => ed.pkg, () => set(s => ({ version: s.version + 1 })));
      const thumbs = new Thumbnails(ed, zip, () => set(s => ({ version: s.version + 1 })));
      const session: Session = { ed, fonts, images, scenes: new SceneCache(ed, fonts), thumbs, name, password: password ?? null };
      ed.subscribe(parts => {
        // while typing, the thumbnail follows 300 ms after the last keystroke; any other change shows at once
        if (get().text) session.thumbs.syncLater(parts);
        else session.thumbs.sync(parts);
        set(s => ({ version: s.version + 1 }));
      });
      set({
        session, slide: ed.doc.slideParts()[0] ?? null, selection: [], entered: null, text: null, cells: null, pending: null, askPassword: null, guides: { x: [], y: [] },
        edits: 0, savedEdits: 0, view: 'normal', masterReturn: null, folded: [],
        version: get().version + 1,
        message: ed.mainKind === 'template' ? { kind: 'info', text: 'Opened a template: saving writes a presentation (.pptx).' }
          : ed.mainKind === 'macro' || pkg.partNames().some(n => /vbaProject\.bin$/i.test(n))
            ? { kind: 'info', text: 'This presentation contains macros. They are kept in the file but never run here.' } : null,
      });
      performance.mark('open-end');
    } catch (e) {
      if (e instanceof OfficeCryptoError && e.code === 'wrong-password') {
        set({ askPassword: { name, bytes, retry: true } });
        return;
      }
      const text = e instanceof OpcError || e instanceof OfficeCryptoError ? e.message : `This file could not be opened: ${e instanceof Error ? e.message : String(e)}`;
      set({ message: { kind: 'error', text }, askPassword: null });
    }
  },

  run(cmd, after) {
    const s = get().session;
    if (!s) return;
    try {
      s.ed.run(cmd);
      set(st => ({ ...(after ?? {}), edits: st.edits + 1 }));
    } catch (e) {
      get().notify('error', `${cmd.label} failed: ${e instanceof Error ? e.message : String(e)}`);
    }
  },

  undo() {
    const s = get().session;
    if (!s?.ed.undo()) return;
    s.thumbs.flush();
    set(st => ({ selection: attached(st.selection), text: fitted(s, st.text), pending: null, edits: st.edits + 1 }));
    const slides = s.ed.doc.slideParts();
    if (!slides.includes(get().slide ?? '')) set({ slide: slides[0] ?? null });
  },

  redo() {
    const s = get().session;
    if (!s?.ed.redo()) return;
    s.thumbs.flush();
    set(st => ({ selection: attached(st.selection), text: fitted(s, st.text), pending: null, edits: st.edits + 1 }));
    const slides = s.ed.doc.slideParts();
    if (!slides.includes(get().slide ?? '')) set({ slide: slides[0] ?? null });
  },

  setSlide(part) {
    const s = get().session;
    s?.ed.log.seal();
    s?.thumbs.flush();
    set({ slide: part, selection: [], entered: null, text: null, cells: null, cropping: null, playing: null });
  },

  select(els) {
    get().session?.ed.log.seal();
    set({ selection: els, text: null, cells: null, ...(get().cropping && !els.includes(get().cropping!) ? { cropping: null } : {}), ...(get().playing && !els.includes(get().playing!) ? { playing: null } : {}) });
  },

  setState(patch) { set(patch); },

  notify(kind, text) { set({ message: { kind, text } }); },

  async save() {
    const s = get().session;
    if (!s) return undefined;
    prepareForSave(s.ed, { asPresentation: true });
    const entries = s.ed.pkg.entries();
    const worker = new Worker(new URL('../engine/save.worker.ts', import.meta.url), { type: 'module' });
    try {
      const result = await new Promise<{ ok: boolean; bytes?: Uint8Array; error?: string }>(resolve => {
        worker.onmessage = e => resolve(e.data);
        worker.onerror = e => resolve({ ok: false, error: e.message });
        worker.postMessage({ entries, ...(s.password ? { password: s.password } : {}) });
      });
      if (!result.ok || !result.bytes) throw new Error(result.error ?? 'saving failed');
      return result.bytes;
    } finally {
      worker.terminate();
    }
  },
}));

// when editing ends, however it ends, the slide typed in reaches the thumbnail worker without the typing pause
useEditor.subscribe((st, prev) => { if (prev.text && !st.text) st.session?.thumbs.flush(); });

/** The session, for code outside React (event handlers). */
export const session = () => useEditor.getState().session;

/** Whether the deck has changes that have not been saved. */
export const unsaved = (s: Pick<EditorState, 'session' | 'edits' | 'savedEdits'> = useEditor.getState()) => !!s.session && s.edits !== s.savedEdits;

/** The deck has just been written somewhere: nothing is unsaved, and the copy kept for recovery is no longer needed. */
export function markSaved(): void {
  useEditor.setState(s => ({ savedEdits: s.edits }));
  void autosave.clear();
}

/**
 * Autosave: a few seconds after the last change the deck is written (as a save would write it, password and all)
 * to this browser's own database, where the welcome screen finds it if the tab is closed before a real save. The
 * wait grows with the time a save takes, so a very large deck is not being written all the time.
 */
let timer: ReturnType<typeof setTimeout> | undefined;
let writing = false;
let wait = 2500;
/** Decks larger than this are not autosaved: writing one would stall the editor again and again. */
export const AUTOSAVE_LIMIT = 48 << 20;
function tooLarge(s: Session): boolean {
  let total = 0;
  for (const n of s.ed.pkg.partNames()) {
    total += s.ed.pkg.getPart(n)?.bytes.length ?? 0;
    if (total > AUTOSAVE_LIMIT) return true;
  }
  return false;
}
async function keep(): Promise<void> {
  const s = useEditor.getState();
  if (!s.session || writing || !autosave.available()) return;
  if (!unsaved(s)) { void autosave.clear(); return; }
  if (tooLarge(s.session)) return;
  writing = true;
  const edits = s.edits, session = s.session, started = performance.now();
  try {
    const bytes = await s.save();
    // only when this is still the deck, and nothing saved it meanwhile
    if (bytes && useEditor.getState().session === session && unsaved()) await autosave.put({ name: session.name, bytes, at: Date.now(), encrypted: !!session.password });
  } catch {
    // a deck that cannot be written now is tried again after the next change
  } finally {
    writing = false;
    wait = Math.max(2500, Math.min(60_000, (performance.now() - started) * 10));
    if (useEditor.getState().edits !== edits) schedule();
  }
}
function schedule(): void {
  if (timer) clearTimeout(timer);
  timer = setTimeout(() => void keep(), wait);
}
useEditor.subscribe((st, prev) => { if (st.edits !== prev.edits || st.session !== prev.session) schedule(); });
/** Write the recovery copy now (tests, and before the page goes away where there is time). */
export const autosaveNow = () => { if (timer) clearTimeout(timer); return keep(); };

// leaving with unsaved changes asks first, as browsers let a page ask
if (typeof window !== 'undefined') {
  window.addEventListener('beforeunload', e => {
    if (!unsaved()) return;
    e.preventDefault();
    e.returnValue = '';
  });
}

