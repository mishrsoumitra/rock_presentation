// apps/web/src/state/store.ts  -- the editor's state: the open deck, the slide on screen, selection, text editing
import { create } from 'zustand';
import { isEncryptedPackage, OfficeCryptoError } from '@pptx/crypto';
import { Editor, paragraphLength, paragraphs, prepareForSave, type Command, type ShapeClip, type RunPatch, type TextPos } from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { OpcError, OpcPackage } from '@pptx/opc';
import type { FontRegistry } from '@pptx/text';
import type { XmlElement } from '@pptx/xml';
import { loadFonts, registerCssFonts } from '../engine/fonts.ts';
import { ImageCache } from '../engine/images.ts';
import { SceneCache } from '../engine/scene.ts';
import { Thumbnails } from '../engine/thumbs.ts';
import { decryptInWorker } from '../engine/workers.ts';

/** What text editing knows: the shape, the caret and the other end of the selection, in file offsets. */
export interface TextEditing { part: string; el: XmlElement; caret: TextPos; anchor: TextPos; }

export type Pane = 'none' | 'animation' | 'selection';

export interface Session {
  ed: Editor;
  scenes: SceneCache;
  images: ImageCache;
  thumbs: Thumbnails;
  fonts: FontRegistry;
  name: string;
  /** The password the deck was opened with: saving encrypts again with it. It stays in this tab. */
  password: string | null;
}

export interface EditorState {
  session: Session | null;
  /** Bumped on every change, so components re-render; per-part versions decide what actually redraws. */
  version: number;
  slide: string | null;
  selection: XmlElement[];
  /** The group being edited inside (double-click on a group), or null at the top of the slide. */
  entered: XmlElement | null;
  text: TextEditing | null;
  /** Formatting to apply to the next typed text (Ctrl+B with a caret but no selection). */
  pending: RunPatch | null;
  zoom: number | 'fit';
  /** A preset picked from the shape gallery, waiting to be drawn by dragging on the slide. */
  drawing: string | null;
  mode: 'edit' | 'present';
  pane: Pane;
  clipboard: ShapeClip | null;
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
  session: null, version: 0, slide: null, selection: [], entered: null, text: null, pending: null,
  zoom: 'fit', drawing: null, mode: 'edit', pane: 'none', clipboard: null, message: null, askPassword: null,

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
        session, slide: ed.doc.slideParts()[0] ?? null, selection: [], entered: null, text: null, pending: null, askPassword: null,
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
      if (after) set(after);
    } catch (e) {
      get().notify('error', `${cmd.label} failed: ${e instanceof Error ? e.message : String(e)}`);
    }
  },

  undo() {
    const s = get().session;
    if (!s?.ed.undo()) return;
    s.thumbs.flush();
    set(st => ({ selection: attached(st.selection), text: fitted(s, st.text), pending: null }));
    const slides = s.ed.doc.slideParts();
    if (!slides.includes(get().slide ?? '')) set({ slide: slides[0] ?? null });
  },

  redo() {
    const s = get().session;
    if (!s?.ed.redo()) return;
    s.thumbs.flush();
    set(st => ({ selection: attached(st.selection), text: fitted(s, st.text), pending: null }));
    const slides = s.ed.doc.slideParts();
    if (!slides.includes(get().slide ?? '')) set({ slide: slides[0] ?? null });
  },

  setSlide(part) {
    const s = get().session;
    s?.ed.log.seal();
    s?.thumbs.flush();
    set({ slide: part, selection: [], entered: null, text: null });
  },

  select(els) {
    get().session?.ed.log.seal();
    set({ selection: els, text: null });
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
