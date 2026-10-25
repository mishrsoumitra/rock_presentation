// apps/web/src/ui/shortcuts.ts  -- keyboard shortcuts on the slide (text editing has its own, in TextInput)
import { useEffect } from 'react';
import { paragraphLength, paragraphs } from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { useEditor } from '../state/store.ts';
import * as act from './actions.ts';
import { typeChars } from './textActions.ts';

const NUDGE = 76200;          // 1/12 inch, PowerPoint's arrow-key step
const FINE = 12700;           // one point, with Ctrl or Alt

export function useShortcuts(): void {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const s = useEditor.getState();
      const target = e.target as HTMLElement | null;
      if (!s.session || s.mode !== 'edit' || s.text) return;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT' || target.closest('.slides'))) return;
      const mod = e.metaKey || e.ctrlKey;
      const k = e.key.toLowerCase();
      const handled = () => { e.preventDefault(); e.stopPropagation(); };
      if (mod && k === 'z') { handled(); if (e.shiftKey) s.redo(); else s.undo(); }
      else if (mod && k === 'y') { handled(); s.redo(); }
      else if (mod && k === 'c') { handled(); act.copySelection(); }
      else if (mod && k === 'x') { handled(); act.copySelection(true); }
      else if (mod && k === 'v') { handled(); void act.paste(); }
      else if (mod && k === 'd') { handled(); act.duplicateSelection(); }
      else if (mod && k === 'g') { handled(); if (e.shiftKey) act.ungroupSelection(); else act.groupSelection(); }
      else if (mod && k === 'm') { handled(); act.newSlide(); }
      else if (mod && k === 's') { handled(); void act.saveAndDownload(); }
      else if (mod && k === 'a') {
        handled();
        const scene = s.slide ? s.session.scenes.get(s.slide) : undefined;
        s.select(scene ? scene.slide.shapes.filter(x => x.layer === 'slide' && x.source).map(x => x.source.view.el) : []);
      }
      else if (e.key === 'Delete' || e.key === 'Backspace') { if (s.selection.length) { handled(); act.deleteSelection(); } }
      else if (e.key === 'Escape') { handled(); if (s.drawing) s.setState({ drawing: null }); else if (s.entered) s.setState({ entered: null, selection: [] }); else s.select([]); }
      else if (e.key.startsWith('Arrow') && s.selection.length) {
        handled();
        const d = mod || e.altKey ? FINE : NUDGE;
        act.nudge(e.key === 'ArrowLeft' ? -d : e.key === 'ArrowRight' ? d : 0, e.key === 'ArrowUp' ? -d : e.key === 'ArrowDown' ? d : 0);
      }
      else if ((e.key === 'Enter' || e.key === 'F2') && s.selection.length === 1 && s.slide) {
        // start editing the selected shape's text, caret at the end
        const v = new ShapeView(s.selection[0]!, s.session.ed.doc.owner(s.slide));
        const body = v.txBody;
        if (!body) return;
        handled();
        const ps = paragraphs(body);
        const end = { p: Math.max(0, ps.length - 1), offset: ps.length ? paragraphLength(ps[ps.length - 1]!) : 0 };
        s.setState({ text: { part: s.slide, el: v.el, caret: end, anchor: end } });
      }
      else if (e.key === 'Tab' && s.slide) {
        handled();
        const scene = s.session.scenes.get(s.slide);
        const els = scene.slide.shapes.filter(x => x.layer === 'slide' && x.source).map(x => x.source.view.el);
        if (!els.length) return;
        const i = els.indexOf(s.selection[0]!);
        s.select([els[(i + (e.shiftKey ? -1 : 1) + els.length) % els.length]!]);
      }
      else if (e.key.length === 1 && !mod && !e.altKey && s.selection.length === 1 && s.slide) {
        // typing with a shape selected adds to its text, as PowerPoint does
        const v = new ShapeView(s.selection[0]!, s.session.ed.doc.owner(s.slide));
        const body = v.txBody;
        if (!body) return;
        handled();
        const ps = paragraphs(body);
        const end = { p: Math.max(0, ps.length - 1), offset: ps.length ? paragraphLength(ps[ps.length - 1]!) : 0 };
        s.setState({ text: { part: s.slide, el: v.el, caret: end, anchor: end } });
        typeChars(e.key);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);
}
