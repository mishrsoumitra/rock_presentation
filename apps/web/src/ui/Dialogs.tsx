// apps/web/src/ui/Dialogs.tsx  -- password prompts and the slide size dialog
import { useEffect, useState } from 'react';
import { useEditor } from '../state/store.ts';
import type { LinkTarget } from '@pptx/editor-core';
import { applyFooter, footerHere, resizeSlides, todayText } from './actions.ts';
import { applyLink, linkContext, slideChoices } from './linkActions.ts';
import { normalizeUrl } from './links.ts';
import { refocusText } from './widgets.tsx';

function Dialog({ title, children, onClose }: { title: string; children: React.ReactNode; onClose: () => void }) {
  useEffect(() => {
    const esc = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', esc);
    return () => window.removeEventListener('keydown', esc);
  }, [onClose]);
  return (
    <div className="backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="dialog" role="dialog" aria-modal="true" aria-label={title}>
        <h2>{title}</h2>
        {children}
      </div>
    </div>
  );
}

/** Asking for the password of an encrypted deck. It is used here, in the browser, and never sent anywhere. */
export function OpenPasswordDialog() {
  const ask = useEditor(s => s.askPassword);
  const [pw, setPw] = useState('');
  const [busy, setBusy] = useState(false);
  if (!ask) return null;
  const close = () => useEditor.setState({ askPassword: null });
  return (
    <Dialog title="Password" onClose={close}>
      <p>{ask.retry ? 'The password is not correct. ' : ''}{ask.name} is protected with a password.</p>
      <form onSubmit={async e => {
        e.preventDefault();
        setBusy(true);
        await useEditor.getState().open(ask.bytes, ask.name, pw);
        setBusy(false);
        setPw('');
      }}>
        <input type="password" autoFocus aria-label="Password" value={pw} onChange={e => setPw(e.target.value)} data-testid="password" />
        <div className="dialog-buttons"><button type="button" onClick={close}>Cancel</button><button type="submit" disabled={busy || !pw}>{busy ? 'Opening…' : 'Open'}</button></div>
      </form>
    </Dialog>
  );
}

export function NewPasswordDialog() {
  const [open, setOpen] = useState(false);
  const [pw, setPw] = useState('');
  const [again, setAgain] = useState('');
  useEffect(() => {
    const on = () => setOpen(true);
    window.addEventListener('ask-new-password', on);
    return () => window.removeEventListener('ask-new-password', on);
  }, []);
  if (!open) return null;
  const close = () => { setOpen(false); setPw(''); setAgain(''); };
  return (
    <Dialog title="Encrypt Document" onClose={close}>
      <p>Saving encrypts the presentation with this password (AES-256, as PowerPoint does). If you lose it, the file cannot be recovered.</p>
      <form onSubmit={e => {
        e.preventDefault();
        const s = useEditor.getState().session;
        if (s && pw && pw === again) {
          s.password = pw;
          useEditor.getState().notify('info', 'The next save is encrypted with your password.');
          close();
        }
      }}>
        <input type="password" autoFocus aria-label="Password" value={pw} onChange={e => setPw(e.target.value)} />
        <input type="password" aria-label="Reenter password" placeholder="Reenter password" value={again} onChange={e => setAgain(e.target.value)} />
        <div className="dialog-buttons"><button type="button" onClick={close}>Cancel</button><button type="submit" disabled={!pw || pw !== again}>OK</button></div>
      </form>
    </Dialog>
  );
}

/** Slide size: standard sizes or a custom one, then PowerPoint's question: maximize, or ensure fit. */
export function SlideSizeDialog() {
  const [target, setTarget] = useState<{ cx: number; cy: number } | null | undefined>(undefined);
  const [w, setW] = useState('13.333');
  const [h, setH] = useState('7.5');
  useEffect(() => {
    const on = (e: Event) => {
      const d = (e as CustomEvent<{ cx: number; cy: number } | null>).detail;
      setTarget(d);
      const size = useEditor.getState().session?.ed.doc.presentation.slideSize;
      if (size) { setW((size.cx / 914400).toFixed(3)); setH((size.cy / 914400).toFixed(3)); }
    };
    window.addEventListener('slide-size', on);
    return () => window.removeEventListener('slide-size', on);
  }, []);
  if (target === undefined) return null;
  const close = () => setTarget(undefined);
  const size = target ?? { cx: Math.round(Number(w) * 914400), cy: Math.round(Number(h) * 914400) };
  const valid = size.cx >= 914400 && size.cy >= 914400 && size.cx <= 51206400 && size.cy <= 51206400;
  return (
    <Dialog title="Slide Size" onClose={close}>
      {target === null && (
        <div className="row">
          <label>Width (in) <input type="number" step="0.01" value={w} onChange={e => setW(e.target.value)} /></label>
          <label>Height (in) <input type="number" step="0.01" value={h} onChange={e => setH(e.target.value)} /></label>
        </div>
      )}
      <p>You are scaling to a new slide size. Do you want to maximize the size of your content, or scale it down to make sure it fits on the new slide?</p>
      <div className="dialog-buttons">
        <button type="button" onClick={close}>Cancel</button>
        <button type="button" disabled={!valid} onClick={() => { resizeSlides(size.cx, size.cy, 'maximize'); close(); }}>Maximize</button>
        <button type="button" disabled={!valid} onClick={() => { resizeSlides(size.cx, size.cy, 'ensureFit'); close(); }}>Ensure Fit</button>
      </div>
    </Dialog>
  );
}

/**
 * The hyperlink dialog (Insert > Link, Ctrl/Cmd+K, the right-click menu): a web or mail address, or a slide of this
 * deck, for the selected text or shapes. It opens showing the link that is there, to change or remove.
 */
export function LinkDialog() {
  const open = useEditor(s => s.linking);
  useEffect(() => {
    // Ctrl/Cmd+K anywhere in the editor, the text being typed in included
    const onKey = (e: KeyboardEvent) => {
      const s = useEditor.getState();
      if (!(e.metaKey || e.ctrlKey) || e.shiftKey || e.altKey || e.key.toLowerCase() !== 'k') return;
      if (!s.session || s.mode !== 'edit' || s.linking || !linkContext()) return;
      e.preventDefault();
      e.stopPropagation();
      useEditor.setState({ linking: true });
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, []);

  // the form is made anew each time the dialog opens, so it starts from the link that is there then
  return open ? <LinkForm /> : null;
}

function LinkForm() {
  const slides = slideChoices();
  const [target] = useState(() => linkContext()?.target);
  const [kind, setKind] = useState<'url' | 'slide'>(target && 'slide' in target ? 'slide' : 'url');
  const [url, setUrl] = useState(target && 'url' in target ? target.url : '');
  const [slide, setSlide] = useState(target && 'slide' in target ? target.slide : slides[0]?.part ?? '');
  const [error, setError] = useState('');
  const existing = !!target;
  const close = () => { useEditor.setState({ linking: false }); refocusText(); };
  const apply = () => {
    let target: LinkTarget;
    if (kind === 'url') {
      const address = normalizeUrl(url);
      if (!address) { setError('Enter a web address (example.com) or an email address.'); return; }
      target = { url: address };
    } else {
      if (!slides.some(x => x.part === slide)) return;
      target = { slide };
    }
    applyLink(target);
    close();
  };
  return (
    <Dialog title="Link" onClose={close}>
      <form data-testid="link-dialog" onSubmit={e => { e.preventDefault(); apply(); }}>
        <div className="row link-kinds">
          <label><input type="radio" name="link-kind" checked={kind === 'url'} onChange={() => setKind('url')} data-testid="link-kind-url" /> Web or email address</label>
          <label><input type="radio" name="link-kind" checked={kind === 'slide'} onChange={() => setKind('slide')} data-testid="link-kind-slide" /> Slide in this presentation</label>
        </div>
        {kind === 'url'
          ? <input type="text" autoFocus aria-label="Address" placeholder="example.com" value={url} data-testid="link-url" onChange={e => { setUrl(e.target.value); setError(''); }} />
          : <select aria-label="Slide" value={slide} onChange={e => setSlide(e.target.value)} data-testid="link-slide">{slides.map(x => <option key={x.part} value={x.part}>{x.label}</option>)}</select>}
        {error && <p className="dialog-error" role="alert" data-testid="link-error">{error}</p>}
        <div className="dialog-buttons">
          {existing && <button type="button" className="plain" onClick={() => { applyLink(null); close(); }} data-testid="link-remove">Remove Link</button>}
          <button type="button" onClick={close}>Cancel</button>
          <button type="submit" data-testid="link-apply">Apply</button>
        </div>
      </form>
    </Dialog>
  );
}

/**
 * Header & Footer: the date (kept up to date, or fixed text), the slide number and the footer text, for this slide
 * (Apply) or every slide (Apply to All). It opens showing what the current slide has.
 */
export function FooterDialog() {
  const open = useEditor(s => s.footerEditing);
  return open ? <FooterForm /> : null;
}

function FooterForm() {
  const [now] = useState(() => footerHere() ?? { date: null, footer: null, slideNumber: false });
  const [date, setDate] = useState(now.date !== null);
  const [auto, setAuto] = useState(now.date ? now.date.auto : true);
  const [fixed, setFixed] = useState(now.date && !now.date.auto ? now.date.text : '');
  const [number, setNumber] = useState(now.slideNumber);
  const [footer, setFooter] = useState(now.footer !== null);
  const [text, setText] = useState(now.footer ?? '');
  const [skipTitle, setSkipTitle] = useState(false);
  const close = () => useEditor.setState({ footerEditing: false });
  const apply = (all: boolean) => {
    applyFooter({ date: date ? { auto, text: auto ? todayText() : fixed } : null, footer: footer ? text : null, slideNumber: number }, all, skipTitle);
    close();
  };
  return (
    <Dialog title="Header & Footer" onClose={close}>
      <form className="footer-dialog" data-testid="footer-dialog" onSubmit={e => { e.preventDefault(); apply(false); }}>
        <label><input type="checkbox" checked={date} onChange={e => setDate(e.target.checked)} data-testid="footer-date" /> Date and time</label>
        <div className="footer-sub">
          <label><input type="radio" name="footer-date-kind" disabled={!date} checked={auto} onChange={() => setAuto(true)} data-testid="footer-date-auto" /> Update automatically <span className="muted">({todayText()})</span></label>
          <label><input type="radio" name="footer-date-kind" disabled={!date} checked={!auto} onChange={() => setAuto(false)} data-testid="footer-date-fixed" /> Fixed</label>
          <input type="text" aria-label="Fixed date" disabled={!date || auto} value={fixed} onChange={e => setFixed(e.target.value)} data-testid="footer-date-text" />
        </div>
        <label><input type="checkbox" checked={number} onChange={e => setNumber(e.target.checked)} data-testid="footer-number" /> Slide number</label>
        <label><input type="checkbox" checked={footer} onChange={e => setFooter(e.target.checked)} data-testid="footer-footer" /> Footer</label>
        <div className="footer-sub">
          <input type="text" aria-label="Footer text" disabled={!footer} value={text} onChange={e => setText(e.target.value)} data-testid="footer-text" />
        </div>
        <label className="footer-title"><input type="checkbox" checked={skipTitle} onChange={e => setSkipTitle(e.target.checked)} data-testid="footer-skip-title" /> Don't show on title slide</label>
        <div className="dialog-buttons">
          <button type="button" onClick={close}>Cancel</button>
          <button type="button" onClick={() => apply(true)} data-testid="footer-apply-all">Apply to All</button>
          <button type="submit" data-testid="footer-apply">Apply</button>
        </div>
      </form>
    </Dialog>
  );
}

