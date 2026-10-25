// apps/web/src/ui/Dialogs.tsx  -- password prompts and the slide size dialog
import { useEffect, useState } from 'react';
import { useEditor } from '../state/store.ts';
import { resizeSlides } from './actions.ts';

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
