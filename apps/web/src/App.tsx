// apps/web/src/App.tsx  -- the editor's window: ribbon, slide panel, the slide, status bar
import { useEffect, useState } from 'react';
import { useEditor } from './state/store.ts';
import { newPresentation, openFile } from './ui/actions.ts';
import { NewPasswordDialog, OpenPasswordDialog, SlideSizeDialog } from './ui/Dialogs.tsx';
import { Presenter } from './ui/Presenter.tsx';
import { AnimationPane, AnimationsTab, TransitionsTab } from './ui/AnimationTabs.tsx';
import { Ribbon } from './ui/Ribbon.tsx';
import { SlideCanvas } from './ui/SlideCanvas.tsx';
import { SlidePanel } from './ui/SlidePanel.tsx';
import { useShortcuts } from './ui/shortcuts.ts';

function StatusBar() {
  const session = useEditor(s => s.session);
  const zoom = useEditor(s => s.zoom);
  const message = useEditor(s => s.message);
  // "Slide n of m" changes with the presentation part (the list of slides) and the current slide, not with typing
  const place = useEditor(s => {
    if (!s.session) return '';
    const parts = s.session.ed.doc.slideParts();
    return `${s.slide ? parts.indexOf(s.slide) + 1 : 0}/${parts.length}`;
  });
  const [n, count] = place.split('/').map(Number);
  const setZoom = (z: number | 'fit') => useEditor.setState({ zoom: z });
  const current = typeof zoom === 'number' ? zoom * 914400 / 96 : undefined;
  return (
    <footer className="status">
      <span data-testid="slide-count">{session ? `Slide ${n} of ${count}` : 'No presentation'}</span>
      {message && <span className={`message ${message.kind}`} role={message.kind === 'error' ? 'alert' : 'status'}>{message.text}
        <button type="button" aria-label="Dismiss" onClick={() => useEditor.setState({ message: null })}>×</button></span>}
      <span className="zoom">
        <button type="button" onClick={() => setZoom('fit')} aria-pressed={zoom === 'fit'}>Fit</button>
        {[0.5, 0.75, 1, 1.5, 2].map(z => <button key={z} type="button" aria-pressed={current !== undefined && Math.abs(current - z) < 0.01} onClick={() => setZoom((z * 96) / 914400)}>{z * 100}%</button>)}
      </span>
    </footer>
  );
}

function Welcome() {
  const [over, setOver] = useState(false);
  return (
    <div className={`welcome${over ? ' over' : ''}`}
      onDragOver={e => { e.preventDefault(); setOver(true); }} onDragLeave={() => setOver(false)}
      onDrop={e => { e.preventDefault(); setOver(false); const f = e.dataTransfer.files[0]; if (f) void openFile(f); }}>
      <h1>Presentations</h1>
      <p>Open a PowerPoint file (.pptx, .potx, .ppsx), or drop one here. Everything stays in this browser tab.</p>
      <div className="welcome-buttons">
        <button type="button" onClick={() => void newPresentation()} data-testid="welcome-new">New presentation</button>
        <label className="button">Open…<input type="file" hidden accept=".pptx,.pptm,.potx,.ppsx,.ppt" data-testid="welcome-open"
          onChange={e => { const f = e.target.files?.[0]; if (f) void openFile(f); }} /></label>
      </div>
    </div>
  );
}

export function App() {
  const session = useEditor(s => s.session);
  const mode = useEditor(s => s.mode);
  const pane = useEditor(s => s.pane);
  useShortcuts();
  useEffect(() => {
    // a file dropped anywhere opens it
    const prevent = (e: DragEvent) => { if (e.dataTransfer?.types.includes('Files')) e.preventDefault(); };
    const drop = (e: DragEvent) => {
      const f = e.dataTransfer?.files[0];
      if (!f) return;
      e.preventDefault();
      if (/\.(pptx|pptm|potx|ppsx|ppt)$/i.test(f.name)) void openFile(f);
    };
    window.addEventListener('dragover', prevent);
    window.addEventListener('drop', drop);
    return () => { window.removeEventListener('dragover', prevent); window.removeEventListener('drop', drop); };
  }, []);
  if (session && mode === 'present') return <Presenter />;
  return (
    <div className="app">
      <Ribbon extraTabs={{ Transitions: <TransitionsTab />, Animations: <AnimationsTab /> }} />
      <main className={`workspace${session && pane === 'animation' ? ' with-pane' : ''}`}>
        {session ? <><SlidePanel /><SlideCanvas />{pane === 'animation' && <AnimationPane />}</> : <Welcome />}
      </main>
      <StatusBar />
      <OpenPasswordDialog />
      <NewPasswordDialog />
      <SlideSizeDialog />
    </div>
  );
}
