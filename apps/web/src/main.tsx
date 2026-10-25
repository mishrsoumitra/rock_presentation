// apps/web/src/main.tsx  -- start the editor
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App.tsx';
import { useEditor } from './state/store.ts';
import './styles.css';

createRoot(document.getElementById('root')!).render(<StrictMode><App /></StrictMode>);

// for end-to-end tests and debugging: the store, read-only in spirit
(globalThis as unknown as { __editor: typeof useEditor }).__editor = useEditor;
