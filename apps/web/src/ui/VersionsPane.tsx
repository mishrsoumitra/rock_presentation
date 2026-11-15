// apps/web/src/ui/VersionsPane.tsx  -- Version History: earlier states of this deck, to restore, download or delete
import { useEffect, useState } from 'react';
import { sizeText, whenText, type VersionInfo } from '../engine/versions.ts';
import { useEditor } from '../state/store.ts';
import { deleteVersion, downloadVersion, listVersions, restoreVersion, saveVersion } from './actions.ts';

/**
 * The versions of the deck being edited, newest first. A version is kept at each save, every ten minutes while
 * the deck is being changed, before a restore, and whenever one is saved here by name. Restoring puts that state
 * in the editor as unsaved changes; what was there is kept as a version first, so a restore can be undone.
 */
export function VersionsPane() {
  const session = useEditor(st => st.session);
  const tick = useEditor(st => st.versionsTick);
  const [list, setList] = useState<VersionInfo[] | null>(null);
  const [name, setName] = useState('');
  const [busy, setBusy] = useState(false);
  const deck = session?.name;
  useEffect(() => {
    let live = true;
    void listVersions().then(found => { if (live) setList(found); });
    return () => { live = false; };
  }, [deck, tick]);
  if (!session) return null;
  const now = Date.now();
  const save = async () => {
    setBusy(true);
    try { await saveVersion(name.trim() || undefined); setName(''); } finally { setBusy(false); }
  };
  return (
    <aside className="pane versions-pane" aria-label="Version History" data-testid="versions-pane">
      <header><strong>Version History</strong><button type="button" aria-label="Close" onClick={() => useEditor.setState({ pane: 'none' })}>×</button></header>
      <div className="version-new">
        <input type="text" aria-label="Version name" placeholder="Name this version (optional)" value={name} maxLength={80} data-testid="version-name"
          onChange={e => setName(e.target.value)} onKeyDown={e => { e.stopPropagation(); if (e.key === 'Enter' && !busy) void save(); }} />
        <button type="button" disabled={busy} onClick={() => void save()} data-testid="version-save">Save Version</button>
      </div>
      {list && list.length === 0 && <p className="hint" data-testid="versions-empty">No versions of {session.name} yet. One is kept each time you save, and every ten minutes while you work.</p>}
      <ol>
        {(list ?? []).map(v => (
          <li key={v.id} className="version" data-testid="version" data-id={v.id}>
            <div className="version-what">
              <strong className={v.auto ? 'auto' : ''} data-testid="version-label">{v.label}</strong>
              <span data-testid="version-when">{whenText(v.at, now)}</span>
              <span className="version-meta">{v.slides} {v.slides === 1 ? 'slide' : 'slides'} · {sizeText(v.size)}{v.encrypted ? ' · encrypted' : ''}</span>
            </div>
            <div className="version-tools">
              <button type="button" onClick={() => void restoreVersion(v.id)} data-testid="version-restore">Restore</button>
              <button type="button" onClick={() => void downloadVersion(v.id)} data-testid="version-download">Download</button>
              <button type="button" aria-label={`Delete version ${v.label}`} onClick={() => void deleteVersion(v.id)} data-testid="version-delete">Delete</button>
            </div>
          </li>
        ))}
      </ol>
      <p className="hint">Versions are kept in this browser only, up to twenty for each presentation.</p>
    </aside>
  );
}
