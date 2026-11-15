// apps/web/src/ui/ContextMenu.tsx  -- the right-click menu: on a shape, on the empty slide, on a slide thumbnail
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { ShapeView } from '@pptx/model';
import { useEditor, type EditorState } from '../state/store.ts';
import * as act from './actions.ts';
import * as tables from './tables.ts';
import { refocusText } from './widgets.tsx';

export type MenuEntry = { label: string; run: () => void; disabled?: boolean; keys?: string; testid?: string } | 'separator';

const MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);
/** A shortcut as this platform writes it: "Ctrl+C", or "⌘C" on a Mac. */
export const keysText = (keys: string) => (MAC ? keys.replace(/Ctrl\+/g, '⌘').replace(/Shift\+/g, '⇧').replace(/Alt\+/g, '⌥') : keys);

/** What the menu offers for where it was opened. The commands are the ribbon's and the keyboard's own. */
export function menuEntries(s: EditorState): MenuEntry[] {
  const menu = s.menu;
  const ed = s.session?.ed;
  if (!menu || !ed || !s.slide) return [];
  const slides = ed.doc.slideParts();
  const index = slides.indexOf(s.slide);
  const hidden = ed.doc.slide(s.slide).hidden;
  const slideEntries: MenuEntry[] = [
    { label: 'New Slide', run: () => act.newSlide(), keys: 'Ctrl+M', testid: 'menu-new-slide' },
    { label: 'Duplicate Slide', run: act.duplicateCurrentSlide, testid: 'menu-duplicate-slide' },
    { label: 'Delete Slide', run: act.deleteCurrentSlide, disabled: slides.length < 2, testid: 'menu-delete-slide' },
    { label: hidden ? 'Show Slide' : 'Hide Slide', run: act.toggleHidden, testid: 'menu-hide-slide' },
  ];
  if (menu.kind === 'thumb') {
    return [
      { label: 'Cut Slide', run: () => void act.cutCurrentSlide(), disabled: slides.length < 2, keys: 'Ctrl+X', testid: 'menu-cut-slide' },
      { label: 'Copy Slide', run: () => void act.copyCurrentSlide(), keys: 'Ctrl+C', testid: 'menu-copy-slide' },
      { label: 'Paste Slide', run: act.pasteSlideClip, disabled: !s.slideClip, keys: 'Ctrl+V', testid: 'menu-paste-slide' },
      'separator',
      ...slideEntries, 'separator',
      { label: 'Move Slide Up', run: () => act.moveCurrentSlide(index - 1), disabled: index <= 0, testid: 'menu-slide-up' },
      { label: 'Move Slide Down', run: () => act.moveCurrentSlide(index + 1), disabled: index >= slides.length - 1, testid: 'menu-slide-down' },
    ];
  }
  const paste: MenuEntry = { label: 'Paste', run: () => void act.paste(), disabled: !s.clipboard, keys: 'Ctrl+V', testid: 'menu-paste' };
  if (menu.kind === 'slide') {
    return [paste, { label: 'Select All', run: act.selectAllShapes, keys: 'Ctrl+A', testid: 'menu-select-all' }, 'separator', ...slideEntries];
  }
  const n = s.selection.length;
  const views = s.selection.map(el => new ShapeView(el, ed.doc.owner(s.slide!)));
  const hasText = n === 1 && !!views[0]!.txBody;
  const hasGroup = views.some(v => v.kind === 'grpSp');
  const table = tables.tableAbilities();
  const tableEntries: MenuEntry[] = table.cells ? [
    { label: 'Insert Row Above', run: () => tables.insertRowBeside('above'), testid: 'menu-row-above' },
    { label: 'Insert Row Below', run: () => tables.insertRowBeside('below'), testid: 'menu-row-below' },
    { label: 'Insert Column Left', run: () => tables.insertColumnBeside('left'), testid: 'menu-col-left' },
    { label: 'Insert Column Right', run: () => tables.insertColumnBeside('right'), testid: 'menu-col-right' },
    { label: 'Delete Row', run: tables.deleteSelectedRows, disabled: !table.deleteRow, testid: 'menu-delete-row' },
    { label: 'Delete Column', run: tables.deleteSelectedColumns, disabled: !table.deleteColumn, testid: 'menu-delete-col' },
    { label: 'Merge Cells', run: tables.mergeSelectedCells, disabled: !table.merge, testid: 'menu-merge' },
    { label: 'Split Cell', run: tables.splitSelectedCell, disabled: !table.split, testid: 'menu-split' },
    'separator',
  ] : [];
  return [
    { label: 'Cut', run: () => act.copySelection(true), keys: 'Ctrl+X', testid: 'menu-cut' },
    { label: 'Copy', run: () => act.copySelection(), keys: 'Ctrl+C', testid: 'menu-copy' },
    paste,
    { label: 'Duplicate', run: act.duplicateSelection, keys: 'Ctrl+D', testid: 'menu-duplicate' },
    { label: 'Delete', run: act.deleteSelection, keys: 'Delete', testid: 'menu-delete' },
    'separator',
    ...tableEntries,
    ...(hasText ? [{ label: 'Edit Text', run: act.editSelectedText, keys: 'Enter', testid: 'menu-edit-text' } satisfies MenuEntry, 'separator' as const] : []),
    { label: 'Bring to Front', run: () => act.order('front'), testid: 'menu-front' },
    { label: 'Bring Forward', run: () => act.order('forward'), testid: 'menu-forward' },
    { label: 'Send Backward', run: () => act.order('backward'), testid: 'menu-backward' },
    { label: 'Send to Back', run: () => act.order('back'), testid: 'menu-back' },
    'separator',
    { label: 'Group', run: act.groupSelection, disabled: n < 2, keys: 'Ctrl+G', testid: 'menu-group' },
    { label: 'Ungroup', run: act.ungroupSelection, disabled: !hasGroup, keys: 'Ctrl+Shift+G', testid: 'menu-ungroup' },
    'separator',
    { label: 'Alt Text…', run: act.editAltText, disabled: n !== 1, testid: 'menu-alt-text' },
    { label: 'Selection Pane…', run: () => useEditor.setState({ pane: 'selection' }), testid: 'menu-selection-pane' },
  ];
}

/**
 * The menu itself, at the pointer and kept inside the window. It closes when a command is chosen, on Escape, on a
 * press anywhere else, and when the window scrolls or resizes. Up and Down move through the commands.
 */
export function ContextMenu() {
  const menu = useEditor(s => s.menu);
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  // closing gives the keyboard back to the text being edited, if any
  const close = () => { useEditor.setState({ menu: null }); if (useEditor.getState().text) refocusText(); };

  useLayoutEffect(() => {
    if (!menu || !ref.current) { setPos(null); return; }
    const r = ref.current.getBoundingClientRect();
    setPos({ left: Math.max(4, Math.min(menu.x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(menu.y, window.innerHeight - r.height - 4)) });
  }, [menu]);
  // once it is placed and showing, the first command takes the keyboard
  useLayoutEffect(() => {
    if (pos) ref.current?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus({ preventScroll: true });
  }, [pos]);

  useEffect(() => {
    if (!menu) return;
    const away = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) close(); };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); return; }
      if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
      e.preventDefault();
      e.stopPropagation();
      const items = [...(ref.current?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') ?? [])];
      const i = items.indexOf(document.activeElement as HTMLButtonElement);
      items[(i + (e.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length]?.focus();
    };
    window.addEventListener('mousedown', away, true);
    window.addEventListener('keydown', key, true);
    window.addEventListener('resize', close);
    window.addEventListener('wheel', close, { passive: true });
    return () => {
      window.removeEventListener('mousedown', away, true);
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('resize', close);
      window.removeEventListener('wheel', close);
    };
  }, [menu]);

  if (!menu) return null;
  const entries = menuEntries(useEditor.getState());
  return (
    <div ref={ref} className="context-menu" role="menu" aria-label="Context menu" data-testid="context-menu" data-kind={menu.kind}
      style={{ left: pos?.left ?? menu.x, top: pos?.top ?? menu.y, visibility: pos ? 'visible' : 'hidden' }} onContextMenu={e => e.preventDefault()}>
      {entries.map((e, i) => (e === 'separator' ? <hr key={i} /> : (
        <button key={e.label} type="button" role="menuitem" className="menu-item" disabled={e.disabled} data-testid={e.testid}
          onClick={() => { close(); e.run(); }}>
          <span>{e.label}</span>{e.keys && <kbd>{keysText(e.keys)}</kbd>}
        </button>
      )))}
    </div>
  );
}
