// apps/web/src/ui/Ribbon.tsx  -- the commands, in tabs as PowerPoint arranges them
import { useRef, useState } from 'react';
import { useEditor } from '../state/store.ts';
import * as act from './actions.ts';
import { formatPara, formatText, formattingHere } from './textActions.ts';
import { Btn, ColorGrid, Menu, MenuItem, ShapeGallery } from './widgets.tsx';

const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Courier New', 'Aptos', 'Georgia', 'Verdana', 'Segoe UI'];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
type Tab = 'File' | 'Home' | 'Insert' | 'Design' | 'Transitions' | 'Animations' | 'Slide Show';

function FilePicker({ accept, onFile, children }: { accept: string; onFile: (f: File) => void; children: (open: () => void) => React.ReactNode }) {
  const ref = useRef<HTMLInputElement>(null);
  return (
    <>
      <input ref={ref} type="file" accept={accept} hidden onChange={e => { const f = e.target.files?.[0]; e.target.value = ''; if (f) onFile(f); }} />
      {children(() => ref.current?.click())}
    </>
  );
}

function LayoutMenu({ label, icon, onPick, testid }: { label: string; icon: string; onPick: (part: string) => void; testid?: string }) {
  return (
    <Menu icon={icon} label={label} testid={testid}>
      {close => act.layouts().map(l => <MenuItem key={l.part} label={l.name} onClick={() => { close(); onPick(l.part); }} />)}
    </Menu>
  );
}

export function Ribbon({ extraTabs }: { extraTabs?: Partial<Record<Tab, React.ReactNode>> }) {
  const [tab, setTab] = useState<Tab>('Home');
  const session = useEditor(s => s.session);
  useEditor(s => s.version);
  useEditor(s => s.text);
  const selection = useEditor(s => s.selection);
  const log = session?.ed.log;
  const here = formattingHere();
  const run = here.run;
  const para = here.para?.props;
  const hasSel = selection.length > 0;
  const isPicture = selection.length === 1 && selection[0]!.local === 'pic';
  const tabs: Tab[] = ['File', 'Home', 'Insert', 'Design', 'Transitions', 'Animations', 'Slide Show'];

  return (
    <header className="ribbon">
      <nav className="tabs" role="tablist">
        {tabs.map(t => <button key={t} type="button" role="tab" aria-selected={tab === t} className={tab === t ? 'tab active' : 'tab'} onClick={() => setTab(t)}>{t}</button>)}
        <span className="doc-name">{session?.name ?? ''}</span>
      </nav>
      <div className="groups" role="toolbar">
        {tab === 'File' && (
          <>
            <section className="group"><Btn icon="file" label="New" wide onClick={() => void act.newPresentation()} testid="new" />
              <FilePicker accept=".pptx,.pptm,.potx,.ppsx,.ppt" onFile={f => void act.openFile(f)}>{open => <Btn icon="open" label="Open…" wide onClick={open} testid="open" />}</FilePicker>
              <Btn icon="save" label="Save a copy" wide disabled={!session} onClick={() => void act.saveAndDownload()} testid="save" /></section>
            <section className="group"><Btn icon="lock" label={session?.password ? 'Change password…' : 'Encrypt with password…'} wide disabled={!session}
              onClick={() => window.dispatchEvent(new CustomEvent('ask-new-password'))} />
              {session?.password && <Btn label="Remove password" wide onClick={() => { session.password = null; useEditor.getState().notify('info', 'The next save is not encrypted.'); }} />}</section>
          </>
        )}
        {tab === 'Home' && (
          <>
            <section className="group" aria-label="Undo">
              <Btn icon="undo" label={log?.undoLabel ? `Undo ${log.undoLabel}` : 'Undo'} disabled={!log?.canUndo} onClick={() => useEditor.getState().undo()} testid="undo" />
              <Btn icon="redo" label={log?.redoLabel ? `Redo ${log.redoLabel}` : 'Redo'} disabled={!log?.canRedo} onClick={() => useEditor.getState().redo()} testid="redo" />
            </section>
            <section className="group" aria-label="Clipboard">
              <Btn icon="paste" label="Paste" disabled={!useEditor.getState().clipboard} onClick={() => void act.paste()} />
              <Btn icon="cut" label="Cut" disabled={!hasSel} onClick={() => act.copySelection(true)} />
              <Btn icon="copy" label="Copy" disabled={!hasSel} onClick={() => act.copySelection()} />
            </section>
            <section className="group" aria-label="Slides">
              <Btn icon="plus" label="New Slide" wide disabled={!session} onClick={() => act.newSlide()} testid="new-slide" />
              <LayoutMenu icon="layout" label="Layout" onPick={act.layoutCurrentSlide} />
            </section>
            <section className="group" aria-label="Font">
              <select aria-label="Font" value={run?.latin ?? ''} onChange={e => formatText({ latin: e.target.value }, 'Font')} disabled={!session}>
                {!run && <option value="" />}
                {[...new Set([run?.latin ?? '', ...FONTS])].filter(Boolean).map(f => <option key={f}>{f}</option>)}
              </select>
              <select aria-label="Font size" value={run ? String(run.size) : ''} onChange={e => formatText({ sz: Math.round(Number(e.target.value) * 100) }, 'Font Size')} disabled={!session} data-testid="font-size">
                {!run && <option value="" />}
                {[...new Set([...(run ? [run.size] : []), ...SIZES])].sort((a, b) => a - b).map(sz => <option key={sz} value={String(sz)}>{sz}</option>)}
              </select>
              <Btn icon="bold" label="Bold" active={run?.bold} onClick={() => formatText({ b: !run?.bold }, 'Bold')} testid="bold" />
              <Btn icon="italic" label="Italic" active={run?.italic} onClick={() => formatText({ i: !run?.italic }, 'Italic')} />
              <Btn icon="underline" label="Underline" active={!!run?.underline && run.underline !== 'none'} onClick={() => formatText({ u: run?.underline && run.underline !== 'none' ? 'none' : 'sng' }, 'Underline')} />
              <Btn icon="strike" label="Strikethrough" active={run?.strike === 'sngStrike'} onClick={() => formatText({ strike: run?.strike === 'sngStrike' ? 'noStrike' : 'sngStrike' }, 'Strikethrough')} />
              <Menu icon="text" label="Font Color" wide={false}>{close => <ColorGrid onPick={c => { close(); formatText({ fill: c === 'none' ? null : c }, 'Font Color'); }} />}</Menu>
            </section>
            <section className="group" aria-label="Paragraph">
              <Btn icon="bullets" label="Bullets" active={para?.bullet?.kind === 'char'} onClick={() => formatPara({ bullet: para?.bullet?.kind === 'char' ? { kind: 'none' } : { kind: 'char', char: '•', font: 'Arial' } }, 'Bullets')} />
              <Btn icon="numbers" label="Numbering" active={para?.bullet?.kind === 'autonum'} onClick={() => formatPara({ bullet: para?.bullet?.kind === 'autonum' ? { kind: 'none' } : { kind: 'autonum', scheme: 'arabicPeriod' } }, 'Numbering')} />
              <Btn icon="outdent" label="Decrease List Level" onClick={() => formatPara({ lvl: Math.max(0, (para?.level ?? 0) - 1) }, 'Decrease Indent')} />
              <Btn icon="indent" label="Increase List Level" onClick={() => formatPara({ lvl: Math.min(8, (para?.level ?? 0) + 1) }, 'Increase Indent')} />
              <Btn icon="left" label="Align Left" active={para?.align === 'l'} onClick={() => formatPara({ algn: 'l' }, 'Align Left')} />
              <Btn icon="center" label="Center" active={para?.align === 'ctr'} onClick={() => formatPara({ algn: 'ctr' }, 'Center')} testid="center" />
              <Btn icon="right" label="Align Right" active={para?.align === 'r'} onClick={() => formatPara({ algn: 'r' }, 'Align Right')} />
              <Btn icon="justify" label="Justify" active={para?.align === 'just'} onClick={() => formatPara({ algn: 'just' }, 'Justify')} />
            </section>
            {isPicture && (
              <section className="group" aria-label="Picture">
                <Menu icon="image" label="Crop">{close => (<>
                  {([['Square 1:1', 1], ['4:3', 4 / 3], ['3:4', 3 / 4], ['16:9', 16 / 9], ['9:16', 9 / 16], ['3:2', 3 / 2]] as const).map(([l, r]) =>
                    <MenuItem key={l} label={`Aspect Ratio ${l}`} onClick={() => { close(); act.cropToAspect(r); }} />)}
                  <MenuItem label="Reset Crop" onClick={() => { close(); act.cropToAspect(null); }} />
                </>)}</Menu>
                <FilePicker accept="image/png,image/jpeg,image/gif,image/bmp,image/webp,.emf,.wmf" onFile={f => void act.changePicture(f)}>
                  {open => <Btn icon="image" label="Change Picture" wide onClick={open} />}
                </FilePicker>
              </section>
            )}
            <section className="group" aria-label="Drawing">
              <Menu icon="outline" label="Shapes" testid="shapes">{close => <ShapeGallery onPick={p => { close(); act.startDrawing(p); }} />}</Menu>
              <Menu icon="fill" label="Shape Fill" wide={false}>{close => <ColorGrid onPick={c => { close(); act.fillSelection(c); }} extra={<>
                <MenuItem label="No Fill" onClick={() => { close(); act.fillSelection('none'); }} />
                <MenuItem label="Theme default" onClick={() => { close(); act.fillSelection(null); }} /></>} />}</Menu>
              <Menu icon="outline" label="Shape Outline" wide={false}>{close => <ColorGrid onPick={c => { close(); act.outlineSelection({ color: c }); }} extra={<>
                <MenuItem label="No Outline" onClick={() => { close(); act.outlineSelection({ color: 'none' }); }} />
                {[0.75, 1.5, 3, 6].map(w => <MenuItem key={w} label={`Weight ${w} pt`} onClick={() => { close(); act.outlineSelection({ width: w * 12700 }); }} />)}
                {['solid', 'sysDash', 'dash', 'lgDash', 'sysDot'].map(d => <MenuItem key={d} label={`Dashes: ${d}`} onClick={() => { close(); act.outlineSelection({ dash: d }); }} />)}</>} />}</Menu>
              <Menu icon="front" label="Arrange" testid="arrange">{close => (<>
                <MenuItem label="Bring to Front" disabled={!hasSel} onClick={() => { close(); act.order('front'); }} />
                <MenuItem label="Bring Forward" disabled={!hasSel} onClick={() => { close(); act.order('forward'); }} />
                <MenuItem label="Send Backward" disabled={!hasSel} onClick={() => { close(); act.order('backward'); }} />
                <MenuItem label="Send to Back" disabled={!hasSel} onClick={() => { close(); act.order('back'); }} />
                <hr />
                <MenuItem label="Group" disabled={selection.length < 2} onClick={() => { close(); act.groupSelection(); }} />
                <MenuItem label="Ungroup" disabled={!hasSel} onClick={() => { close(); act.ungroupSelection(); }} />
                <hr />
                {(['left', 'center', 'right', 'top', 'middle', 'bottom'] as const).map(a => <MenuItem key={a} label={`Align ${a[0]!.toUpperCase()}${a.slice(1)}`} disabled={!hasSel} onClick={() => { close(); act.alignSelection(a); }} />)}
                <MenuItem label="Distribute Horizontally" disabled={selection.length < 3} onClick={() => { close(); act.distributeSelection('h'); }} />
                <MenuItem label="Distribute Vertically" disabled={selection.length < 3} onClick={() => { close(); act.distributeSelection('v'); }} />
                <hr />
                <MenuItem label="Rotate Right 90°" disabled={!hasSel} onClick={() => { close(); act.rotateSelection(1); }} />
                <MenuItem label="Rotate Left 90°" disabled={!hasSel} onClick={() => { close(); act.rotateSelection(-1); }} />
                <MenuItem label="Flip Vertical" disabled={!hasSel} onClick={() => { close(); act.flipSelection('v'); }} />
                <MenuItem label="Flip Horizontal" disabled={!hasSel} onClick={() => { close(); act.flipSelection('h'); }} />
              </>)}</Menu>
            </section>
          </>
        )}
        {tab === 'Insert' && (
          <>
            <section className="group"><LayoutMenu icon="plus" label="New Slide" onPick={p => act.newSlide(p)} /></section>
            <section className="group">
              <FilePicker accept="image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml,.emf,.wmf" onFile={f => void act.insertPictureFile(f)}>
                {open => <Btn icon="image" label="Pictures" wide disabled={!session} onClick={open} testid="insert-picture" />}
              </FilePicker>
              <Menu icon="outline" label="Shapes">{close => <ShapeGallery onPick={p => { close(); act.startDrawing(p); }} />}</Menu>
              <Btn icon="text" label="Text Box" wide disabled={!session} onClick={act.insertTextBoxHere} testid="text-box" />
            </section>
          </>
        )}
        {tab === 'Design' && (
          <>
            <section className="group">
              <Menu icon="size" label="Slide Size">{close => (<>
                <MenuItem label="Standard (4:3)" onClick={() => { close(); window.dispatchEvent(new CustomEvent('slide-size', { detail: { cx: 9144000, cy: 6858000 } })); }} />
                <MenuItem label="Widescreen (16:9)" onClick={() => { close(); window.dispatchEvent(new CustomEvent('slide-size', { detail: { cx: 12192000, cy: 6858000 } })); }} />
                <MenuItem label="Custom Slide Size…" onClick={() => { close(); window.dispatchEvent(new CustomEvent('slide-size', { detail: null })); }} />
              </>)}</Menu>
              <Menu icon="fill" label="Format Background">{close => <ColorGrid onPick={c => {
                close();
                const hex = c && c !== 'none' && c.kind === 'srgb' ? c.hex : null;
                if (c && c !== 'none' && c.kind === 'scheme') useEditor.getState().notify('info', 'Background: pick a standard colour (theme backgrounds keep following the master).');
                act.backgroundCurrentSlide(hex);
              }} extra={<MenuItem label="Reset Background" onClick={() => { close(); act.backgroundCurrentSlide(null); }} />} />}</Menu>
              <FilePicker accept=".pptx,.potx,.pptm,.potm" onFile={f => void act.importTemplate(f)}>{open => <Btn icon="layout" label="Import Slide Masters…" wide disabled={!session} onClick={open} />}</FilePicker>
            </section>
          </>
        )}
        {(tab === 'Transitions' || tab === 'Animations') && (extraTabs?.[tab] ?? <section className="group"><span className="hint">Select a slide or shape.</span></section>)}
        {tab === 'Slide Show' && (
          <>
            <section className="group">
              <Btn icon="play" label="From Beginning" wide disabled={!session} onClick={() => useEditor.setState({ mode: 'present', slide: session?.ed.doc.slideParts()[0] ?? null })} testid="present" />
              <Btn icon="play" label="From Current Slide" wide disabled={!session} onClick={() => useEditor.setState({ mode: 'present' })} />
              <Btn icon="hide" label="Hide Slide" wide disabled={!session} onClick={act.toggleHidden} />
            </section>
          </>
        )}
      </div>
    </header>
  );
}
