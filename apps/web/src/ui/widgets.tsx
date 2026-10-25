// apps/web/src/ui/widgets.tsx  -- small building blocks: icons, buttons, drop-down menus, the colour picker
import { useEffect, useRef, useState, type ReactNode } from 'react';
import { STANDARD_COLORS, themePalette, type ColorSpec } from '@pptx/editor-core';
import { evaluateGeometry, PRESET_NAMES, presetGeometry } from '@pptx/geometry';
import { useEditor } from '../state/store.ts';

const PATHS: Record<string, string> = {
  undo: 'M7 7H3V3M3.5 7A8 8 0 1 1 5 17', redo: 'M17 7h4V3M20.5 7A8 8 0 1 0 19 17',
  bold: 'M7 4h6a4 4 0 0 1 0 8H7zM7 12h7a4 4 0 0 1 0 8H7z', italic: 'M10 4h8M6 20h8M14 4l-4 16', underline: 'M7 4v7a5 5 0 0 0 10 0V4M5 20h14',
  strike: 'M5 12h14M16 6a4 4 0 0 0-4-2c-2.5 0-4 1.3-4 3s1.5 2.5 4 3M8 18a4 4 0 0 0 4 2c2.5 0 4-1.3 4-3',
  bullets: 'M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01', numbers: 'M10 6h10M10 12h10M10 18h10M4 4v4M3 18h2.5l-2.5 3h3',
  indent: 'M3 4h18M10 9h11M10 14h11M3 19h18M3 9l4 2.5L3 14', outdent: 'M3 4h18M10 9h11M10 14h11M3 19h18M7 9l-4 2.5L7 14',
  left: 'M3 5h18M3 10h12M3 15h18M3 20h12', center: 'M3 5h18M6 10h12M3 15h18M6 20h12', right: 'M3 5h18M9 10h12M3 15h18M9 20h12', justify: 'M3 5h18M3 10h18M3 15h18M3 20h18',
  copy: 'M8 8h12v12H8zM4 16V4h12', cut: 'M6 6l12 12M18 6L6 18M4 20a2 2 0 1 0 0-.1M20 20a2 2 0 1 0 0-.1', paste: 'M8 4h8v3H8zM6 6H4v15h16V6h-2',
  trash: 'M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3', plus: 'M12 5v14M5 12h14', image: 'M3 5h18v14H3zM3 16l5-5 4 4 3-3 6 6M15.5 9.5h.01',
  text: 'M5 5h14M12 5v14M8 19h8', play: 'M7 5l12 7-12 7z', fill: 'M5 13l7-7 7 7-7 7zM19 15s2 2.5 2 4a2 2 0 0 1-4 0c0-1.5 2-4 2-4', outline: 'M4 4h16v16H4z',
  front: 'M8 8h12v12H8zM4 4h12v12', back: 'M4 4h12v12H4zM8 8h12v12', group: 'M3 3h7v7H3zM14 14h7v7h-7zM10 6h4M18 10v4', rotate: 'M20 12a8 8 0 1 1-2.3-5.6M20 4v5h-5',
  flipH: 'M12 3v18M8 7l-5 5 5 5zM16 7l5 5-5 5z', flipV: 'M3 12h18M7 8l5-5 5 5zM7 16l5 5 5-5z', open: 'M3 6h6l2 2h10v11H3z', save: 'M5 3h11l3 3v15H5zM8 3v5h7V3M8 21v-7h8v7',
  file: 'M6 2h9l5 5v15H6zM14 2v6h6', size: 'M3 6h18v12H3zM8 10v4M16 10v4', hide: 'M3 3l18 18M10.6 10.6a2 2 0 0 0 2.8 2.8M9.9 5.1A10 10 0 0 1 22 12c-.6 1.3-1.4 2.4-2.4 3.4M6.1 6.1A10 10 0 0 0 2 12c1.8 4 5.6 7 10 7 1.6 0 3.2-.4 4.6-1.1',
  lock: 'M6 11h12v10H6zM8 11V7a4 4 0 0 1 8 0v4', anim: 'M4 12h4l3-7 3 14 3-7h3', layout: 'M3 4h18v16H3zM3 9h18M9 9v11',
};

/**
 * Ribbon controls keep the focus where it is: a mouse-down on a button must not blur the text input, so the text
 * selection the command applies to stays, and typing goes on in the box afterwards (as in PowerPoint's ribbon).
 */
export const keepFocus = (e: React.MouseEvent) => e.preventDefault();

/** Put the focus back in the text being edited, after a control that had to take it (a select). */
export function refocusText(): void {
  const el = document.querySelector<HTMLTextAreaElement>('.text-input');
  if (el && document.activeElement !== el) el.focus({ preventScroll: true });
}

export function Icon({ name, size = 16 }: { name: string; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d={PATHS[name] ?? PATHS.plus} />
    </svg>
  );
}

export function Btn({ icon, label, onClick, active, disabled, wide, testid }: { icon?: string; label: string; onClick?: () => void; active?: boolean; disabled?: boolean; wide?: boolean; testid?: string }) {
  return (
    <button type="button" className={`btn${active ? ' active' : ''}${wide ? ' wide' : ''}`} title={label} aria-label={label} aria-pressed={active} disabled={disabled} onClick={onClick} onMouseDown={keepFocus} data-testid={testid}>
      {icon && <Icon name={icon} />}
      {(wide || !icon) && <span>{label}</span>}
    </button>
  );
}

/** A button that opens a panel below it; clicking outside or choosing closes it. */
export function Menu({ icon, label, children, wide = true, testid }: { icon?: string; label: string; children: (close: () => void) => ReactNode; wide?: boolean; testid?: string }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const off = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) setOpen(false); };
    document.addEventListener('mousedown', off);
    return () => document.removeEventListener('mousedown', off);
  }, [open]);
  return (
    <div className="menu" ref={ref}>
      <button type="button" className={`btn${wide ? ' wide' : ''}`} aria-haspopup="true" aria-expanded={open} title={label} onClick={() => setOpen(o => !o)} onMouseDown={keepFocus} data-testid={testid}>
        {icon && <Icon name={icon} />}{wide && <span>{label}</span>}<span className="caret">▾</span>
      </button>
      {open && <div className="menu-panel" role="menu">{children(() => setOpen(false))}</div>}
    </div>
  );
}

export function MenuItem({ label, onClick, disabled }: { label: string; onClick: () => void; disabled?: boolean }) {
  return <button type="button" role="menuitem" className="menu-item" disabled={disabled} onClick={onClick} onMouseDown={keepFocus}>{label}</button>;
}

/** The theme colour picker: the slide's theme colours with PowerPoint's tints and shades, then standard colours. */
export function ColorGrid({ onPick, extra }: { onPick: (c: ColorSpec | 'none' | null) => void; extra?: ReactNode }) {
  const s = useEditor(st => st.session);
  const slide = useEditor(st => st.slide);
  if (!s || !slide) return null;
  const cols = themePalette(s.ed, slide);
  return (
    <div className="colors">
      <div className="colors-title">Theme Colors</div>
      <div className="colors-grid">
        {cols.map(c => <button key={c.slot} type="button" className="swatch" style={{ background: c.base.hex }} title={c.base.label} onClick={() => onPick(c.base.color)} onMouseDown={keepFocus} />)}
      </div>
      {[0, 1, 2, 3, 4].map(row => (
        <div key={row} className="colors-grid tight">
          {cols.map(c => {
            const v = c.variants[row]!;
            return <button key={c.slot} type="button" className="swatch" style={{ background: v.hex }} title={v.label} onClick={() => onPick(v.color)} onMouseDown={keepFocus} />;
          })}
        </div>
      ))}
      <div className="colors-title">Standard Colors</div>
      <div className="colors-grid">
        {STANDARD_COLORS.map(c => <button key={c.hex} type="button" className="swatch" style={{ background: c.hex }} title={c.label} onClick={() => onPick(c.color)} onMouseDown={keepFocus} />)}
      </div>
      {extra}
    </div>
  );
}

/** A small drawing of a preset shape, from its own geometry. */
export function PresetIcon({ prst, size = 22 }: { prst: string; size?: number }) {
  const g = presetGeometry(prst);
  if (!g) return null;
  const w = 100, h = /Arrow|Connector|line/i.test(prst) ? 60 : 100;
  let d = '';
  try {
    for (const p of evaluateGeometry(g, w, h).paths) {
      for (const c of p.commands) {
        if (c.op === 'M') d += `M${c.x} ${c.y}`;
        else if (c.op === 'L') d += `L${c.x} ${c.y}`;
        else if (c.op === 'C') d += `C${c.x1} ${c.y1} ${c.x2} ${c.y2} ${c.x} ${c.y}`;
        else if (c.op === 'Q') d += `Q${c.x1} ${c.y1} ${c.x} ${c.y}`;
        else d += 'Z';
      }
    }
  } catch {
    return null;
  }
  return (
    <svg width={size} height={size} viewBox={`-6 -6 ${w + 12} ${h + 12}`} aria-hidden="true">
      <path d={d} fill="none" stroke="currentColor" strokeWidth="5" strokeLinejoin="round" />
    </svg>
  );
}

/** The shapes the gallery offers, in PowerPoint's groups. */
export const SHAPE_GROUPS: [string, string[]][] = [
  ['Lines', ['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3']],
  ['Rectangles', ['rect', 'roundRect', 'snip1Rect', 'snip2SameRect', 'round1Rect', 'round2SameRect']],
  ['Basic Shapes', ['ellipse', 'triangle', 'rtTriangle', 'parallelogram', 'trapezoid', 'diamond', 'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon',
    'dodecagon', 'pie', 'chord', 'teardrop', 'frame', 'halfFrame', 'corner', 'diagStripe', 'plus', 'plaque', 'can', 'cube', 'bevel', 'donut', 'noSmoking',
    'blockArc', 'foldedCorner', 'smileyFace', 'heart', 'lightningBolt', 'sun', 'moon', 'cloud', 'arc', 'bracketPair', 'bracePair']],
  ['Block Arrows', ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow', 'leftRightUpArrow', 'bentArrow', 'uturnArrow',
    'leftUpArrow', 'bentUpArrow', 'curvedRightArrow', 'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'stripedRightArrow', 'notchedRightArrow',
    'homePlate', 'chevron', 'rightArrowCallout', 'circularArrow']],
  ['Equation Shapes', ['mathPlus', 'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual']],
  ['Flowchart', ['flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision', 'flowChartInputOutput', 'flowChartPredefinedProcess',
    'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation', 'flowChartManualInput', 'flowChartConnector',
    'flowChartMerge', 'flowChartExtract', 'flowChartMagneticDisk']],
  ['Stars and Banners', ['irregularSeal1', 'irregularSeal2', 'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32',
    'ribbon2', 'ribbon', 'ellipseRibbon2', 'ellipseRibbon', 'verticalScroll', 'horizontalScroll', 'wave', 'doubleWave']],
  ['Callouts', ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout', 'borderCallout1', 'borderCallout2', 'accentCallout1', 'callout1']],
];

/** Every other preset shape the specification defines (presetShapeDefinitions.xml), so all of them can be drawn. */
const MORE: string[] = (() => {
  const shown = new Set(SHAPE_GROUPS.flatMap(([, n]) => n));
  return PRESET_NAMES.filter(n => !shown.has(n));
})();

export function ShapeGallery({ onPick }: { onPick: (prst: string) => void }) {
  return (
    <div className="gallery" data-testid="shape-gallery">
      {[...SHAPE_GROUPS, ['More Shapes', MORE] as [string, string[]]].map(([title, names]) => (
        <div key={title}>
          <div className="colors-title">{title}</div>
          <div className="gallery-grid">
            {names.map(n => <button key={n} type="button" className="gallery-item" title={n} aria-label={n} data-prst={n} onClick={() => onPick(n)} onMouseDown={keepFocus}><PresetIcon prst={n} /></button>)}
          </div>
        </div>
      ))}
    </div>
  );
}
