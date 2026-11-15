// apps/web/src/ui/ChartDialog.tsx  -- a chart's data as a small sheet: categories down the side, a series per column
import { useEffect, useState } from 'react';
import type { ChartData } from '@pptx/editor-core';
import { useEditor } from '../state/store.ts';
import { applyChartData, selectedChartData } from './actions.ts';

/** The sheet as it is typed: every cell text, so a half-typed number is not lost. */
interface Sheet { categories: string[]; names: string[]; cells: string[][]; }

const toSheet = (d: ChartData): Sheet => ({
  categories: [...d.categories], names: d.series.map(s => s.name),
  cells: d.categories.map((_, r) => d.series.map(s => (s.values[r] === null || s.values[r] === undefined ? '' : String(s.values[r])))),
});
/** A cell's number: blank is no value; anything that is not a number is an error (NaN). */
const numberOf = (text: string) => (text.trim() === '' ? null : Number(text.trim().replace(',', '.')));
const toData = (s: Sheet): ChartData => ({
  categories: s.categories.map((c, i) => c.trim() || String(i + 1)),
  series: s.names.map((name, c) => ({ name: name.trim() || `Series ${c + 1}`, values: s.cells.map(row => numberOf(row[c] ?? '')) })),
});

/**
 * Edit Data: the selected chart's categories, series names and values in a grid. Rows and columns are added and
 * removed at the end. Apply writes the whole sheet to the chart as one undo step; a cell that is not a number
 * is marked and holds Apply back.
 */
export function ChartDialog() {
  const open = useEditor(s => s.chartEditing);
  return open ? <ChartSheet /> : null;
}

function ChartSheet() {
  const [sheet, setSheet] = useState<Sheet | null>(() => { const d = selectedChartData(); return d ? toSheet(d) : null; });
  const close = () => useEditor.setState({ chartEditing: false });
  useEffect(() => {
    const esc = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
    window.addEventListener('keydown', esc, true);
    return () => window.removeEventListener('keydown', esc, true);
  }, []);
  if (!sheet) return null;
  const bad = (text: string) => { const n = numberOf(text); return n !== null && !Number.isFinite(n); };
  const invalid = sheet.cells.some(row => row.some(bad));
  const set = (change: (s: Sheet) => void) => setSheet(prev => { const next: Sheet = { categories: [...prev!.categories], names: [...prev!.names], cells: prev!.cells.map(r => [...r]) }; change(next); return next; });
  const rows = sheet.categories.length, cols = sheet.names.length;
  return (
    <div className="backdrop" onMouseDown={e => { if (e.target === e.currentTarget) close(); }}>
      <div className="dialog chart-dialog" role="dialog" aria-modal="true" aria-label="Chart Data" data-testid="chart-dialog" onKeyDown={e => e.stopPropagation()}>
        <h2>Chart Data</h2>
        <div className="chart-sheet">
          <table>
            <thead>
              <tr>
                <th aria-label="Categories" />
                {sheet.names.map((n, c) => <th key={c}><input aria-label={`Series ${c + 1} name`} value={n} data-testid={`series-${c}`} onChange={e => set(s => { s.names[c] = e.target.value; })} /></th>)}
              </tr>
            </thead>
            <tbody>
              {sheet.categories.map((cat, r) => (
                <tr key={r}>
                  <th><input aria-label={`Category ${r + 1}`} value={cat} data-testid={`category-${r}`} onChange={e => set(s => { s.categories[r] = e.target.value; })} /></th>
                  {sheet.names.map((_, c) => (
                    <td key={c}><input aria-label={`${sheet.names[c]} for ${cat}`} inputMode="decimal" className={bad(sheet.cells[r]![c] ?? '') ? 'bad' : ''} value={sheet.cells[r]![c] ?? ''} data-testid={`cell-${r}-${c}`}
                      onChange={e => set(s => { s.cells[r]![c] = e.target.value; })} /></td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="chart-tools">
          <button type="button" onClick={() => set(s => { s.categories.push(`Category ${rows + 1}`); s.cells.push(s.names.map(() => '')); })} data-testid="chart-add-row">Add Row</button>
          <button type="button" disabled={rows < 2} onClick={() => set(s => { s.categories.pop(); s.cells.pop(); })} data-testid="chart-remove-row">Remove Row</button>
          <button type="button" onClick={() => set(s => { s.names.push(`Series ${cols + 1}`); for (const row of s.cells) row.push(''); })} data-testid="chart-add-column">Add Column</button>
          <button type="button" disabled={cols < 2} onClick={() => set(s => { s.names.pop(); for (const row of s.cells) row.pop(); })} data-testid="chart-remove-column">Remove Column</button>
        </div>
        {invalid && <p className="dialog-error" role="alert" data-testid="chart-error">Values must be numbers. Leave a cell empty for no value.</p>}
        <div className="dialog-buttons">
          <button type="button" onClick={close}>Cancel</button>
          <button type="button" disabled={invalid} onClick={() => { applyChartData(toData(sheet)); close(); }} data-testid="chart-apply">Apply</button>
        </div>
      </div>
    </div>
  );
}
