import React, { useMemo, useRef, useState } from 'react';
import { useSpreadsheet } from '../hooks/useSpreadsheet';
import { colIndexToLabel, cellCoordsToId } from '../utils/cellIds';
import type { CellId } from '../types';
import './spreadsheet.css';

export default function Spreadsheet() {
  const { state, setCell, selectCell, setDimensions, undo, redo, exportCSV, importCSV, getDisplay } = useSpreadsheet(30, 20);
  const [formula, setFormula] = useState<string>(state.selected ? (state.sheet[state.selected]?.raw ?? '') : '');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const onSelect = (id: CellId) => {
    selectCell(id, false);
    setFormula(state.sheet[id]?.raw ?? '');
  };

  const onEdit = (id: CellId, raw: string) => {
    setCell(id, raw);
    setFormula(raw);
  };

  const headers = useMemo(() => Array.from({ length: state.cols }, (_, i) => colIndexToLabel(i)), [state.cols]);

  const handleKeyDown: React.KeyboardEventHandler<HTMLDivElement> = (e) => {
    if (!state.selected) return;
    const { row, col } = idToCoords(state.selected);
    if (e.key === 'ArrowDown') { e.preventDefault(); const r = Math.min(state.rows - 1, row + 1); onSelect(cellCoordsToId(r, col)); }
    if (e.key === 'ArrowUp') { e.preventDefault(); const r = Math.max(0, row - 1); onSelect(cellCoordsToId(r, col)); }
    if (e.key === 'ArrowRight') { e.preventDefault(); const c = Math.min(state.cols - 1, col + 1); onSelect(cellCoordsToId(row, c)); }
    if (e.key === 'ArrowLeft') { e.preventDefault(); const c = Math.max(0, col - 1); onSelect(cellCoordsToId(row, c)); }
    if (e.ctrlKey && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); }
    if (e.ctrlKey && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) { e.preventDefault(); redo(); }
  };

  return (
    <div className="sheet-root" tabIndex={0} onKeyDown={handleKeyDown}>
      <div className="toolbar">
        <button onClick={() => undo()}>Undo</button>
        <button onClick={() => redo()}>Redo</button>
        <button onClick={() => exportCSV()}>Export CSV</button>
        <button onClick={() => fileInputRef.current?.click()}>Import CSV</button>
        <input ref={fileInputRef} type="file" accept=".csv,text/csv" style={{ display: 'none' }} onChange={async (e) => {
          const f = e.target.files?.[0]; if (!f) return;
          const text = await f.text();
          importCSV(text);
        }} />
        <div className="spacer" />
        <label>Rows <input type="number" min={1} value={state.rows} onChange={e => setDimensions(parseInt(e.target.value || '1'), state.cols)} /></label>
        <label>Cols <input type="number" min={1} value={state.cols} onChange={e => setDimensions(state.rows, parseInt(e.target.value || '1'))} /></label>
      </div>
      <div className="formula">
        <span>fx</span>
        <input
          value={formula}
          onChange={e => setFormula(e.target.value)}
          onKeyDown={e => {
            if (e.key === 'Enter' && state.selected) {
              onEdit(state.selected, formula);
            }
          }}
          onBlur={() => { if (state.selected) onEdit(state.selected, formula); }}
        />
      </div>
      <div className="grid" style={{ gridTemplateColumns: `48px repeat(${state.cols}, 120px)` }}>
        <div className="corner" />
        {headers.map(h => (
          <div className="col-header" key={h}>{h}</div>
        ))}
        {Array.from({ length: state.rows }, (_, r) => (
          <React.Fragment key={r}>
            <div className="row-header">{r + 1}</div>
            {Array.from({ length: state.cols }, (_, c) => {
              const id = cellCoordsToId(r, c);
              const selected = id === state.selected;
              const display = getDisplay(id);
              const raw = state.sheet[id]?.raw ?? '';
              return (
                <div
                  key={id}
                  className={`cell ${selected ? 'selected' : ''}`}
                  onClick={() => onSelect(id)}
                  onDoubleClick={() => onSelect(id)}
                >
                  <input
                    className="cell-input"
                    value={selected ? (formula) : (raw.startsWith('=') ? display : raw || display)}
                    onChange={e => { if (selected) setFormula(e.target.value); }}
                    onBlur={() => { if (selected) onEdit(id, formula); }}
                    onKeyDown={e => {
                      if (e.key === 'Enter') { if (selected) onEdit(id, formula); const next = Math.min(state.rows - 1, r + 1); onSelect(cellCoordsToId(next, c)); }
                    }}
                  />
                </div>
              );
            })}
          </React.Fragment>
        ))}
      </div>
    </div>
  );
}

function idToCoords(id: string): { row: number; col: number } {
  const m = id.match(/^([A-Z]+)(\d+)$/);
  if (!m) return { row: 0, col: 0 };
  const col = lettersToIndex(m[1]);
  const row = parseInt(m[2], 10) - 1;
  return { row, col };
}

function lettersToIndex(label: string): number {
  let n = 0; for (let i = 0; i < label.length; i++) { n = n * 26 + (label.charCodeAt(i) - 64); } return n - 1;
}