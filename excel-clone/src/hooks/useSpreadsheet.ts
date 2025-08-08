import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { SpreadsheetState, SheetData, CellId } from '../types';
import { evaluateAll } from '../utils/evaluator';
import { toCSV, fromCSV } from '../utils/csv';

const STORAGE_KEY = 'excel-clone-state-v1';

function createEmptySheet(): SheetData { return {}; }

export function useSpreadsheet(initialRows = 30, initialCols = 20) {
  const [state, setState] = useState<SpreadsheetState>(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try { return JSON.parse(saved) as SpreadsheetState; } catch {}
    }
    return { sheet: createEmptySheet(), selected: 'A1', editing: false, rows: initialRows, cols: initialCols };
  });

  const undoStack = useRef<SpreadsheetState[]>([]);
  const redoStack = useRef<SpreadsheetState[]>([]);

  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }, [state]);

  const evaluated = useMemo(() => evaluateAll(state.sheet, state.rows, state.cols), [state.sheet, state.rows, state.cols]);

  const setCell = useCallback((id: CellId, raw: string) => {
    setState(prev => {
      undoStack.current.push(prev);
      redoStack.current = [];
      const next: SpreadsheetState = {
        ...prev,
        sheet: { ...prev.sheet, [id]: { raw } },
      };
      if (raw === '') {
        // Clean up empty cell
        const copy = { ...next.sheet };
        delete copy[id];
        next.sheet = copy;
      }
      return next;
    });
  }, []);

  const selectCell = useCallback((id: CellId, startEditing = false) => {
    setState(prev => ({ ...prev, selected: id, editing: startEditing }));
  }, []);

  const setDimensions = useCallback((rows: number, cols: number) => {
    setState(prev => ({ ...prev, rows, cols }));
  }, []);

  const undo = useCallback(() => {
    setState(prev => {
      const last = undoStack.current.pop();
      if (!last) return prev;
      redoStack.current.push(prev);
      return last;
    });
  }, []);

  const redo = useCallback(() => {
    setState(prev => {
      const next = redoStack.current.pop();
      if (!next) return prev;
      undoStack.current.push(prev);
      return next;
    });
  }, []);

  const exportCSV = useCallback(() => {
    const rows: string[][] = [];
    for (let r = 0; r < state.rows; r++) {
      const row: string[] = [];
      for (let c = 0; c < state.cols; c++) {
        const id = letters(c) + String(r + 1);
        row.push(state.sheet[id]?.raw ?? '');
      }
      rows.push(row);
    }
    const csv = toCSV(rows);
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'sheet.csv'; a.click();
    URL.revokeObjectURL(url);
  }, [state]);

  const importCSV = useCallback((csvText: string) => {
    const rows = fromCSV(csvText);
    setState(prev => {
      undoStack.current.push(prev);
      redoStack.current = [];
      const sheet: SheetData = {};
      const rowsCount = Math.max(prev.rows, rows.length);
      let colsCount = prev.cols;
      for (let r = 0; r < rows.length; r++) {
        colsCount = Math.max(colsCount, rows[r].length);
        for (let c = 0; c < rows[r].length; c++) {
          const id = letters(c) + String(r + 1);
          const raw = rows[r][c] ?? '';
          if (raw !== '') sheet[id] = { raw };
        }
      }
      return { ...prev, sheet, rows: rowsCount, cols: colsCount };
    });
  }, []);

  const getDisplay = useCallback((id: CellId): string => {
    const ev = evaluated.get(id);
    if (!ev) return '';
    if (ev.error) return ev.error;
    if (ev.value == null) return '';
    return String(ev.value);
  }, [evaluated]);

  return {
    state,
    setCell,
    selectCell,
    setDimensions,
    undo,
    redo,
    exportCSV,
    importCSV,
    getDisplay,
  };
}

function letters(index: number): string {
  const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let n = index + 1, label = '';
  while (n > 0) { const rem = (n - 1) % 26; label = LETTERS[rem] + label; n = Math.floor((n - 1) / 26); }
  return label;
}