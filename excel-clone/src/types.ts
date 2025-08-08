export type CellId = string; // e.g., "A1"

export type CellRaw = string; // user-entered text, may start with '=' for formulas

export interface CellData {
  raw: CellRaw;
}

export type SheetData = Record<CellId, CellData>;

export interface EvaluatedCellValue {
  value: number | string | null;
  error?: string; // e.g., #REF, #CYCLE, #ERROR
}

export interface SpreadsheetState {
  sheet: SheetData;
  selected: CellId | null;
  editing: boolean;
  rows: number;
  cols: number;
}