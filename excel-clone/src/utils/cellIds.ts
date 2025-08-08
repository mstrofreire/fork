const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

export function colIndexToLabel(index: number): string {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = LETTERS[rem] + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

export function colLabelToIndex(label: string): number {
  let n = 0;
  for (let i = 0; i < label.length; i++) {
    n = n * 26 + (label.charCodeAt(i) - 64);
  }
  return n - 1;
}

export function cellCoordsToId(row: number, col: number): string {
  return `${colIndexToLabel(col)}${row + 1}`;
}

export function cellIdToCoords(cellId: string): { row: number; col: number } {
  const match = cellId.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error("Invalid cell id");
  const [, colLabel, rowStr] = match as unknown as [string, string, string, string?];
  const col = colLabelToIndex(colLabel.toUpperCase());
  const row = parseInt(rowStr, 10) - 1;
  return { row, col };
}

export function expandRange(range: string): string[] {
  const [start, end] = range.split(":");
  if (!start || !end) return [];
  const a = cellIdToCoords(start);
  const b = cellIdToCoords(end);
  const rowStart = Math.min(a.row, b.row);
  const rowEnd = Math.max(a.row, b.row);
  const colStart = Math.min(a.col, b.col);
  const colEnd = Math.max(a.col, b.col);
  const cells: string[] = [];
  for (let r = rowStart; r <= rowEnd; r++) {
    for (let c = colStart; c <= colEnd; c++) {
      cells.push(cellCoordsToId(r, c));
    }
  }
  return cells;
}