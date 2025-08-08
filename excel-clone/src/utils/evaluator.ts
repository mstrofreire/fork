import { Parser } from 'expr-eval';
import { expandRange } from './cellIds';
import type { SheetData, EvaluatedCellValue, CellId } from '../types';

export interface EvaluateOptions {
  rows: number;
  cols: number;
}

const parser = new Parser({ allowMemberAccess: false });

function isNumericString(text: string): boolean {
  return /^\s*-?\d+(?:\.\d+)?\s*$/.test(text);
}

function preprocessFormula(formulaBody: string): { body: string; refs: Set<string>; ranges: string[] } {
  // Collect bare refs like A1, B12, ZZ100
  const refRegex = /\b([A-Za-z]+[1-9][0-9]*)\b/g;
  const refs = new Set<string>();
  let m: RegExpExecArray | null;
  while ((m = refRegex.exec(formulaBody)) !== null) {
    refs.add(m[1].toUpperCase());
  }
  // Wrap range tokens inside functions into RANGE("A1:B2"). We only replace the token A1:B2
  const rangeRegex = /([A-Za-z]+[1-9][0-9]*:[A-Za-z]+[1-9][0-9]*)/g;
  const ranges: string[] = [];
  const body = formulaBody.replace(rangeRegex, (tok) => {
    ranges.push(tok.toUpperCase());
    return `RANGE("${tok.toUpperCase()}")`;
  });
  return { body, refs, ranges };
}

export function evaluateCell(cellId: CellId, sheet: SheetData, opts: EvaluateOptions, memo: Map<string, EvaluatedCellValue>, visiting: Set<string>): EvaluatedCellValue {
  if (memo.has(cellId)) return memo.get(cellId)!;
  if (visiting.has(cellId)) {
    const cyc: EvaluatedCellValue = { value: null, error: '#CYCLE' };
    memo.set(cellId, cyc);
    return cyc;
  }
  visiting.add(cellId);
  try {
    const cell = sheet[cellId];
    const raw = cell?.raw ?? '';
    if (raw.length === 0) {
      const v = { value: null } as EvaluatedCellValue; memo.set(cellId, v); return v;
    }
    if (raw.startsWith("'")) {
      const v = { value: raw.slice(1) } as EvaluatedCellValue; memo.set(cellId, v); return v;
    }
    if (raw.startsWith('=')) {
      const body0 = raw.slice(1);
      const { body, refs } = preprocessFormula(body0);

      const variables: Record<string, any> = {};

      // Functions
      const numify = (x: unknown): number => {
        if (typeof x === 'number') return x;
        if (typeof x === 'string' && isNumericString(x)) return parseFloat(x);
        return 0;
      };
      const flat = (args: unknown[]): number[] => args.flatMap(a => Array.isArray(a) ? a.map(numify) : [numify(a)]);
      variables['SUM'] = (...args: unknown[]) => flat(args).reduce((a, b) => a + b, 0);
      variables['AVG'] = (...args: unknown[]) => { const arr = flat(args); return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0; };
      variables['MIN'] = (...args: unknown[]) => { const arr = flat(args); return arr.length ? Math.min(...arr) : 0; };
      variables['MAX'] = (...args: unknown[]) => { const arr = flat(args); return arr.length ? Math.max(...arr) : 0; };
      variables['COUNT'] = (...args: unknown[]) => flat(args).length;
      variables['RANGE'] = (range: string) => {
        const cells = expandRange(range);
        return cells.map(id => {
          const ev = evaluateCell(id, sheet, opts, memo, visiting);
          return typeof ev.value === 'number' ? ev.value : (isNumericString(String(ev.value)) ? parseFloat(String(ev.value)) : 0);
        });
      };

      // Inject referenced variables as evaluated values
      refs.forEach(ref => {
        const ev = evaluateCell(ref, sheet, opts, memo, visiting);
        variables[ref] = typeof ev.value === 'number' ? ev.value : ev.value ?? 0;
      });

      try {
        const expr = parser.parse(body);
        const out = expr.evaluate(variables);
        const val = typeof out === 'number' || typeof out === 'string' ? out : null;
        const v = { value: val } as EvaluatedCellValue;
        memo.set(cellId, v);
        return v;
      } catch (e: any) {
        const v = { value: null, error: '#ERROR' } as EvaluatedCellValue;
        memo.set(cellId, v);
        return v;
      }
    }
    // Not a formula: try number else text
    if (isNumericString(raw)) {
      const v = { value: parseFloat(raw) } as EvaluatedCellValue; memo.set(cellId, v); return v;
    }
    const v = { value: raw } as EvaluatedCellValue; memo.set(cellId, v); return v;
  } finally {
    visiting.delete(cellId);
  }
}

export function evaluateAll(sheet: SheetData, rows: number, cols: number): Map<string, EvaluatedCellValue> {
  const memo = new Map<string, EvaluatedCellValue>();
  const visiting = new Set<string>();
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const colLabel = indexToLetters(c);
      const id = `${colLabel}${r + 1}`;
      if (!memo.has(id)) evaluateCell(id, sheet, { rows, cols }, memo, visiting);
    }
  }
  return memo;
}

function indexToLetters(index: number): string {
  const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let n = index + 1;
  let label = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = LETTERS[rem] + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}