import { DetailedCellError, HyperFormula } from 'hyperformula';
import type { CellAddr, CellRange } from '../../shared/types';

/**
 * Thin wrapper over HyperFormula.
 *
 * HyperFormula owns values, formulas and the dependency graph; it knows nothing
 * about presentation. Styles, column widths and merges live in the zustand
 * store instead (see ../store.ts).
 *
 * NOTE: HyperFormula is GPLv3 under this license key. Fine for an open or
 * personal project; swap the engine before shipping anything closed-source.
 */
export const hf = HyperFormula.buildEmpty({
  licenseKey: 'gpl-v3',
  useColumnIndex: true,
  undoLimit: 200,
  // Excel-compatible string comparison and date handling.
  caseSensitive: false,
  smartRounding: true,
});

export type RawValue = string | number | boolean | null;

export interface CellSnapshot {
  /** Evaluated value, or an error string like "#DIV/0!". */
  value: RawValue;
  /** Raw user input: "=SUM(A1:A3)" for formulas, else the literal. */
  raw: string;
  isFormula: boolean;
  isError: boolean;
  errorMessage?: string;
}

const EMPTY: CellSnapshot = {
  value: null,
  raw: '',
  isFormula: false,
  isError: false,
};

function hfAddr(sheet: number, a: CellAddr) {
  return { sheet, row: a.row, col: a.col };
}

export function readCell(sheet: number, a: CellAddr): CellSnapshot {
  const addr = hfAddr(sheet, a);
  const serialized = hf.getCellSerialized(addr);
  if (serialized === null || serialized === undefined || serialized === '') return EMPTY;

  const value = hf.getCellValue(addr);
  const raw = String(serialized);
  const isFormula = raw.startsWith('=');

  if (value instanceof DetailedCellError) {
    return {
      value: value.value,
      raw,
      isFormula,
      isError: true,
      errorMessage: value.message,
    };
  }

  return { value: value as RawValue, raw, isFormula, isError: false };
}

/** Raw input for the formula bar — formula text when present, else the literal. */
export function readRaw(sheet: number, a: CellAddr): string {
  const s = hf.getCellSerialized(hfAddr(sheet, a));
  return s === null || s === undefined ? '' : String(s);
}

export function writeCell(sheet: number, a: CellAddr, input: string): void {
  hf.setCellContents(hfAddr(sheet, a), coerceInput(input));
}

/** Writes a rectangular block in one engine call so recalc happens once. */
export function writeBlock(sheet: number, topLeft: CellAddr, rows: string[][]): void {
  const coerced = rows.map((r) => r.map((cell) => coerceInput(cell)));
  hf.setCellContents(hfAddr(sheet, topLeft), coerced);
}

export function clearRange(sheet: number, r: CellRange): void {
  const height = r.r2 - r.r1 + 1;
  const width = r.c2 - r.c1 + 1;
  const blank: null[][] = Array.from({ length: height }, () =>
    Array.from({ length: width }, () => null),
  );
  hf.setCellContents(hfAddr(sheet, { row: r.r1, col: r.c1 }), blank);
}

/**
 * Turns editor text into something HyperFormula stores sensibly.
 * Empty -> null (truly empty cell). Numeric-looking text -> number.
 * Everything else passes through as a string, including "=..." formulas.
 */
export function coerceInput(input: string): RawValue {
  const t = input.trim();
  if (t === '') return null;
  if (t.startsWith('=')) return t;
  if (t.startsWith("'")) return t; // explicit text escape, as in Excel
  if (/^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) {
    const n = Number(t);
    if (Number.isFinite(n)) return n;
  }
  if (/^(TRUE|FALSE)$/i.test(t)) return t.toUpperCase() === 'TRUE';
  return input;
}

export function rangeValues(sheet: number, r: CellRange): RawValue[][] {
  const out: RawValue[][] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const line: RawValue[] = [];
    for (let col = r.c1; col <= r.c2; col++) {
      const snap = readCell(sheet, { row, col });
      line.push(snap.value);
    }
    out.push(line);
  }
  return out;
}

export function rawRange(sheet: number, r: CellRange): string[][] {
  const out: string[][] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const line: string[] = [];
    for (let col = r.c1; col <= r.c2; col++) {
      line.push(readRaw(sheet, { row, col }));
    }
    out.push(line);
  }
  return out;
}

/**
 * Shifts relative A1 refs in a formula by (dRow, dCol) — used by fill/paste.
 *
 * The lookaround is what keeps function names intact: without it `LOG10(` and
 * `ATAN2(` parse as the refs LOG10 / TAN2 and get renumbered. A ref may not sit
 * next to an identifier character, and may not be followed by "(".
 */
const A1_REF =
  /"(?:[^"]|"")*"|(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_.(])/g;

export function offsetFormula(formula: string, dRow: number, dCol: number): string {
  if (!formula.startsWith('=')) return formula;
  // Quoted string literals match the first branch and pass through untouched,
  // so "A1" inside text is never rewritten.
  return formula.replace(
    A1_REF,
    (match, colAbs: string, colRef: string, rowAbs: string, rowRef: string) => {
      if (match.startsWith('"')) return match;
      let col = 0;
      for (const ch of colRef.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
      col -= 1;
      let row = parseInt(rowRef, 10) - 1;
      if (!colAbs) col += dCol;
      if (!rowAbs) row += dRow;
      if (col < 0 || row < 0) return '#REF!';
      let label = '';
      let n = col;
      while (n >= 0) {
        label = String.fromCharCode(65 + (n % 26)) + label;
        n = Math.floor(n / 26) - 1;
      }
      return `${colAbs}${label}${rowAbs}${row + 1}`;
    },
  );
}

export function functionNames(): string[] {
  return hf.getRegisteredFunctionNames().sort();
}
