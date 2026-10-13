/** Shared domain types used by the spreadsheet (and later the word processor). */

export interface CellAddr {
  row: number;
  col: number;
}

/** Normalised rectangular range. r1/c1 are always <= r2/c2. */
export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export type HorizAlign = 'left' | 'center' | 'right';
export type VertAlign = 'top' | 'middle' | 'bottom';

export type NumFmt =
  | 'general'
  | 'number'
  | 'number0'
  | 'currency'
  | 'percent'
  | 'date'
  | 'datetime'
  | 'text';

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  align?: HorizAlign;
  valign?: VertAlign;
  color?: string;
  bg?: string;
  fontSize?: number;
  fontFamily?: string;
  numFmt?: NumFmt;
  wrap?: boolean;
  border?: BorderSet;
}

export interface BorderSet {
  top?: boolean;
  right?: boolean;
  bottom?: boolean;
  left?: boolean;
}

/** Style map key: `${row}:${col}` within a single sheet. */
export type StyleMap = Record<string, CellStyle>;

export interface SheetMeta {
  id: number;
  name: string;
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  styles: StyleMap;
  frozenRows: number;
  frozenCols: number;
  merges: CellRange[];
}

export const DEFAULT_COL_WIDTH = 96;
export const DEFAULT_ROW_HEIGHT = 24;
export const ROW_HEADER_WIDTH = 48;
export const COL_HEADER_HEIGHT = 24;
export const MAX_ROWS = 5000;
export const MAX_COLS = 200;
