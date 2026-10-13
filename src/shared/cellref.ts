import type { CellAddr, CellRange } from './types';

/** 0 -> "A", 25 -> "Z", 26 -> "AA" */
export function colLabel(col: number): string {
  let s = '';
  let n = col;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

/** "A" -> 0, "AA" -> 26. Returns -1 on bad input. */
export function colIndex(label: string): number {
  if (!/^[A-Za-z]+$/.test(label)) return -1;
  let n = 0;
  for (const ch of label.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

export function addrLabel(a: CellAddr): string {
  return `${colLabel(a.col)}${a.row + 1}`;
}

/** Parses "B7" / "b7" into an address. Returns null if it is not a plain A1 ref. */
export function parseAddr(text: string): CellAddr | null {
  const m = /^([A-Za-z]+)([0-9]+)$/.exec(text.trim());
  if (!m) return null;
  const col = colIndex(m[1]);
  const row = parseInt(m[2], 10) - 1;
  if (col < 0 || row < 0) return null;
  return { row, col };
}

export function rangeLabel(r: CellRange): string {
  if (r.r1 === r.r2 && r.c1 === r.c2) return addrLabel({ row: r.r1, col: r.c1 });
  return `${addrLabel({ row: r.r1, col: r.c1 })}:${addrLabel({ row: r.r2, col: r.c2 })}`;
}

export function normaliseRange(anchor: CellAddr, focus: CellAddr): CellRange {
  return {
    r1: Math.min(anchor.row, focus.row),
    c1: Math.min(anchor.col, focus.col),
    r2: Math.max(anchor.row, focus.row),
    c2: Math.max(anchor.col, focus.col),
  };
}

export function rangeContains(r: CellRange, a: CellAddr): boolean {
  return a.row >= r.r1 && a.row <= r.r2 && a.col >= r.c1 && a.col <= r.c2;
}

export function rangeCellCount(r: CellRange): number {
  return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
}

export function styleKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export function* iterRange(r: CellRange): Generator<CellAddr> {
  for (let row = r.r1; row <= r.r2; row++) {
    for (let col = r.c1; col <= r.c2; col++) {
      yield { row, col };
    }
  }
}
