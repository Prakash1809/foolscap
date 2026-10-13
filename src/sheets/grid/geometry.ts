import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, MAX_COLS, MAX_ROWS } from '../../shared/types';

export interface Axis {
  /** offsets[i] = pixel offset of track i; length is count + 1 (last = total size). */
  offsets: Float64Array;
  total: number;
  count: number;
}

export function buildAxis(
  count: number,
  sizes: Record<number, number>,
  fallback: number,
): Axis {
  const offsets = new Float64Array(count + 1);
  let acc = 0;
  for (let i = 0; i < count; i++) {
    offsets[i] = acc;
    acc += sizes[i] ?? fallback;
  }
  offsets[count] = acc;
  return { offsets, total: acc, count };
}

export function buildColAxis(colWidths: Record<number, number>): Axis {
  return buildAxis(MAX_COLS, colWidths, DEFAULT_COL_WIDTH);
}

export function buildRowAxis(rowHeights: Record<number, number>): Axis {
  return buildAxis(MAX_ROWS, rowHeights, DEFAULT_ROW_HEIGHT);
}

/** Largest index whose offset is <= pos. Binary search over the offsets array. */
export function indexAt(axis: Axis, pos: number): number {
  const { offsets, count } = axis;
  if (pos <= 0) return 0;
  if (pos >= offsets[count]) return count - 1;
  let lo = 0;
  let hi = count - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (offsets[mid] <= pos) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

export function sizeOf(axis: Axis, index: number): number {
  return axis.offsets[index + 1] - axis.offsets[index];
}

export interface VisibleRange {
  start: number;
  end: number;
}

/** Index window covering [pos, pos + extent), padded by `overscan` tracks. */
export function visibleRange(
  axis: Axis,
  pos: number,
  extent: number,
  overscan = 3,
): VisibleRange {
  const start = Math.max(0, indexAt(axis, pos) - overscan);
  const end = Math.min(axis.count - 1, indexAt(axis, pos + extent) + overscan);
  return { start, end };
}
