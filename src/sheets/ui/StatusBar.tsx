import { useMemo } from 'react';
import { useSheets } from '../store';
import { readCell } from '../engine/engine';
import { rangeCellCount, rangeLabel } from '../../shared/cellref';

/** Scanning a select-all would read a million cells; cap the live stats instead. */
const STAT_CELL_LIMIT = 20_000;

export function StatusBar() {
  const anchor = useSheets((s) => s.anchor);
  const focus = useSheets((s) => s.focus);
  const activeSheetId = useSheets((s) => s.activeSheetId);
  const rev = useSheets((s) => s.rev);
  const dirty = useSheets((s) => s.dirty);
  const sheets = useSheets((s) => s.sheets);

  const selection = useMemo(
    () => ({
      r1: Math.min(anchor.row, focus.row),
      c1: Math.min(anchor.col, focus.col),
      r2: Math.max(anchor.row, focus.row),
      c2: Math.max(anchor.col, focus.col),
    }),
    [anchor, focus],
  );

  const stats = useMemo(() => {
    // `rev` is not read here, but cell values live in HyperFormula outside
    // React — it is what tells this memo the engine has changed underneath.
    void rev;
    if (rangeCellCount(selection) > STAT_CELL_LIMIT) return null;
    let sum = 0;
    let numeric = 0;
    let filled = 0;
    let min = Infinity;
    let max = -Infinity;
    for (let row = selection.r1; row <= selection.r2; row++) {
      for (let col = selection.c1; col <= selection.c2; col++) {
        const snap = readCell(activeSheetId, { row, col });
        if (snap.value === null) continue;
        filled++;
        if (typeof snap.value === 'number') {
          numeric++;
          sum += snap.value;
          min = Math.min(min, snap.value);
          max = Math.max(max, snap.value);
        }
      }
    }
    return { sum, numeric, filled, min, max };
  }, [selection, activeSheetId, rev]);

  const fmt = (n: number) => parseFloat(n.toPrecision(12)).toLocaleString();

  return (
    <div className="status-bar">
      <span className="status-item">{rangeLabel(selection)}</span>
      {stats && stats.filled > 0 && (
        <>
          <span className="status-item">Count: {stats.filled}</span>
          {stats.numeric > 0 && (
            <>
              <span className="status-item">Sum: {fmt(stats.sum)}</span>
              <span className="status-item">Average: {fmt(stats.sum / stats.numeric)}</span>
              <span className="status-item">Min: {fmt(stats.min)}</span>
              <span className="status-item">Max: {fmt(stats.max)}</span>
            </>
          )}
        </>
      )}
      <span className="status-spacer" />
      <span className="status-item">{sheets.length} sheet{sheets.length === 1 ? '' : 's'}</span>
      <span className="status-item">{dirty ? '● Unsaved' : '✓ Saved'}</span>
    </div>
  );
}
