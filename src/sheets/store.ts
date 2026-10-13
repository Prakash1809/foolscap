import { create } from 'zustand';
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  MAX_COLS,
  MAX_ROWS,
  type CellAddr,
  type CellRange,
  type CellStyle,
  type SheetMeta,
  type StyleMap,
} from '../shared/types';
import { normaliseRange, styleKey } from '../shared/cellref';
import {
  clearRange,
  hf,
  offsetFormula,
  rawRange,
  readRaw,
  writeBlock,
  writeCell,
} from './engine/engine';

interface Clipboard {
  rows: string[][];
  /** Cut clears the source on paste; copy does not. */
  cut: boolean;
  origin: CellRange;
  sheetId: number;
}

interface SheetsState {
  sheets: SheetMeta[];
  activeSheetId: number;
  anchor: CellAddr;
  focus: CellAddr;
  editing: { addr: CellAddr; draft: string; selectAll: boolean } | null;
  clipboard: Clipboard | null;
  /** Bumped on every engine mutation so the grid re-reads HyperFormula. */
  rev: number;
  fileName: string;
  dirty: boolean;

  // --- selection ---
  select: (addr: CellAddr, extend?: boolean) => void;
  selectRange: (r: CellRange) => void;
  moveSelection: (dRow: number, dCol: number, extend: boolean) => void;
  selectAll: () => void;
  selectRow: (row: number, extend?: boolean) => void;
  selectCol: (col: number, extend?: boolean) => void;

  // --- editing ---
  beginEdit: (addr: CellAddr, seed?: string, selectAll?: boolean) => void;
  updateDraft: (draft: string) => void;
  commitEdit: (move?: { dRow: number; dCol: number }) => void;
  cancelEdit: () => void;

  // --- values ---
  setCell: (addr: CellAddr, input: string) => void;
  clearSelection: () => void;

  // --- styles ---
  applyStyle: (patch: CellStyle) => void;
  toggleStyle: (key: 'bold' | 'italic' | 'underline' | 'strike' | 'wrap') => void;
  clearFormatting: () => void;

  // --- structure ---
  setColWidth: (col: number, width: number) => void;
  setRowHeight: (row: number, height: number) => void;
  insertRows: (at: number, count: number) => void;
  deleteRows: (at: number, count: number) => void;
  insertCols: (at: number, count: number) => void;
  deleteCols: (at: number, count: number) => void;
  toggleFreeze: () => void;

  // --- sheets ---
  addSheet: () => void;
  renameSheet: (id: number, name: string) => void;
  deleteSheet: (id: number) => void;
  setActiveSheet: (id: number) => void;
  replaceWorkbook: (fileName: string, sheets: SheetMeta[]) => void;

  // --- clipboard & history ---
  copy: () => void;
  cut: () => void;
  paste: () => void;
  pasteText: (text: string) => void;
  fillDown: () => void;
  fillFromHandle: (target: CellRange) => void;
  undo: () => void;
  redo: () => void;

  // --- helpers ---
  selection: () => CellRange;
  activeSheet: () => SheetMeta;
  styleAt: (addr: CellAddr) => CellStyle | undefined;
  colWidth: (col: number) => number;
  rowHeight: (row: number) => number;
}

function blankSheet(id: number, name: string): SheetMeta {
  return {
    id,
    name,
    colWidths: {},
    rowHeights: {},
    styles: {},
    frozenRows: 0,
    frozenCols: 0,
    merges: [],
  };
}

/** Creates the first sheet in HyperFormula so ids line up from the start. */
function bootstrap(): SheetMeta[] {
  const name = hf.addSheet('Sheet1');
  const id = hf.getSheetId(name)!;
  return [blankSheet(id, 'Sheet1')];
}

const clamp = (n: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, n));

/** Cell count above which a whole-sheet operation is trimmed to the used area. */
const BULK_LIMIT = 50_000;

/**
 * Trims a range to the sheet's populated area.
 *
 * Selections can legitimately cover the whole 5000x200 grid (Ctrl+A), and
 * writing a million blanks through the engine exhausts the heap. Everything
 * outside the used area is already empty, so clipping is lossless.
 */
function clampToUsed(sheetId: number, r: CellRange): CellRange | null {
  if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) <= BULK_LIMIT) return r;
  const dim = hf.getSheetDimensions(sheetId);
  const r2 = Math.min(r.r2, Math.max(dim.height - 1, r.r1));
  const c2 = Math.min(r.c2, Math.max(dim.width - 1, r.c1));
  if (r2 < r.r1 || c2 < r.c1) return null;
  return { r1: r.r1, c1: r.c1, r2, c2 };
}

/** Rebuilds a style map after rows/cols are inserted or removed. */
function shiftStyles(
  styles: StyleMap,
  axis: 'row' | 'col',
  at: number,
  delta: number,
): StyleMap {
  const next: StyleMap = {};
  for (const [key, style] of Object.entries(styles)) {
    const [rowStr, colStr] = key.split(':');
    let row = Number(rowStr);
    let col = Number(colStr);
    const pos = axis === 'row' ? row : col;
    if (delta < 0 && pos >= at && pos < at - delta) continue; // deleted
    if (pos >= at) {
      if (axis === 'row') row += delta;
      else col += delta;
    }
    if (row < 0 || col < 0) continue;
    next[styleKey(row, col)] = style;
  }
  return next;
}

function shiftSizes(
  sizes: Record<number, number>,
  at: number,
  delta: number,
): Record<number, number> {
  const next: Record<number, number> = {};
  for (const [k, v] of Object.entries(sizes)) {
    const pos = Number(k);
    if (delta < 0 && pos >= at && pos < at - delta) continue;
    next[pos >= at ? pos + delta : pos] = v;
  }
  return next;
}

export const useSheets = create<SheetsState>((set, get) => ({
  sheets: bootstrap(),
  activeSheetId: 0,
  anchor: { row: 0, col: 0 },
  focus: { row: 0, col: 0 },
  editing: null,
  clipboard: null,
  rev: 0,
  fileName: 'Book1',
  dirty: false,

  selection: () => normaliseRange(get().anchor, get().focus),
  activeSheet: () => {
    const s = get();
    return s.sheets.find((sh) => sh.id === s.activeSheetId) ?? s.sheets[0];
  },
  styleAt: (addr) => get().activeSheet().styles[styleKey(addr.row, addr.col)],
  colWidth: (col) => get().activeSheet().colWidths[col] ?? DEFAULT_COL_WIDTH,
  rowHeight: (row) => get().activeSheet().rowHeights[row] ?? DEFAULT_ROW_HEIGHT,

  select: (addr, extend = false) =>
    set((s) => ({
      anchor: extend ? s.anchor : addr,
      focus: addr,
      editing: null,
    })),

  selectRange: (r) =>
    set({ anchor: { row: r.r1, col: r.c1 }, focus: { row: r.r2, col: r.c2 }, editing: null }),

  moveSelection: (dRow, dCol, extend) =>
    set((s) => {
      const next = {
        row: clamp(s.focus.row + dRow, 0, MAX_ROWS - 1),
        col: clamp(s.focus.col + dCol, 0, MAX_COLS - 1),
      };
      return { focus: next, anchor: extend ? s.anchor : next, editing: null };
    }),

  selectAll: () =>
    set({ anchor: { row: 0, col: 0 }, focus: { row: MAX_ROWS - 1, col: MAX_COLS - 1 } }),

  selectRow: (row, extend = false) =>
    set((s) => ({
      anchor: extend ? { row: s.anchor.row, col: 0 } : { row, col: 0 },
      focus: { row, col: MAX_COLS - 1 },
      editing: null,
    })),

  selectCol: (col, extend = false) =>
    set((s) => ({
      anchor: extend ? { row: 0, col: s.anchor.col } : { row: 0, col },
      focus: { row: MAX_ROWS - 1, col },
      editing: null,
    })),

  beginEdit: (addr, seed, selectAll = false) =>
    set((s) => ({
      anchor: addr,
      focus: addr,
      editing: {
        addr,
        draft: seed !== undefined ? seed : readRaw(s.activeSheetId, addr),
        selectAll,
      },
    })),

  updateDraft: (draft) =>
    set((s) => (s.editing ? { editing: { ...s.editing, draft } } : {})),

  commitEdit: (move) => {
    const s = get();
    if (!s.editing) return;
    const { addr, draft } = s.editing;
    writeCell(s.activeSheetId, addr, draft);
    set({ editing: null, rev: s.rev + 1, dirty: true });
    if (move) get().moveSelection(move.dRow, move.dCol, false);
  },

  cancelEdit: () => set({ editing: null }),

  setCell: (addr, input) => {
    const s = get();
    writeCell(s.activeSheetId, addr, input);
    set({ rev: s.rev + 1, dirty: true });
  },

  clearSelection: () => {
    const s = get();
    const target = clampToUsed(s.activeSheetId, s.selection());
    if (!target) return;
    clearRange(s.activeSheetId, target);
    set({ rev: s.rev + 1, dirty: true });
  },

  applyStyle: (patch) =>
    set((s) => {
      const sel = clampToUsed(s.activeSheetId, s.selection());
      if (!sel) return {};
      const sheets = s.sheets.map((sh) => {
        if (sh.id !== s.activeSheetId) return sh;
        const styles = { ...sh.styles };
        for (let row = sel.r1; row <= sel.r2; row++) {
          for (let col = sel.c1; col <= sel.c2; col++) {
            const k = styleKey(row, col);
            styles[k] = { ...styles[k], ...patch };
          }
        }
        return { ...sh, styles };
      });
      return { sheets, dirty: true };
    }),

  toggleStyle: (key) => {
    const s = get();
    const sel = s.selection();
    const current = s.activeSheet().styles[styleKey(sel.r1, sel.c1)]?.[key];
    s.applyStyle({ [key]: !current } as CellStyle);
  },

  // Filters the style map rather than walking the selection, so a select-all
  // costs O(styled cells) instead of O(grid).
  clearFormatting: () =>
    set((s) => {
      const sel = s.selection();
      const sheets = s.sheets.map((sh) => {
        if (sh.id !== s.activeSheetId) return sh;
        const styles: StyleMap = {};
        for (const [key, style] of Object.entries(sh.styles)) {
          const [row, col] = key.split(':').map(Number);
          const inside = row >= sel.r1 && row <= sel.r2 && col >= sel.c1 && col <= sel.c2;
          if (!inside) styles[key] = style;
        }
        return { ...sh, styles };
      });
      return { sheets, dirty: true };
    }),

  setColWidth: (col, width) =>
    set((s) => ({
      sheets: s.sheets.map((sh) =>
        sh.id === s.activeSheetId
          ? { ...sh, colWidths: { ...sh.colWidths, [col]: Math.max(24, Math.round(width)) } }
          : sh,
      ),
      dirty: true,
    })),

  setRowHeight: (row, height) =>
    set((s) => ({
      sheets: s.sheets.map((sh) =>
        sh.id === s.activeSheetId
          ? { ...sh, rowHeights: { ...sh.rowHeights, [row]: Math.max(16, Math.round(height)) } }
          : sh,
      ),
      dirty: true,
    })),

  insertRows: (at, count) => {
    const s = get();
    hf.addRows(s.activeSheetId, [at, count]);
    set({
      rev: s.rev + 1,
      dirty: true,
      sheets: s.sheets.map((sh) =>
        sh.id === s.activeSheetId
          ? {
              ...sh,
              styles: shiftStyles(sh.styles, 'row', at, count),
              rowHeights: shiftSizes(sh.rowHeights, at, count),
            }
          : sh,
      ),
    });
  },

  deleteRows: (at, count) => {
    const s = get();
    hf.removeRows(s.activeSheetId, [at, count]);
    set({
      rev: s.rev + 1,
      dirty: true,
      sheets: s.sheets.map((sh) =>
        sh.id === s.activeSheetId
          ? {
              ...sh,
              styles: shiftStyles(sh.styles, 'row', at, -count),
              rowHeights: shiftSizes(sh.rowHeights, at, -count),
            }
          : sh,
      ),
    });
  },

  insertCols: (at, count) => {
    const s = get();
    hf.addColumns(s.activeSheetId, [at, count]);
    set({
      rev: s.rev + 1,
      dirty: true,
      sheets: s.sheets.map((sh) =>
        sh.id === s.activeSheetId
          ? {
              ...sh,
              styles: shiftStyles(sh.styles, 'col', at, count),
              colWidths: shiftSizes(sh.colWidths, at, count),
            }
          : sh,
      ),
    });
  },

  deleteCols: (at, count) => {
    const s = get();
    hf.removeColumns(s.activeSheetId, [at, count]);
    set({
      rev: s.rev + 1,
      dirty: true,
      sheets: s.sheets.map((sh) =>
        sh.id === s.activeSheetId
          ? {
              ...sh,
              styles: shiftStyles(sh.styles, 'col', at, -count),
              colWidths: shiftSizes(sh.colWidths, at, -count),
            }
          : sh,
      ),
    });
  },

  toggleFreeze: () =>
    set((s) => {
      const sh = s.activeSheet();
      const frozen = sh.frozenRows > 0 || sh.frozenCols > 0;
      const sel = s.selection();
      return {
        sheets: s.sheets.map((x) =>
          x.id === s.activeSheetId
            ? {
                ...x,
                frozenRows: frozen ? 0 : sel.r1,
                frozenCols: frozen ? 0 : sel.c1,
              }
            : x,
        ),
      };
    }),

  addSheet: () => {
    const s = get();
    let n = s.sheets.length + 1;
    while (s.sheets.some((sh) => sh.name === `Sheet${n}`)) n++;
    const name = hf.addSheet(`Sheet${n}`);
    const id = hf.getSheetId(name)!;
    set({
      sheets: [...s.sheets, blankSheet(id, name)],
      activeSheetId: id,
      anchor: { row: 0, col: 0 },
      focus: { row: 0, col: 0 },
      dirty: true,
    });
  },

  renameSheet: (id, name) => {
    const trimmed = name.trim();
    const s = get();
    if (!trimmed || s.sheets.some((sh) => sh.id !== id && sh.name === trimmed)) return;
    hf.renameSheet(id, trimmed);
    set({
      sheets: s.sheets.map((sh) => (sh.id === id ? { ...sh, name: trimmed } : sh)),
      dirty: true,
    });
  },

  deleteSheet: (id) => {
    const s = get();
    if (s.sheets.length === 1) return;
    hf.removeSheet(id);
    const sheets = s.sheets.filter((sh) => sh.id !== id);
    set({
      sheets,
      activeSheetId: s.activeSheetId === id ? sheets[0].id : s.activeSheetId,
      rev: s.rev + 1,
      dirty: true,
    });
  },

  setActiveSheet: (id) =>
    set({ activeSheetId: id, anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 }, editing: null }),

  replaceWorkbook: (fileName, sheets) =>
    set({
      fileName,
      sheets,
      activeSheetId: sheets[0]?.id ?? 0,
      anchor: { row: 0, col: 0 },
      focus: { row: 0, col: 0 },
      editing: null,
      clipboard: null,
      rev: get().rev + 1,
      dirty: false,
    }),

  copy: () => {
    const s = get();
    const sel = s.selection();
    set({
      clipboard: { rows: rawRange(s.activeSheetId, sel), cut: false, origin: sel, sheetId: s.activeSheetId },
    });
  },

  cut: () => {
    const s = get();
    const sel = s.selection();
    set({
      clipboard: { rows: rawRange(s.activeSheetId, sel), cut: true, origin: sel, sheetId: s.activeSheetId },
    });
  },

  paste: () => {
    const s = get();
    const clip = s.clipboard;
    if (!clip) return;
    const target = s.selection();
    if (clip.cut) clearRange(clip.sheetId, clip.origin);
    const dRow = target.r1 - clip.origin.r1;
    const dCol = target.c1 - clip.origin.c1;
    const shifted = clip.rows.map((row) =>
      row.map((cell) => (cell.startsWith('=') ? offsetFormula(cell, dRow, dCol) : cell)),
    );
    writeBlock(s.activeSheetId, { row: target.r1, col: target.c1 }, shifted);
    set({
      rev: s.rev + 1,
      dirty: true,
      clipboard: clip.cut ? null : clip,
      focus: {
        row: target.r1 + shifted.length - 1,
        col: target.c1 + (shifted[0]?.length ?? 1) - 1,
      },
      anchor: { row: target.r1, col: target.c1 },
    });
  },

  /** Paste from the OS clipboard: TSV (Excel/Sheets) or newline-separated text. */
  pasteText: (text) => {
    const s = get();
    const rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((line) => line.split('\t'));
    if (rows.length === 0) return;
    const target = s.selection();
    writeBlock(s.activeSheetId, { row: target.r1, col: target.c1 }, rows);
    set({
      rev: s.rev + 1,
      dirty: true,
      anchor: { row: target.r1, col: target.c1 },
      focus: {
        row: target.r1 + rows.length - 1,
        col: target.c1 + (rows[0]?.length ?? 1) - 1,
      },
    });
  },

  fillDown: () => {
    const s = get();
    const sel = s.selection();
    if (sel.r2 === sel.r1) return;
    const source = rawRange(s.activeSheetId, { ...sel, r2: sel.r1 });
    const out: string[][] = [];
    for (let row = sel.r1 + 1; row <= sel.r2; row++) {
      out.push(
        source[0].map((cell) => (cell.startsWith('=') ? offsetFormula(cell, row - sel.r1, 0) : cell)),
      );
    }
    writeBlock(s.activeSheetId, { row: sel.r1 + 1, col: sel.c1 }, out);
    set({ rev: s.rev + 1, dirty: true });
  },

  /** Extends the current selection's contents over `target` (fill handle drag). */
  fillFromHandle: (target) => {
    const s = get();
    const src = s.selection();
    const source = rawRange(s.activeSheetId, src);
    const srcH = src.r2 - src.r1 + 1;
    const srcW = src.c2 - src.c1 + 1;
    const out: string[][] = [];
    for (let row = target.r1; row <= target.r2; row++) {
      const line: string[] = [];
      for (let col = target.c1; col <= target.c2; col++) {
        const sr = (((row - src.r1) % srcH) + srcH) % srcH;
        const sc = (((col - src.c1) % srcW) + srcW) % srcW;
        const cell = source[sr][sc];
        const dRow = row - (src.r1 + sr);
        const dCol = col - (src.c1 + sc);
        line.push(cell.startsWith('=') ? offsetFormula(cell, dRow, dCol) : cell);
      }
      out.push(line);
    }
    writeBlock(s.activeSheetId, { row: target.r1, col: target.c1 }, out);
    set({ rev: s.rev + 1, dirty: true, anchor: { row: target.r1, col: target.c1 }, focus: { row: target.r2, col: target.c2 } });
  },

  undo: () => {
    if (!hf.isThereSomethingToUndo()) return;
    hf.undo();
    set({ rev: get().rev + 1, dirty: true });
  },

  redo: () => {
    if (!hf.isThereSomethingToRedo()) return;
    hf.redo();
    set({ rev: get().rev + 1, dirty: true });
  },
}));
