import ExcelJS from 'exceljs';
import { hf, readCell, readRaw } from '../engine/engine';
import { dateToSerial } from '../engine/format';
import { styleKey } from '../../shared/cellref';
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  MAX_COLS,
  MAX_ROWS,
  type CellStyle,
  type NumFmt,
  type SheetMeta,
  type StyleMap,
} from '../../shared/types';

/** Excel column width is in "characters"; ~7px per character at the default font. */
const PX_PER_CHAR = 7;
/** Row height is in points. */
const PX_PER_POINT = 4 / 3;

function argbToHex(argb: string | undefined): string | undefined {
  if (!argb) return undefined;
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  return `#${hex.toLowerCase()}`;
}

function hexToArgb(hex: string | undefined): string | undefined {
  if (!hex) return undefined;
  const h = hex.replace('#', '');
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return undefined;
  return `FF${h.toUpperCase()}`;
}

/** Best-effort mapping from an Excel number-format code to our fixed set. */
function numFmtFromCode(code: string | undefined): NumFmt | undefined {
  if (!code || code === 'General') return undefined;
  const c = code.toLowerCase();
  if (c.includes('%')) return 'percent';
  if (/[$€£¥]/.test(code)) return 'currency';
  if (c.includes('h') && (c.includes('y') || c.includes('d'))) return 'datetime';
  if (c.includes('yy') || c.includes('dd') || c.includes('mmm')) return 'date';
  if (c === '@') return 'text';
  if (c.includes('0.00')) return 'number';
  if (c.includes('#,##0') || c === '0') return 'number0';
  return undefined;
}

const NUM_FMT_CODES: Record<NumFmt, string> = {
  general: 'General',
  number0: '#,##0',
  number: '#,##0.00',
  currency: '"$"#,##0.00',
  percent: '0.00%',
  date: 'mm/dd/yyyy',
  datetime: 'mm/dd/yyyy hh:mm',
  text: '@',
};

function readCellStyle(cell: ExcelJS.Cell): CellStyle | undefined {
  const style: CellStyle = {};
  const font = cell.font;
  if (font?.bold) style.bold = true;
  if (font?.italic) style.italic = true;
  if (font?.underline) style.underline = true;
  if (font?.strike) style.strike = true;
  if (font?.size && font.size !== 11) style.fontSize = Math.round(font.size * 1.18);
  if (font?.name && font.name !== 'Calibri') style.fontFamily = font.name;
  const fontColor = argbToHex((font?.color as { argb?: string } | undefined)?.argb);
  if (fontColor && fontColor !== '#000000') style.color = fontColor;

  const fill = cell.fill as ExcelJS.FillPattern | undefined;
  const bg = argbToHex((fill?.fgColor as { argb?: string } | undefined)?.argb);
  if (fill?.type === 'pattern' && bg && bg !== '#ffffff') style.bg = bg;

  const al = cell.alignment;
  if (al?.horizontal === 'left' || al?.horizontal === 'center' || al?.horizontal === 'right') {
    style.align = al.horizontal;
  }
  if (al?.vertical === 'top' || al?.vertical === 'middle' || al?.vertical === 'bottom') {
    style.valign = al.vertical;
  }
  if (al?.wrapText) style.wrap = true;

  const fmt = numFmtFromCode(cell.numFmt);
  if (fmt) style.numFmt = fmt;

  const b = cell.border;
  if (b && (b.top || b.right || b.bottom || b.left)) {
    style.border = {
      top: Boolean(b.top),
      right: Boolean(b.right),
      bottom: Boolean(b.bottom),
      left: Boolean(b.left),
    };
  }

  return Object.keys(style).length > 0 ? style : undefined;
}

/** ExcelJS cell value -> raw input string for HyperFormula. */
function toRawInput(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return '';

  if (typeof v === 'object') {
    if ('formula' in v && v.formula) return `=${v.formula}`;
    if ('sharedFormula' in v && v.sharedFormula) {
      // A shared-formula cell carries no expression of its own. Fall back to the
      // cached result; HyperFormula cannot re-derive the master's formula here.
      const result = (v as { result?: unknown }).result;
      return result === undefined || result === null ? '' : String(result);
    }
    if ('richText' in v && Array.isArray(v.richText)) {
      return v.richText.map((r) => r.text).join('');
    }
    if ('error' in v) return String(v.error);
    if ('text' in v) return String((v as { text: unknown }).text);
    if (v instanceof Date) return String(dateToSerial(v));
  }

  if (v instanceof Date) return String(dateToSerial(v));
  return String(v);
}

export interface LoadResult {
  fileName: string;
  sheets: SheetMeta[];
}

/** Replaces every sheet in the HyperFormula instance with the workbook's contents. */
export async function loadXlsx(file: File): Promise<LoadResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());

  // Drop existing sheets only after the parse succeeds, so a bad file is non-destructive.
  for (const name of hf.getSheetNames()) {
    hf.removeSheet(hf.getSheetId(name)!);
  }

  const sheets: SheetMeta[] = [];

  wb.eachSheet((ws) => {
    const safeName = (ws.name || `Sheet${sheets.length + 1}`).slice(0, 31);
    const created = hf.addSheet(safeName);
    const id = hf.getSheetId(created)!;

    const styles: StyleMap = {};
    const colWidths: Record<number, number> = {};
    const rowHeights: Record<number, number> = {};
    const contents: (string | number | boolean | null)[][] = [];

    ws.columns?.forEach((col, i) => {
      if (col?.width) colWidths[i] = Math.round(col.width * PX_PER_CHAR);
    });

    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const r = rowNumber - 1;
      if (r >= MAX_ROWS) return;
      if (row.height) rowHeights[r] = Math.round(row.height * PX_PER_POINT);

      const line: (string | number | boolean | null)[] = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const c = colNumber - 1;
        if (c >= MAX_COLS) return;
        const raw = toRawInput(cell);
        line[c] = raw === '' ? null : raw.startsWith('=') ? raw : coerce(raw);
        const st = readCellStyle(cell);
        if (st) styles[styleKey(r, c)] = st;
      });
      contents[r] = line;
    });

    // Normalise ragged rows so setSheetContent gets a rectangle.
    const width = contents.reduce((m, row) => Math.max(m, row?.length ?? 0), 1);
    for (let r = 0; r < contents.length; r++) {
      if (!contents[r]) contents[r] = [];
      for (let c = 0; c < width; c++) if (contents[r][c] === undefined) contents[r][c] = null;
    }
    if (contents.length > 0) hf.setSheetContent(id, contents);

    const frozen = ws.views?.[0];
    sheets.push({
      id,
      name: created,
      colWidths,
      rowHeights,
      styles,
      frozenRows: frozen && 'ySplit' in frozen ? (frozen.ySplit ?? 0) : 0,
      frozenCols: frozen && 'xSplit' in frozen ? (frozen.xSplit ?? 0) : 0,
      merges: [],
    });
  });

  if (sheets.length === 0) {
    const created = hf.addSheet('Sheet1');
    sheets.push({
      id: hf.getSheetId(created)!,
      name: created,
      colWidths: {},
      rowHeights: {},
      styles: {},
      frozenRows: 0,
      frozenCols: 0,
      merges: [],
    });
  }

  return { fileName: file.name.replace(/\.xlsx?$/i, ''), sheets };
}

function coerce(text: string): string | number | boolean {
  if (/^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(text)) {
    const n = Number(text);
    if (Number.isFinite(n)) return n;
  }
  if (/^(TRUE|FALSE)$/i.test(text)) return text.toUpperCase() === 'TRUE';
  return text;
}

/** Widest populated row/col per sheet, so export does not emit 5000 blank rows. */
function usedBounds(sheetId: number): { rows: number; cols: number } {
  const dim = hf.getSheetDimensions(sheetId);
  return {
    rows: Math.min(dim.height, MAX_ROWS),
    cols: Math.min(dim.width, MAX_COLS),
  };
}

export async function saveXlsx(fileName: string, sheets: SheetMeta[]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Foolscap Sheets';
  wb.created = new Date();

  for (const sheet of sheets) {
    const ws = wb.addWorksheet(sheet.name);
    const { rows, cols } = usedBounds(sheet.id);

    const styledKeys = Object.keys(sheet.styles);
    const maxStyledRow = styledKeys.reduce((m, k) => Math.max(m, Number(k.split(':')[0]) + 1), 0);
    const maxStyledCol = styledKeys.reduce((m, k) => Math.max(m, Number(k.split(':')[1]) + 1), 0);
    const rowCount = Math.max(rows, maxStyledRow);
    const colCount = Math.max(cols, maxStyledCol);

    for (let c = 0; c < colCount; c++) {
      const px = sheet.colWidths[c] ?? DEFAULT_COL_WIDTH;
      ws.getColumn(c + 1).width = px / PX_PER_CHAR;
    }

    for (let r = 0; r < rowCount; r++) {
      const px = sheet.rowHeights[r];
      if (px && px !== DEFAULT_ROW_HEIGHT) ws.getRow(r + 1).height = px / PX_PER_POINT;

      for (let c = 0; c < colCount; c++) {
        const raw = readRaw(sheet.id, { row: r, col: c });
        const style = sheet.styles[styleKey(r, c)];
        if (raw === '' && !style) continue;

        const cell = ws.getCell(r + 1, c + 1);
        if (raw.startsWith('=')) {
          const snap = readCell(sheet.id, { row: r, col: c });
          cell.value = { formula: raw.slice(1), result: snap.isError ? undefined : (snap.value as never) };
        } else if (raw !== '') {
          cell.value = coerce(raw) as ExcelJS.CellValue;
        }

        if (!style) continue;
        applyStyleToCell(cell, style);
      }
    }

    if (sheet.frozenRows > 0 || sheet.frozenCols > 0) {
      ws.views = [
        {
          state: 'frozen',
          xSplit: sheet.frozenCols,
          ySplit: sheet.frozenRows,
        },
      ];
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  download(
    new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    `${fileName}.xlsx`,
  );
}

function applyStyleToCell(cell: ExcelJS.Cell, style: CellStyle): void {
  const font: Partial<ExcelJS.Font> = {};
  if (style.bold) font.bold = true;
  if (style.italic) font.italic = true;
  if (style.underline) font.underline = true;
  if (style.strike) font.strike = true;
  if (style.fontSize) font.size = Math.round(style.fontSize / 1.18);
  if (style.fontFamily) font.name = style.fontFamily;
  const colorArgb = hexToArgb(style.color);
  if (colorArgb) font.color = { argb: colorArgb };
  if (Object.keys(font).length > 0) cell.font = font as ExcelJS.Font;

  const bgArgb = hexToArgb(style.bg);
  if (bgArgb) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
  }

  if (style.align || style.valign || style.wrap) {
    cell.alignment = {
      horizontal: style.align,
      vertical: style.valign === 'middle' ? 'middle' : style.valign,
      wrapText: style.wrap,
    };
  }

  if (style.numFmt && style.numFmt !== 'general') cell.numFmt = NUM_FMT_CODES[style.numFmt];

  if (style.border) {
    const thin: ExcelJS.Border = { style: 'thin', color: { argb: 'FF4B5563' } };
    cell.border = {
      top: style.border.top ? thin : undefined,
      right: style.border.right ? thin : undefined,
      bottom: style.border.bottom ? thin : undefined,
      left: style.border.left ? thin : undefined,
    };
  }
}

export function exportCsv(fileName: string, sheet: SheetMeta): void {
  const { rows, cols } = usedBounds(sheet.id);
  const lines: string[] = [];
  for (let r = 0; r < rows; r++) {
    const cells: string[] = [];
    for (let c = 0; c < cols; c++) {
      const snap = readCell(sheet.id, { row: r, col: c });
      const text = snap.value === null ? '' : String(snap.value);
      cells.push(/[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
    }
    lines.push(cells.join(','));
  }
  download(new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), `${fileName}.csv`);
}

function download(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}
