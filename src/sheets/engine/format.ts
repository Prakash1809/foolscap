import type { CellStyle, NumFmt } from '../../shared/types';

/** Excel serial-date epoch (1899-12-30 UTC) in ms, matching HyperFormula's default nullDate. */
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

export function serialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH_MS + Math.round(serial * MS_PER_DAY));
}

export function dateToSerial(d: Date): number {
  return (d.getTime() - EXCEL_EPOCH_MS) / MS_PER_DAY;
}

const groupNum = (n: number, decimals: number) =>
  n.toLocaleString(undefined, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });

/** Trims a float to at most 10 significant decimals so 0.1+0.2 renders as 0.3. */
function tidyFloat(n: number): string {
  if (Number.isInteger(n)) return String(n);
  const r = parseFloat(n.toPrecision(12));
  return String(r);
}

function pad(n: number): string {
  return String(n).padStart(2, '0');
}

export function formatValue(value: unknown, style: CellStyle | undefined): string {
  if (value === null || value === undefined || value === '') return '';

  const fmt: NumFmt = style?.numFmt ?? 'general';

  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return '#NUM!';
    switch (fmt) {
      case 'number':
        return groupNum(value, 2);
      case 'number0':
        return groupNum(value, 0);
      case 'currency':
        return value.toLocaleString(undefined, {
          style: 'currency',
          currency: 'USD',
        });
      case 'percent':
        return `${groupNum(value * 100, 2)}%`;
      case 'date': {
        const d = serialToDate(value);
        return `${pad(d.getUTCMonth() + 1)}/${pad(d.getUTCDate())}/${d.getUTCFullYear()}`;
      }
      case 'datetime': {
        const d = serialToDate(value);
        return (
          `${pad(d.getUTCMonth() + 1)}/${pad(d.getUTCDate())}/${d.getUTCFullYear()} ` +
          `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`
        );
      }
      case 'text':
        return tidyFloat(value);
      default:
        return tidyFloat(value);
    }
  }

  return String(value);
}

/** Numbers and dates right-align by default, like Excel. Text left-aligns. */
export function defaultAlign(value: unknown): 'left' | 'right' | 'center' {
  if (typeof value === 'number') return 'right';
  if (typeof value === 'boolean') return 'center';
  return 'left';
}

export const NUM_FMT_LABELS: Record<NumFmt, string> = {
  general: 'General',
  number0: 'Number (0 dp)',
  number: 'Number (2 dp)',
  currency: 'Currency',
  percent: 'Percent',
  date: 'Short date',
  datetime: 'Date & time',
  text: 'Text',
};
