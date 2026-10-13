import { describe, expect, it } from 'vitest';
import { coerceInput, offsetFormula } from './engine';
import { formatValue, dateToSerial, serialToDate } from './format';
import { colIndex, colLabel, parseAddr, rangeLabel } from '../../shared/cellref';

describe('colLabel / colIndex', () => {
  it('round-trips the first few hundred columns', () => {
    for (const col of [0, 1, 25, 26, 27, 51, 52, 701, 702]) {
      expect(colIndex(colLabel(col))).toBe(col);
    }
  });

  it('matches Excel labels', () => {
    expect(colLabel(0)).toBe('A');
    expect(colLabel(25)).toBe('Z');
    expect(colLabel(26)).toBe('AA');
    expect(colLabel(27)).toBe('AB');
    expect(colLabel(701)).toBe('ZZ');
  });
});

describe('parseAddr', () => {
  it('parses A1 refs case-insensitively', () => {
    expect(parseAddr('B7')).toEqual({ row: 6, col: 1 });
    expect(parseAddr('  aa10 ')).toEqual({ row: 9, col: 26 });
  });

  it('rejects anything that is not a plain ref', () => {
    expect(parseAddr('A0')).toBeNull();
    expect(parseAddr('1A')).toBeNull();
    expect(parseAddr('A1:B2')).toBeNull();
    expect(parseAddr('')).toBeNull();
  });
});

describe('rangeLabel', () => {
  it('collapses single cells', () => {
    expect(rangeLabel({ r1: 0, c1: 0, r2: 0, c2: 0 })).toBe('A1');
    expect(rangeLabel({ r1: 0, c1: 0, r2: 4, c2: 2 })).toBe('A1:C5');
  });
});

describe('coerceInput', () => {
  it('keeps formulas as text', () => {
    expect(coerceInput('=SUM(A1:A3)')).toBe('=SUM(A1:A3)');
  });

  it('converts numeric strings to numbers', () => {
    expect(coerceInput('42')).toBe(42);
    expect(coerceInput('-3.5')).toBe(-3.5);
    expect(coerceInput('.25')).toBe(0.25);
    expect(coerceInput('1e3')).toBe(1000);
  });

  it('leaves non-numeric text alone', () => {
    expect(coerceInput('42abc')).toBe('42abc');
    expect(coerceInput('hello')).toBe('hello');
    expect(coerceInput('  ')).toBeNull();
  });

  it('recognises booleans', () => {
    expect(coerceInput('true')).toBe(true);
    expect(coerceInput('FALSE')).toBe(false);
  });
});

describe('offsetFormula', () => {
  it('shifts relative refs', () => {
    expect(offsetFormula('=A1+B2', 1, 0)).toBe('=A2+B3');
    expect(offsetFormula('=A1', 0, 2)).toBe('=C1');
    expect(offsetFormula('=SUM(A1:A3)', 2, 0)).toBe('=SUM(A3:A5)');
  });

  it('honours absolute markers', () => {
    expect(offsetFormula('=$A$1', 5, 5)).toBe('=$A$1');
    expect(offsetFormula('=$A1', 1, 1)).toBe('=$A2');
    expect(offsetFormula('=A$1', 1, 1)).toBe('=B$1');
  });

  it('does not renumber function names that end in digits', () => {
    expect(offsetFormula('=LOG10(A1)', 1, 0)).toBe('=LOG10(A2)');
    expect(offsetFormula('=ATAN2(A1,B1)', 1, 0)).toBe('=ATAN2(A2,B2)');
  });

  it('leaves string literals untouched', () => {
    expect(offsetFormula('=IF(A1>0,"A1","B2")', 1, 0)).toBe('=IF(A2>0,"A1","B2")');
  });

  it('does not treat exponents as refs', () => {
    expect(offsetFormula('=1E2+A1', 1, 0)).toBe('=1E2+A2');
  });

  it('shifts sheet-qualified refs', () => {
    expect(offsetFormula('=Sheet2!A1', 1, 0)).toBe('=Sheet2!A2');
  });

  it('yields #REF! when a shift goes off-grid', () => {
    expect(offsetFormula('=A1', -1, 0)).toBe('=#REF!');
  });

  it('passes literals through unchanged', () => {
    expect(offsetFormula('plain text', 3, 3)).toBe('plain text');
  });
});

describe('formatValue', () => {
  it('renders blanks for empty values', () => {
    expect(formatValue(null, undefined)).toBe('');
    expect(formatValue('', undefined)).toBe('');
  });

  it('tidies float noise under the general format', () => {
    expect(formatValue(0.1 + 0.2, undefined)).toBe('0.3');
    expect(formatValue(7, undefined)).toBe('7');
  });

  it('applies percent and currency formats', () => {
    expect(formatValue(0.25, { numFmt: 'percent' })).toBe('25.00%');
    expect(formatValue(1234.5, { numFmt: 'currency' })).toContain('1,234.50');
  });

  it('renders booleans like Excel', () => {
    expect(formatValue(true, undefined)).toBe('TRUE');
    expect(formatValue(false, undefined)).toBe('FALSE');
  });

  it('formats serial dates', () => {
    // 1 Jan 2024 is serial 45292 on the 1899-12-30 epoch.
    expect(formatValue(45292, { numFmt: 'date' })).toBe('01/01/2024');
  });
});

describe('date serials', () => {
  it('round-trips through the Excel epoch', () => {
    const d = new Date(Date.UTC(2024, 5, 15));
    expect(serialToDate(dateToSerial(d)).toISOString()).toBe(d.toISOString());
  });
});
