import { beforeEach, describe, expect, it } from 'vitest';
import { useSheets } from './store';
import { readCell, readRaw } from './engine/engine';

const store = () => useSheets.getState();

function reset() {
  const s = store();
  // Keep one sheet and wipe it, so each test starts from a clean A1.
  for (const sheet of s.sheets.slice(1)) s.deleteSheet(sheet.id);
  s.selectAll();
  s.clearSelection();
  s.clearFormatting();
  s.select({ row: 0, col: 0 });
  useSheets.setState({ clipboard: null, dirty: false });
}

const valueAt = (row: number, col: number) =>
  readCell(store().activeSheetId, { row, col }).value;
const rawAt = (row: number, col: number) => readRaw(store().activeSheetId, { row, col });

describe('cell values and formulas', () => {
  beforeEach(reset);

  it('stores literals and recalculates dependents', () => {
    store().setCell({ row: 0, col: 0 }, '2');
    store().setCell({ row: 1, col: 0 }, '3');
    store().setCell({ row: 2, col: 0 }, '=SUM(A1:A2)');
    expect(valueAt(2, 0)).toBe(5);

    store().setCell({ row: 0, col: 0 }, '10');
    expect(valueAt(2, 0)).toBe(13);
  });

  it('surfaces errors with an Excel-style code', () => {
    store().setCell({ row: 0, col: 0 }, '=1/0');
    const snap = readCell(store().activeSheetId, { row: 0, col: 0 });
    expect(snap.isError).toBe(true);
    expect(snap.value).toBe('#DIV/0!');
  });

  it('keeps the raw formula for the formula bar', () => {
    store().setCell({ row: 0, col: 0 }, '=1+1');
    expect(rawAt(0, 0)).toBe('=1+1');
    expect(valueAt(0, 0)).toBe(2);
  });
});

describe('selection', () => {
  beforeEach(reset);

  it('normalises a backwards drag', () => {
    store().select({ row: 5, col: 5 });
    store().select({ row: 1, col: 2 }, true);
    expect(store().selection()).toEqual({ r1: 1, c1: 2, r2: 5, c2: 5 });
  });

  it('clamps movement to the sheet bounds', () => {
    store().select({ row: 0, col: 0 });
    store().moveSelection(-5, -5, false);
    expect(store().focus).toEqual({ row: 0, col: 0 });
  });
});

describe('fill', () => {
  beforeEach(reset);

  it('fills a formula down and rewrites relative refs', () => {
    store().setCell({ row: 0, col: 0 }, '1');
    store().setCell({ row: 1, col: 0 }, '2');
    store().setCell({ row: 0, col: 1 }, '=A1*10');
    store().selectRange({ r1: 0, c1: 1, r2: 1, c2: 1 });
    store().fillDown();
    expect(rawAt(1, 1)).toBe('=A2*10');
    expect(valueAt(1, 1)).toBe(20);
  });

  it('repeats the source block when dragged past it', () => {
    store().setCell({ row: 0, col: 0 }, 'x');
    store().selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
    store().fillFromHandle({ r1: 0, c1: 0, r2: 2, c2: 0 });
    expect(valueAt(1, 0)).toBe('x');
    expect(valueAt(2, 0)).toBe('x');
  });
});

describe('clipboard', () => {
  beforeEach(reset);

  it('copy/paste offsets relative formulas', () => {
    store().setCell({ row: 0, col: 0 }, '5');
    store().setCell({ row: 0, col: 1 }, '=A1*2');
    store().select({ row: 0, col: 1 });
    store().copy();
    store().setCell({ row: 1, col: 0 }, '7');
    store().select({ row: 1, col: 1 });
    store().paste();
    expect(rawAt(1, 1)).toBe('=A2*2');
    expect(valueAt(1, 1)).toBe(14);
  });

  it('cut clears the source', () => {
    store().setCell({ row: 0, col: 0 }, 'move me');
    store().select({ row: 0, col: 0 });
    store().cut();
    store().select({ row: 3, col: 0 });
    store().paste();
    expect(valueAt(0, 0)).toBeNull();
    expect(valueAt(3, 0)).toBe('move me');
  });

  it('pastes TSV from the OS clipboard', () => {
    store().select({ row: 0, col: 0 });
    store().pasteText('a\tb\n1\t2');
    expect(valueAt(0, 0)).toBe('a');
    expect(valueAt(0, 1)).toBe('b');
    expect(valueAt(1, 0)).toBe(1);
    expect(valueAt(1, 1)).toBe(2);
  });
});

describe('styles', () => {
  beforeEach(reset);

  it('applies and toggles over the selection', () => {
    store().selectRange({ r1: 0, c1: 0, r2: 1, c2: 1 });
    store().toggleStyle('bold');
    expect(store().styleAt({ row: 1, col: 1 })?.bold).toBe(true);
    store().toggleStyle('bold');
    expect(store().styleAt({ row: 1, col: 1 })?.bold).toBe(false);
  });

  it('clears formatting without touching values', () => {
    store().setCell({ row: 0, col: 0 }, 'keep');
    store().selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
    store().applyStyle({ bg: '#ff0000' });
    store().clearFormatting();
    expect(store().styleAt({ row: 0, col: 0 })).toBeUndefined();
    expect(valueAt(0, 0)).toBe('keep');
  });
});

describe('rows and columns', () => {
  beforeEach(reset);

  it('insert shifts values, formulas and styles together', () => {
    store().setCell({ row: 0, col: 0 }, '1');
    store().setCell({ row: 1, col: 0 }, '=A1+1');
    store().selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
    store().applyStyle({ bold: true });

    store().insertRows(0, 1);

    expect(valueAt(0, 0)).toBeNull();
    expect(valueAt(1, 0)).toBe(1);
    expect(rawAt(2, 0)).toBe('=A2+1');
    expect(store().styleAt({ row: 1, col: 0 })?.bold).toBe(true);
    expect(store().styleAt({ row: 0, col: 0 })).toBeUndefined();
  });

  it('delete drops the styles of removed rows', () => {
    store().selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
    store().applyStyle({ bold: true });
    store().setCell({ row: 1, col: 0 }, 'second');
    store().deleteRows(0, 1);
    expect(store().styleAt({ row: 0, col: 0 })).toBeUndefined();
    expect(valueAt(0, 0)).toBe('second');
  });

  it('insert shifts column widths', () => {
    store().setColWidth(0, 200);
    store().insertCols(0, 1);
    expect(store().colWidth(1)).toBe(200);
  });
});

describe('sheets', () => {
  beforeEach(reset);

  it('adds sheets with unique names and switches to them', () => {
    const before = store().sheets.length;
    store().addSheet();
    expect(store().sheets.length).toBe(before + 1);
    expect(store().activeSheetId).toBe(store().sheets.at(-1)!.id);
  });

  it('refuses a duplicate sheet name', () => {
    store().addSheet();
    const [first, second] = store().sheets;
    store().renameSheet(second.id, first.name);
    expect(store().sheets[1].name).toBe(second.name);
  });

  it('never deletes the last sheet', () => {
    while (store().sheets.length > 1) store().deleteSheet(store().sheets.at(-1)!.id);
    store().deleteSheet(store().sheets[0].id);
    expect(store().sheets.length).toBe(1);
  });

  it('supports cross-sheet formulas', () => {
    store().setCell({ row: 0, col: 0 }, '4');
    const firstName = store().sheets[0].name;
    store().addSheet();
    store().setCell({ row: 0, col: 0 }, `=${firstName}!A1*2`);
    expect(valueAt(0, 0)).toBe(8);
  });
});

describe('undo / redo', () => {
  beforeEach(reset);

  it('restores the previous value', () => {
    store().setCell({ row: 0, col: 0 }, 'first');
    store().setCell({ row: 0, col: 0 }, 'second');
    store().undo();
    expect(valueAt(0, 0)).toBe('first');
    store().redo();
    expect(valueAt(0, 0)).toBe('second');
  });
});
