import { useRef, useState } from 'react';
import { useSheets } from '../store';
import { NUM_FMT_LABELS } from '../engine/format';
import { styleKey } from '../../shared/cellref';
import type { HorizAlign, NumFmt } from '../../shared/types';

const TEXT_COLORS = ['#111827', '#b91c1c', '#15803d', '#1d4ed8', '#b45309', '#7e22ce', '#0f766e', '#ffffff'];
const FILL_COLORS = ['transparent', '#fee2e2', '#dcfce7', '#dbeafe', '#fef9c3', '#f3e8ff', '#ccfbf1', '#e5e7eb'];
const FONT_SIZES = [10, 11, 12, 13, 14, 16, 18, 20, 24, 32];

export function Ribbon() {
  const fileInput = useRef<HTMLInputElement>(null);
  const [busy, setBusy] = useState<string | null>(null);

  const fileName = useSheets((s) => s.fileName);
  const anchor = useSheets((s) => s.anchor);
  const focus = useSheets((s) => s.focus);
  const sheet = useSheets((s) => s.activeSheet());

  const activeStyle = sheet.styles[styleKey(Math.min(anchor.row, focus.row), Math.min(anchor.col, focus.col))];
  const store = () => useSheets.getState();

  // ExcelJS is ~1.4 MB; keep it out of the initial bundle.
  const io = () => import('../io/xlsx');

  const openFile = async (file: File) => {
    setBusy('Opening…');
    try {
      const { loadXlsx } = await io();
      const result = await loadXlsx(file);
      store().replaceWorkbook(result.fileName, result.sheets);
    } catch (err) {
      alert(`Could not open ${file.name}:\n${err instanceof Error ? err.message : String(err)}`);
    } finally {
      setBusy(null);
    }
  };

  const save = async () => {
    setBusy('Saving…');
    try {
      const { saveXlsx } = await io();
      await saveXlsx(fileName, store().sheets);
      useSheets.setState({ dirty: false });
    } catch (err) {
      alert(`Export failed:\n${err instanceof Error ? err.message : String(err)}`);
    } finally {
      setBusy(null);
    }
  };

  const csv = async () => {
    const { exportCsv } = await io();
    exportCsv(fileName, sheet);
  };

  const insertFormula = (fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN') => {
    const s = store();
    const sel = s.selection();
    // Place the result just past the selection, summing the selected block.
    const target =
      sel.r1 === sel.r2 && sel.c1 === sel.c2
        ? { row: sel.r1, col: sel.c1 }
        : { row: sel.r2 + 1, col: sel.c1 };
    const refStart = `${colName(sel.c1)}${sel.r1 + 1}`;
    const refEnd = `${colName(sel.c2)}${sel.r2 + 1}`;
    s.beginEdit(target, `=${fn}(${refStart}:${refEnd})`, false);
  };

  return (
    <div className="ribbon">
      <Group label="File">
        <button onClick={() => fileInput.current?.click()} title="Open an .xlsx file">
          📂 Open
        </button>
        <button onClick={save} title="Download as .xlsx">
          💾 Save
        </button>
        <button onClick={csv} title="Export active sheet as CSV">
          ⤓ CSV
        </button>
        <input
          ref={fileInput}
          type="file"
          accept=".xlsx,.xls"
          hidden
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) void openFile(f);
            e.target.value = '';
          }}
        />
      </Group>

      <Group label="Undo">
        <button onClick={() => store().undo()} title="Undo (Ctrl+Z)">
          ↶
        </button>
        <button onClick={() => store().redo()} title="Redo (Ctrl+Y)">
          ↷
        </button>
      </Group>

      <Group label="Clipboard">
        <button onClick={() => store().cut()} title="Cut (Ctrl+X)">
          ✂
        </button>
        <button onClick={() => store().copy()} title="Copy (Ctrl+C)">
          ⧉
        </button>
        <button onClick={() => store().paste()} title="Paste (Ctrl+V)">
          📋
        </button>
      </Group>

      <Group label="Font">
        <select
          className="font-size"
          value={activeStyle?.fontSize ?? 13}
          onChange={(e) => store().applyStyle({ fontSize: Number(e.target.value) })}
          title="Font size"
        >
          {FONT_SIZES.map((n) => (
            <option key={n} value={n}>
              {n}
            </option>
          ))}
        </select>
        <button
          className={activeStyle?.bold ? 'active' : ''}
          onClick={() => store().toggleStyle('bold')}
          title="Bold (Ctrl+B)"
        >
          <b>B</b>
        </button>
        <button
          className={activeStyle?.italic ? 'active' : ''}
          onClick={() => store().toggleStyle('italic')}
          title="Italic (Ctrl+I)"
        >
          <i>I</i>
        </button>
        <button
          className={activeStyle?.underline ? 'active' : ''}
          onClick={() => store().toggleStyle('underline')}
          title="Underline (Ctrl+U)"
        >
          <u>U</u>
        </button>
        <button
          className={activeStyle?.strike ? 'active' : ''}
          onClick={() => store().toggleStyle('strike')}
          title="Strikethrough"
        >
          <s>S</s>
        </button>
        <Swatches label="A" title="Text colour" colors={TEXT_COLORS} onPick={(c) => store().applyStyle({ color: c })} />
        <Swatches label="▢" title="Fill colour" colors={FILL_COLORS} onPick={(c) => store().applyStyle({ bg: c === 'transparent' ? undefined : c })} />
      </Group>

      <Group label="Alignment">
        {(['left', 'center', 'right'] as HorizAlign[]).map((a) => (
          <button
            key={a}
            className={activeStyle?.align === a ? 'active' : ''}
            onClick={() => store().applyStyle({ align: a })}
            title={`Align ${a}`}
          >
            {a === 'left' ? '⬅' : a === 'center' ? '↔' : '➡'}
          </button>
        ))}
        <button
          className={activeStyle?.wrap ? 'active' : ''}
          onClick={() => store().toggleStyle('wrap')}
          title="Wrap text"
        >
          ⏎
        </button>
      </Group>

      <Group label="Number">
        <select
          value={activeStyle?.numFmt ?? 'general'}
          onChange={(e) => store().applyStyle({ numFmt: e.target.value as NumFmt })}
          title="Number format"
        >
          {(Object.keys(NUM_FMT_LABELS) as NumFmt[]).map((k) => (
            <option key={k} value={k}>
              {NUM_FMT_LABELS[k]}
            </option>
          ))}
        </select>
      </Group>

      <Group label="Cells">
        <button onClick={() => store().insertRows(store().selection().r1, rowsInSelection())} title="Insert rows">
          +Row
        </button>
        <button onClick={() => store().deleteRows(store().selection().r1, rowsInSelection())} title="Delete rows">
          −Row
        </button>
        <button onClick={() => store().insertCols(store().selection().c1, colsInSelection())} title="Insert columns">
          +Col
        </button>
        <button onClick={() => store().deleteCols(store().selection().c1, colsInSelection())} title="Delete columns">
          −Col
        </button>
        <button onClick={() => store().toggleFreeze()} title="Freeze panes at selection">
          ❄ Freeze
        </button>
      </Group>

      <Group label="Formulas">
        {(['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as const).map((fn) => (
          <button key={fn} onClick={() => insertFormula(fn)} title={`Insert ${fn} over the selection`}>
            {fn === 'SUM' ? 'Σ' : fn[0] + fn.slice(1).toLowerCase()}
          </button>
        ))}
      </Group>

      <Group label="Editing">
        <button onClick={() => store().fillDown()} title="Fill down (Ctrl+D)">
          ↓ Fill
        </button>
        <button onClick={() => store().clearSelection()} title="Clear contents (Delete)">
          Clear
        </button>
        <button onClick={() => store().clearFormatting()} title="Clear formatting">
          Clear fmt
        </button>
      </Group>

      {busy && <span className="ribbon-busy">{busy}</span>}
    </div>
  );
}

function rowsInSelection(): number {
  const sel = useSheets.getState().selection();
  return Math.min(sel.r2 - sel.r1 + 1, 100);
}

function colsInSelection(): number {
  const sel = useSheets.getState().selection();
  return Math.min(sel.c2 - sel.c1 + 1, 50);
}

function colName(col: number): string {
  let s = '';
  let n = col;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function Group({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="ribbon-group">
      <div className="ribbon-buttons">{children}</div>
      <div className="ribbon-group-label">{label}</div>
    </div>
  );
}

function Swatches({
  label,
  title,
  colors,
  onPick,
}: {
  label: string;
  title: string;
  colors: string[];
  onPick: (color: string) => void;
}) {
  const [open, setOpen] = useState(false);
  return (
    <div className="swatch-wrap">
      <button onClick={() => setOpen((o) => !o)} title={title}>
        {label} ▾
      </button>
      {open && (
        <div className="swatch-grid" onMouseLeave={() => setOpen(false)}>
          {colors.map((c) => (
            <button
              key={c}
              className="swatch"
              style={{ background: c === 'transparent' ? 'repeating-conic-gradient(#ddd 0% 25%, #fff 0% 50%) 50% / 8px 8px' : c }}
              onClick={() => {
                onPick(c);
                setOpen(false);
              }}
              title={c}
            />
          ))}
        </div>
      )}
    </div>
  );
}
