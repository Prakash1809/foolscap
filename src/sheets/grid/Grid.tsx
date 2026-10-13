import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
} from 'react';
import { useSheets } from '../store';
import { readCell } from '../engine/engine';
import { defaultAlign, formatValue } from '../engine/format';
import { colLabel, normaliseRange, styleKey } from '../../shared/cellref';
import { MAX_COLS, MAX_ROWS, type CellAddr, type CellRange, type CellStyle } from '../../shared/types';
import { buildColAxis, buildRowAxis, indexAt, sizeOf, visibleRange, type Axis } from './geometry';
import { CellEditor } from './CellEditor';

type ResizeDrag =
  | { kind: 'col'; index: number; startPos: number; startSize: number }
  | { kind: 'row'; index: number; startPos: number; startSize: number };

const SPILL_LIMIT = 12;

const JUSTIFY: Record<string, CSSProperties['justifyContent']> = {
  left: 'flex-start',
  center: 'center',
  right: 'flex-end',
};

function styleToCss(style: CellStyle | undefined, value: unknown): CSSProperties {
  const align = style?.align ?? defaultAlign(value);
  const css: CSSProperties = {
    textAlign: align,
    justifyContent: JUSTIFY[align],
  };
  if (style?.bold) css.fontWeight = 700;
  if (style?.italic) css.fontStyle = 'italic';
  if (style?.underline || style?.strike) {
    css.textDecoration = [style.underline && 'underline', style.strike && 'line-through']
      .filter(Boolean)
      .join(' ');
  }
  if (style?.color) css.color = style.color;
  if (style?.bg) css.background = style.bg;
  if (style?.fontSize) css.fontSize = `${style.fontSize}px`;
  if (style?.fontFamily) css.fontFamily = style.fontFamily;
  if (style?.wrap) {
    css.whiteSpace = 'pre-wrap';
    css.overflowWrap = 'anywhere';
    css.alignItems = 'flex-start';
  }
  if (style?.valign === 'top') css.alignItems = 'flex-start';
  if (style?.valign === 'bottom') css.alignItems = 'flex-end';
  if (style?.border) {
    const b = style.border;
    if (b.top) css.borderTop = '1px solid #4b5563';
    if (b.right) css.borderRight = '1px solid #4b5563';
    if (b.bottom) css.borderBottom = '1px solid #4b5563';
    if (b.left) css.borderLeft = '1px solid #4b5563';
  }
  return css;
}

export function Grid() {
  const scrollRef = useRef<HTMLDivElement>(null);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [viewport, setViewport] = useState({ w: 800, h: 600 });
  const [dragSelecting, setDragSelecting] = useState(false);
  const [fillTarget, setFillTarget] = useState<CellRange | null>(null);
  const fillDragRef = useRef(false);
  const resizeRef = useRef<ResizeDrag | null>(null);

  const sheet = useSheets((s) => s.activeSheet());
  const activeSheetId = useSheets((s) => s.activeSheetId);
  const anchor = useSheets((s) => s.anchor);
  const focus = useSheets((s) => s.focus);
  const editing = useSheets((s) => s.editing);
  const rev = useSheets((s) => s.rev);
  const clipboard = useSheets((s) => s.clipboard);

  const colAxis = useMemo(() => buildColAxis(sheet.colWidths), [sheet.colWidths]);
  const rowAxis = useMemo(() => buildRowAxis(sheet.rowHeights), [sheet.rowHeights]);
  const selection = useMemo(() => normaliseRange(anchor, focus), [anchor, focus]);

  const cols = useMemo(
    () => visibleRange(colAxis, scroll.left, viewport.w),
    [colAxis, scroll.left, viewport.w],
  );
  const rows = useMemo(
    () => visibleRange(rowAxis, scroll.top, viewport.h),
    [rowAxis, scroll.top, viewport.h],
  );

  // --- viewport measurement -------------------------------------------------
  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const measure = () => setViewport({ w: el.clientWidth, h: el.clientHeight });
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // --- keep the focused cell on screen --------------------------------------
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const x = colAxis.offsets[focus.col];
    const w = sizeOf(colAxis, focus.col);
    const y = rowAxis.offsets[focus.row];
    const h = sizeOf(rowAxis, focus.row);
    let left = el.scrollLeft;
    let top = el.scrollTop;
    if (x < left) left = x;
    else if (x + w > left + el.clientWidth) left = x + w - el.clientWidth;
    if (y < top) top = y;
    else if (y + h > top + el.clientHeight) top = y + h - el.clientHeight;
    if (left !== el.scrollLeft || top !== el.scrollTop) el.scrollTo({ left, top });
  }, [focus, colAxis, rowAxis]);

  // --- pointer helpers ------------------------------------------------------
  const addrFromEvent = useCallback(
    (clientX: number, clientY: number): CellAddr => {
      const el = scrollRef.current!;
      const box = el.getBoundingClientRect();
      const x = clientX - box.left + el.scrollLeft;
      const y = clientY - box.top + el.scrollTop;
      return { row: indexAt(rowAxis, y), col: indexAt(colAxis, x) };
    },
    [colAxis, rowAxis],
  );

  const onCellMouseDown = (e: ReactMouseEvent, addr: CellAddr) => {
    if (e.button !== 0) return;
    const store = useSheets.getState();
    if (store.editing) store.commitEdit();
    store.select(addr, e.shiftKey);
    setDragSelecting(true);
    scrollRef.current?.focus();
  };

  useEffect(() => {
    if (!dragSelecting) return;
    const onMove = (e: MouseEvent) => {
      useSheets.getState().select(addrFromEvent(e.clientX, e.clientY), true);
    };
    const onUp = () => setDragSelecting(false);
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [dragSelecting, addrFromEvent]);

  // --- fill handle ----------------------------------------------------------
  const startFillDrag = (e: ReactMouseEvent) => {
    e.stopPropagation();
    e.preventDefault();
    fillDragRef.current = true;
    setFillTarget(selection);
  };

  useEffect(() => {
    if (!fillDragRef.current) return;
    const onMove = (e: MouseEvent) => {
      const at = addrFromEvent(e.clientX, e.clientY);
      const dRow = at.row > selection.r2 ? at.row - selection.r2 : at.row < selection.r1 ? at.row - selection.r1 : 0;
      const dCol = at.col > selection.c2 ? at.col - selection.c2 : at.col < selection.c1 ? at.col - selection.c1 : 0;
      // Excel locks the fill to the dominant axis.
      const target: CellRange =
        Math.abs(dRow) >= Math.abs(dCol)
          ? { ...selection, r1: Math.min(selection.r1, at.row), r2: Math.max(selection.r2, at.row) }
          : { ...selection, c1: Math.min(selection.c1, at.col), c2: Math.max(selection.c2, at.col) };
      setFillTarget(target);
    };
    const onUp = () => {
      fillDragRef.current = false;
      setFillTarget((t) => {
        if (t && (t.r1 !== selection.r1 || t.r2 !== selection.r2 || t.c1 !== selection.c1 || t.c2 !== selection.c2)) {
          useSheets.getState().fillFromHandle(t);
        }
        return null;
      });
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [fillTarget, selection, addrFromEvent]);

  // --- header resize --------------------------------------------------------
  const startResize = (e: ReactMouseEvent, drag: ResizeDrag) => {
    e.stopPropagation();
    e.preventDefault();
    resizeRef.current = drag;
  };

  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      const drag = resizeRef.current;
      if (!drag) return;
      const store = useSheets.getState();
      if (drag.kind === 'col') {
        store.setColWidth(drag.index, drag.startSize + (e.clientX - drag.startPos));
      } else {
        store.setRowHeight(drag.index, drag.startSize + (e.clientY - drag.startPos));
      }
    };
    const onUp = () => {
      resizeRef.current = null;
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, []);

  /** Double-clicking a column edge sizes it to the widest visible value. */
  const autofitCol = (col: number) => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.font = '13px system-ui, sans-serif';
    let widest = 40;
    for (let row = 0; row < Math.min(MAX_ROWS, 500); row++) {
      const snap = readCell(activeSheetId, { row, col });
      if (snap.value === null) continue;
      const text = formatValue(snap.value, sheet.styles[styleKey(row, col)]);
      widest = Math.max(widest, ctx.measureText(text).width + 18);
    }
    useSheets.getState().setColWidth(col, Math.min(widest, 480));
  };

  // --- keyboard -------------------------------------------------------------
  const onKeyDown = (e: ReactKeyboardEvent) => {
    const store = useSheets.getState();
    const ctrl = e.ctrlKey || e.metaKey;

    if (store.editing) {
      // The editor owns most keys; only navigation commits and moves.
      if (e.key === 'Enter' && !e.altKey) {
        e.preventDefault();
        store.commitEdit({ dRow: e.shiftKey ? -1 : 1, dCol: 0 });
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        store.commitEdit({ dRow: 0, dCol: e.shiftKey ? -1 : 1 });
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        store.cancelEdit();
        return;
      }
      return;
    }

    if (ctrl) {
      switch (e.key.toLowerCase()) {
        case 'c':
          e.preventDefault();
          store.copy();
          return;
        case 'x':
          e.preventDefault();
          store.cut();
          return;
        case 'v':
          // Native paste event carries the OS clipboard; only fall back internally.
          if (store.clipboard) {
            e.preventDefault();
            store.paste();
          }
          return;
        case 'z':
          e.preventDefault();
          if (e.shiftKey) store.redo();
          else store.undo();
          return;
        case 'y':
          e.preventDefault();
          store.redo();
          return;
        case 'b':
          e.preventDefault();
          store.toggleStyle('bold');
          return;
        case 'i':
          e.preventDefault();
          store.toggleStyle('italic');
          return;
        case 'u':
          e.preventDefault();
          store.toggleStyle('underline');
          return;
        case 'a':
          e.preventDefault();
          store.selectAll();
          return;
        case 'd':
          e.preventDefault();
          store.fillDown();
          return;
        case 'home':
          e.preventDefault();
          store.select({ row: 0, col: 0 });
          return;
      }
    }

    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        store.moveSelection(-1, 0, e.shiftKey);
        return;
      case 'ArrowDown':
        e.preventDefault();
        store.moveSelection(1, 0, e.shiftKey);
        return;
      case 'ArrowLeft':
        e.preventDefault();
        store.moveSelection(0, -1, e.shiftKey);
        return;
      case 'ArrowRight':
        e.preventDefault();
        store.moveSelection(0, 1, e.shiftKey);
        return;
      case 'Tab':
        e.preventDefault();
        store.moveSelection(0, e.shiftKey ? -1 : 1, false);
        return;
      case 'Enter':
        e.preventDefault();
        store.beginEdit(focus, undefined, true);
        return;
      case 'F2':
        e.preventDefault();
        store.beginEdit(focus, undefined, true);
        return;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        store.clearSelection();
        return;
      case 'Home':
        e.preventDefault();
        store.select({ row: focus.row, col: 0 }, e.shiftKey);
        return;
      case 'PageDown':
        e.preventDefault();
        store.moveSelection(Math.floor(viewport.h / 24), 0, e.shiftKey);
        return;
      case 'PageUp':
        e.preventDefault();
        store.moveSelection(-Math.floor(viewport.h / 24), 0, e.shiftKey);
        return;
      case 'Escape':
        e.preventDefault();
        useSheets.setState({ clipboard: null });
        return;
    }

    // Typing replaces the cell, as in Excel.
    if (!ctrl && !e.altKey && e.key.length === 1) {
      e.preventDefault();
      store.beginEdit(focus, e.key, false);
    }
  };

  // --- render ---------------------------------------------------------------
  const cellNodes: React.ReactNode[] = [];
  for (let row = rows.start; row <= rows.end; row++) {
    for (let col = cols.start; col <= cols.end; col++) {
      const snap = readCell(activeSheetId, { row, col });
      const style = sheet.styles[styleKey(row, col)];
      if (snap.value === null && !style) continue;

      const text = snap.isError ? String(snap.value) : formatValue(snap.value, style);
      let width = sizeOf(colAxis, col);

      // Text spill: borrow width from empty neighbours when not wrapping.
      const spills = typeof snap.value === 'string' && !style?.wrap && !snap.isError;
      if (spills && (style?.align ?? 'left') === 'left') {
        for (let c = col + 1; c <= Math.min(MAX_COLS - 1, col + SPILL_LIMIT); c++) {
          if (readCell(activeSheetId, { row, col: c }).value !== null) break;
          width += sizeOf(colAxis, c);
        }
      }

      cellNodes.push(
        <div
          key={`${row}:${col}`}
          className={`cell${snap.isError ? ' cell-error' : ''}`}
          title={snap.isError ? snap.errorMessage : undefined}
          style={{
            left: colAxis.offsets[col],
            top: rowAxis.offsets[row],
            width,
            height: sizeOf(rowAxis, row),
            ...styleToCss(style, snap.value),
          }}
        >
          {text}
        </div>,
      );
    }
  }

  const selBox = {
    left: colAxis.offsets[selection.c1],
    top: rowAxis.offsets[selection.r1],
    width: colAxis.offsets[Math.min(selection.c2 + 1, MAX_COLS)] - colAxis.offsets[selection.c1],
    height: rowAxis.offsets[Math.min(selection.r2 + 1, MAX_ROWS)] - rowAxis.offsets[selection.r1],
  };
  const focusBox = {
    left: colAxis.offsets[focus.col],
    top: rowAxis.offsets[focus.row],
    width: sizeOf(colAxis, focus.col),
    height: sizeOf(rowAxis, focus.row),
  };

  return (
    <div className="grid-root">
      <div className="grid-corner" onMouseDown={() => useSheets.getState().selectAll()} />

      <div className="grid-col-headers">
        <div className="headers-track" style={{ transform: `translateX(${-scroll.left}px)` }}>
          {range(cols.start, cols.end).map((col) => (
            <div
              key={col}
              className={`col-header${col >= selection.c1 && col <= selection.c2 ? ' header-active' : ''}`}
              style={{ left: colAxis.offsets[col], width: sizeOf(colAxis, col) }}
              onMouseDown={(e) => useSheets.getState().selectCol(col, e.shiftKey)}
            >
              {colLabel(col)}
              <span
                className="col-resizer"
                onMouseDown={(e) =>
                  startResize(e, {
                    kind: 'col',
                    index: col,
                    startPos: e.clientX,
                    startSize: sizeOf(colAxis, col),
                  })
                }
                onDoubleClick={(e) => {
                  e.stopPropagation();
                  autofitCol(col);
                }}
              />
            </div>
          ))}
        </div>
      </div>

      <div className="grid-row-headers">
        <div className="headers-track" style={{ transform: `translateY(${-scroll.top}px)` }}>
          {range(rows.start, rows.end).map((row) => (
            <div
              key={row}
              className={`row-header${row >= selection.r1 && row <= selection.r2 ? ' header-active' : ''}`}
              style={{ top: rowAxis.offsets[row], height: sizeOf(rowAxis, row) }}
              onMouseDown={(e) => useSheets.getState().selectRow(row, e.shiftKey)}
            >
              {row + 1}
              <span
                className="row-resizer"
                onMouseDown={(e) =>
                  startResize(e, {
                    kind: 'row',
                    index: row,
                    startPos: e.clientY,
                    startSize: sizeOf(rowAxis, row),
                  })
                }
              />
            </div>
          ))}
        </div>
      </div>

      <div
        ref={scrollRef}
        className="grid-scroll"
        tabIndex={0}
        onKeyDown={onKeyDown}
        onScroll={(e) =>
          setScroll({ top: e.currentTarget.scrollTop, left: e.currentTarget.scrollLeft })
        }
        onPaste={(e) => {
          const text = e.clipboardData.getData('text/plain');
          if (text) {
            e.preventDefault();
            useSheets.getState().pasteText(text);
          }
        }}
        onCopy={(e) => {
          const store = useSheets.getState();
          store.copy();
          const sel = store.selection();
          const tsv = range(sel.r1, sel.r2)
            .map((row) =>
              range(sel.c1, sel.c2)
                .map((col) => {
                  const snap = readCell(store.activeSheetId, { row, col });
                  return formatValue(snap.value, store.activeSheet().styles[styleKey(row, col)]);
                })
                .join('\t'),
            )
            .join('\n');
          e.clipboardData.setData('text/plain', tsv);
          e.preventDefault();
        }}
      >
        <div
          className="grid-canvas"
          style={{ width: colAxis.total, height: rowAxis.total }}
          onMouseDown={(e) => {
            if (e.target === e.currentTarget || (e.target as HTMLElement).classList.contains('cell')) {
              onCellMouseDown(e, addrFromEvent(e.clientX, e.clientY));
            }
          }}
          onDoubleClick={(e) =>
            useSheets.getState().beginEdit(addrFromEvent(e.clientX, e.clientY), undefined, true)
          }
        >
          <GridLines colAxis={colAxis} rowAxis={rowAxis} cols={cols} rows={rows} />

          {/* The range tint paints *under* the cells so text stays crisp. */}
          <div className="selection-range" style={selBox} />

          {clipboard && clipboard.sheetId === activeSheetId && (
            <div
              className="marching-ants"
              style={{
                left: colAxis.offsets[clipboard.origin.c1],
                top: rowAxis.offsets[clipboard.origin.r1],
                width:
                  colAxis.offsets[Math.min(clipboard.origin.c2 + 1, MAX_COLS)] -
                  colAxis.offsets[clipboard.origin.c1],
                height:
                  rowAxis.offsets[Math.min(clipboard.origin.r2 + 1, MAX_ROWS)] -
                  rowAxis.offsets[clipboard.origin.r1],
              }}
            />
          )}

          {fillTarget && (
            <div
              className="fill-preview"
              style={{
                left: colAxis.offsets[fillTarget.c1],
                top: rowAxis.offsets[fillTarget.r1],
                width:
                  colAxis.offsets[Math.min(fillTarget.c2 + 1, MAX_COLS)] - colAxis.offsets[fillTarget.c1],
                height:
                  rowAxis.offsets[Math.min(fillTarget.r2 + 1, MAX_ROWS)] - rowAxis.offsets[fillTarget.r1],
              }}
            />
          )}

          {cellNodes}

          <div className="selection-focus" style={focusBox} />

          <div
            className="fill-handle"
            style={{ left: selBox.left + selBox.width - 3, top: selBox.top + selBox.height - 3 }}
            onMouseDown={startFillDrag}
          />

          {editing && (
            <CellEditor
              key={`${editing.addr.row}:${editing.addr.col}:${rev}`}
              left={colAxis.offsets[editing.addr.col]}
              top={rowAxis.offsets[editing.addr.row]}
              width={sizeOf(colAxis, editing.addr.col)}
              height={sizeOf(rowAxis, editing.addr.row)}
            />
          )}
        </div>
      </div>
    </div>
  );
}

function range(start: number, end: number): number[] {
  const out: number[] = [];
  for (let i = start; i <= end; i++) out.push(i);
  return out;
}

/** Background gridlines drawn as two repeating-gradient layers per visible band. */
function GridLines({
  colAxis,
  rowAxis,
  cols,
  rows,
}: {
  colAxis: Axis;
  rowAxis: Axis;
  cols: { start: number; end: number };
  rows: { start: number; end: number };
}) {
  const vertical = range(cols.start, cols.end).map((col) => (
    <div
      key={`v${col}`}
      className="gridline-v"
      style={{ left: colAxis.offsets[col + 1] - 1, top: rowAxis.offsets[rows.start], height: rowAxis.offsets[rows.end + 1] - rowAxis.offsets[rows.start] }}
    />
  ));
  const horizontal = range(rows.start, rows.end).map((row) => (
    <div
      key={`h${row}`}
      className="gridline-h"
      style={{ top: rowAxis.offsets[row + 1] - 1, left: colAxis.offsets[cols.start], width: colAxis.offsets[cols.end + 1] - colAxis.offsets[cols.start] }}
    />
  ));
  return (
    <>
      {vertical}
      {horizontal}
    </>
  );
}
