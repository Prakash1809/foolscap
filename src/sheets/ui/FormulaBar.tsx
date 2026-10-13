import { useState } from 'react';
import { useSheets } from '../store';
import { functionNames, readRaw } from '../engine/engine';
import { normaliseRange, parseAddr, rangeLabel } from '../../shared/cellref';

const ALL_FUNCTIONS = functionNames();

/**
 * Both inputs here mirror state that lives elsewhere (the engine, the
 * selection) but must stay editable while the user types. Rather than syncing
 * with effects, each buffer records the identity of what it was opened against
 * and is derived away as soon as that identity changes.
 */
export function FormulaBar() {
  const focus = useSheets((s) => s.focus);
  const anchor = useSheets((s) => s.anchor);
  const activeSheetId = useSheets((s) => s.activeSheetId);
  const editing = useSheets((s) => s.editing);
  const rev = useSheets((s) => s.rev);

  const [nameDraft, setNameDraft] = useState<string | null>(null);
  const [buffer, setBuffer] = useState<{ key: string; text: string } | null>(null);

  const selectionLabel = rangeLabel(normaliseRange(anchor, focus));
  const nameBox = nameDraft ?? selectionLabel;

  // Identity of the cell revision this buffer belongs to.
  const cellKey = `${activeSheetId}:${focus.row}:${focus.col}:${rev}`;
  const committedRaw = readRaw(activeSheetId, focus);
  const value = editing ? editing.draft : buffer?.key === cellKey ? buffer.text : committedRaw;

  const setValue = (text: string) => {
    if (editing) useSheets.getState().updateDraft(text);
    else setBuffer({ key: cellKey, text });
  };

  const suggestions =
    value.startsWith('=') && !value.includes('(')
      ? ALL_FUNCTIONS.filter((fn) => fn.startsWith(value.slice(1).toUpperCase())).slice(0, 8)
      : [];

  return (
    <div className="formula-bar">
      <input
        className="name-box"
        value={nameBox}
        onChange={(e) => setNameDraft(e.target.value)}
        onBlur={() => setNameDraft(null)}
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            setNameDraft(null);
            e.currentTarget.blur();
            return;
          }
          if (e.key !== 'Enter') return;
          const addr = parseAddr(nameBox);
          setNameDraft(null);
          if (addr) useSheets.getState().select(addr);
        }}
        title="Name box — type a cell reference and press Enter"
      />

      <div className="fx-label" title="Insert function">
        fx
      </div>

      <div className="formula-input-wrap">
        <input
          className="formula-input"
          value={value}
          spellCheck={false}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              const store = useSheets.getState();
              if (editing) {
                store.commitEdit({ dRow: 1, dCol: 0 });
              } else {
                store.setCell(focus, e.currentTarget.value);
                store.moveSelection(1, 0, false);
              }
              setBuffer(null);
            } else if (e.key === 'Escape') {
              e.preventDefault();
              if (editing) useSheets.getState().cancelEdit();
              setBuffer(null);
              e.currentTarget.blur();
            }
          }}
        />
        {suggestions.length > 0 && (
          <ul className="fn-suggestions">
            {suggestions.map((fn) => (
              <li
                key={fn}
                onMouseDown={(e) => {
                  e.preventDefault();
                  setValue(`=${fn}(`);
                }}
              >
                {fn}
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}
