import { useState } from 'react';
import { useSheets } from '../store';

export function SheetTabs() {
  const sheets = useSheets((s) => s.sheets);
  const activeSheetId = useSheets((s) => s.activeSheetId);
  const [renaming, setRenaming] = useState<number | null>(null);
  const [draft, setDraft] = useState('');

  return (
    <div className="sheet-tabs">
      <button className="tab-add" onClick={() => useSheets.getState().addSheet()} title="New sheet">
        ＋
      </button>
      {sheets.map((sheet) => (
        <div
          key={sheet.id}
          className={`sheet-tab${sheet.id === activeSheetId ? ' active' : ''}`}
          onMouseDown={() => useSheets.getState().setActiveSheet(sheet.id)}
          onDoubleClick={() => {
            setRenaming(sheet.id);
            setDraft(sheet.name);
          }}
        >
          {renaming === sheet.id ? (
            <input
              autoFocus
              className="tab-rename"
              value={draft}
              onChange={(e) => setDraft(e.target.value)}
              onBlur={() => {
                useSheets.getState().renameSheet(sheet.id, draft);
                setRenaming(null);
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') e.currentTarget.blur();
                if (e.key === 'Escape') setRenaming(null);
              }}
            />
          ) : (
            <>
              <span>{sheet.name}</span>
              {sheets.length > 1 && (
                <button
                  className="tab-close"
                  title={`Delete ${sheet.name}`}
                  onMouseDown={(e) => {
                    e.stopPropagation();
                    if (confirm(`Delete sheet "${sheet.name}"? This cannot be undone.`)) {
                      useSheets.getState().deleteSheet(sheet.id);
                    }
                  }}
                >
                  ×
                </button>
              )}
            </>
          )}
        </div>
      ))}
    </div>
  );
}
