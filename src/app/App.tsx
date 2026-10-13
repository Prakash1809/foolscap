import { useState } from 'react';
import { SheetsApp } from '../sheets/SheetsApp';
import { useSheets } from '../sheets/store';

type Tool = 'sheets' | 'docs';

export default function App() {
  const [tool, setTool] = useState<Tool>('sheets');
  const fileName = useSheets((s) => s.fileName);
  const dirty = useSheets((s) => s.dirty);

  return (
    <div className="shell">
      <header className="title-bar">
        <span className="brand">Foolscap</span>
        <nav className="tool-switch">
          <button className={tool === 'sheets' ? 'active' : ''} onClick={() => setTool('sheets')}>
            Sheets
          </button>
          <button className={tool === 'docs' ? 'active' : ''} onClick={() => setTool('docs')}>
            Docs
          </button>
        </nav>
        <input
          className="file-name"
          value={fileName}
          onChange={(e) => useSheets.setState({ fileName: e.target.value })}
          aria-label="File name"
        />
        {dirty && <span className="dirty-dot" title="Unsaved changes" />}
      </header>

      <main className="shell-body">
        {tool === 'sheets' ? <SheetsApp /> : <DocsPlaceholder />}
      </main>
    </div>
  );
}

function DocsPlaceholder() {
  return (
    <div className="placeholder">
      <h2>Docs</h2>
      <p>Word processor — not built yet.</p>
      <p className="muted">
        Planned: rich-text model, headings and styles, lists, tables, page layout with print view,
        and .docx import/export via <code>docx</code> + <code>mammoth</code>.
      </p>
    </div>
  );
}
