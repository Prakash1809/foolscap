import { useEffect } from 'react';
import { Grid } from './grid/Grid';
import { FormulaBar } from './ui/FormulaBar';
import { Ribbon } from './ui/Ribbon';
import { SheetTabs } from './ui/SheetTabs';
import { StatusBar } from './ui/StatusBar';
import { useSheets } from './store';

export function SheetsApp() {
  const dirty = useSheets((s) => s.dirty);

  // Guard against losing unsaved work on reload/close.
  useEffect(() => {
    if (!dirty) return;
    const warn = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [dirty]);

  return (
    <div className="sheets-app">
      <Ribbon />
      <FormulaBar />
      <Grid />
      <SheetTabs />
      <StatusBar />
    </div>
  );
}
