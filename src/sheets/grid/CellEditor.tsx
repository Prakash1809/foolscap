import { useEffect, useRef } from 'react';
import { useSheets } from '../store';

interface Props {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * In-cell editor. Key handling for Enter/Tab/Escape lives in Grid's onKeyDown —
 * the event bubbles up from this input, so there is one code path for commit.
 */
export function CellEditor({ left, top, width, height }: Props) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const draft = useSheets((s) => s.editing?.draft ?? '');
  const selectAll = useSheets((s) => s.editing?.selectAll ?? false);
  const updateDraft = useSheets((s) => s.updateDraft);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    el.focus();
    if (selectAll) el.select();
    else el.setSelectionRange(el.value.length, el.value.length);
  }, [selectAll]);

  return (
    <textarea
      ref={ref}
      className="cell-editor"
      value={draft}
      spellCheck={false}
      style={{
        left,
        top,
        width: Math.max(width, 120),
        height,
        minHeight: height,
      }}
      onChange={(e) => updateDraft(e.target.value)}
      onBlur={() => useSheets.getState().commitEdit()}
    />
  );
}
