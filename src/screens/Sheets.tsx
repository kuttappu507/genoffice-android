import { useMemo, useRef, useState } from 'react';
import { evalCell } from '../lib/formulas';
import { chatStream, errMsg } from '../lib/ai-client';
import { debounce, downloadText, getDoc, getSettings, putDoc, uid } from '../lib/storage';

const COLS = 26;
const ROWS = 60;

interface SheetData {
  cells: Record<string, string>;
}

function colName(i: number): string {
  return String.fromCharCode(65 + i);
}

export default function Sheets({ initialId }: { initialId?: string }) {
  const sheetId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Spreadsheet' : 'Untitled sheet');
  const [cells, setCells] = useState<Record<string, string>>(() => {
    if (initialId) return getDoc<SheetData>(initialId)?.cells ?? {};
    return {};
  });
  const [sel, setSel] = useState('A1');
  const [editing, setEditing] = useState<string | null>(null);
  const [aiOpen, setAiOpen] = useState(false);
  const [aiPrompt, setAiPrompt] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [toast, setToast] = useState('');

  const save = useMemo(
    () =>
      debounce((c: Record<string, string>, t: string) => {
        putDoc<SheetData>('sheet', sheetId, t, { cells: c });
      }, 700),
    [sheetId],
  );

  const display = (ref: string): string => evalCell(cells[ref] ?? '', (r) => cells[r] ?? '');

  const commit = (ref: string, raw: string) => {
    const next = { ...cells };
    if (raw === '') delete next[ref];
    else next[ref] = raw;
    setCells(next);
    save(next, title);
  };

  const selectCell = (ref: string) => {
    setSel(ref);
    setEditing(cells[ref] ?? '');
  };

  const commitAndMove = () => {
    if (editing !== null) commit(sel, editing);
    const col = sel.charCodeAt(0);
    const row = parseInt(sel.slice(1), 10);
    if (row < ROWS) {
      const next = `${String.fromCharCode(col)}${row + 1}`;
      setSel(next);
      setEditing(cells[next] ?? '');
    }
  };

  const runAi = async () => {
    const s = getSettings();
    if (!s.apiKey) {
      setToast('Add your API key in Settings first.');
      setAiOpen(false);
      return;
    }
    const topic = aiPrompt.trim();
    if (!topic) return;
    setAiBusy(true);
    setToast('');
    try {
      const out = await chatStream(s, [
        {
          role: 'system',
          content:
            'You generate spreadsheet data. Return ONLY a tab-separated table. First row is a header row. No markdown fences, no commentary, max 8 columns and 20 rows. Same language as the request.',
        },
        { role: 'user', content: topic },
      ]);
      const clean = out.replace(/```[a-z]*\n?/gi, '').trim();
      const lines = clean.split('\n').filter((l) => l.trim());
      const next = { ...cells };
      let r = 1;
      for (const line of lines.slice(0, ROWS)) {
        const cols = line.split('\t');
        for (let c = 0; c < Math.min(cols.length, COLS); c++) {
          const v = cols[c].trim();
          if (v) next[`${colName(c)}${r}`] = v;
        }
        r++;
      }
      setCells(next);
      save(next, title);
      setAiOpen(false);
      setToast(`Imported ${Math.min(lines.length, ROWS)} rows starting at A1.`);
    } catch (e) {
      setToast(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  const exportCsv = () => {
    let maxRow = 0;
    let maxCol = 0;
    for (const k of Object.keys(cells)) {
      const r = parseInt(k.slice(1), 10);
      const c = k.charCodeAt(0) - 65;
      if (Number.isFinite(r) && cells[k] !== '') {
        maxRow = Math.max(maxRow, r);
        maxCol = Math.max(maxCol, c);
      }
    }
    const rows: string[] = [];
    for (let r = 1; r <= maxRow; r++) {
      const cols: string[] = [];
      for (let c = 0; c <= maxCol; c++) {
        const v = display(`${colName(c)}${r}`);
        cols.push(/[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
      }
      rows.push(cols.join(','));
    }
    void downloadText(`${title.replace(/[^\w-]+/g, '_') || 'sheet'}.csv`, rows.join('\n'), 'text/csv').then(setToast);
  };

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  };

  const rows = [];
  for (let r = 1; r <= ROWS; r++) {
    const cols = [];
    for (let c = 0; c < COLS; c++) {
      const ref = `${colName(c)}${r}`;
      const isSel = ref === sel;
      cols.push(
        <button
          key={ref}
          className={`cell${isSel ? ' selected' : ''}`}
          onClick={() => selectCell(ref)}
        >
          {display(ref)}
        </button>,
      );
    }
    rows.push(
      <div className="sheet-row" key={r}>
        <div className="cell head">{r}</div>
        {cols}
      </div>,
    );
  }

  return (
    <div className="screen sheet-screen">
      <header className="screen-head">
        <input
          className="title-input"
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            save(cells, e.target.value);
          }}
          placeholder="Sheet title"
        />
        <button className="btn small" onClick={() => setAiOpen(true)}>
          AI Fill
        </button>
        <button className="btn small" onClick={exportCsv}>
          CSV
        </button>
      </header>

      <div className="sheet-grid">
        <div className="sheet-row">
          <div className="cell head corner" />
          {Array.from({ length: COLS }, (_, c) => (
            <div key={c} className="cell head">
              {colName(c)}
            </div>
          ))}
        </div>
        {rows}
      </div>

      <div className="formula-bar">
        <span className="formula-ref">{sel}</span>
        <input
          key={sel}
          className="input"
          value={editing ?? ''}
          placeholder="Value or =A1+B2*2, =SUM(A1:A9)"
          onChange={(e) => setEditing(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') commitAndMove();
          }}
          onBlur={() => {
            if (editing !== null) commit(sel, editing);
          }}
        />
      </div>

      {aiOpen && (
        <div className="modal" onClick={() => !aiBusy && setAiOpen(false)}>
          <div className="modal-body" onClick={(e) => e.stopPropagation()}>
            <h3>Generate a table</h3>
            <p className="hint">Describe the table. It will be written starting at A1.</p>
            <textarea
              className="input"
              rows={3}
              value={aiPrompt}
              placeholder="e.g. Monthly budget for a student with rent, food, transport"
              onChange={(e) => setAiPrompt(e.target.value)}
            />
            <div className="btn-row">
              <button className="btn primary" disabled={aiBusy || !aiPrompt.trim()} onClick={() => void runAi()}>
                {aiBusy ? 'Generating...' : 'Generate'}
              </button>
              <button className="btn" onClick={() => setAiOpen(false)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
