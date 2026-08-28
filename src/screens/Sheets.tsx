import { useMemo, useRef, useState } from 'react';
import { Icon, FileTypeIcon } from '../components/Icon';
import { evalCell } from '../lib/formulas';
import { chatStream, errMsg } from '../lib/ai-client';
import { exportXlsx, importSpreadsheet, openFilePicker, saveBinary, sanitizeName } from '../lib/fileio';
import { debounce, downloadText, getDoc, getSettings, putDoc, uid } from '../lib/storage';

const COLS = 26;
const ROWS = 60;

interface SheetData {
  cells: Record<string, string>;
}

function colName(i: number): string {
  return String.fromCharCode(65 + i);
}

export default function Sheets({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
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
  const [menu, setMenu] = useState(false);
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

  const commitAndMove = (dc: number, dr: number) => {
    if (editing !== null) commit(sel, editing);
    const col = sel.charCodeAt(0) - 65 + dc;
    const row = parseInt(sel.slice(1), 10) + dr;
    if (col >= 0 && col < COLS && row >= 1 && row <= ROWS) {
      const next = `${colName(col)}${row}`;
      setSel(next);
      setEditing(cells[next] ?? '');
    }
  };

  const runAi = async () => {
    const s = getSettings();
    if (!s.apiKey) {
      flash('Add your API key in Settings first.');
      setAiOpen(false);
      return;
    }
    const topic = aiPrompt.trim();
    if (!topic) return;
    setAiBusy(true);
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
      flash(`Imported ${Math.min(lines.length, ROWS)} rows starting at A1.`);
    } catch (e) {
      flash(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  const openFile = async () => {
    setMenu(false);
    const pick = await openFilePicker('.xlsx,.xls,.csv');
    if (!pick) return;
    try {
      const { cells: loaded } = await importSpreadsheet(pick.buf);
      setCells(loaded);
      setSel('A1');
      setEditing(loaded['A1'] ?? '');
      const t = pick.name.replace(/\.[^.]+$/, '');
      setTitle(t);
      save(loaded, t);
      flash(`Opened ${pick.name} (${Object.keys(loaded).length} cells)`);
    } catch (e) {
      flash(`Could not open: ${errMsg(e)}`);
    }
  };

  const saveXlsx = async () => {
    try {
      const bytes = await exportXlsx(title, cells);
      await saveBinary(sanitizeName(title, 'xlsx'), bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      flash('Workbook saved');
    } catch (e) {
      flash(`Save failed: ${errMsg(e)}`);
    }
  };

  const exportCsv = () => {
    setMenu(false);
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
    void downloadText(`${title.replace(/[^\w-]+/g, '_') || 'sheet'}.csv`, rows.join('\n'), 'text/csv').then(flash);
  };

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  };

  const selCol = sel.charCodeAt(0) - 65;
  const selRow = parseInt(sel.slice(1), 10);

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
        <div className={`cell head rowhead${r === selRow ? ' hitsel' : ''}`}>{r}</div>
        {cols}
      </div>,
    );
  }

  return (
    <div className="edscreen" style={{ ['--app' as string]: 'var(--excel)' }}>
      <header className="appbar">
        <button className="icon-btn light" aria-label="Back to Home" onClick={onExit}>
          <Icon name="arrowLeft" size={21} />
        </button>
        <FileTypeIcon kind="sheet" size={26} />
        <input
          className="appbar-title"
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            save(cells, e.target.value);
          }}
          placeholder="Sheet title"
        />
        <button className="icon-btn light" aria-label="Save as .xlsx" onClick={() => void saveXlsx()}>
          <Icon name="save" size={20} />
        </button>
        <div className="menu-wrap">
          <button className="icon-btn light" aria-label="More actions" onClick={() => setMenu(!menu)}>
            <Icon name="more" size={20} />
          </button>
          {menu && (
            <>
              <div className="menu-backdrop" onClick={() => setMenu(false)} />
              <div className="menu">
                <button className="menu-item" onClick={() => void openFile()}>
                  <Icon name="folder" size={18} /> Open .xlsx / .csv
                </button>
                <button className="menu-item" onClick={() => { setMenu(false); setAiOpen(true); }}>
                  <Icon name="sparkle" size={18} /> AI Fill
                </button>
                <button className="menu-item" onClick={exportCsv}>
                  <Icon name="download" size={18} /> Export CSV
                </button>
              </div>
            </>
          )}
        </div>
      </header>

      <div className="formula-bar">
        <span className="namebox">{sel}</span>
        <span className="fx">
          <Icon name="fx" size={18} />
        </span>
        <input
          key={sel}
          className="fx-input"
          value={editing ?? ''}
          placeholder="Value or =A1+B2, =SUM(A1:A9)"
          onChange={(e) => setEditing(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') commitAndMove(0, 1);
            if (e.key === 'Tab') {
              e.preventDefault();
              commitAndMove(1, 0);
            }
          }}
          onBlur={() => {
            if (editing !== null) commit(sel, editing);
          }}
        />
      </div>

      <div className="sheet-grid">
        <div className="sheet-row">
          <div className="cell head corner" />
          {Array.from({ length: COLS }, (_, c) => (
            <div key={c} className={`cell head colhead${c === selCol ? ' hitsel' : ''}`}>
              {colName(c)}
            </div>
          ))}
        </div>
        {rows}
      </div>

      <div className="sheet-tabs">
        <button className="stab active">Sheet1</button>
        <button
          className="stab-add"
          aria-label="Add sheet"
          onClick={() => flash('Multiple sheets are coming in a future update.')}
        >
          <Icon name="plus" size={16} />
        </button>
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
                {aiBusy ? 'Generating…' : 'Generate'}
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
