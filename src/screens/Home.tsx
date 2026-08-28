import { useState } from 'react';
import type { DocKind } from '../types';
import { FileTypeIcon, Icon } from '../components/Icon';
import { getSettings, listDocs, putDoc, removeDoc, uid } from '../lib/storage';
import { errMsg } from '../lib/ai-client';
import { baseName, importDocx, importPptx, importSpreadsheet, openFilePicker, textToHtml } from '../lib/fileio';

const CREATE: { kind: DocKind; label: string; desc: string; go: 'docs' | 'sheets' | 'slides' | 'chat' }[] = [
  { kind: 'doc', label: 'Document', desc: 'Word-style editor + AI', go: 'docs' },
  { kind: 'sheet', label: 'Spreadsheet', desc: 'Grid, formulas, AI tables', go: 'sheets' },
  { kind: 'deck', label: 'Presentation', desc: 'Slides, present mode', go: 'slides' },
  { kind: 'chat', label: 'AI Chat', desc: 'Streaming, any model', go: 'chat' },
];

function ago(ts: number): string {
  const s = Math.floor((Date.now() - ts) / 1000);
  if (s < 60) return 'just now';
  if (s < 3600) return `${Math.floor(s / 60)}m ago`;
  if (s < 86400) return `${Math.floor(s / 3600)}h ago`;
  return `${Math.floor(s / 86400)}d ago`;
}

export default function Home({
  onOpen,
  onGo,
}: {
  onOpen: (kind: DocKind, id: string) => void;
  onGo: (tab: 'chat' | 'docs' | 'sheets' | 'slides' | 'settings') => void;
}) {
  const [recent, setRecent] = useState(() => listDocs().slice(0, 12));
  const settings = useState(() => getSettings())[0];
  const configured = settings.apiKey.trim().length > 0;
  const [opening, setOpening] = useState(false);
  const [openErr, setOpenErr] = useState('');

  const openAny = async () => {
    const pick = await openFilePicker('.docx,.xlsx,.xls,.csv,.pptx,.txt,.md,.html,.htm');
    if (!pick) return;
    setOpening(true);
    setOpenErr('');
    try {
      const ext = (pick.name.split('.').pop() ?? '').toLowerCase();
      const base = baseName(pick.name);
      if (ext === 'docx') {
        const html = await importDocx(pick.buf);
        const id = uid();
        putDoc('doc', id, base, { html });
        onOpen('doc', id);
        return;
      }
      if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
        const { cells } = await importSpreadsheet(pick.buf);
        const id = uid();
        putDoc('sheet', id, base, { cells });
        onOpen('sheet', id);
        return;
      }
      if (ext === 'pptx') {
        const slides = await importPptx(pick.buf);
        if (slides.length === 0) throw new Error('No slides found in the file');
        const id = uid();
        putDoc('deck', id, base, { slides });
        onOpen('deck', id);
        return;
      }
      if (ext === 'txt' || ext === 'md' || ext === 'html' || ext === 'htm') {
        let html: string;
        if (ext === 'html' || ext === 'htm') {
          html = new DOMParser().parseFromString(new TextDecoder().decode(pick.buf), 'text/html').body.innerHTML;
        } else {
          html = textToHtml(new TextDecoder().decode(pick.buf));
        }
        const id = uid();
        putDoc('doc', id, base, { html });
        onOpen('doc', id);
        return;
      }
      setOpenErr(`Unsupported file type: .${ext}`);
    } catch (e) {
      setOpenErr(`Could not open ${pick.name}: ${errMsg(e)}`);
    } finally {
      setOpening(false);
    }
  };

  return (
    <div className="screen">
      <header className="hub-head">
        <div className="brand">
          <FileTypeIcon kind="doc" size={30} />
          <span>
            Gen<span className="brand-thin">Office</span>
          </span>
        </div>
        <button className="icon-btn" aria-label="Settings" onClick={() => onGo('settings')}>
          <Icon name="settings" size={21} />
        </button>
      </header>

      {!configured && (
        <button className="keycard" onClick={() => onGo('settings')}>
          <Icon name="key" size={18} />
          <span>Add your free API key to unlock AI</span>
          <Icon name="chevronRight" size={16} />
        </button>
      )}

      <button className="openfile" onClick={() => void openAny()} disabled={opening}>
        <span className="openfile-icon">≡</span>
        <span className="openfile-text">
          <strong>{opening ? 'Opening…' : 'Open a file'}</strong>
          <span>.docx · .xlsx · .csv · .pptx · .txt · .md</span>
        </span>
        <Icon name="chevronRight" size={18} className="dim" />
      </button>
      {openErr && <p className="err" style={{ marginTop: 8 }}>{openErr}</p>}

      <h2 className="section-title">Create new</h2>
      <section className="tiles">
        {CREATE.map((a) => (
          <button key={a.kind} className="tile" onClick={() => onGo(a.go)}>
            <FileTypeIcon kind={a.kind} size={40} />
            <span className="tile-text">
              <strong>{a.label}</strong>
              <span>{a.desc}</span>
            </span>
          </button>
        ))}
      </section>

      <h2 className="section-title">
        <Icon name="clock" size={15} /> Recent
      </h2>
      {recent.length === 0 ? (
        <p className="empty">Nothing here yet. Create a document, sheet or deck above, or open an Office file.</p>
      ) : (
        <div className="list">
          {recent.map((d) => (
            <div key={d.id} className="list-row">
              <button className="list-main" onClick={() => onOpen(d.kind, d.id)}>
                <FileTypeIcon kind={d.kind} size={32} />
                <span className="list-title">{d.title}</span>
                <span className="list-time">{ago(d.updated)}</span>
              </button>
              <button
                className="icon-btn list-del"
                aria-label={`Delete ${d.title}`}
                onClick={() => {
                  if (window.confirm(`Delete "${d.title}"? This cannot be undone.`)) {
                    removeDoc(d.id);
                    setRecent(listDocs().slice(0, 12));
                  }
                }}
              >
                ×
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
