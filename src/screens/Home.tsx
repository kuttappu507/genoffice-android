import { useMemo, useState } from 'react';
import type { DocKind } from '../types';
import { getSettings, listDocs, putDoc, uid } from '../lib/storage';
import { errMsg } from '../lib/ai-client';
import { baseName, importDocx, importPptx, importSpreadsheet, openFilePicker, textToHtml } from '../lib/fileio';

const ACTIONS: { kind: DocKind; label: string; desc: string }[] = [
  { kind: 'doc', label: 'Document', desc: 'Write with AI assistance' },
  { kind: 'sheet', label: 'Spreadsheet', desc: 'Grid, formulas, AI tables' },
  { kind: 'deck', label: 'Presentation', desc: 'Outlines and present mode' },
  { kind: 'chat', label: 'AI Chat', desc: 'Streaming chat, any model' },
];

const KIND_LABEL: Record<DocKind, string> = {
  doc: 'W',
  sheet: 'X',
  deck: 'P',
  chat: 'AI',
};

const ICON_LETTER = KIND_LABEL;

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
  const recent = useMemo(() => listDocs().slice(0, 12), []);
  const settings = useMemo(() => getSettings(), []);
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
      <header className="hero">
        <div className="office-header">
          <div className="brand">
            <span className="brand-mark">G</span>
            GenOffice
          </div>
          <button className="btn small" onClick={() => onGo('settings')}>
            Settings
          </button>
        </div>
        <p className="tagline">
          Local-first AI office for your phone. Bring your own key - OpenRouter or NVIDIA - and keep
          every file on your device.
        </p>
        {!configured && (
          <button className="chip warn" onClick={() => onGo('settings')}>
            No API key yet - tap to add one in Settings
          </button>
        )}
        <div className="open-row">
          <button className="btn primary" disabled={opening} onClick={() => void openAny()}>
            {opening ? 'Opening...' : 'Open file'}
          </button>
          <span className="hint">.docx .xlsx .xls .csv .pptx .txt .md</span>
        </div>
        {openErr && <p className="err">{openErr}</p>}
      </header>

      <section className="tiles">
        {ACTIONS.map((a) => (
          <button key={a.kind} className="tile" onClick={() => onGo(a.kind === 'chat' ? 'chat' : a.kind === 'doc' ? 'docs' : a.kind === 'sheet' ? 'sheets' : 'slides')}>
            <span className={`tile-icon icon-${a.kind}`}>{ICON_LETTER[a.kind]}</span>
            <span className="tile-text">
              <strong>{a.label}</strong>
              <span>{a.desc}</span>
            </span>
          </button>
        ))}
      </section>

      <section>
        <h2 className="section-title">Recent</h2>
        {recent.length === 0 ? (
          <p className="empty">Nothing yet. Create a document, sheet, deck or chat above.</p>
        ) : (
          <div className="list">
            {recent.map((d) => (
              <button key={d.id} className="list-row" onClick={() => onOpen(d.kind, d.id)}>
                <span className={`badge badge-${d.kind}`}>{KIND_LABEL[d.kind]}</span>
                <span className="list-title">{d.title}</span>
                <span className="list-time">{ago(d.updated)}</span>
              </button>
            ))}
          </div>
        )}
      </section>
    </div>
  );
}
