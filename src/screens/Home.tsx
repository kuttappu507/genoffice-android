import { useMemo } from 'react';
import type { DocKind } from '../types';
import { getSettings, listDocs } from '../lib/storage';

const ACTIONS: { kind: DocKind; label: string; desc: string }[] = [
  { kind: 'doc', label: 'Document', desc: 'Write with AI assistance' },
  { kind: 'sheet', label: 'Spreadsheet', desc: 'Grid, formulas, AI tables' },
  { kind: 'deck', label: 'Presentation', desc: 'Outlines and present mode' },
  { kind: 'chat', label: 'AI Chat', desc: 'Streaming chat, any model' },
];

const KIND_LABEL: Record<DocKind, string> = {
  doc: 'Doc',
  sheet: 'Sheet',
  deck: 'Deck',
  chat: 'Chat',
};

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

  return (
    <div className="screen">
      <header className="hero">
        <h1>GenOffice</h1>
        <p className="tagline">
          Local-first AI office for your phone. Bring your own key - OpenRouter or NVIDIA - and keep
          every file on your device.
        </p>
        {!configured && (
          <button className="chip warn" onClick={() => onGo('settings')}>
            No API key yet - tap to add one in Settings
          </button>
        )}
      </header>

      <section className="tiles">
        {ACTIONS.map((a) => (
          <button key={a.kind} className="tile" onClick={() => onGo(a.kind === 'chat' ? 'chat' : a.kind === 'doc' ? 'docs' : a.kind === 'sheet' ? 'sheets' : 'slides')}>
            <strong>{a.label}</strong>
            <span>{a.desc}</span>
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
