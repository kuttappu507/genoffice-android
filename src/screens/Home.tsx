import { useEffect, useMemo, useState } from 'react';
import type { DocKind, DocMeta } from '../types';
import { FileTypeIcon, Icon } from '../components/Icon';
import { BottomSheet, ConfirmSheet, PromptSheet, SheetMenu, Toast, useToast } from '../components/Sheet';
import { duplicateDoc, getPrefs, getSettings, listDocs, putDoc, removeDoc, renameDoc, storageUsage, togglePin, uid } from '../lib/storage';
import { errMsg } from '../lib/ai-client';
import { baseName, importDocx, importPptx, importWorkbook, openFilePicker, textToHtml } from '../lib/fileio';
import { normalizeDeck } from '../lib/deck-model';
import { tap } from '../lib/native';

type Go = 'chat' | 'docs' | 'sheets' | 'slides' | 'settings';

const CREATE: { kind: DocKind; label: string; desc: string; go: Go }[] = [
  { kind: 'doc', label: 'Document', desc: 'Word-style editor', go: 'docs' },
  { kind: 'sheet', label: 'Spreadsheet', desc: 'Formulas & charts', go: 'sheets' },
  { kind: 'deck', label: 'Presentation', desc: 'Slides & presenter', go: 'slides' },
  { kind: 'chat', label: 'AI Chat', desc: 'Any model, your key', go: 'chat' },
];

const KIND_LABEL: Record<DocKind, string> = { doc: 'Document', sheet: 'Spreadsheet', deck: 'Presentation', chat: 'Chat' };

function ago(ts: number): string {
  const s = Math.floor((Date.now() - ts) / 1000);
  if (s < 60) return 'just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  if (s < 86400 * 7) return `${Math.floor(s / 86400)} d ago`;
  return new Date(ts).toLocaleDateString();
}

function greeting(): string {
  const h = new Date().getHours();
  return h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening';
}

export default function Home({ onOpen, onGo }: { onOpen: (kind: DocKind, id: string) => void; onGo: (tab: Go) => void }) {
  const [docs, setDocs] = useState<DocMeta[]>(() => listDocs());
  const [filter, setFilter] = useState<DocKind | 'all'>('all');
  const [query, setQuery] = useState('');
  const [searching, setSearching] = useState(false);
  const settings = useMemo(() => getSettings(), []);
  const configured = settings.apiKey.trim().length > 0;
  const [opening, setOpening] = useState(false);
  const [toast, flash] = useToast();
  const [menuDoc, setMenuDoc] = useState<DocMeta | null>(null);
  const [renameTarget, setRenameTarget] = useState<DocMeta | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<DocMeta | null>(null);
  const [newOpen, setNewOpen] = useState(false);
  const usage = useMemo(() => storageUsage(), [docs]);
  const prefs = getPrefs();

  const refresh = () => setDocs(listDocs());
  useEffect(() => {
    const h = () => refresh();
    window.addEventListener('focus', h);
    return () => window.removeEventListener('focus', h);
  }, []);

  const shown = useMemo(() => {
    const q = query.trim().toLowerCase();
    return docs.filter((d) => (filter === 'all' || d.kind === filter) && (!q || d.title.toLowerCase().includes(q)));
  }, [docs, filter, query]);
  const pinned = shown.filter((d) => d.pinned);
  const rest = shown.filter((d) => !d.pinned);

  const openAny = async () => {
    const pick = await openFilePicker('.docx,.xlsx,.xls,.csv,.tsv,.ods,.pptx,.txt,.md,.html,.htm');
    if (!pick) return;
    setOpening(true);
    try {
      const ext = (pick.name.split('.').pop() ?? '').toLowerCase();
      const base = baseName(pick.name);
      const id = uid();
      if (ext === 'docx') {
        putDoc('doc', id, base, { html: await importDocx(pick.buf) });
        onOpen('doc', id);
      } else if (['xlsx', 'xls', 'csv', 'tsv', 'ods'].includes(ext)) {
        putDoc('sheet', id, base, await importWorkbook(pick.buf, pick.name));
        onOpen('sheet', id);
      } else if (ext === 'pptx') {
        const slides = normalizeDeck(await importPptx(pick.buf));
        if (slides.length === 0) throw new Error('No slides found in the file');
        putDoc('deck', id, base, { slides });
        onOpen('deck', id);
      } else if (['txt', 'md', 'html', 'htm'].includes(ext)) {
        const text = new TextDecoder().decode(pick.buf);
        const html = ext === 'html' || ext === 'htm' ? new DOMParser().parseFromString(text, 'text/html').body.innerHTML : textToHtml(text);
        putDoc('doc', id, base, { html });
        onOpen('doc', id);
      } else flash(`Unsupported file type: .${ext}`);
    } catch (e) {
      flash(`Could not open ${pick.name}: ${errMsg(e)}`);
    } finally {
      setOpening(false);
    }
  };

  const row = (d: DocMeta) => (
    <div key={d.id} className="list-row">
      <button className="list-main" onClick={() => onOpen(d.kind, d.id)} onContextMenu={(e) => { e.preventDefault(); setMenuDoc(d); }}>
        <FileTypeIcon kind={d.kind} size={32} />
        <span className="list-text">
          <span className="list-title">{d.pinned && <Icon name="pin" size={12} className="pin-ic" />}{d.title || 'Untitled'}</span>
          <span className="list-sub">{KIND_LABEL[d.kind]} · {ago(d.updated)}</span>
        </span>
      </button>
      <button className="icon-btn" aria-label={`Options for ${d.title}`} onClick={() => { setMenuDoc(d); void tap(); }}>
        <Icon name="more" size={18} />
      </button>
    </div>
  );

  return (
    <div className="screen hub">
      <header className="hub-head">
        <div className="brand">
          <span className="brand-mark">G</span>
          <span>Gen<span className="brand-thin">Office</span></span>
        </div>
        <button className={`icon-btn${searching ? ' on' : ''}`} aria-label="Search files" onClick={() => { setSearching(!searching); setQuery(''); }}>
          <Icon name="search" size={21} />
        </button>
        <button className="icon-btn" aria-label="Settings" onClick={() => onGo('settings')}>
          <Icon name="settings" size={21} />
        </button>
      </header>

      {searching ? (
        <div className="hub-search">
          <Icon name="search" size={16} className="dim" />
          <input className="input" autoFocus placeholder="Search your files" value={query} onChange={(e) => setQuery(e.target.value)} />
          {query && <button className="icon-btn" aria-label="Clear" onClick={() => setQuery('')}><Icon name="close" size={16} /></button>}
        </div>
      ) : (
        <p className="hub-greet">{greeting()}{prefs.onboarded ? '' : ' — welcome!'}</p>
      )}

      {!configured && (
        <button className="keycard" onClick={() => onGo('settings')}>
          <Icon name="key" size={18} />
          <span><strong>Add an API key to unlock AI</strong><small>Editing, files and formulas work without one.</small></span>
          <Icon name="chevronRight" size={16} />
        </button>
      )}

      <section className="tiles">
        {CREATE.map((a) => (
          <button key={a.kind} className="tile" data-kind={a.kind} onClick={() => { void tap(); onGo(a.go); }}>
            <FileTypeIcon kind={a.kind} size={38} />
            <span className="tile-text">
              <strong>{a.label}</strong>
              <span>{a.desc}</span>
            </span>
          </button>
        ))}
      </section>

      <button className="openfile" onClick={() => void openAny()} disabled={opening}>
        <span className="openfile-icon"><Icon name="fileOpen" size={22} /></span>
        <span className="openfile-text">
          <strong>{opening ? 'Opening…' : 'Open from device'}</strong>
          <span>Word · Excel · PowerPoint · CSV · Markdown · Text</span>
        </span>
        <Icon name="chevronRight" size={18} className="dim" />
      </button>

      <div className="hub-filter chip-row">
        {(['all', 'doc', 'sheet', 'deck', 'chat'] as const).map((k) => (
          <button key={k} className={`chip${filter === k ? ' on' : ''}`} onClick={() => setFilter(k)}>
            {k === 'all' ? `All (${docs.length})` : `${KIND_LABEL[k]}s (${docs.filter((d) => d.kind === k).length})`}
          </button>
        ))}
      </div>

      {pinned.length > 0 && (
        <>
          <h2 className="section-title"><Icon name="pin" size={15} /> Pinned</h2>
          <div className="list">{pinned.map(row)}</div>
        </>
      )}

      <h2 className="section-title"><Icon name="clock" size={15} /> Recent</h2>
      {rest.length === 0 ? (
        <p className="empty">{query ? `No files match “${query}”.` : docs.length ? 'Nothing else here.' : 'Nothing here yet. Create a document, sheet or presentation above, or open a file from your device.'}</p>
      ) : (
        <div className="list">{rest.map(row)}</div>
      )}

      <p className="hub-foot">{usage.docs} file{usage.docs === 1 ? '' : 's'} · {(usage.bytes / 1024).toFixed(0)} KB stored on this device</p>

      <button className="fab" aria-label="Create new" onClick={() => setNewOpen(true)}>
        <Icon name="plus" size={24} />
      </button>

      <BottomSheet open={newOpen} onClose={() => setNewOpen(false)} title="Create new">
        <SheetMenu onClose={() => setNewOpen(false)} items={CREATE.map((a) => ({ icon: a.kind === 'doc' ? 'fileText' : a.kind === 'sheet' ? 'grid' : a.kind === 'deck' ? 'monitor' : 'chat', label: a.label, hint: a.desc, onRun: () => onGo(a.go) }))} />
      </BottomSheet>

      <BottomSheet open={!!menuDoc} onClose={() => setMenuDoc(null)} title={menuDoc?.title || 'Untitled'}>
        {menuDoc && (
          <SheetMenu
            onClose={() => setMenuDoc(null)}
            items={[
              { icon: 'externalLink', label: 'Open', onRun: () => onOpen(menuDoc.kind, menuDoc.id) },
              { icon: 'pin', label: menuDoc.pinned ? 'Unpin' : 'Pin to top', onRun: () => { togglePin(menuDoc.id); refresh(); } },
              { icon: 'rename', label: 'Rename', onRun: () => setRenameTarget(menuDoc) },
              { icon: 'duplicate', label: 'Duplicate', onRun: () => { const c = duplicateDoc(menuDoc.id); refresh(); flash(c ? `Created “${c.title}”` : 'Could not duplicate.'); } },
              'divider',
              { icon: 'trash', label: 'Delete', danger: true, onRun: () => setDeleteTarget(menuDoc) },
            ]}
          />
        )}
      </BottomSheet>

      <PromptSheet open={!!renameTarget} title="Rename" initial={renameTarget?.title ?? ''} confirmLabel="Rename" onSubmit={(v) => { if (renameTarget && v.trim()) { renameDoc(renameTarget.id, v.trim()); refresh(); } }} onClose={() => setRenameTarget(null)} />
      <ConfirmSheet open={!!deleteTarget} title={`Delete “${deleteTarget?.title}”?`} message="This removes the file from this device. It cannot be undone." onConfirm={() => { if (deleteTarget) { removeDoc(deleteTarget.id); refresh(); flash('Deleted.'); } }} onClose={() => setDeleteTarget(null)} />
      <Toast msg={toast} />
    </div>
  );
}
