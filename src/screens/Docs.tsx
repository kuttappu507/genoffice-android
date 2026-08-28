import { useEffect, useMemo, useRef, useState } from 'react';
import { Icon, FileTypeIcon } from '../components/Icon';
import { chatStream, errMsg } from '../lib/ai-client';
import { debounce, getDoc, getSettings, putDoc, downloadText, uid } from '../lib/storage';
import { exportDocx, importDocx, openFilePicker, saveBinary, sanitizeName, textToHtml } from '../lib/fileio';

interface DocData {
  html: string;
}

type AiMode = 'continue' | 'summarize' | 'rewrite';
type RibbonTab = 'home' | 'insert' | 'ai';

const AI_PROMPTS: Record<AiMode, { sys: string; label: string; icon: string }> = {
  continue: {
    sys: 'You are a writing assistant. Continue the user\'s text naturally in the same voice and language. Output ONLY the continuation, no heading, no commentary.',
    label: 'Continue',
    icon: 'sparkle',
  },
  summarize: {
    sys: 'Summarize the given text into tight bullet points, in the same language. Output only the bullets.',
    label: 'Summarize',
    icon: 'listBullet',
  },
  rewrite: {
    sys: 'Rewrite the given text: clearer, tighter, same meaning and language. Output only the rewritten text.',
    label: 'Rewrite',
    icon: 'edit',
  },
};

/** Office-standard palette. */
const PALETTE = [
  '#000000', '#404040', '#8B0000', '#C00000', '#FFC000', '#FFFF00',
  '#92D050', '#00B050', '#00B0F0', '#0070C0', '#1F4E79', '#7030A0',
];

interface RBtn {
  icon: string;
  label?: string;
  cmd?: string;
  arg?: string;
  state?: string;
}

const HOME_RIBBON: (RBtn | 'div')[] = [
  { icon: 'undo', cmd: 'undo', label: 'Undo' },
  { icon: 'redo', cmd: 'redo', label: 'Redo' },
  'div',
  { icon: 'bold', cmd: 'bold', state: 'bold', label: 'Bold' },
  { icon: 'italic', cmd: 'italic', state: 'italic', label: 'Italic' },
  { icon: 'underline', cmd: 'underline', state: 'underline', label: 'Underline' },
  { icon: 'strike', cmd: 'strikeThrough', state: 'strikeThrough', label: 'Strikethrough' },
  'div',
  { icon: 'h1', cmd: 'formatBlock', arg: '<h1>', label: 'Heading 1' },
  { icon: 'h2', cmd: 'formatBlock', arg: '<h2>', label: 'Heading 2' },
  'div',
  { icon: 'listBullet', cmd: 'insertUnorderedList', state: 'insertUnorderedList', label: 'Bullets' },
  { icon: 'listOrdered', cmd: 'insertOrderedList', state: 'insertOrderedList', label: 'Numbering' },
  { icon: 'outdent', cmd: 'outdent', label: 'Decrease indent' },
  { icon: 'indent', cmd: 'indent', label: 'Increase indent' },
  'div',
  { icon: 'alignLeft', cmd: 'justifyLeft', state: 'justifyLeft', label: 'Align left' },
  { icon: 'alignCenter', cmd: 'justifyCenter', state: 'justifyCenter', label: 'Center' },
  { icon: 'alignRight', cmd: 'justifyRight', state: 'justifyRight', label: 'Align right' },
  { icon: 'alignJustify', cmd: 'justifyFull', state: 'justifyFull', label: 'Justify' },
  'div',
  { icon: 'clearFormat', cmd: 'removeFormat', label: 'Clear formatting' },
];

export default function Docs({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const docId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Document' : 'Untitled document');
  const [words, setWords] = useState(0);
  const [aiOut, setAiOut] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [toast, setToast] = useState('');
  const [rTab, setRTab] = useState<RibbonTab>('home');
  const [palette, setPalette] = useState<'text' | 'hilite' | null>(null);
  const [menu, setMenu] = useState(false);
  const [fmt, setFmt] = useState<Record<string, boolean>>({});
  const editorRef = useRef<HTMLDivElement>(null);
  const loaded = useRef(false);

  useEffect(() => {
    document.execCommand('styleWithCSS', false, 'true');
    if (loaded.current) return;
    loaded.current = true;
    if (initialId) {
      const d = getDoc<DocData>(initialId);
      if (d && editorRef.current) {
        editorRef.current.innerHTML = d.html;
        setWords(d.html ? (editorRef.current.innerText ?? '').trim().split(/\s+/).length : 0);
      }
    }
  }, [initialId]);

  // live ribbon states, like Word's active format highlight
  useEffect(() => {
    const h = () => {
      try {
        const active = document.activeElement === editorRef.current;
        if (!active) return;
        setFmt({
          bold: document.queryCommandState('bold'),
          italic: document.queryCommandState('italic'),
          underline: document.queryCommandState('underline'),
          strikeThrough: document.queryCommandState('strikeThrough'),
          insertUnorderedList: document.queryCommandState('insertUnorderedList'),
          insertOrderedList: document.queryCommandState('insertOrderedList'),
          justifyLeft: document.queryCommandState('justifyLeft'),
          justifyCenter: document.queryCommandState('justifyCenter'),
          justifyRight: document.queryCommandState('justifyRight'),
          justifyFull: document.queryCommandState('justifyFull'),
        });
      } catch {
        /* noop */
      }
    };
    document.addEventListener('selectionchange', h);
    return () => document.removeEventListener('selectionchange', h);
  }, []);

  const save = useMemo(
    () =>
      debounce(() => {
        if (editorRef.current) putDoc<DocData>('doc', docId, title, { html: editorRef.current.innerHTML });
      }, 800),
    [docId, title],
  );

  const recount = () => {
    const t = editorRef.current?.innerText ?? '';
    setWords(t.trim() ? t.trim().split(/\s+/).length : 0);
  };

  const exec = (cmd: string, value?: string) => {
    document.execCommand(cmd, false, value);
    editorRef.current?.focus();
    save();
  };

  const insertAtEnd = (text: string) => {
    const el = editorRef.current;
    if (!el) return;
    el.focus();
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    document.execCommand('insertHTML', false, `<p><br></p><p>${text.replace(/\n/g, '<br>')}</p>`);
    recount();
    save();
  };

  const insertAtCaret = (text: string) => {
    editorRef.current?.focus();
    document.execCommand('insertText', false, text);
    recount();
    save();
  };

  const runAi = async (mode: AiMode) => {
    const s = getSettings();
    if (!s.apiKey) {
      flash('Add your API key in Settings first.');
      return;
    }
    const full = (editorRef.current?.innerText ?? '').trim();
    if (!full) {
      flash('Write something first.');
      return;
    }
    const selText = window.getSelection?.()?.toString() ?? '';
    const body = mode === 'continue' ? full.slice(-4000) : (selText || full).slice(0, 6000);
    setAiBusy(true);
    setAiOut('');
    try {
      const out = await chatStream(s, [
        { role: 'system', content: AI_PROMPTS[mode].sys },
        { role: 'user', content: body },
      ]);
      setAiOut(out);
      if (mode === 'continue') {
        insertAtEnd(out);
        flash('Continuation inserted at the end.');
      }
    } catch (e) {
      setAiOut(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  const openFile = async () => {
    setMenu(false);
    const pick = await openFilePicker('.docx,.txt,.md,.html,.htm');
    if (!pick) return;
    try {
      const ext = (pick.name.split('.').pop() ?? '').toLowerCase();
      let html = '';
      if (ext === 'docx') html = await importDocx(pick.buf);
      else if (ext === 'html' || ext === 'htm') {
        html = new DOMParser().parseFromString(new TextDecoder().decode(pick.buf), 'text/html').body.innerHTML;
      } else html = textToHtml(new TextDecoder().decode(pick.buf));
      if (editorRef.current) editorRef.current.innerHTML = html;
      setTitle(pick.name.replace(/\.[^.]+$/, ''));
      recount();
      save();
      flash(`Opened ${pick.name}`);
    } catch (e) {
      flash(`Could not open: ${errMsg(e)}`);
    }
  };

  const saveDocx = async () => {
    try {
      const blob = await exportDocx(title, editorRef.current?.innerHTML ?? '');
      flash(await saveBinary(sanitizeName(title, 'docx'), blob, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'));
    } catch (e) {
      flash(`Save failed: ${errMsg(e)}`);
    }
  };

  const saveHtml = () => {
    setMenu(false);
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body>${editorRef.current?.innerHTML ?? ''}</body></html>`;
    void downloadText(`${title.replace(/[^\w-]+/g, '_') || 'document'}.html`, html, 'text/html').then(flash);
  };

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  };

  const homeBtn = (b: RBtn, i: number) => (
    <button
      key={`${b.cmd}-${i}`}
      className={`rbtn${b.state && fmt[b.state] ? ' active' : ''}`}
      title={b.label}
      aria-label={b.label}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => b.cmd && exec(b.cmd, b.arg)}
    >
      <Icon name={b.icon} size={19} />
    </button>
  );

  return (
    <div className="edscreen" style={{ ['--app' as string]: 'var(--word)' }}>
      <header className="appbar">
        <button className="icon-btn light" aria-label="Back to Home" onClick={onExit}>
          <Icon name="arrowLeft" size={21} />
        </button>
        <FileTypeIcon kind="doc" size={26} />
        <input
          className="appbar-title"
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            save();
          }}
          placeholder="Document title"
        />
        <button className="icon-btn light" aria-label="Save as .docx" onClick={() => void saveDocx()}>
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
                  <Icon name="folder" size={18} /> Open file
                </button>
                <button className="menu-item" onClick={saveHtml}>
                  <Icon name="download" size={18} /> Save as HTML
                </button>
              </div>
            </>
          )}
        </div>
      </header>

      <div className="edbody">
        <div className="paper-wrap">
          <div
            className="paper"
            ref={editorRef}
            contentEditable
            suppressContentEditableWarning
            onInput={() => {
              recount();
              save();
            }}
            data-placeholder="Start writing…"
          />
          <div className="ed-status">
            <span>{words} words</span>
            <span>Saved on device</span>
          </div>
        </div>
      </div>

      {(aiBusy || aiOut) && (
        <div className="ai-sheet">
          <div className="ai-sheet-head">
            <Icon name="sparkle" size={16} /> AI result
            <button className="icon-btn" aria-label="Dismiss" onClick={() => { setAiOut(''); setAiBusy(false); }}>
              <Icon name="close" size={17} />
            </button>
          </div>
          <div className="ai-sheet-body">{aiBusy ? 'Generating…' : aiOut}</div>
          {!aiBusy && aiOut && (
            <div className="btn-row">
              <button className="btn small primary" onClick={() => { insertAtCaret(aiOut); flash('Inserted at cursor.'); }}>
                Insert at cursor
              </button>
              <button className="btn small" onClick={() => { setAiOut(''); }}>Dismiss</button>
            </div>
          )}
        </div>
      )}

      <div className="ribbon">
        <div className="ribbon-tabs">
          {(['home', 'insert', 'ai'] as RibbonTab[]).map((t) => (
            <button key={t} className={`ribbon-tab${rTab === t ? ' active' : ''}`} onClick={() => { setRTab(t); setPalette(null); }}>
              {t === 'ai' ? 'AI' : t === 'home' ? 'Home' : 'Insert'}
            </button>
          ))}
        </div>

        {palette && (
          <div className="palette-row">
            {PALETTE.map((c) => (
              <button
                key={c}
                className="swatch"
                style={{ background: c }}
                aria-label={`Color ${c}`}
                onClick={() => {
                  exec(palette === 'text' ? 'foreColor' : 'hiliteColor', c);
                  setPalette(null);
                }}
              />
            ))}
          </div>
        )}

        {rTab === 'home' && (
          <div className="ribbon-row">
            {HOME_RIBBON.map((b, i) =>
              b === 'div' ? <span key={`d${i}`} className="rdiv" /> : homeBtn(b, i),
            )}
            <span className="rdiv" />
            <button className="rbtn" title="Font color" aria-label="Font color" onClick={() => setPalette(palette === 'text' ? null : 'text')}>
              <Icon name="fontColor" size={19} />
            </button>
            <button className="rbtn" title="Highlight" aria-label="Highlight" onClick={() => setPalette(palette === 'hilite' ? null : 'hilite')}>
              <Icon name="highlight" size={19} />
            </button>
          </div>
        )}

        {rTab === 'insert' && (
          <div className="ribbon-row">
            <button
              className="rbtn"
              title="Insert table"
              aria-label="Insert table"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() =>
                exec(
                  'insertHTML',
                  '<table><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr></table><p><br></p>',
                )
              }
            >
              <Icon name="table" size={19} />
            </button>
            <button className="rbtn" title="Divider line" aria-label="Insert divider" onMouseDown={(e) => e.preventDefault()} onClick={() => exec('insertHTML', '<hr><p><br></p>')}>
              <Icon name="hr" size={19} />
            </button>
            <button
              className="rbtn"
              title="Insert link"
              aria-label="Insert link"
              onClick={() => {
                const url = window.prompt('Link URL', 'https://');
                if (url) exec('createLink', url);
              }}
            >
              <Icon name="link" size={19} />
            </button>
            <button
              className="rbtn"
              title="Insert date"
              aria-label="Insert date"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insertAtCaret(new Date().toLocaleDateString())}
            >
              <Icon name="calendar" size={19} />
            </button>
          </div>
        )}

        {rTab === 'ai' && (
          <div className="ribbon-row">
            {(Object.keys(AI_PROMPTS) as AiMode[]).map((m) => (
              <button key={m} className="rbtn wide" disabled={aiBusy} onClick={() => void runAi(m)}>
                <Icon name={AI_PROMPTS[m].icon} size={19} />
                <span>{AI_PROMPTS[m].label}</span>
              </button>
            ))}
          </div>
        )}
      </div>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
