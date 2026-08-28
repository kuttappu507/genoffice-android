import { useEffect, useMemo, useRef, useState } from 'react';
import { chatStream, errMsg } from '../lib/ai-client';
import { debounce, getDoc, getSettings, putDoc, downloadText, uid } from '../lib/storage';
import { exportDocx, importDocx, openFilePicker, saveBinary, sanitizeName, textToHtml } from '../lib/fileio';

interface DocData {
  html: string;
}

type AiMode = 'continue' | 'summarize' | 'rewrite';

const AI_PROMPTS: Record<AiMode, { sys: string; label: string }> = {
  continue: {
    sys: 'You are a writing assistant. Continue the user\'s text naturally in the same voice and language. Output ONLY the continuation, no heading, no commentary.',
    label: 'Continue writing',
  },
  summarize: {
    sys: 'Summarize the given text into tight bullet points, in the same language. Output only the bullets.',
    label: 'Summarize selection',
  },
  rewrite: {
    sys: 'Rewrite the given text: clearer, tighter, same meaning and language. Output only the rewritten text.',
    label: 'Rewrite selection',
  },
};

export default function Docs({ initialId }: { initialId?: string }) {
  const docId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Document' : 'Untitled document');
  const [words, setWords] = useState(0);
  const [aiOut, setAiOut] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [toast, setToast] = useState('');
  const editorRef = useRef<HTMLDivElement>(null);
  const loaded = useRef(false);

  useEffect(() => {
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
    document.execCommand('insertText', false, `\n\n${text}`);
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
      setToast('Add your API key in Settings first.');
      return;
    }
    const full = (editorRef.current?.innerText ?? '').trim();
    if (!full) {
      setToast('Write something first.');
      return;
    }
    const selText = window.getSelection?.()?.toString() ?? '';
    const body = mode === 'continue' ? full.slice(-4000) : (selText || full).slice(0, 6000);
    setAiBusy(true);
    setAiOut('');
    setToast('');
    try {
      const out = await chatStream(s, [
        { role: 'system', content: AI_PROMPTS[mode].sys },
        { role: 'user', content: body },
      ]);
      setAiOut(out);
      if (mode === 'continue') {
        insertAtEnd(out);
        setToast('Continuation inserted at the end.');
      }
    } catch (e) {
      setAiOut(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  const openFile = async () => {
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

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  };

  return (
    <div className="screen">
      <header className="screen-head">
        <input
          className="title-input"
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            save();
          }}
          placeholder="Document title"
        />
        <button className="btn small" onClick={() => void openFile()}>
          Open
        </button>
        <button className="btn small primary" onClick={() => void saveDocx()}>
          Docx
        </button>
        <button
          className="btn small"
          onClick={() => {
            const html = `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body>${editorRef.current?.innerHTML ?? ''}</body></html>`;
            void downloadText(`${title.replace(/[^\w-]+/g, '_') || 'document'}.html`, html, 'text/html').then(flash);
          }}
        >
          Html
        </button>
      </header>

      <div className="toolbar">
        <button onClick={() => exec('bold')}><b>B</b></button>
        <button onClick={() => exec('italic')}><i>I</i></button>
        <button onClick={() => exec('underline')}><u>U</u></button>
        <button onClick={() => exec('formatBlock', '<h1>')}>H1</button>
        <button onClick={() => exec('formatBlock', '<h2>')}>H2</button>
        <button onClick={() => exec('insertUnorderedList')}>List</button>
        <button onClick={() => exec('removeFormat')}>Clear</button>
      </div>

      <div
        className="editor"
        ref={editorRef}
        contentEditable
        suppressContentEditableWarning
        onInput={() => {
          recount();
          save();
        }}
        data-placeholder="Start writing..."
      />

      <div className="status-row">
        <span>{words} words</span>
        <span>Saved on this device automatically</span>
      </div>

      <div className="ai-panel">
        <div className="ai-panel-head">AI actions</div>
        <div className="btn-row">
          <button className="btn" disabled={aiBusy} onClick={() => void runAi('continue')}>
            Continue
          </button>
          <button className="btn" disabled={aiBusy} onClick={() => void runAi('summarize')}>
            Summarize
          </button>
          <button className="btn" disabled={aiBusy} onClick={() => void runAi('rewrite')}>
            Rewrite
          </button>
        </div>
        {(aiBusy || aiOut) && (
          <div className="ai-result">
            {aiBusy ? 'Generating...' : (
              <>
                <div className="ai-result-text">{aiOut}</div>
                <div className="btn-row">
                  <button className="btn small primary" onClick={() => { insertAtCaret(aiOut); flash('Inserted at cursor.'); }}>
                    Insert at cursor
                  </button>
                  <button className="btn small" onClick={() => { setAiOut(''); }}>Dismiss</button>
                </div>
              </>
            )}
          </div>
        )}
      </div>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
