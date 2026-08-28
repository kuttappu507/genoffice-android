import { useEffect, useMemo, useRef, useState } from 'react';
import { FileTypeIcon } from '../components/Icon';
import { Palette, RBtn, RGroup, RSelect, RWide } from '../components/Ribbon';
import { chatStream, errMsg } from '../lib/ai-client';
import { debounce, getDoc, getSettings, putDoc, downloadText, uid } from '../lib/storage';
import { exportDocx, importDocx, openFilePicker, pickImage, saveBinary, sanitizeName, textToHtml } from '../lib/fileio';

interface DocData {
  html: string;
}

type AiMode = 'continue' | 'summarize' | 'rewrite';
type RibbonTab = 'home' | 'insert' | 'review';

const AI_PROMPTS: Record<AiMode, { sys: string; label: string; icon: string }> = {
  continue: {
    sys: 'You are a writing assistant. Continue the user\'s text naturally in the same voice and language. Output ONLY the continuation, no heading, no commentary.',
    label: 'Continue writing',
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

const FONTS = ['Calibri', 'Segoe UI', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Courier New'].map((f) => ({ v: f, t: f }));
const SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32', '36', '48'].map((s) => ({ v: s, t: s }));
const STYLES = [
  { v: 'p', t: 'Normal text' },
  { v: 'h1', t: 'Heading 1' },
  { v: 'h2', t: 'Heading 2' },
  { v: 'h3', t: 'Heading 3' },
  { v: 'blockquote', t: 'Quote' },
];

const DEFAULT_HILITE = '#FFFF00';
const DEFAULT_TEXT = '#C00000';

export default function Docs({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const docId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Document' : 'Untitled document');
  const [words, setWords] = useState(0);
  const [chars, setChars] = useState(0);
  const [aiOut, setAiOut] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [toast, setToast] = useState('');
  const [rTab, setRTab] = useState<RibbonTab>('home');
  const [palette, setPalette] = useState<'text' | 'hilite' | null>(null);
  const [menu, setMenu] = useState(false);
  const [fontFam, setFontFam] = useState('Calibri');
  const [fontSize, setFontSize] = useState('11');
  const [block, setBlock] = useState('p');
  const [fmt, setFmt] = useState<Record<string, boolean>>({});
  const [selBar, setSelBar] = useState<{ top: number; left: number } | null>(null);

  // find & replace
  const [findOpen, setFindOpen] = useState(false);
  const [fq, setFq] = useState('');
  const [fr, setFr] = useState('');
  const [hitCount, setHitCount] = useState(0);
  const [curHit, setCurHit] = useState(0);

  const editorRef = useRef<HTMLDivElement>(null);
  const loaded = useRef(false);
  /** Last known selection inside the editor — restored before commands run. */
  const savedRange = useRef<Range | null>(null);

  useEffect(() => {
    document.execCommand('styleWithCSS', false, 'true');
    if (loaded.current) return;
    loaded.current = true;
    if (initialId) {
      const d = getDoc<DocData>(initialId);
      if (d && editorRef.current) {
        editorRef.current.innerHTML = d.html;
        recount();
      }
    }
  }, [initialId]);

  const recount = () => {
    const t = editorRef.current?.innerText ?? '';
    setWords(t.trim() ? t.trim().split(/\s+/).length : 0);
    setChars(t.replace(/\n/g, '').length);
  };

  // live ribbon states + floating selection toolbar, like Word mobile
  useEffect(() => {
    const h = () => {
      const ed = editorRef.current;
      const sel = window.getSelection();
      try {
        const active = document.activeElement === ed;
        // remember the exact range so ribbon taps (which blur the editor on
        // touch devices) can restore it before execCommand
        if (active && sel && sel.rangeCount > 0 && ed) {
          const r = sel.getRangeAt(0);
          if (ed.contains(r.commonAncestorContainer)) savedRange.current = r.cloneRange();
        }
        if (active) {
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
          const b = (document.queryCommandValue('formatBlock') ?? '').toLowerCase().replace(/[<>]/g, '');
          if (b) setBlock(STYLES.some((s) => s.v === b) ? b : 'p');
        }
      } catch {
        /* noop */
      }
      // floating Aa toolbar while text is selected
      if (sel && !sel.isCollapsed && ed && (ed.contains(sel.anchorNode) || ed.contains(sel.focusNode))) {
        const r = sel.getRangeAt(0).getBoundingClientRect();
        if (r.width > 0 || r.height > 0) {
          setSelBar({
            top: Math.max(8, r.top - 50),
            left: Math.min(window.innerWidth - 90, Math.max(90, r.left + r.width / 2)),
          });
          return;
        }
      }
      setSelBar(null);
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

  /** Focus the editor and put the selection back where the user left it. */
  const restoreSel = (): boolean => {
    const ed = editorRef.current;
    if (!ed) return false;
    ed.focus();
    const sel = window.getSelection();
    if (!sel) return false;
    if (sel.rangeCount > 0 && ed.contains(sel.getRangeAt(0).commonAncestorContainer)) return true;
    const r = savedRange.current;
    if (!r) return false;
    try {
      sel.removeAllRanges();
      sel.addRange(r);
      return true;
    } catch {
      return false;
    }
  };

  const exec = (cmd: string, value?: string) => {
    try {
      document.execCommand('styleWithCSS', false, 'true');
    } catch {
      /* older engines */
    }
    restoreSel();
    document.execCommand(cmd, false, value);
    recount();
    save();
  };

  /** Word-style font size: wraps the selection in a span with an exact px size. */
  const applyFontSize = (px: string) => {
    setFontSize(px);
    const ed = editorRef.current;
    if (!ed) return;
    ed.focus();
    const sel = window.getSelection();
    let range: Range | null =
      sel && sel.rangeCount > 0 && ed.contains(sel.getRangeAt(0).commonAncestorContainer)
        ? sel.getRangeAt(0)
        : savedRange.current;
    if (!range) return; // nothing selected anywhere — nothing to resize
    if (sel && sel.rangeCount === 0) {
      sel.removeAllRanges();
      sel.addRange(range);
    }
    if (range.collapsed) {
      document.execCommand('insertHTML', false, `<span style="font-size:${px}px">\u200b</span>`);
    } else {
      const span = document.createElement('span');
      span.style.fontSize = `${px}px`;
      try {
        span.appendChild(range.extractContents());
        range.insertNode(span);
        sel?.removeAllRanges();
        const r2 = document.createRange();
        r2.selectNodeContents(span);
        sel?.addRange(r2);
        savedRange.current = r2.cloneRange();
      } catch {
        document.execCommand('fontSize', false, '4');
      }
    }
    recount();
    save();
  };

  const applyFont = (f: string) => {
    setFontFam(f);
    exec('fontName', f);
  };
  // ------------------------------------------------------------------ find & replace
  const clearHits = () => {
    const ed = editorRef.current;
    if (!ed) return;
    ed.querySelectorAll('mark.find-hit').forEach((m) => {
      const p = m.parentNode;
      if (p) {
        p.replaceChild(document.createTextNode(m.textContent ?? ''), m);
        p.normalize();
      }
    });
  };

  const runFind = () => {
    const ed = editorRef.current;
    if (!ed) return;
    clearHits();
    setCurHit(0);
    if (!fq.trim()) {
      setHitCount(0);
      return;
    }
    const q = fq.toLowerCase();
    const walker = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
    const targets: Text[] = [];
    let n: Node | null;
    while ((n = walker.nextNode())) {
      const t = n as Text;
      const inMark = t.parentElement?.closest('mark.find-hit');
      if (!inMark && t.textContent && t.textContent.toLowerCase().includes(q)) targets.push(t);
    }
    let count = 0;
    for (const t of targets) {
      const text = t.textContent ?? '';
      const lower = text.toLowerCase();
      const frag = document.createDocumentFragment();
      let i = 0;
      for (;;) {
        const idx = lower.indexOf(q, i);
        if (idx < 0) {
          if (i < text.length) frag.appendChild(document.createTextNode(text.slice(i)));
          break;
        }
        if (idx > i) frag.appendChild(document.createTextNode(text.slice(i, idx)));
        const mk = document.createElement('mark');
        mk.className = 'find-hit';
        mk.textContent = text.slice(idx, idx + fq.length);
        frag.appendChild(mk);
        count++;
        i = idx + q.length;
      }
      t.replaceWith(frag);
    }
    setHitCount(count);
    if (count > 0) {
      setCurHit(1);
      ed.querySelector('mark.find-hit')?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  };

  const gotoHit = (dir: 1 | -1) => {
    const ed = editorRef.current;
    if (!ed) return;
    const ms = Array.from(ed.querySelectorAll('mark.find-hit'));
    if (ms.length === 0) return;
    const idx = (curHit - 1 + dir + ms.length) % ms.length;
    ms.forEach((m) => m.classList.remove('cur'));
    ms[idx].classList.add('cur');
    ms[idx].scrollIntoView({ block: 'center', behavior: 'smooth' });
    setCurHit(idx + 1);
  };

  const replaceOne = () => {
    const ed = editorRef.current;
    if (!ed) return;
    const ms = Array.from(ed.querySelectorAll('mark.find-hit'));
    if (ms.length === 0) return;
    const target = ed.querySelector('mark.find-hit.cur') ?? ms[0];
    target.replaceWith(document.createTextNode(fr));
    ed.normalize();
    runFind();
    save();
  };

  const replaceAll = () => {
    const ed = editorRef.current;
    if (!ed) return;
    const ms = Array.from(ed.querySelectorAll('mark.find-hit'));
    if (ms.length === 0) return;
    ms.forEach((m) => m.replaceWith(document.createTextNode(fr)));
    ed.normalize();
    clearHits();
    setHitCount(0);
    setCurHit(0);
    recount();
    save();
    flash(`Replaced ${ms.length} occurrence(s).`);
  };

  const closeFind = () => {
    clearHits();
    setFindOpen(false);
    setHitCount(0);
    setCurHit(0);
  };

  // ------------------------------------------------------------------ inserts
  const insertAtCaret = (text: string) => {
    editorRef.current?.focus();
    document.execCommand('insertText', false, text);
    recount();
    save();
  };

  const insertHtml = (html: string) => {
    editorRef.current?.focus();
    document.execCommand('insertHTML', false, html);
    recount();
    save();
  };

  const insertImage = async () => {
    try {
      const data = await pickImage(1400);
      if (data) insertHtml(`<img src="${data}" alt=""><p><br></p>`);
    } catch (e) {
      flash(`Could not insert image: ${errMsg(e)}`);
    }
  };

  const insertTable = (rows: number, cols: number) => {
    const trs = Array.from({ length: rows }, () => `<tr>${'<td>&nbsp;</td>'.repeat(cols)}</tr>`).join('');
    insertHtml(`<table>${trs}</table><p><br></p>`);
  };

  // ------------------------------------------------------------------ AI
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

  // ------------------------------------------------------------------ files
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

  return (
    <div className="edscreen" style={{ ['--app' as string]: 'var(--word)' }}>
      <header className="appbar">
        <button className="icon-btn light" aria-label="Back to Home" onClick={onExit}>
          <FileTypeIcon kind="doc" size={22} />
        </button>
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
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" /><polyline points="17 21 17 13 7 13 7 21" /><polyline points="7 3 7 8 15 8" /></svg>
        </button>
        <div className="menu-wrap">
          <button className="icon-btn light" aria-label="More actions" onClick={() => setMenu(!menu)}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true"><circle cx="12" cy="5" r="1.7" fill="currentColor" /><circle cx="12" cy="12" r="1.7" fill="currentColor" /><circle cx="12" cy="19" r="1.7" fill="currentColor" /></svg>
          </button>
          {menu && (
            <>
              <div className="menu-backdrop" onClick={() => setMenu(false)} />
              <div className="menu">
                <button className="menu-item" onClick={() => void openFile()}>
                  Open file (.docx / .txt)
                </button>
                <button className="menu-item" onClick={saveHtml}>
                  Save as HTML
                </button>
              </div>
            </>
          )}
        </div>
      </header>

      {findOpen && (
        <div className="findbar">
          <input className="input find-input" value={fq} placeholder="Find" onChange={(e) => setFq(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && runFind()} />
          <input className="input find-input" value={fr} placeholder="Replace with" onChange={(e) => setFr(e.target.value)} />
          <div className="find-ops">
            <button className="btn small" onClick={runFind}>Find all</button>
            <button className="btn small" disabled={!hitCount} onClick={() => gotoHit(-1)} aria-label="Previous match">↑</button>
            <button className="btn small" disabled={!hitCount} onClick={() => gotoHit(1)} aria-label="Next match">↓</button>
            <button className="btn small" disabled={!hitCount} onClick={replaceOne}>Replace</button>
            <button className="btn small" disabled={!hitCount} onClick={replaceAll}>All</button>
            <span className="find-count">{hitCount ? `${curHit}/${hitCount}` : ''}</span>
            <button className="btn small" onClick={closeFind} aria-label="Close find and replace">✕</button>
          </div>
        </div>
      )}

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
            <span>{words} words · {chars} characters</span>
            <span>Saved on device</span>
          </div>
        </div>
      </div>

      {/* floating selection toolbar (Word mobile "Aa" bar) */}
      {selBar && (
        <div className="selbar" style={{ top: selBar.top, left: selBar.left }}>
          <button className="icon-btn" aria-label="Bold" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('bold')}><b>B</b></button>
          <button className="icon-btn" aria-label="Italic" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('italic')}><i>I</i></button>
          <button className="icon-btn" aria-label="Underline" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('underline')}><u>U</u></button>
          <button className="icon-btn" aria-label="Strikethrough" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('strikeThrough')}><s>S</s></button>
          <button className="icon-btn aa" aria-label="Show all formatting tools" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => { setRTab('home'); setSelBar(null); }}>Aa</button>
        </div>
      )}

      {(aiBusy || aiOut) && (
        <div className="ai-sheet">
          <div className="ai-sheet-head">
            AI result
            <button className="icon-btn" aria-label="Dismiss" onClick={() => { setAiOut(''); setAiBusy(false); }}>
              ✕
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
          {(['home', 'insert', 'review'] as RibbonTab[]).map((t) => (
            <button key={t} className={`ribbon-tab${rTab === t ? ' active' : ''}`} onClick={() => { setRTab(t); setPalette(null); }}>
              {t === 'home' ? 'Home' : t === 'insert' ? 'Insert' : 'Review'}
            </button>
          ))}
        </div>

        {palette && (
          <Palette
            onPick={(c) => {
              exec(palette === 'text' ? 'foreColor' : 'hiliteColor', c);
              setPalette(null);
            }}
            auto={() => {
              exec(palette === 'text' ? 'foreColor' : 'hiliteColor', palette === 'text' ? '#1b1b1b' : 'transparent');
              setPalette(null);
            }}
          />
        )}

        {rTab === 'home' && (
          <div className="ribbon-row">
            <RGroup label="Undo">
              <RBtn icon="undo" label="Undo" keepFocus onRun={() => exec('undo')} />
              <RBtn icon="redo" label="Redo" keepFocus onRun={() => exec('redo')} />
            </RGroup>
            <RGroup label="Font">
              <RSelect value={fontFam} options={FONTS} onChange={applyFont} width={104} title="Font" />
              <RSelect value={fontSize} options={SIZES} onChange={applyFontSize} width={58} title="Font size" />
              <RBtn icon="bold" label="Bold" active={fmt.bold} keepFocus onRun={() => exec('bold')} />
              <RBtn icon="italic" label="Italic" active={fmt.italic} keepFocus onRun={() => exec('italic')} />
              <RBtn icon="underline" label="Underline" active={fmt.underline} keepFocus onRun={() => exec('underline')} />
              <RBtn icon="strike" label="Strike" active={fmt.strikeThrough} keepFocus onRun={() => exec('strikeThrough')} />
              <RBtn icon="fontColor" label="Color" colorBar={DEFAULT_TEXT} keepFocus onRun={() => setPalette(palette === 'text' ? null : 'text')} />
              <RBtn icon="highlight" label="Highlight" colorBar={DEFAULT_HILITE} keepFocus onRun={() => setPalette(palette === 'hilite' ? null : 'hilite')} />
              <RBtn icon="clearFormat" label="Clear" keepFocus onRun={() => exec('removeFormat')} />
            </RGroup>
            <RGroup label="Paragraph">
              <RBtn icon="listBullet" label="Bullets" active={fmt.insertUnorderedList} keepFocus onRun={() => exec('insertUnorderedList')} />
              <RBtn icon="listOrdered" label="Numbering" active={fmt.insertOrderedList} keepFocus onRun={() => exec('insertOrderedList')} />
              <RBtn icon="outdent" label="Less indent" keepFocus onRun={() => exec('outdent')} />
              <RBtn icon="indent" label="More indent" keepFocus onRun={() => exec('indent')} />
              <RBtn icon="alignLeft" label="Left" active={fmt.justifyLeft} keepFocus onRun={() => exec('justifyLeft')} />
              <RBtn icon="alignCenter" label="Center" active={fmt.justifyCenter} keepFocus onRun={() => exec('justifyCenter')} />
              <RBtn icon="alignRight" label="Right" active={fmt.justifyRight} keepFocus onRun={() => exec('justifyRight')} />
              <RBtn icon="alignJustify" label="Justify" active={fmt.justifyFull} keepFocus onRun={() => exec('justifyFull')} />
            </RGroup>
            <RGroup label="Styles">
              <RSelect
                value={block}
                options={STYLES}
                onChange={(v) => {
                  setBlock(v);
                  exec('formatBlock', `<${v}>`);
                }}
                width={118}
                title="Paragraph style"
              />
            </RGroup>
            <RGroup label="Editing">
              <RBtn icon="search" label="Find" active={findOpen} onRun={() => setFindOpen(!findOpen)} />
            </RGroup>
          </div>
        )}

        {rTab === 'insert' && (
          <div className="ribbon-row">
            <RGroup label="Media">
              <RBtn icon="image" label="Picture" keepFocus onRun={() => void insertImage()} />
            </RGroup>
            <RGroup label="Tables">
              <RBtn icon="table" label="2 × 3" keepFocus onRun={() => insertTable(2, 3)} />
              <RBtn icon="table" label="3 × 3" keepFocus onRun={() => insertTable(3, 3)} />
              <RBtn icon="table" label="4 × 4" keepFocus onRun={() => insertTable(4, 4)} />
            </RGroup>
            <RGroup label="Links">
              <RBtn
                icon="link"
                label="Link"
                keepFocus
                onRun={() => {
                  const url = window.prompt('Link URL', 'https://');
                  if (url) exec('createLink', url);
                }}
              />
              <RBtn icon="hr" label="Divider" keepFocus onRun={() => insertHtml('<hr><p><br></p>')} />
            </RGroup>
            <RGroup label="Date">
              <RBtn icon="calendar" label="Today" keepFocus onRun={() => insertAtCaret(new Date().toLocaleDateString())} />
            </RGroup>
          </div>
        )}

        {rTab === 'review' && (
          <div className="ribbon-row">
            <RGroup label="AI writing">
              {(Object.keys(AI_PROMPTS) as AiMode[]).map((m) => (
                <RWide key={m} icon={AI_PROMPTS[m].icon} label={AI_PROMPTS[m].label} disabled={aiBusy} onRun={() => void runAi(m)} />
              ))}
            </RGroup>
            <RGroup label="Proofing">
              <RBtn icon="search" label="Find & Replace" active={findOpen} onRun={() => setFindOpen(!findOpen)} />
            </RGroup>
          </div>
        )}
      </div>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
