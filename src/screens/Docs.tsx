import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { FileTypeIcon, Icon } from '../components/Icon';
import { AppBar, Palette, RBtn, RGroup, RSeg, RSelect, RStepper, RWide, RibbonPanel, RibbonTabs } from '../components/Ribbon';
import { BottomSheet, ConfirmSheet, PromptSheet, SheetMenu, Toast, useToast } from '../components/Sheet';
import { chatStream, errMsg } from '../lib/ai-client';
import { debounce, getDoc, getMeta, getPrefs, getSettings, putDoc, downloadText, uid } from '../lib/storage';
import { exportDocx, exportPdfFromHtml, importDocx, openFilePicker, pickImage, saveBinary, sanitizeName, textToHtml, htmlToMarkdown } from '../lib/fileio';
import { onBack, shareText, tap } from '../lib/native';

// ---------------------------------------------------------------------------
// document model
// ---------------------------------------------------------------------------

interface PageSetup {
  size: 'A4' | 'Letter' | 'Legal';
  orientation: 'portrait' | 'landscape';
  margins: 'normal' | 'narrow' | 'wide';
  header?: string;
  footer?: string;
  pageNumbers?: boolean;
  lineSpacing?: number;
  paraSpacing?: number;
  columns?: 1 | 2;
}

interface DocData {
  html: string;
  page?: PageSetup;
  comments?: DocComment[];
  font?: string;
}

interface DocComment {
  id: string;
  quote: string;
  text: string;
  created: number;
  resolved?: boolean;
}

type AiMode = 'continue' | 'summarize' | 'rewrite' | 'expand' | 'shorten' | 'formal' | 'casual' | 'fixGrammar' | 'translate' | 'outline' | 'draft';
type RibbonTab = 'home' | 'insert' | 'layout' | 'review' | 'view';
type Panel = 'textColor' | 'hilite' | 'table' | 'spacing' | 'symbols' | 'styles' | null;

const AI_PROMPTS: Record<AiMode, { sys: string; label: string; icon: string; scope: 'selection' | 'document' | 'prompt' }> = {
  continue: {
    sys: "You are a writing assistant. Continue the user's text naturally in the same voice and language. Output ONLY the continuation, no heading, no commentary.",
    label: 'Continue writing', icon: 'sparkle', scope: 'document',
  },
  summarize: { sys: 'Summarize the given text into tight bullet points, in the same language. Output only the bullets, one per line starting with "- ".', label: 'Summarize', icon: 'listBullet', scope: 'selection' },
  rewrite: { sys: 'Rewrite the given text: clearer, tighter, same meaning and language. Output only the rewritten text.', label: 'Rewrite', icon: 'edit', scope: 'selection' },
  expand: { sys: 'Expand the given text with more detail and examples, same voice and language. Output only the expanded text.', label: 'Expand', icon: 'zoomIn', scope: 'selection' },
  shorten: { sys: 'Shorten the given text to about half its length without losing key points. Same language. Output only the shortened text.', label: 'Shorten', icon: 'zoomOut', scope: 'selection' },
  formal: { sys: 'Rewrite the given text in a professional, formal tone. Same language and meaning. Output only the rewritten text.', label: 'Make formal', icon: 'readMode', scope: 'selection' },
  casual: { sys: 'Rewrite the given text in a friendly, conversational tone. Same language and meaning. Output only the rewritten text.', label: 'Make casual', icon: 'chat', scope: 'selection' },
  fixGrammar: { sys: 'Fix spelling, grammar and punctuation in the given text. Keep wording and language otherwise unchanged. Output only the corrected text.', label: 'Fix grammar', icon: 'spell', scope: 'selection' },
  translate: { sys: 'Translate the given text into the language named on the first line of the user message (the text follows after a blank line). Output only the translation.', label: 'Translate', icon: 'symbol', scope: 'prompt' },
  outline: { sys: 'Create a document outline for the topic. Use "# " for the title, "## " for section headings and "- " for bullet points. No commentary.', label: 'Outline from topic', icon: 'toc', scope: 'prompt' },
  draft: { sys: 'Write a complete, well-structured document about the topic. Use markdown headings (#, ##) and paragraphs. Same language as the topic. No commentary.', label: 'Draft document', icon: 'blankDoc', scope: 'prompt' },
};

const FONTS = ['Calibri', 'Segoe UI', 'Arial', 'Helvetica', 'Times New Roman', 'Georgia', 'Cambria', 'Verdana', 'Trebuchet MS', 'Garamond', 'Courier New', 'Roboto', 'Noto Sans', 'Noto Serif'].map((f) => ({ v: f, t: f }));
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 40, 48, 72];
const STYLES = [
  { v: 'p', t: 'Normal', preview: 'Normal text', cls: 'st-normal' },
  { v: 'h1', t: 'Heading 1', preview: 'Heading 1', cls: 'st-h1' },
  { v: 'h2', t: 'Heading 2', preview: 'Heading 2', cls: 'st-h2' },
  { v: 'h3', t: 'Heading 3', preview: 'Heading 3', cls: 'st-h3' },
  { v: 'h4', t: 'Title', preview: 'Title', cls: 'st-title' },
  { v: 'blockquote', t: 'Quote', preview: '“Quote”', cls: 'st-quote' },
  { v: 'pre', t: 'Code', preview: 'code()', cls: 'st-code' },
];
const SYMBOLS = ['—', '–', '…', '•', '·', '©', '®', '™', '°', '±', '×', '÷', '≠', '≤', '≥', '≈', '∞', '√', 'π', '∑', '€', '£', '¥', '₹', '¢', '§', '¶', '†', '‡', '→', '←', '↑', '↓', '⇒', '⇔', '★', '☆', '✓', '✗', '♥', '♦', '½', '¼', '¾', 'α', 'β', 'γ', 'δ', 'λ', 'μ', 'Ω'];
const PAGE_SIZES: Record<PageSetup['size'], { w: number; h: number }> = { A4: { w: 8.27, h: 11.69 }, Letter: { w: 8.5, h: 11 }, Legal: { w: 8.5, h: 14 } };
const MARGIN_IN: Record<PageSetup['margins'], number> = { normal: 1, narrow: 0.5, wide: 1.5 };
const DEFAULT_PAGE: PageSetup = { size: 'A4', orientation: 'portrait', margins: 'normal', lineSpacing: 1.15, paraSpacing: 8, columns: 1, pageNumbers: false };

const TEMPLATES: { id: string; label: string; desc: string; html: () => string }[] = [
  { id: 'blank', label: 'Blank', desc: 'Empty page', html: () => '<p><br></p>' },
  {
    id: 'letter', label: 'Letter', desc: 'Formal letter',
    html: () => `<p style="text-align:right">${new Date().toLocaleDateString()}</p><p>Dear [Name],</p><p><br></p><p>I am writing to …</p><p><br></p><p>Thank you for your time and consideration.</p><p><br></p><p>Sincerely,</p><p>[Your name]</p>`,
  },
  {
    id: 'report', label: 'Report', desc: 'Title, sections, summary',
    html: () => '<h1>Report title</h1><p><i>Prepared by … · ' + new Date().toLocaleDateString() + '</i></p><h2>Summary</h2><p>One paragraph overview.</p><h2>Background</h2><p>…</p><h2>Findings</h2><ul><li>Finding one</li><li>Finding two</li></ul><h2>Recommendations</h2><ol><li>…</li></ol>',
  },
  {
    id: 'meeting', label: 'Meeting notes', desc: 'Agenda, notes, actions',
    html: () => `<h1>Meeting notes</h1><p><b>Date:</b> ${new Date().toLocaleDateString()}<br><b>Attendees:</b> </p><h2>Agenda</h2><ol><li></li></ol><h2>Notes</h2><p></p><h2>Action items</h2><table><tr><td><b>Owner</b></td><td><b>Task</b></td><td><b>Due</b></td></tr><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr></table>`,
  },
  {
    id: 'resume', label: 'Resume', desc: 'CV skeleton',
    html: () => '<h1 style="text-align:center">Your Name</h1><p style="text-align:center">City · phone · email</p><h2>Experience</h2><p><b>Role</b> — Company (2022 – present)</p><ul><li>Achievement</li></ul><h2>Education</h2><p><b>Degree</b> — School (year)</p><h2>Skills</h2><p>Skill one · Skill two · Skill three</p>',
  },
  { id: 'todo', label: 'To-do list', desc: 'Checklist', html: () => '<h2>To do</h2><ul><li>☐ First task</li><li>☐ Second task</li><li>☐ Third task</li></ul>' },
];

function wordsAndChars(text: string): { words: number; chars: number; charsNoSpace: number; sentences: number; readMin: number } {
  const t = text.trim();
  const words = t ? t.split(/\s+/).length : 0;
  const sentences = t ? (t.match(/[.!?]+(\s|$)/g) ?? []).length || (t ? 1 : 0) : 0;
  return { words, chars: text.replace(/\n/g, '').length, charsNoSpace: text.replace(/\s/g, '').length, sentences, readMin: Math.max(1, Math.round(words / 200)) };
}

function mdToHtml(md: string): string {
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const inline = (s: string) => esc(s).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').replace(/\*(.+?)\*/g, '<i>$1</i>').replace(/`(.+?)`/g, '<code>$1</code>');
  const out: string[] = [];
  let list: 'ul' | 'ol' | null = null;
  const close = () => { if (list) { out.push(`</${list}>`); list = null; } };
  for (const raw of md.split('\n')) {
    const line = raw.trimEnd();
    if (!line.trim()) { close(); continue; }
    const h = /^(#{1,3})\s+(.*)$/.exec(line);
    if (h) { close(); out.push(`<h${h[1].length}>${inline(h[2])}</h${h[1].length}>`); continue; }
    const b = /^\s*[-*•]\s+(.*)$/.exec(line);
    if (b) { if (list !== 'ul') { close(); out.push('<ul>'); list = 'ul'; } out.push(`<li>${inline(b[1])}</li>`); continue; }
    const n = /^\s*\d+[.)]\s+(.*)$/.exec(line);
    if (n) { if (list !== 'ol') { close(); out.push('<ol>'); list = 'ol'; } out.push(`<li>${inline(n[1])}</li>`); continue; }
    close();
    out.push(`<p>${inline(line)}</p>`);
  }
  close();
  return out.join('');
}

// ---------------------------------------------------------------------------
// component
// ---------------------------------------------------------------------------

export default function Docs({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const docId = useRef(initialId ?? uid()).current;
  const prefs = useMemo(() => getPrefs(), []);
  const [title, setTitle] = useState(() => (initialId ? getMeta(initialId)?.title ?? 'Document' : 'Untitled document'));
  const [stats, setStats] = useState(() => wordsAndChars(''));
  const [aiOut, setAiOut] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [aiMode, setAiMode] = useState<AiMode | null>(null);
  const [aiPromptOpen, setAiPromptOpen] = useState<AiMode | null>(null);
  const [toast, flash] = useToast();
  const [rTab, setRTab] = useState<RibbonTab>('home');
  const [panel, setPanel] = useState<Panel>(null);
  const [menu, setMenu] = useState(false);
  const [fontFam, setFontFam] = useState(prefs.docFont);
  const [fontSize, setFontSize] = useState(11);
  const [block, setBlock] = useState('p');
  const [fmt, setFmt] = useState<Record<string, boolean>>({});
  const [selBar, setSelBar] = useState<{ top: number; left: number } | null>(null);
  const [page, setPage] = useState<PageSetup>(DEFAULT_PAGE);
  const [view, setView] = useState<'mobile' | 'print' | 'read'>(prefs.docView);
  const [zoom, setZoom] = useState(100);
  const [dark, setDark] = useState(false);
  const [comments, setComments] = useState<DocComment[]>([]);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [statsOpen, setStatsOpen] = useState(false);
  const [templatesOpen, setTemplatesOpen] = useState(!initialId);
  const [saved, setSaved] = useState<'saved' | 'saving' | 'dirty'>('saved');
  const [linkOpen, setLinkOpen] = useState(false);
  const [commentPrompt, setCommentPrompt] = useState(false);
  const [hfPrompt, setHfPrompt] = useState<'header' | 'footer' | null>(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [fq, setFq] = useState('');
  const [fr, setFr] = useState('');
  const [hitCount, setHitCount] = useState(0);
  const [curHit, setCurHit] = useState(0);
  const [outlineOpen, setOutlineOpen] = useState(false);
  const [ribbonHidden, setRibbonHidden] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const caseCycle = useRef(0);

  const editorRef = useRef<HTMLDivElement>(null);
  const loaded = useRef(false);
  const savedRange = useRef<Range | null>(null);
  const lastSelText = useRef('');

  // ------------------------------------------------------------------ load
  useEffect(() => {
    try { document.execCommand('styleWithCSS', false, 'true'); } catch { /* noop */ }
    if (loaded.current) return;
    loaded.current = true;
    if (initialId) {
      const d = getDoc<DocData>(initialId);
      if (d && editorRef.current) {
        editorRef.current.innerHTML = d.html;
        if (d.page) setPage({ ...DEFAULT_PAGE, ...d.page });
        if (d.comments) setComments(d.comments);
        if (d.font) setFontFam(d.font);
        recount();
      }
    }
  }, [initialId]);

  // Android back: close panels first, then exit the editor.
  useEffect(
    () =>
      onBack(() => {
        if (panel) { setPanel(null); return true; }
        if (findOpen) { closeFind(); return true; }
        if (aiOut || aiBusy) { setAiOut(''); return true; }
        if (view === 'read') { setView(prefs.docView); return true; }
        onExit?.();
        return true;
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [panel, findOpen, aiOut, aiBusy, view, onExit],
  );

  const recount = useCallback(() => {
    setStats(wordsAndChars(editorRef.current?.innerText ?? ''));
  }, []);

  const persist = useMemo(
    () =>
      debounce((t: string, pg: PageSetup, cm: DocComment[], font: string) => {
        if (!editorRef.current) return;
        putDoc<DocData>('doc', docId, t, { html: editorRef.current.innerHTML, page: pg, comments: cm, font });
        setSaved('saved');
      }, 700),
    [docId],
  );
  const save = useCallback(() => {
    setSaved('saving');
    persist(title, page, comments, fontFam);
  }, [persist, title, page, comments, fontFam]);

  useEffect(() => { if (loaded.current) save(); }, [page, comments, title]); // eslint-disable-line react-hooks/exhaustive-deps

  // ------------------------------------------------------------------ selection tracking
  useEffect(() => {
    const h = () => {
      const ed = editorRef.current;
      const sel = window.getSelection();
      try {
        const active = document.activeElement === ed;
        if (active && sel && sel.rangeCount > 0 && ed) {
          const r = sel.getRangeAt(0);
          if (ed.contains(r.commonAncestorContainer)) savedRange.current = r.cloneRange();
        }
        if (sel && !sel.isCollapsed) lastSelText.current = sel.toString();
        if (active) {
          setFmt({
            bold: document.queryCommandState('bold'),
            italic: document.queryCommandState('italic'),
            underline: document.queryCommandState('underline'),
            strikeThrough: document.queryCommandState('strikeThrough'),
            superscript: document.queryCommandState('superscript'),
            subscript: document.queryCommandState('subscript'),
            insertUnorderedList: document.queryCommandState('insertUnorderedList'),
            insertOrderedList: document.queryCommandState('insertOrderedList'),
            justifyLeft: document.queryCommandState('justifyLeft'),
            justifyCenter: document.queryCommandState('justifyCenter'),
            justifyRight: document.queryCommandState('justifyRight'),
            justifyFull: document.queryCommandState('justifyFull'),
          });
          const b = (document.queryCommandValue('formatBlock') ?? '').toLowerCase().replace(/[<>]/g, '');
          setBlock(STYLES.some((s) => s.v === b) ? b : 'p');
          const fs = document.queryCommandValue('fontSize');
          const node = sel?.anchorNode;
          const el = node ? (node.nodeType === 3 ? node.parentElement : (node as Element)) : null;
          if (el && ed?.contains(el)) {
            const px = parseFloat(getComputedStyle(el).fontSize);
            if (Number.isFinite(px)) setFontSize(Math.round((px * 72) / 96));
            const ff = getComputedStyle(el).fontFamily.split(',')[0].replace(/["']/g, '').trim();
            const known = FONTS.find((f) => f.v.toLowerCase() === ff.toLowerCase());
            if (known) setFontFam(known.v);
          } else if (fs) setFontSize(11);
        }
      } catch { /* noop */ }
      if (sel && !sel.isCollapsed && ed && (ed.contains(sel.anchorNode) || ed.contains(sel.focusNode))) {
        const r = sel.getRangeAt(0).getBoundingClientRect();
        if (r.width > 0 || r.height > 0) {
          setSelBar({ top: Math.max(8, r.top - 52), left: Math.min(window.innerWidth - 120, Math.max(120, r.left + r.width / 2)) });
          return;
        }
      }
      setSelBar(null);
    };
    document.addEventListener('selectionchange', h);
    return () => document.removeEventListener('selectionchange', h);
  }, []);

  // ------------------------------------------------------------------ commands
  const restoreSel = (): boolean => {
    const ed = editorRef.current;
    if (!ed) return false;
    ed.focus({ preventScroll: true });
    const sel = window.getSelection();
    if (!sel) return false;
    if (sel.rangeCount > 0 && ed.contains(sel.getRangeAt(0).commonAncestorContainer)) return true;
    const r = savedRange.current;
    if (!r) {
      const range = document.createRange();
      range.selectNodeContents(ed);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
      return true;
    }
    try { sel.removeAllRanges(); sel.addRange(r); return true; } catch { return false; }
  };

  const exec = (cmd: string, value?: string) => {
    try { document.execCommand('styleWithCSS', false, 'true'); } catch { /* older engines */ }
    restoreSel();
    document.execCommand(cmd, false, value);
    recount();
    setSaved('dirty');
    save();
  };

  /** Wrap the current selection in a span with inline style (font size / spacing). */
  const wrapSelection = (style: Partial<CSSStyleDeclaration>) => {
    const ed = editorRef.current;
    if (!ed) return;
    restoreSel();
    const sel = window.getSelection();
    const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : savedRange.current;
    if (!range) return;
    if (range.collapsed) {
      const span = document.createElement('span');
      Object.assign(span.style, style);
      span.appendChild(document.createTextNode('\u200b'));
      range.insertNode(span);
      const r2 = document.createRange();
      r2.setStart(span.firstChild!, 1);
      r2.collapse(true);
      sel?.removeAllRanges();
      sel?.addRange(r2);
    } else {
      const span = document.createElement('span');
      Object.assign(span.style, style);
      try {
        span.appendChild(range.extractContents());
        range.insertNode(span);
        // unwrap nested spans with the same property to avoid piling up
        span.querySelectorAll('span').forEach((inner) => {
          for (const k of Object.keys(style)) (inner.style as unknown as Record<string, string>)[k] = '';
          if (!inner.getAttribute('style')) inner.replaceWith(...Array.from(inner.childNodes));
        });
        const r2 = document.createRange();
        r2.selectNodeContents(span);
        sel?.removeAllRanges();
        sel?.addRange(r2);
        savedRange.current = r2.cloneRange();
      } catch {
        document.execCommand('fontSize', false, '4');
      }
    }
    recount();
    save();
  };

  const applyFontSize = (pt: number) => {
    const clamped = Math.max(6, Math.min(96, pt));
    setFontSize(clamped);
    wrapSelection({ fontSize: `${clamped}pt` });
  };
  const stepSize = (dir: 1 | -1) => {
    const idx = SIZES.findIndex((s) => s >= fontSize);
    const next = dir === 1 ? SIZES[Math.min(SIZES.length - 1, (idx < 0 ? SIZES.length - 1 : idx) + 1)] : SIZES[Math.max(0, (idx < 0 ? SIZES.length : idx) - 1)];
    applyFontSize(next);
  };

  const applyFont = (f: string) => {
    setFontFam(f);
    exec('fontName', f);
  };

  const applyBlock = (v: string) => {
    setBlock(v);
    exec('formatBlock', `<${v}>`);
    setPanel(null);
  };

  const changeCase = (mode: 'upper' | 'lower' | 'title' | 'sentence') => {
    restoreSel();
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed) { flash('Select some text first.'); return; }
    const t = sel.toString();
    const out =
      mode === 'upper' ? t.toUpperCase()
        : mode === 'lower' ? t.toLowerCase()
          : mode === 'title' ? t.toLowerCase().replace(/(^|\s)(\S)/g, (m, p, c: string) => p + c.toUpperCase())
            : t.toLowerCase().replace(/(^\s*\S|[.!?]\s+\S)/g, (m) => m.toUpperCase());
    document.execCommand('insertText', false, out);
    recount();
    save();
  };

  const lineSpacing = (v: number) => {
    setPage((p) => ({ ...p, lineSpacing: v }));
  };

  const paraSpacing = (v: number) => setPage((p) => ({ ...p, paraSpacing: Math.max(0, Math.min(40, v)) }));

  // ------------------------------------------------------------------ inserts
  const insertHtml = (html: string) => {
    restoreSel();
    document.execCommand('insertHTML', false, html);
    recount();
    save();
  };
  const insertAtCaret = (text: string) => {
    restoreSel();
    document.execCommand('insertText', false, text);
    recount();
    save();
  };
  const insertImage = async () => {
    try {
      const data = await pickImage(1400);
      if (data) insertHtml(`<img src="${data}" alt=""><p><br></p>`);
    } catch (e) { flash(`Could not insert image: ${errMsg(e)}`); }
  };
  const insertTable = (rows: number, cols: number, header = true) => {
    const cell = (i: number) => (header && i === 0 ? '<th>&nbsp;</th>' : '<td>&nbsp;</td>');
    const trs = Array.from({ length: rows }, (_, i) => `<tr>${cell(i).repeat(cols)}</tr>`).join('');
    insertHtml(`<table>${trs}</table><p><br></p>`);
    setPanel(null);
  };
  const tableOp = (op: 'rowBelow' | 'rowAbove' | 'colRight' | 'colLeft' | 'delRow' | 'delCol' | 'delTable') => {
    const sel = window.getSelection();
    const node = sel?.anchorNode;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as Element)) : null;
    const td = el?.closest('td,th') as HTMLTableCellElement | null;
    const tr = td?.parentElement as HTMLTableRowElement | null;
    const table = tr?.closest('table');
    if (!td || !tr || !table) { flash('Tap inside a table cell first.'); return; }
    const ci = td.cellIndex;
    const rows = Array.from(table.rows);
    if (op === 'delTable') table.remove();
    else if (op === 'delRow') { if (rows.length > 1) tr.remove(); else table.remove(); }
    else if (op === 'delCol') { rows.forEach((r) => r.cells[ci]?.remove()); if (rows[0]?.cells.length === 0) table.remove(); }
    else if (op === 'rowBelow' || op === 'rowAbove') {
      const nr = tr.cloneNode(true) as HTMLTableRowElement;
      Array.from(nr.cells).forEach((c) => { c.innerHTML = '&nbsp;'; if (c.tagName === 'TH') { const d = document.createElement('td'); d.innerHTML = '&nbsp;'; c.replaceWith(d); } });
      tr.insertAdjacentElement(op === 'rowBelow' ? 'afterend' : 'beforebegin', nr);
    } else {
      rows.forEach((r) => {
        const ref = r.cells[ci];
        const nc = document.createElement(ref?.tagName === 'TH' ? 'th' : 'td');
        nc.innerHTML = '&nbsp;';
        if (!ref) r.appendChild(nc);
        else ref.insertAdjacentElement(op === 'colRight' ? 'afterend' : 'beforebegin', nc);
      });
    }
    recount();
    save();
  };
  const insertLink = (url: string) => {
    if (!url) return;
    const href = /^(https?:|mailto:|tel:)/i.test(url) ? url : `https://${url}`;
    restoreSel();
    const sel = window.getSelection();
    if (sel && sel.isCollapsed) insertHtml(`<a href="${href}">${href}</a>`);
    else exec('createLink', href);
  };
  const insertToc = () => {
    const ed = editorRef.current;
    if (!ed) return;
    const hs = Array.from(ed.querySelectorAll('h1,h2,h3'));
    if (hs.length === 0) { flash('Add headings first (Home → Styles).'); return; }
    const items = hs.map((h, i) => {
      if (!h.id) h.id = `h-${i}-${uid()}`;
      const lvl = h.tagName === 'H1' ? 0 : h.tagName === 'H2' ? 1 : 2;
      return `<p style="margin-left:${lvl * 18}px"><a href="#${h.id}">${h.textContent}</a></p>`;
    });
    // insert at top
    ed.insertAdjacentHTML('afterbegin', `<div class="toc"><p><b>Contents</b></p>${items.join('')}</div><p><br></p>`);
    recount();
    save();
    flash('Table of contents inserted at the top.');
  };
  const insertPageBreak = () => insertHtml('<hr class="page-break"><p><br></p>');
  const insertFootnote = () => {
    const ed = editorRef.current;
    if (!ed) return;
    const n = ed.querySelectorAll('sup.fn').length + 1;
    insertHtml(`<sup class="fn">[${n}]</sup>`);
    ed.insertAdjacentHTML('beforeend', `<p class="footnote"><small>[${n}] </small></p>`);
    flash(`Footnote ${n} added at the end.`);
  };
  const insertCheckbox = () => insertHtml('☐ ');
  const insertDate = (fmtKind: 'short' | 'long' | 'time') => {
    const d = new Date();
    insertAtCaret(fmtKind === 'short' ? d.toLocaleDateString() : fmtKind === 'long' ? d.toLocaleDateString(undefined, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }) : d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }));
  };

  // ------------------------------------------------------------------ find & replace
  const clearHits = () => {
    const ed = editorRef.current;
    if (!ed) return;
    ed.querySelectorAll('mark.find-hit').forEach((m) => {
      const p = m.parentNode;
      if (p) { p.replaceChild(document.createTextNode(m.textContent ?? ''), m); p.normalize(); }
    });
  };
  const runFind = () => {
    const ed = editorRef.current;
    if (!ed) return;
    clearHits();
    setCurHit(0);
    if (!fq.trim()) { setHitCount(0); return; }
    const q = fq.toLowerCase();
    const walker = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
    const targets: Text[] = [];
    let n: Node | null;
    while ((n = walker.nextNode())) {
      const t = n as Text;
      if (t.textContent && t.textContent.toLowerCase().includes(q)) targets.push(t);
    }
    let count = 0;
    for (const t of targets) {
      const text = t.textContent ?? '';
      const lower = text.toLowerCase();
      const frag = document.createDocumentFragment();
      let i = 0;
      for (;;) {
        const idx = lower.indexOf(q, i);
        if (idx < 0) { if (i < text.length) frag.appendChild(document.createTextNode(text.slice(i))); break; }
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
      const first = ed.querySelector('mark.find-hit');
      first?.classList.add('cur');
      first?.scrollIntoView({ block: 'center', behavior: 'smooth' });
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
    const target = ed.querySelector('mark.find-hit.cur') ?? ed.querySelector('mark.find-hit');
    if (!target) return;
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
    setHitCount(0);
    setCurHit(0);
    recount();
    save();
    flash(`Replaced ${ms.length} occurrence(s).`);
  };
  const closeFind = () => { clearHits(); setFindOpen(false); setHitCount(0); setCurHit(0); };

  // ------------------------------------------------------------------ comments
  const addComment = (text: string) => {
    const quote = lastSelText.current.trim().slice(0, 80) || '(document)';
    const c: DocComment = { id: uid(), quote, text, created: Date.now() };
    setComments((cs) => [...cs, c]);
    restoreSel();
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed) {
      const mark = document.createElement('mark');
      mark.className = 'cmt';
      mark.dataset.cid = c.id;
      try { mark.appendChild(sel.getRangeAt(0).extractContents()); sel.getRangeAt(0).insertNode(mark); } catch { /* skip highlight */ }
    }
    setCommentsOpen(true);
    save();
  };
  const resolveComment = (id: string) => {
    setComments((cs) => cs.map((c) => (c.id === id ? { ...c, resolved: !c.resolved } : c)));
  };
  const deleteComment = (id: string) => {
    editorRef.current?.querySelectorAll(`mark.cmt[data-cid="${id}"]`).forEach((m) => m.replaceWith(...Array.from(m.childNodes)));
    setComments((cs) => cs.filter((c) => c.id !== id));
  };
  const jumpToComment = (id: string) => {
    const m = editorRef.current?.querySelector(`mark.cmt[data-cid="${id}"]`);
    m?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    setCommentsOpen(false);
  };

  // ------------------------------------------------------------------ AI
  const insertAtEnd = (html: string) => {
    const el = editorRef.current;
    if (!el) return;
    el.insertAdjacentHTML('beforeend', html);
    el.lastElementChild?.scrollIntoView({ block: 'end', behavior: 'smooth' });
    recount();
    save();
  };

  const runAi = async (mode: AiMode, extra?: string) => {
    const s = getSettings();
    if (!s.apiKey) { flash('Add your API key in Settings first.'); return; }
    const conf = AI_PROMPTS[mode];
    const full = (editorRef.current?.innerText ?? '').trim();
    let body = '';
    if (conf.scope === 'prompt') {
      if (mode === 'translate') body = `${extra}\n\n${(lastSelText.current || full).slice(0, 6000)}`;
      else body = extra ?? '';
      if (!body.trim()) return;
    } else {
      if (!full) { flash('Write something first.'); return; }
      const selText = lastSelText.current;
      body = mode === 'continue' ? full.slice(-4000) : (selText || full).slice(0, 6000);
    }
    setAiMode(mode);
    setAiBusy(true);
    setAiOut('');
    try {
      let acc = '';
      const out = await chatStream(s, [{ role: 'system', content: conf.sys }, { role: 'user', content: body }], {
        onDelta: (d) => { acc += d; setAiOut(acc); },
      });
      setAiOut(out);
      if (mode === 'continue') {
        insertAtEnd(mdToHtml(out));
        setAiOut('');
        flash('Continuation added at the end.');
      } else if (mode === 'outline' || mode === 'draft') {
        const ed = editorRef.current;
        if (ed && !full) { ed.innerHTML = mdToHtml(out); recount(); save(); setAiOut(''); flash('Document created.'); }
      }
    } catch (e) {
      setAiOut(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  const replaceSelectionWith = (text: string) => {
    restoreSel();
    const sel = window.getSelection();
    const html = mdToHtml(text);
    if (sel && !sel.isCollapsed) document.execCommand('insertHTML', false, html);
    else insertAtEnd(html);
    setAiOut('');
    recount();
    save();
    flash('Applied.');
  };

  // ------------------------------------------------------------------ files
  const openFile = async () => {
    const pick = await openFilePicker('.docx,.txt,.md,.html,.htm');
    if (!pick) return;
    try {
      const ext = (pick.name.split('.').pop() ?? '').toLowerCase();
      let html = '';
      if (ext === 'docx') html = await importDocx(pick.buf);
      else if (ext === 'html' || ext === 'htm') html = new DOMParser().parseFromString(new TextDecoder().decode(pick.buf), 'text/html').body.innerHTML;
      else if (ext === 'md') html = mdToHtml(new TextDecoder().decode(pick.buf));
      else html = textToHtml(new TextDecoder().decode(pick.buf));
      if (editorRef.current) editorRef.current.innerHTML = html;
      setTitle(pick.name.replace(/\.[^.]+$/, ''));
      recount();
      save();
      flash(`Opened ${pick.name}`);
    } catch (e) { flash(`Could not open: ${errMsg(e)}`); }
  };
  const saveDocx = async () => {
    try {
      const blob = await exportDocx(title, editorRef.current?.innerHTML ?? '', { page, header: page.header, footer: page.footer, pageNumbers: page.pageNumbers, font: fontFam, lineSpacing: page.lineSpacing });
      flash(await saveBinary(sanitizeName(title, 'docx'), blob, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'));
    } catch (e) { flash(`Save failed: ${errMsg(e)}`); }
  };
  const savePdf = async () => {
    try {
      flash('Rendering PDF…', 8000);
      const bytes = await exportPdfFromHtml(title, editorRef.current?.innerHTML ?? '', { page, header: page.header, footer: page.footer, pageNumbers: page.pageNumbers, font: fontFam, lineSpacing: page.lineSpacing });
      flash(await saveBinary(sanitizeName(title, 'pdf'), bytes, 'application/pdf'));
    } catch (e) { flash(`PDF failed: ${errMsg(e)}`); }
  };
  const saveHtml = () => {
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:${fontFam},sans-serif">${editorRef.current?.innerHTML ?? ''}</body></html>`;
    void downloadText(`${sanitizeName(title, 'html')}`, html, 'text/html').then(flash);
  };
  const saveMd = () => void downloadText(sanitizeName(title, 'md'), htmlToMarkdown(editorRef.current?.innerHTML ?? ''), 'text/markdown').then(flash);
  const saveTxt = () => void downloadText(sanitizeName(title, 'txt'), editorRef.current?.innerText ?? '', 'text/plain').then(flash);
  const share = () => void shareText(title, editorRef.current?.innerText ?? '').then(flash);
  const printDoc = () => {
    const w = window.open('', '_blank');
    if (!w) { flash('Pop-ups blocked. Use Save as PDF instead.'); return; }
    w.document.write(`<!doctype html><html><head><title>${title}</title><style>body{font-family:${fontFam},sans-serif;line-height:${page.lineSpacing};padding:${MARGIN_IN[page.margins]}in;} table{border-collapse:collapse} td,th{border:1px solid #999;padding:4px 8px} hr.page-break{page-break-after:always;border:0}</style></head><body>${editorRef.current?.innerHTML ?? ''}</body></html>`);
    w.document.close();
    w.focus();
    setTimeout(() => w.print(), 300);
  };
  const applyTemplate = (id: string) => {
    const t = TEMPLATES.find((x) => x.id === id);
    if (!t || !editorRef.current) return;
    editorRef.current.innerHTML = t.html();
    setTemplatesOpen(false);
    recount();
    save();
    setTimeout(() => editorRef.current?.focus(), 50);
  };
  const clearDoc = () => {
    if (editorRef.current) editorRef.current.innerHTML = '<p><br></p>';
    setComments([]);
    recount();
    save();
  };

  // ------------------------------------------------------------------ outline (navigation pane)
  const headings = useMemo(() => {
    if (!outlineOpen) return [] as { text: string; lvl: number; el: Element }[];
    return Array.from(editorRef.current?.querySelectorAll('h1,h2,h3') ?? []).map((el) => ({ text: el.textContent ?? '', lvl: parseInt(el.tagName[1], 10), el }));
  }, [outlineOpen]);

  // ------------------------------------------------------------------ derived styles
  const pageDims = PAGE_SIZES[page.size];
  const pw = page.orientation === 'portrait' ? pageDims.w : pageDims.h;
  const ph = page.orientation === 'portrait' ? pageDims.h : pageDims.w;
  const margin = MARGIN_IN[page.margins];
  const paperStyle: React.CSSProperties =
    view === 'print'
      ? { width: `${pw}in`, minHeight: `${ph}in`, padding: `${margin}in`, fontFamily: fontFam, lineHeight: page.lineSpacing, ['--para-gap' as string]: `${page.paraSpacing ?? 8}px`, columnCount: page.columns === 2 ? 2 : undefined, columnGap: page.columns === 2 ? '0.4in' : undefined, zoom: zoom / 100 }
      : { fontFamily: fontFam, lineHeight: page.lineSpacing, ['--para-gap' as string]: `${page.paraSpacing ?? 8}px`, fontSize: `${(15.5 * zoom) / 100}px`, columnCount: page.columns === 2 ? 2 : undefined, columnGap: page.columns === 2 ? '18px' : undefined };

  const openComments = comments.filter((c) => !c.resolved).length;

  return (
    <div className={`edscreen${dark ? ' dark-canvas' : ''}${view === 'read' ? ' reading' : ''}`} style={{ ['--app' as string]: 'var(--word)' }}>
      {view !== 'read' && (
        <AppBar kindIcon={<FileTypeIcon kind="doc" size={24} light />} title={title} onTitle={(t) => { setTitle(t); setSaved('dirty'); }} placeholder="Document title" onBack={onExit} saved={saved}>
          <button className="icon-btn light" aria-label="Undo" onPointerDown={(e) => e.preventDefault()} onClick={() => exec('undo')}>
            <Icon name="undo" size={20} />
          </button>
          <button className="icon-btn light" aria-label="Save as .docx" onClick={() => void saveDocx()}>
            <Icon name="save" size={20} />
          </button>
          <button className="icon-btn light" aria-label="More actions" onClick={() => setMenu(true)}>
            <Icon name="more" size={20} />
          </button>
        </AppBar>
      )}

      {view === 'read' && (
        <header className="appbar readbar">
          <button className="icon-btn light" aria-label="Exit reading view" onClick={() => setView(prefs.docView)}>
            <Icon name="arrowLeft" size={22} />
          </button>
          <span className="appbar-title" style={{ padding: '8px 4px' }}>{title}</span>
          <button className="icon-btn light" aria-label="Smaller text" onClick={() => setZoom((z) => Math.max(70, z - 10))}><Icon name="zoomOut" size={20} /></button>
          <button className="icon-btn light" aria-label="Larger text" onClick={() => setZoom((z) => Math.min(200, z + 10))}><Icon name="zoomIn" size={20} /></button>
          <button className={`icon-btn light${dark ? ' on' : ''}`} aria-label="Toggle dark page" onClick={() => setDark(!dark)}><Icon name={dark ? 'sun' : 'moon'} size={20} /></button>
        </header>
      )}

      {findOpen && (
        <div className="findbar">
          <div className="find-row">
            <Icon name="search" size={16} className="dim" />
            <input className="input find-input" value={fq} placeholder="Find" onChange={(e) => setFq(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && runFind()} />
            <span className="find-count">{hitCount ? `${curHit}/${hitCount}` : fq ? '0' : ''}</span>
            <button className="icon-btn" disabled={!hitCount} onClick={() => gotoHit(-1)} aria-label="Previous match"><Icon name="chevronUp" size={18} /></button>
            <button className="icon-btn" disabled={!hitCount} onClick={() => gotoHit(1)} aria-label="Next match"><Icon name="chevronDown" size={18} /></button>
            <button className="icon-btn" onClick={closeFind} aria-label="Close find"><Icon name="close" size={18} /></button>
          </div>
          <div className="find-row">
            <Icon name="replace" size={16} className="dim" />
            <input className="input find-input" value={fr} placeholder="Replace with" onChange={(e) => setFr(e.target.value)} />
            <button className="btn small" onClick={runFind}>Find</button>
            <button className="btn small" disabled={!hitCount} onClick={replaceOne}>Replace</button>
            <button className="btn small" disabled={!hitCount} onClick={replaceAll}>All</button>
          </div>
        </div>
      )}

      <div className={`edbody${view === 'print' ? ' print-view' : ''}`} onClick={(e) => { if (e.target === e.currentTarget) editorRef.current?.focus(); }}>
        <div className="paper-wrap">
          {view === 'print' && (page.header || page.pageNumbers) && (
            <div className="paper-hf top" style={{ width: `${pw}in`, zoom: zoom / 100 }}>{page.header}</div>
          )}
          <div
            className={`paper${view === 'print' ? ' print' : ''}${dark ? ' dark' : ''}`}
            ref={editorRef}
            contentEditable={view !== 'read'}
            suppressContentEditableWarning
            spellCheck
            style={paperStyle}
            onInput={() => { recount(); setSaved('dirty'); save(); }}
            onPaste={(e) => {
              // paste as clean text unless it's rich HTML from another editor
              const html = e.clipboardData.getData('text/html');
              const text = e.clipboardData.getData('text/plain');
              if (!html || html.length > 200000) {
                e.preventDefault();
                document.execCommand('insertText', false, text);
              }
            }}
            data-placeholder="Start writing, or pick a template from ⋮ → Templates…"
          />
          {view === 'print' && (page.footer || page.pageNumbers) && (
            <div className="paper-hf bottom" style={{ width: `${pw}in`, zoom: zoom / 100 }}>
              <span>{page.footer}</span>
              {page.pageNumbers && <span>Page 1</span>}
            </div>
          )}
          <div className="ed-status">
            <button className="status-btn" onClick={() => setStatsOpen(true)}>
              {stats.words} words · {stats.chars} characters
            </button>
            <span>{page.size} · {page.orientation} · {zoom}%</span>
          </div>
        </div>
      </div>

      {/* floating selection toolbar (Word mobile "Aa" bar) */}
      {selBar && view !== 'read' && (
        <div className="selbar" style={{ top: selBar.top, left: selBar.left }}>
          <button className={`icon-btn${fmt.bold ? ' on' : ''}`} aria-label="Bold" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('bold')}><b>B</b></button>
          <button className={`icon-btn${fmt.italic ? ' on' : ''}`} aria-label="Italic" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('italic')}><i>I</i></button>
          <button className={`icon-btn${fmt.underline ? ' on' : ''}`} aria-label="Underline" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('underline')}><u>U</u></button>
          <button className="icon-btn" aria-label="Highlight" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => exec('hiliteColor', '#FFFF00')}><Icon name="highlight" size={17} /></button>
          <button className="icon-btn" aria-label="Comment" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => setCommentPrompt(true)}><Icon name="comment" size={17} /></button>
          <button className="icon-btn aa" aria-label="AI rewrite" onPointerDown={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()} onClick={() => { setRTab('review'); setSelBar(null); }}><Icon name="ai" size={17} /></button>
        </div>
      )}

      {(aiBusy || aiOut) && (
        <div className="ai-sheet">
          <div className="ai-sheet-head">
            <Icon name="ai" size={14} /> {aiMode ? AI_PROMPTS[aiMode].label : 'AI'}
            {aiBusy && <span className="ai-dot" />}
            <button className="icon-btn" aria-label="Dismiss" onClick={() => { setAiOut(''); setAiBusy(false); }}>
              <Icon name="close" size={16} />
            </button>
          </div>
          <div className="ai-sheet-body">{aiOut || 'Thinking…'}</div>
          {!aiBusy && aiOut && !aiOut.startsWith('Error:') && (
            <div className="btn-row">
              <button className="btn small primary" onClick={() => replaceSelectionWith(aiOut)}>
                {lastSelText.current ? 'Replace selection' : 'Insert'}
              </button>
              <button className="btn small" onClick={() => { insertAtEnd(mdToHtml(aiOut)); setAiOut(''); }}>Add at end</button>
              <button className="btn small" onClick={() => void navigator.clipboard?.writeText(aiOut).then(() => flash('Copied.'))}>Copy</button>
              <button className="btn small" onClick={() => aiMode && void runAi(aiMode)}>Retry</button>
            </div>
          )}
        </div>
      )}

      {view !== 'read' && !ribbonHidden && (
        <div className="ribbon">
          {panel === 'textColor' && <Palette onPick={(c) => { exec('foreColor', c); setPanel(null); }} auto={() => { exec('foreColor', dark ? '#f2f2f2' : '#1b1b1b'); setPanel(null); }} autoLabel="Automatic" />}
          {panel === 'hilite' && <Palette onPick={(c) => { exec('hiliteColor', c); setPanel(null); }} auto={() => { exec('hiliteColor', 'transparent'); setPanel(null); }} autoLabel="No highlight" />}
          {panel === 'styles' && (
            <RibbonPanel title="Styles" onClose={() => setPanel(null)}>
              <div className="style-gallery">
                {STYLES.map((s) => (
                  <button key={s.v} className={`style-card ${s.cls}${block === s.v ? ' on' : ''}`} onPointerDown={(e) => e.preventDefault()} onClick={() => applyBlock(s.v)}>
                    {s.preview}
                  </button>
                ))}
              </div>
            </RibbonPanel>
          )}
          {panel === 'table' && (
            <RibbonPanel title="Insert table" onClose={() => setPanel(null)}>
              <TableGrid onPick={(r, c) => insertTable(r, c)} />
              <div className="rpanel-actions">
                <span className="hint">In a table:</span>
                <button className="btn small" onPointerDown={(e) => e.preventDefault()} onClick={() => tableOp('rowBelow')}>+ Row</button>
                <button className="btn small" onPointerDown={(e) => e.preventDefault()} onClick={() => tableOp('colRight')}>+ Column</button>
                <button className="btn small" onPointerDown={(e) => e.preventDefault()} onClick={() => tableOp('delRow')}>− Row</button>
                <button className="btn small" onPointerDown={(e) => e.preventDefault()} onClick={() => tableOp('delCol')}>− Column</button>
                <button className="btn small danger" onPointerDown={(e) => e.preventDefault()} onClick={() => tableOp('delTable')}>Delete table</button>
              </div>
            </RibbonPanel>
          )}
          {panel === 'spacing' && (
            <RibbonPanel title="Line & paragraph spacing" onClose={() => setPanel(null)}>
              <div className="rpanel-row">
                <span className="hint">Line spacing</span>
                <RSeg value={String(page.lineSpacing ?? 1.15)} options={[{ v: '1', t: '1.0' }, { v: '1.15', t: '1.15' }, { v: '1.5', t: '1.5' }, { v: '2', t: '2.0' }]} onChange={(v) => lineSpacing(parseFloat(v))} />
              </div>
              <div className="rpanel-row">
                <span className="hint">Space between paragraphs</span>
                <RStepper value={`${page.paraSpacing ?? 8}px`} title="Paragraph spacing" onDec={() => paraSpacing((page.paraSpacing ?? 8) - 2)} onInc={() => paraSpacing((page.paraSpacing ?? 8) + 2)} width={44} />
              </div>
            </RibbonPanel>
          )}
          {panel === 'symbols' && (
            <RibbonPanel title="Symbols" onClose={() => setPanel(null)}>
              <div className="symbol-grid">
                {SYMBOLS.map((s) => (
                  <button key={s} className="symbol" onPointerDown={(e) => e.preventDefault()} onClick={() => insertAtCaret(s)}>{s}</button>
                ))}
              </div>
            </RibbonPanel>
          )}

          <RibbonTabs
            tabs={[{ id: 'home', label: 'Home' }, { id: 'insert', label: 'Insert' }, { id: 'layout', label: 'Layout' }, { id: 'review', label: 'Review' }, { id: 'view', label: 'View' }]}
            value={rTab}
            onChange={(t) => { setRTab(t); setPanel(null); void tap(); }}
          />

          {rTab === 'home' && (
            <div className="ribbon-row">
              <RGroup label="Clipboard">
                <RBtn icon="undo" label="Undo" keepFocus onRun={() => exec('undo')} />
                <RBtn icon="redo" label="Redo" keepFocus onRun={() => exec('redo')} />
                <RBtn icon="cut" label="Cut" keepFocus onRun={() => exec('cut')} />
                <RBtn icon="copy" label="Copy" keepFocus onRun={() => exec('copy')} />
                <RBtn icon="paste" label="Paste" keepFocus onRun={() => { restoreSel(); navigator.clipboard?.readText().then((t) => t && document.execCommand('insertText', false, t)).catch(() => flash('Long-press in the text to paste.')); }} />
              </RGroup>
              <RGroup label="Font">
                <RSelect value={fontFam} options={FONTS} onChange={applyFont} width={112} title="Font" keepFocus />
                <RStepper value={String(fontSize)} title="Font size" keepFocus onDec={() => stepSize(-1)} onInc={() => stepSize(1)} />
                <RBtn icon="bold" label="Bold" active={fmt.bold} keepFocus onRun={() => exec('bold')} />
                <RBtn icon="italic" label="Italic" active={fmt.italic} keepFocus onRun={() => exec('italic')} />
                <RBtn icon="underline" label="Underline" active={fmt.underline} keepFocus onRun={() => exec('underline')} />
                <RBtn icon="strike" label="Strike" active={fmt.strikeThrough} keepFocus onRun={() => exec('strikeThrough')} />
                <RBtn icon="superscript" label="Super" active={fmt.superscript} keepFocus onRun={() => exec('superscript')} />
                <RBtn icon="subscript" label="Sub" active={fmt.subscript} keepFocus onRun={() => exec('subscript')} />
                <RBtn icon="fontColor" label="Color" colorBar="#C00000" keepFocus menu active={panel === 'textColor'} onRun={() => setPanel(panel === 'textColor' ? null : 'textColor')} />
                <RBtn icon="highlight" label="Highlight" colorBar="#FFFF00" keepFocus menu active={panel === 'hilite'} onRun={() => setPanel(panel === 'hilite' ? null : 'hilite')} />
                <RBtn icon="caseChange" label="Case" keepFocus onRun={() => { const modes = ['upper', 'lower', 'title', 'sentence'] as const; changeCase(modes[caseCycle.current++ % modes.length]); }} />
                <RBtn icon="clearFormat" label="Clear" keepFocus onRun={() => { exec('removeFormat'); exec('formatBlock', '<p>'); }} />
              </RGroup>
              <RGroup label="Paragraph">
                <RBtn icon="listBullet" label="Bullets" active={fmt.insertUnorderedList} keepFocus onRun={() => exec('insertUnorderedList')} />
                <RBtn icon="listOrdered" label="Numbering" active={fmt.insertOrderedList} keepFocus onRun={() => exec('insertOrderedList')} />
                <RBtn icon="outdent" label="Outdent" keepFocus onRun={() => exec('outdent')} />
                <RBtn icon="indent" label="Indent" keepFocus onRun={() => exec('indent')} />
                <RBtn icon="alignLeft" label="Left" active={fmt.justifyLeft} keepFocus onRun={() => exec('justifyLeft')} />
                <RBtn icon="alignCenter" label="Center" active={fmt.justifyCenter} keepFocus onRun={() => exec('justifyCenter')} />
                <RBtn icon="alignRight" label="Right" active={fmt.justifyRight} keepFocus onRun={() => exec('justifyRight')} />
                <RBtn icon="alignJustify" label="Justify" active={fmt.justifyFull} keepFocus onRun={() => exec('justifyFull')} />
                <RBtn icon="lineSpacing" label="Spacing" keepFocus menu active={panel === 'spacing'} onRun={() => setPanel(panel === 'spacing' ? null : 'spacing')} />
              </RGroup>
              <RGroup label="Styles">
                <button className="style-current" onPointerDown={(e) => e.preventDefault()} onClick={() => setPanel(panel === 'styles' ? null : 'styles')}>
                  <span className={STYLES.find((s) => s.v === block)?.cls}>{STYLES.find((s) => s.v === block)?.t ?? 'Normal'}</span>
                  <Icon name="chevronDown" size={14} />
                </button>
              </RGroup>
              <RGroup label="Editing">
                <RBtn icon="search" label="Find" active={findOpen} onRun={() => (findOpen ? closeFind() : setFindOpen(true))} />
                <RBtn icon="ai" label="AI" onRun={() => setRTab('review')} />
              </RGroup>
            </div>
          )}

          {rTab === 'insert' && (
            <div className="ribbon-row">
              <RGroup label="Tables">
                <RBtn icon="table" label="Table" keepFocus menu active={panel === 'table'} onRun={() => setPanel(panel === 'table' ? null : 'table')} />
              </RGroup>
              <RGroup label="Media">
                <RBtn icon="image" label="Picture" keepFocus onRun={() => void insertImage()} />
                <RBtn icon="link" label="Link" keepFocus onRun={() => setLinkOpen(true)} />
                <RBtn icon="symbol" label="Symbol" keepFocus menu active={panel === 'symbols'} onRun={() => setPanel(panel === 'symbols' ? null : 'symbols')} />
              </RGroup>
              <RGroup label="Pages">
                <RBtn icon="pageBreak" label="Page break" keepFocus onRun={insertPageBreak} />
                <RBtn icon="hr" label="Divider" keepFocus onRun={() => insertHtml('<hr><p><br></p>')} />
                <RBtn icon="toc" label="Contents" keepFocus onRun={insertToc} />
                <RBtn icon="bookmark" label="Footnote" keepFocus onRun={insertFootnote} />
              </RGroup>
              <RGroup label="Text">
                <RBtn icon="quote" label="Quote" keepFocus onRun={() => exec('formatBlock', '<blockquote>')} />
                <RBtn icon="check" label="Checkbox" keepFocus onRun={insertCheckbox} />
                <RBtn icon="calendar" label="Date" keepFocus onRun={() => insertDate('short')} />
                <RBtn icon="clock" label="Time" keepFocus onRun={() => insertDate('time')} />
                <RBtn icon="fileText" label="Long date" keepFocus onRun={() => insertDate('long')} />
              </RGroup>
            </div>
          )}

          {rTab === 'layout' && (
            <div className="ribbon-row">
              <RGroup label="Page setup">
                <RSelect value={page.size} options={[{ v: 'A4', t: 'A4' }, { v: 'Letter', t: 'Letter' }, { v: 'Legal', t: 'Legal' }]} onChange={(v) => setPage({ ...page, size: v as PageSetup['size'] })} width={78} title="Paper size" />
                <RSeg value={page.orientation} options={[{ v: 'portrait', t: 'Portrait', icon: 'portrait' }, { v: 'landscape', t: 'Landscape', icon: 'landscape' }]} onChange={(v) => setPage({ ...page, orientation: v })} />
              </RGroup>
              <RGroup label="Margins">
                <RSeg value={page.margins} options={[{ v: 'narrow', t: 'Narrow' }, { v: 'normal', t: 'Normal' }, { v: 'wide', t: 'Wide' }]} onChange={(v) => setPage({ ...page, margins: v })} />
              </RGroup>
              <RGroup label="Columns">
                <RSeg value={String(page.columns ?? 1)} options={[{ v: '1', t: 'One' }, { v: '2', t: 'Two' }]} onChange={(v) => setPage({ ...page, columns: v === '2' ? 2 : 1 })} />
              </RGroup>
              <RGroup label="Header & footer">
                <RBtn icon="header" label="Header" active={!!page.header} onRun={() => setHfPrompt('header')} />
                <RBtn icon="footer" label="Footer" active={!!page.footer} onRun={() => setHfPrompt('footer')} />
                <RBtn icon="wordCount" label="Page #" active={!!page.pageNumbers} onRun={() => setPage({ ...page, pageNumbers: !page.pageNumbers })} />
              </RGroup>
              <RGroup label="Spacing">
                <RBtn icon="lineSpacing" label="Line" menu active={panel === 'spacing'} onRun={() => setPanel(panel === 'spacing' ? null : 'spacing')} />
              </RGroup>
            </div>
          )}

          {rTab === 'review' && (
            <div className="ribbon-row">
              <RGroup label="AI writing">
                {(['continue', 'rewrite', 'summarize', 'expand', 'shorten', 'fixGrammar'] as AiMode[]).map((m) => (
                  <RBtn key={m} icon={AI_PROMPTS[m].icon} label={AI_PROMPTS[m].label.replace('Continue writing', 'Continue')} disabled={aiBusy} keepFocus onRun={() => void runAi(m)} />
                ))}
              </RGroup>
              <RGroup label="Tone">
                <RBtn icon="readMode" label="Formal" disabled={aiBusy} keepFocus onRun={() => void runAi('formal')} />
                <RBtn icon="chat" label="Casual" disabled={aiBusy} keepFocus onRun={() => void runAi('casual')} />
                <RBtn icon="symbol" label="Translate" disabled={aiBusy} keepFocus onRun={() => setAiPromptOpen('translate')} />
              </RGroup>
              <RGroup label="Create">
                <RBtn icon="toc" label="Outline" disabled={aiBusy} onRun={() => setAiPromptOpen('outline')} />
                <RBtn icon="blankDoc" label="Draft" disabled={aiBusy} onRun={() => setAiPromptOpen('draft')} />
              </RGroup>
              <RGroup label="Comments">
                <RBtn icon="comment" label="New" keepFocus onRun={() => setCommentPrompt(true)} />
                <RBtn icon="listBullet" label={openComments ? `Show (${openComments})` : 'Show'} active={commentsOpen} onRun={() => setCommentsOpen(true)} />
              </RGroup>
              <RGroup label="Proofing">
                <RBtn icon="wordCount" label="Word count" onRun={() => setStatsOpen(true)} />
                <RBtn icon="search" label="Find" active={findOpen} onRun={() => (findOpen ? closeFind() : setFindOpen(true))} />
                <RBtn icon="speaker" label="Read aloud" onRun={() => readAloud(lastSelText.current || editorRef.current?.innerText || '', flash)} />
              </RGroup>
            </div>
          )}

          {rTab === 'view' && (
            <div className="ribbon-row">
              <RGroup label="Views">
                <RSeg value={view} options={[{ v: 'mobile', t: 'Mobile', icon: 'phone' }, { v: 'print', t: 'Print layout', icon: 'pageSize' }, { v: 'read', t: 'Read', icon: 'readMode' }]} onChange={(v) => setView(v)} />
              </RGroup>
              <RGroup label="Zoom">
                <RStepper value={`${zoom}%`} title="Zoom" width={44} onDec={() => setZoom((z) => Math.max(50, z - 10))} onInc={() => setZoom((z) => Math.min(200, z + 10))} />
                <RBtn icon="fitWidth" label="100%" onRun={() => setZoom(100)} />
              </RGroup>
              <RGroup label="Show">
                <RBtn icon="toc" label="Outline" active={outlineOpen} onRun={() => setOutlineOpen(true)} />
                <RBtn icon={dark ? 'sun' : 'moon'} label={dark ? 'Light page' : 'Dark page'} active={dark} onRun={() => setDark(!dark)} />
                <RBtn icon="fullscreen" label="Focus" onRun={() => { setRibbonHidden(true); flash('Tap the ✎ button to bring the toolbar back.'); }} />
              </RGroup>
            </div>
          )}
        </div>
      )}

      {ribbonHidden && view !== 'read' && (
        <button className="fab" aria-label="Show toolbar" onClick={() => setRibbonHidden(false)}>
          <Icon name="edit" size={22} />
        </button>
      )}

      {/* ---------------- sheets & dialogs ---------------- */}
      <BottomSheet open={menu} onClose={() => setMenu(false)} title={title}>
        <SheetMenu
          onClose={() => setMenu(false)}
          items={[
            { icon: 'fileOpen', label: 'Open file', hint: '.docx · .txt · .md · .html', onRun: () => void openFile() },
            { icon: 'template', label: 'Templates', hint: 'Letter, report, notes, resume…', onRun: () => setTemplatesOpen(true) },
            'divider',
            { icon: 'save', label: 'Save as Word (.docx)', onRun: () => void saveDocx() },
            { icon: 'pdf', label: 'Save as PDF', onRun: () => void savePdf() },
            { icon: 'export', label: 'Export…', hint: 'HTML · Markdown · plain text', onRun: () => setExportOpen(true) },
            { icon: 'print', label: 'Print', onRun: printDoc },
            { icon: 'share', label: 'Share text', onRun: share },
            'divider',
            { icon: 'wordCount', label: 'Document statistics', onRun: () => setStatsOpen(true) },
            { icon: 'trash', label: 'Clear document', danger: true, onRun: () => setConfirmClear(true) },
          ]}
        />
      </BottomSheet>

      <BottomSheet open={exportOpen} onClose={() => setExportOpen(false)} title="Export">
        <SheetMenu
          onClose={() => setExportOpen(false)}
          items={[
            { icon: 'fileText', label: 'HTML (.html)', onRun: saveHtml },
            { icon: 'fileText', label: 'Markdown (.md)', onRun: saveMd },
            { icon: 'fileText', label: 'Plain text (.txt)', onRun: saveTxt },
          ]}
        />
      </BottomSheet>

      <BottomSheet open={templatesOpen} onClose={() => setTemplatesOpen(false)} title="New document">
        <div className="template-grid">
          {TEMPLATES.map((t) => (
            <button key={t.id} className="template-card" onClick={() => applyTemplate(t.id)}>
              <span className="template-thumb" data-kind={t.id}>
                <i /><i /><i />
              </span>
              <strong>{t.label}</strong>
              <small>{t.desc}</small>
            </button>
          ))}
          <button className="template-card ai" onClick={() => { setTemplatesOpen(false); setAiPromptOpen('draft'); }}>
            <span className="template-thumb ai"><Icon name="ai" size={22} /></span>
            <strong>AI draft</strong>
            <small>Describe it, get a document</small>
          </button>
        </div>
      </BottomSheet>

      <BottomSheet open={statsOpen} onClose={() => setStatsOpen(false)} title="Document statistics">
        <div className="stats-grid">
          <div><b>{stats.words}</b><span>Words</span></div>
          <div><b>{stats.chars}</b><span>Characters</span></div>
          <div><b>{stats.charsNoSpace}</b><span>No spaces</span></div>
          <div><b>{stats.sentences}</b><span>Sentences</span></div>
          <div><b>{editorRef.current?.querySelectorAll('p,h1,h2,h3,li').length ?? 0}</b><span>Paragraphs</span></div>
          <div><b>{stats.readMin} min</b><span>Reading time</span></div>
        </div>
      </BottomSheet>

      <BottomSheet open={commentsOpen} onClose={() => setCommentsOpen(false)} title={`Comments (${comments.length})`} tall>
        {comments.length === 0 && <p className="empty">No comments yet. Select text and tap the comment icon.</p>}
        <div className="comment-list">
          {comments.map((c) => (
            <div key={c.id} className={`comment${c.resolved ? ' resolved' : ''}`}>
              <button className="comment-quote" onClick={() => jumpToComment(c.id)}>“{c.quote}”</button>
              <p>{c.text}</p>
              <div className="comment-meta">
                <span>{new Date(c.created).toLocaleString()}</span>
                <button className="btn small" onClick={() => resolveComment(c.id)}>{c.resolved ? 'Reopen' : 'Resolve'}</button>
                <button className="btn small danger" onClick={() => deleteComment(c.id)}>Delete</button>
              </div>
            </div>
          ))}
        </div>
        <div className="btn-row end">
          <button className="btn primary" onClick={() => { setCommentsOpen(false); setCommentPrompt(true); }}>New comment</button>
        </div>
      </BottomSheet>

      <BottomSheet open={outlineOpen} onClose={() => setOutlineOpen(false)} title="Outline" tall>
        {headings.length === 0 && <p className="empty">No headings yet. Use Home → Styles → Heading 1/2/3.</p>}
        <div className="outline-list">
          {headings.map((h, i) => (
            <button key={i} className={`outline-item lvl${h.lvl}`} onClick={() => { h.el.scrollIntoView({ block: 'start', behavior: 'smooth' }); setOutlineOpen(false); }}>
              {h.text || '(empty heading)'}
            </button>
          ))}
        </div>
      </BottomSheet>

      <PromptSheet open={linkOpen} title="Insert link" label="URL" placeholder="https://example.com" initial="https://" confirmLabel="Insert" onSubmit={insertLink} onClose={() => setLinkOpen(false)} />
      <PromptSheet open={commentPrompt} title="New comment" label={lastSelText.current ? `On: “${lastSelText.current.slice(0, 60)}”` : 'Comment on the document'} placeholder="Write your comment" confirmLabel="Add" multiline onSubmit={addComment} onClose={() => setCommentPrompt(false)} />
      <PromptSheet
        open={hfPrompt !== null}
        title={hfPrompt === 'header' ? 'Header text' : 'Footer text'}
        label="Shown on every page of the exported .docx / PDF"
        initial={(hfPrompt === 'header' ? page.header : page.footer) ?? ''}
        placeholder={hfPrompt === 'header' ? 'e.g. Company · Confidential' : 'e.g. Prepared by …'}
        confirmLabel="Apply"
        onSubmit={(v) => setPage({ ...page, [hfPrompt === 'header' ? 'header' : 'footer']: v || undefined })}
        onClose={() => setHfPrompt(null)}
      />
      <PromptSheet
        open={aiPromptOpen !== null}
        title={aiPromptOpen ? AI_PROMPTS[aiPromptOpen].label : ''}
        label={aiPromptOpen === 'translate' ? 'Translate the selection (or whole document) into…' : 'Describe the topic'}
        placeholder={aiPromptOpen === 'translate' ? 'e.g. Hindi, Kannada, Spanish' : 'e.g. A one-page proposal for a school science fair'}
        confirmLabel="Generate"
        multiline={aiPromptOpen !== 'translate'}
        onSubmit={(v) => aiPromptOpen && void runAi(aiPromptOpen, v)}
        onClose={() => setAiPromptOpen(null)}
      />
      <ConfirmSheet open={confirmClear} title="Clear the whole document?" message="All text, images and comments in this document will be removed. You can still undo with the Undo button right after." confirmLabel="Clear" onConfirm={clearDoc} onClose={() => setConfirmClear(false)} />

      <Toast msg={toast} />
    </div>
  );
}

/** Speech synthesis "Read aloud" (Word mobile feature). */
function readAloud(text: string, flash: (m: string) => void) {
  const synth = window.speechSynthesis;
  if (!synth) { flash('Read aloud is not supported on this device.'); return; }
  if (synth.speaking) { synth.cancel(); flash('Stopped reading.'); return; }
  if (!text.trim()) { flash('Nothing to read.'); return; }
  const u = new SpeechSynthesisUtterance(text.slice(0, 5000));
  u.rate = 1;
  synth.speak(u);
  flash('Reading aloud… tap again to stop.');
}

/** 8x6 hover/tap grid to pick a table size, like Word's Insert Table. */
function TableGrid({ onPick }: { onPick: (rows: number, cols: number) => void }) {
  const [hover, setHover] = useState<{ r: number; c: number }>({ r: 0, c: 0 });
  const cells = [];
  for (let r = 1; r <= 6; r++) for (let c = 1; c <= 8; c++) cells.push({ r, c });
  return (
    <div className="tgrid-wrap">
      <div className="tgrid" onPointerLeave={() => setHover({ r: 0, c: 0 })}>
        {cells.map(({ r, c }) => (
          <button
            key={`${r}-${c}`}
            className={`tgrid-cell${r <= hover.r && c <= hover.c ? ' on' : ''}`}
            aria-label={`${r} by ${c} table`}
            onPointerDown={(e) => e.preventDefault()}
            onPointerEnter={() => setHover({ r, c })}
            onClick={() => onPick(r, c)}
          />
        ))}
      </div>
      <span className="hint">{hover.r ? `${hover.r} × ${hover.c} table` : 'Tap a size'}</span>
    </div>
  );
}
