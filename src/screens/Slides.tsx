import { useMemo, useRef, useState } from 'react';
import { Icon, FileTypeIcon } from '../components/Icon';
import { chatStream, errMsg } from '../lib/ai-client';
import { exportPptx, importPptx, openFilePicker, saveBinary, sanitizeName } from '../lib/fileio';
import { debounce, getDoc, getSettings, putDoc, uid } from '../lib/storage';

interface Slide {
  title: string;
  bullets: string[];
}

interface DeckData {
  slides: Slide[];
}

export default function Slides({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const deckId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Presentation' : 'Untitled deck');
  const [slides, setSlides] = useState<Slide[]>(() => (initialId ? getDoc<DeckData>(initialId)?.slides ?? [] : []));
  const [sel, setSel] = useState(0);
  const [presenting, setPresenting] = useState(false);
  const [presIdx, setPresIdx] = useState(0);
  const [aiOpen, setAiOpen] = useState(false);
  const [aiTopic, setAiTopic] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [menu, setMenu] = useState(false);
  const [toast, setToast] = useState('');

  const save = useMemo(
    () =>
      debounce((sl: Slide[], t: string) => {
        putDoc<DeckData>('deck', deckId, t, { slides: sl });
      }, 700),
    [deckId],
  );

  const update = (next: Slide[]) => {
    setSlides(next);
    save(next, title);
  };

  const editSel = (patch: Partial<Slide>) => {
    const next = slides.map((s, i) => (i === sel ? { ...s, ...patch } : s));
    update(next);
  };

  const addSlide = () => {
    const next = [...slides, { title: 'New slide', bullets: [''] }];
    update(next);
    setSel(next.length - 1);
  };

  const duplicateSlide = () => {
    if (!cur) return;
    const next = [...slides];
    next.splice(sel + 1, 0, { title: cur.title, bullets: [...cur.bullets] });
    update(next);
    setSel(sel + 1);
  };

  const move = (from: number, dir: -1 | 1) => {
    const to = from + dir;
    if (to < 0 || to >= slides.length) return;
    const next = [...slides];
    const [s] = next.splice(from, 1);
    next.splice(to, 0, s);
    update(next);
    setSel(to);
  };

  const removeSlide = (idx: number) => {
    const next = slides.filter((_, i) => i !== idx);
    update(next);
    setSel(Math.max(0, Math.min(sel, next.length - 1)));
  };

  const runAi = async () => {
    const s = getSettings();
    if (!s.apiKey) {
      flash('Add your API key in Settings first.');
      setAiOpen(false);
      return;
    }
    const topic = aiTopic.trim();
    if (!topic) return;
    setAiBusy(true);
    try {
      const out = await chatStream(s, [
        {
          role: 'system',
          content:
            'You create presentation outlines. Format each slide exactly like this: a line starting with "### " followed by the slide title, then 3-4 lines each starting with "- " as bullets. 5-8 slides. No extra commentary, same language as the request.',
        },
        { role: 'user', content: topic },
      ]);
      const parsed: Slide[] = [];
      let curP: Slide | null = null;
      for (const line of out.split('\n')) {
        const t = line.trim();
        const h = /^###\s+(.*)/.exec(t);
        if (h) {
          if (curP) parsed.push(curP);
          curP = { title: h[1], bullets: [] };
        } else if (curP && /^[-*]\s+/.test(t)) {
          curP.bullets.push(t.replace(/^[-*]\s+/, ''));
        }
      }
      if (curP) parsed.push(curP);
      if (parsed.length === 0) throw new Error('Could not parse the outline; try again.');
      update(parsed);
      setSel(0);
      setAiOpen(false);
      flash(`Created ${parsed.length} slides.`);
    } catch (e) {
      flash(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  const openFile = async () => {
    setMenu(false);
    const pick = await openFilePicker('.pptx');
    if (!pick) return;
    try {
      const loaded = await importPptx(pick.buf);
      if (loaded.length === 0) throw new Error('No slides found in the file');
      update(loaded);
      setSel(0);
      setTitle(pick.name.replace(/\.[^.]+$/, ''));
      flash(`Opened ${pick.name} (${loaded.length} slides)`);
    } catch (e) {
      flash(`Could not open: ${errMsg(e)}`);
    }
  };

  const savePptx = async () => {
    if (slides.length === 0) return;
    try {
      const bytes = await exportPptx(title, slides);
      await saveBinary(sanitizeName(title, 'pptx'), bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      flash('Presentation saved');
    } catch (e) {
      flash(`Save failed: ${errMsg(e)}`);
    }
  };

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  };

  if (presenting && slides.length > 0) {
    const s = slides[presIdx];
    const next = () => setPresIdx((i) => Math.min(i + 1, slides.length - 1));
    const prev = () => setPresIdx((i) => Math.max(i - 1, 0));
    return (
      <div className="present">
        <div className="present-zones">
          <button className="present-zone" onClick={prev} aria-label="Previous slide" />
          <button className="present-zone wide" onClick={next} aria-label="Next slide" />
        </div>
        <div className="present-content">
          <h2>{s.title}</h2>
          <ul>
            {s.bullets.filter((b) => b.trim()).map((b, i) => (
              <li key={i}>{b}</li>
            ))}
          </ul>
        </div>
        <div className="present-bar">
          <span>
            {presIdx + 1} / {slides.length}
          </span>
          <button className="icon-btn light" aria-label="Exit presentation" onClick={() => setPresenting(false)}>
            <Icon name="close" size={20} />
          </button>
        </div>
      </div>
    );
  }

  const cur = slides[sel];

  return (
    <div className="edscreen" style={{ ['--app' as string]: 'var(--ppt)' }}>
      <header className="appbar">
        <button className="icon-btn light" aria-label="Back to Home" onClick={onExit}>
          <Icon name="arrowLeft" size={21} />
        </button>
        <FileTypeIcon kind="deck" size={26} />
        <input
          className="appbar-title"
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            save(slides, e.target.value);
          }}
          placeholder="Deck title"
        />
        <button
          className="icon-btn light"
          aria-label="Present"
          disabled={slides.length === 0}
          onClick={() => { setPresIdx(sel); setPresenting(true); }}
        >
          <Icon name="play" size={19} />
        </button>
        <button className="icon-btn light" aria-label="Save as .pptx" disabled={slides.length === 0} onClick={() => void savePptx()}>
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
                  <Icon name="folder" size={18} /> Open .pptx
                </button>
                <button className="menu-item" onClick={() => { setMenu(false); setAiOpen(true); }}>
                  <Icon name="sparkle" size={18} /> AI outline
                </button>
              </div>
            </>
          )}
        </div>
      </header>

      <div className="edbody deck-body">
        {slides.length === 0 ? (
          <div className="deck-empty">
            <FileTypeIcon kind="deck" size={54} />
            <p>No slides yet. Add one below or generate an AI outline from the ⋮ menu.</p>
          </div>
        ) : (
          <>
            {cur && (
              <div className="slide-canvas">
                <span className="slide-canvas-num">{sel + 1}</span>
                <input
                  className="canvas-title"
                  value={cur.title}
                  onChange={(e) => editSel({ title: e.target.value })}
                  placeholder="Slide title"
                />
                <div className="canvas-bullets">
                  {cur.bullets.map((b, i) => (
                    <div key={i} className="canvas-bullet">
                      <span className="dot">•</span>
                      <input
                        value={b}
                        placeholder="Bullet"
                        onChange={(e) => {
                          const next = [...cur.bullets];
                          next[i] = e.target.value;
                          editSel({ bullets: next });
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            const next = [...cur.bullets, ''];
                            editSel({ bullets: next });
                          }
                        }}
                      />
                      <button
                        className="bullet-x"
                        aria-label="Remove bullet"
                        onClick={() => editSel({ bullets: cur.bullets.filter((_, j) => j !== i) })}
                      >
                        <Icon name="close" size={12} />
                      </button>
                    </div>
                  ))}
                  <button className="add-bullet" onClick={() => editSel({ bullets: [...cur.bullets, ''] })}>
                    <Icon name="plus" size={13} /> Add bullet
                  </button>
                </div>
              </div>
            )}

            <div className="slide-ops">
              <button className="icon-btn" aria-label="Move up" disabled={sel === 0} onClick={() => move(sel, -1)}>
                <Icon name="chevronLeft" size={19} />
              </button>
              <button className="icon-btn" aria-label="Move down" disabled={sel === slides.length - 1} onClick={() => move(sel, 1)}>
                <Icon name="chevronRight" size={19} />
              </button>
              <button className="icon-btn" aria-label="Duplicate slide" onClick={duplicateSlide}>
                <Icon name="copy" size={18} />
              </button>
              <button className="icon-btn danger" aria-label="Delete slide" onClick={() => removeSlide(sel)}>
                <Icon name="trash" size={18} />
              </button>
              <span className="slide-count">{slides.length} slides</span>
            </div>
          </>
        )}
      </div>

      <div className="filmstrip">
        {slides.map((s, i) => (
          <button key={i} className={`slide-thumb${i === sel ? ' selected' : ''}`} onClick={() => setSel(i)}>
            <span className="slide-num">{i + 1}</span>
            <span className="slide-thumb-title">{s.title || 'Untitled'}</span>
          </button>
        ))}
        <button className="slide-thumb add" aria-label="Add slide" onClick={addSlide}>
          <Icon name="plus" size={22} />
        </button>
      </div>

      {aiOpen && (
        <div className="modal" onClick={() => !aiBusy && setAiOpen(false)}>
          <div className="modal-body" onClick={(e) => e.stopPropagation()}>
            <h3>Generate an outline</h3>
            <p className="hint">The deck will be replaced with the generated slides.</p>
            <input
              className="input"
              value={aiTopic}
              placeholder="e.g. Introduction to renewable energy for students"
              onChange={(e) => setAiTopic(e.target.value)}
            />
            <div className="btn-row">
              <button className="btn primary" disabled={aiBusy || !aiTopic.trim()} onClick={() => void runAi()}>
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
