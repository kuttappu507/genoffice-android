import { useMemo, useRef, useState } from 'react';
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

export default function Slides({ initialId }: { initialId?: string }) {
  const deckId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Presentation' : 'Untitled deck');
  const [slides, setSlides] = useState<Slide[]>(() => (initialId ? getDoc<DeckData>(initialId)?.slides ?? [] : []));
  const [sel, setSel] = useState(0);
  const [presenting, setPresenting] = useState(false);
  const [presIdx, setPresIdx] = useState(0);
  const [aiOpen, setAiOpen] = useState(false);
  const [aiTopic, setAiTopic] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
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
      setToast('Add your API key in Settings first.');
      setAiOpen(false);
      return;
    }
    const topic = aiTopic.trim();
    if (!topic) return;
    setAiBusy(true);
    setToast('');
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
      let cur: Slide | null = null;
      for (const line of out.split('\n')) {
        const t = line.trim();
        const h = /^###\s+(.*)/.exec(t);
        if (h) {
          if (cur) parsed.push(cur);
          cur = { title: h[1], bullets: [] };
        } else if (cur && /^[-*]\s+/.test(t)) {
          cur.bullets.push(t.replace(/^[-*]\s+/, ''));
        }
      }
      if (cur) parsed.push(cur);
      if (parsed.length === 0) throw new Error('Could not parse the outline; try again.');
      update(parsed);
      setSel(0);
      setAiOpen(false);
      setToast(`Created ${parsed.length} slides.`);
    } catch (e) {
      setToast(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  const openFile = async () => {
    const pick = await openFilePicker('.pptx');
    if (!pick) return;
    try {
      const slides = await importPptx(pick.buf);
      if (slides.length === 0) throw new Error('No slides found in the file');
      update(slides);
      setSel(0);
      setTitle(pick.name.replace(/\.[^.]+$/, ''));
      flash(`Opened ${pick.name} (${slides.length} slides)`);
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
          <button className="btn small" onClick={() => setPresenting(false)}>
            Exit
          </button>
        </div>
      </div>
    );
  }

  const cur = slides[sel];

  return (
    <div className="screen">
      <header className="screen-head">
        <input
          className="title-input"
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            save(slides, e.target.value);
          }}
          placeholder="Deck title"
        />
        <button className="btn small primary" disabled={slides.length === 0} onClick={() => { setPresIdx(0); setPresenting(true); }}>
          Present
        </button>
        <button className="btn small" onClick={() => setAiOpen(true)}>
          AI Outline
        </button>
        <button className="btn small" onClick={() => void openFile()}>
          Open
        </button>
        <button className="btn small primary" disabled={slides.length === 0} onClick={() => void savePptx()}>
          Pptx
        </button>
      </header>

      {slides.length === 0 ? (
        <p className="empty">No slides yet. Use AI Outline or tap Add slide.</p>
      ) : (
        <>
          <div className="slide-filmstrip">
            {slides.map((s, i) => (
              <button key={i} className={`slide-thumb${i === sel ? ' selected' : ''}`} onClick={() => setSel(i)}>
                <span className="slide-num">{i + 1}</span>
                <span className="slide-thumb-title">{s.title || 'Untitled'}</span>
              </button>
            ))}
          </div>

          {cur && (
            <div className="slide-editor card">
              <input
                className="input slide-title-input"
                value={cur.title}
                onChange={(e) => editSel({ title: e.target.value })}
                placeholder="Slide title"
              />
              <textarea
                className="input"
                rows={6}
                value={cur.bullets.join('\n')}
                onChange={(e) => editSel({ bullets: e.target.value.split('\n') })}
                placeholder="One bullet per line"
              />
              <div className="btn-row">
                <button className="btn small" onClick={() => move(sel, -1)} disabled={sel === 0}>
                  Up
                </button>
                <button className="btn small" onClick={() => move(sel, 1)} disabled={sel === slides.length - 1}>
                  Down
                </button>
                <button className="btn small danger" onClick={() => removeSlide(sel)}>
                  Delete
                </button>
                <button className="btn small primary" onClick={addSlide}>
                  Add slide
                </button>
              </div>
            </div>
          )}
        </>
      )}

      {slides.length === 0 && (
        <div className="btn-row">
          <button className="btn primary" onClick={addSlide}>
            Add slide
          </button>
        </div>
      )}

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
                {aiBusy ? 'Generating...' : 'Generate'}
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
