import { useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { FileTypeIcon } from '../components/Icon';
import { Palette, RBtn, RGroup, RWide } from '../components/Ribbon';
import { chatStream, errMsg } from '../lib/ai-client';
import { exportPptx, importPptx, openFilePicker, pickImage, saveBinary, sanitizeName } from '../lib/fileio';
import type { DeckSlide, ShapePara } from '../lib/fileio';
import { debounce, getDoc, getSettings, putDoc, uid } from '../lib/storage';

const alignCss = (a: ShapePara['align']): CSSProperties['textAlign'] =>
  a === 'center' ? 'center' : a === 'right' ? 'right' : a === 'justify' ? 'justify' : 'left';

/**
 * Faithful slide renderer: absolutely-positioned shapes on a cw x ch inch
 * canvas, scaled from inches to pixels. Used by the editor canvas, the
 * filmstrip thumbnails and present mode so all views match the real deck.
 */
export function SlideView({
  slide,
  width,
  editShapes,
  onShapeTap,
  selectedShape,
}: {
  slide: DeckSlide;
  width: number;
  editShapes?: boolean;
  onShapeTap?: (idx: number) => void;
  selectedShape?: number | null;
}) {
  const cw = slide.cw ?? 10;
  const ch = slide.ch ?? 5.63;
  const dark = /^#[0-9a-fA-F]{6}$/.test(slide.bg ?? '')
    ? (0.299 * parseInt(slide.bg!.slice(1, 3), 16) + 0.587 * parseInt(slide.bg!.slice(3, 5), 16) + 0.114 * parseInt(slide.bg!.slice(5, 7), 16)) < 140
    : false;
  if (width <= 0) return null;
  const ptScale = width / (cw * 72); // px per point
  return (
    <div className="sv" style={{ width, height: (width * ch) / cw, background: slide.bg ?? '#FFFFFF' }}>
      {(slide.shapes ?? []).map((sh, i) => {
        const style: CSSProperties = {
          left: `${(sh.x / cw) * 100}%`,
          top: `${(sh.y / ch) * 100}%`,
          width: `${(sh.w / cw) * 100}%`,
          height: `${(sh.h / ch) * 100}%`,
          background: sh.fill,
          border: sh.line ? `1px solid ${sh.line}` : undefined,
        };
        if (sh.kind === 'image' && sh.img) {
          return (
            <div key={i} className="sv-sh img" style={style}>
              <img src={sh.img} alt="" draggable={false} />
            </div>
          );
        }
        return (
          <div
            key={i}
            className={`sv-sh${editShapes && onShapeTap ? ' tappable' : ''}${selectedShape === i ? ' sel' : ''}`}
            style={style}
            onClick={editShapes && onShapeTap ? () => onShapeTap(i) : undefined}
          >
            {sh.paras.map((p, j) => {
              const text = p.runs.map((r) => r.text).join('');
              const first = p.runs.find((r) => r.text.trim()) ?? p.runs[0];
              if (!text.trim()) return <p key={j} style={{ fontSize: Math.max(5, (first?.sz ?? 14) * ptScale) }}>\u00a0</p>;
              return (
                <p
                  key={j}
                  style={{
                    fontSize: Math.max(5, (first?.sz ?? 14) * ptScale),
                    color: first?.color ?? (dark ? '#F2F2F2' : '#333333'),
                    fontWeight: p.runs.some((r) => r.b) ? 700 : 400,
                    fontStyle: p.runs.some((r) => r.i) ? 'italic' : undefined,
                    textDecoration: p.runs.some((r) => r.u) ? 'underline' : undefined,
                    textAlign: alignCss(p.align),
                    paddingLeft: p.bullet ? `${1 + (p.level ?? 0) * 0.9}em` : undefined,
                    textIndent: p.bullet ? '-0.75em' : undefined,
                  }}
                >
                  {p.bullet ? '\u2022 ' : ''}
                  {text}
                </p>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}

type RibbonTab = 'home' | 'insert' | 'design' | 'ai';

interface DeckData {
  slides: DeckSlide[];
}

const THEMES = [
  { name: 'Classic', bg: '#FFFFFF', accent: '#C43E1C' },
  { name: 'Ivory', bg: '#FDF6EC', accent: '#B7791F' },
  { name: 'Forest', bg: '#F0F7F0', accent: '#107C41' },
  { name: 'Ocean', bg: '#EAF3FB', accent: '#185ABD' },
  { name: 'Plum', bg: '#F6EEF9', accent: '#7030A0' },
  { name: 'Ink', bg: '#1E2430', accent: '#7EB8DA' },
];

const LAYOUTS = [
  { id: 'title', icon: 'layoutTitle', label: 'Title' },
  { id: 'content', icon: 'layoutContent', label: 'Content' },
  { id: 'section', icon: 'layoutSection', label: 'Section' },
  { id: 'blank', icon: 'layoutBlank', label: 'Blank' },
] as const;

export default function Slides({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const deckId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Presentation' : 'Untitled deck');
  const [slides, setSlides] = useState<DeckSlide[]>(() => (initialId ? getDoc<DeckData>(initialId)?.slides ?? [] : []));
  const [sel, setSel] = useState(0);
  const [presenting, setPresenting] = useState(false);
  const [presIdx, setPresIdx] = useState(0);
  const [rTab, setRTab] = useState<RibbonTab>('home');
  const [bgPalette, setBgPalette] = useState(false);
  const [menu, setMenu] = useState(false);
  const [aiOpen, setAiOpen] = useState(false);
  const [aiTopic, setAiTopic] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [toast, setToast] = useState('');
  const [shapeEdit, setShapeEdit] = useState<number | null>(null);
  const [shapeText, setShapeText] = useState('');
  const [canvasW, setCanvasW] = useState(0);
  const [presW, setPresW] = useState(0);
  const canvasHolderRef = useRef<HTMLDivElement>(null);

  // keep the shape canvas sized to its container
  useEffect(() => {
    const el = canvasHolderRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setCanvasW(el.clientWidth));
    ro.observe(el);
    setCanvasW(el.clientWidth);
    return () => ro.disconnect();
  }, [sel, presenting, slides.length]);

  // fit the present-mode slide to the viewport
  useEffect(() => {
    if (!presenting) return;
    const calc = () => {
      const s = slides[presIdx];
      const cw = s?.cw ?? 10;
      const ch = s?.ch ?? 5.63;
      const w = Math.min(window.innerWidth - 10, ((window.innerHeight - 70) * cw) / ch);
      setPresW(Math.max(120, w));
    };
    calc();
    window.addEventListener('resize', calc);
    return () => window.removeEventListener('resize', calc);
  }, [presenting, presIdx, slides]);

  const save = useMemo(
    () =>
      debounce((sl: DeckSlide[], t: string) => {
        putDoc<DeckData>('deck', deckId, t, { slides: sl });
      }, 700),
    [deckId],
  );

  const update = (next: DeckSlide[]) => {
    setSlides(next);
    save(next, title);
  };

  const editSel = (patch: Partial<DeckSlide>) => {
    update(slides.map((s, i) => (i === sel ? { ...s, ...patch } : s)));
  };

  const addSlide = (layout: DeckSlide['layout'] = 'content') => {
    const next = [...slides, { title: layout === 'section' ? 'Section title' : 'New slide', bullets: layout === 'title' ? ['Subtitle'] : [''], layout }];
    update(next);
    setSel(next.length - 1);
  };

  const duplicateSlide = () => {
    const cur = slides[sel];
    if (!cur) return;
    const copy: DeckSlide = {
      ...cur,
      bullets: [...cur.bullets],
      shapes: cur.shapes?.map((sh) => ({ ...sh, paras: sh.paras.map((p) => ({ ...p, runs: p.runs.map((r) => ({ ...r })) })) })),
    };
    const next = [...slides];
    next.splice(sel + 1, 0, copy);
    update(next);
    setSel(sel + 1);
  };

  const openShapeEdit = (idx: number) => {
    const sh = slides[sel]?.shapes?.[idx];
    if (!sh) return;
    setShapeText(sh.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n'));
    setShapeEdit(idx);
  };

  const saveShapeEdit = () => {
    const cur = slides[sel];
    if (shapeEdit === null || !cur?.shapes) return;
    const sh = cur.shapes[shapeEdit];
    const lines = shapeText.split('\n');
    const paras: ShapePara[] = lines.map((ln, i) => {
      const old = sh.paras[i] ?? sh.paras[0];
      const base = old?.runs[0];
      return {
        align: old?.align,
        bullet: old?.bullet,
        level: old?.level,
        runs: [{ ...(base ?? { sz: 14 }), text: ln } as ShapePara['runs'][number]],
      };
    });
    editSel({ shapes: cur.shapes.map((s, i) => (i === shapeEdit ? { ...s, paras } : s)) });
    setShapeEdit(null);
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

  const insertImage = async () => {
    try {
      const data = await pickImage(1200);
      if (data) editSel({ image: data });
    } catch (e) {
      flash(`Could not insert image: ${errMsg(e)}`);
    }
  };

  const applyTheme = (t: (typeof THEMES)[number]) => editSel({ bg: t.bg, accent: t.accent });

  const applyThemeAll = (t: (typeof THEMES)[number]) => {
    update(slides.map((s) => ({ ...s, bg: t.bg, accent: t.accent })));
    flash(`${t.name} theme applied to all slides.`);
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
      const parsed: DeckSlide[] = [];
      let curSlide: DeckSlide | null = null;
      for (const line of out.split('\n')) {
        const t = line.trim();
        const h = /^###\s+(.*)/.exec(t);
        if (h) {
          if (curSlide) parsed.push(curSlide);
          curSlide = { title: h[1], bullets: [], layout: 'content' };
        } else if (curSlide && /^[-*]\s+/.test(t)) {
          curSlide.bullets.push(t.replace(/^[-*]\s+/, ''));
        }
      }
      if (curSlide) parsed.push(curSlide);
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
    const dark = (s.bg ?? '#FFFFFF').toLowerCase().match(/^#([0-9a-f]{6})$/) ? isDark(s.bg!) : false;
    const accent = s.accent ?? '#C43E1C';
    return (
      <div className="present">
        <div className="present-zones">
          <button className="present-zone" onClick={prev} aria-label="Previous slide" />
          <button className="present-zone wide" onClick={next} aria-label="Next slide" />
        </div>
        <div
          className="present-content"
          style={{ background: s.bg ?? '#FFFFFF', color: dark ? '#F2F2F2' : '#1B1B1B' }}
        >
          {s.shapes && s.shapes.length > 0 ? (
            <SlideView slide={s} width={presW} />
          ) : (
            <>
              {(s.layout === 'section') && <div className="present-band" style={{ background: accent }} />}
              <h2 style={s.layout === 'title' || s.layout === 'section' ? { textAlign: 'center', color: dark ? '#FFFFFF' : '#1B1B1B' } : undefined}>
                {s.title}
              </h2>
              {s.layout !== 'section' && (
                <ul>
                  {s.bullets.filter((b) => b.trim()).map((b, i) => (
                    <li key={i} style={s.layout === 'title' ? { listStyle: 'none', textAlign: 'center', color: accent, marginLeft: -20 } : undefined}>
                      {b}
                    </li>
                  ))}
                </ul>
              )}
            </>
          )}
        </div>
        <div className="present-bar">
          <span>
            {presIdx + 1} / {slides.length}
          </span>
          <button className="icon-btn light" aria-label="Exit presentation" onClick={() => setPresenting(false)}>
            ✕
          </button>
        </div>
      </div>
    );
  }

  function isDark(hex: string): boolean {
    const m = /^#([0-9a-fA-F]{6})$/.exec(hex);
    if (!m) return false;
    const v = parseInt(m[1], 16);
    const r = (v >> 16) & 255;
    const g = (v >> 8) & 255;
    const b = v & 255;
    return 0.299 * r + 0.587 * g + 0.114 * b < 140;
  }

  const cur = slides[sel];
  const curTheme = cur ? THEMES.find((t) => t.bg === cur.bg && t.accent === cur.accent) : undefined;

  return (
    <div className="edscreen" style={{ ['--app' as string]: 'var(--ppt)' }}>
      <header className="appbar">
        <button className="icon-btn light" aria-label="Back to Home" onClick={onExit}>
          <FileTypeIcon kind="deck" size={22} />
        </button>
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
          ▶
        </button>
        <button className="icon-btn light" aria-label="Save as .pptx" disabled={slides.length === 0} onClick={() => void savePptx()}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" /><polyline points="17 21 17 13 7 13 7 21" /><polyline points="7 3 7 8 15 8" /></svg>
        </button>
        <div className="menu-wrap">
          <button className="icon-btn light" aria-label="More actions" onClick={() => setMenu(!menu)}>
            ⋮
          </button>
          {menu && (
            <>
              <div className="menu-backdrop" onClick={() => setMenu(false)} />
              <div className="menu">
                <button className="menu-item" onClick={() => void openFile()}>
                  Open .pptx
                </button>
                <button className="menu-item" onClick={() => { setMenu(false); setAiOpen(true); }}>
                  AI outline
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
            <p>No slides yet. Use Home → New slide, or generate an AI outline.</p>
          </div>
        ) : cur?.shapes && cur.shapes.length > 0 ? (
          <div
            className="shape-canvas-holder"
            ref={canvasHolderRef}
            style={{ borderColor: cur.accent ?? 'var(--ppt)' }}
          >
            <SlideView slide={cur} width={Math.min(canvasW, 620)} editShapes onShapeTap={openShapeEdit} />
            <p className="shape-hint">Tap any text box to edit it — layout, colors and images stay exactly as designed.</p>
          </div>
        ) : (
          cur && (
            <div
              className={`slide-canvas${cur.image ? ' imgright' : ''}`}
              style={{ background: cur.bg ?? '#FFFFFF', borderColor: cur.accent ?? 'var(--ppt)' }}
            >
              <span className="slide-canvas-num" style={{ color: isDark(cur.bg ?? '#FFFFFF') ? 'rgba(255,255,255,0.7)' : '#9a9a9a' }}>
                {sel + 1}
              </span>

              {cur.layout === 'section' && (
                <div className="section-band" style={{ background: cur.accent ?? 'var(--ppt)' }}>
                  <input
                    className="canvas-title onband"
                    value={cur.title}
                    onChange={(e) => editSel({ title: e.target.value })}
                    placeholder="Section title"
                  />
                </div>
              )}

              {cur.layout === 'title' && (
                <div className="title-wrap">
                  <input
                    className="canvas-title xl"
                    style={{ color: isDark(cur.bg ?? '#FFFFFF') ? '#FFFFFF' : '#1B1B1B' }}
                    value={cur.title}
                    onChange={(e) => editSel({ title: e.target.value })}
                    placeholder="Presentation title"
                  />
                  <input
                    className="canvas-sub"
                    style={{ color: cur.accent ?? 'var(--ppt)' }}
                    value={cur.bullets[0] ?? ''}
                    onChange={(e) => {
                      const b = [...cur.bullets];
                      b[0] = e.target.value;
                      editSel({ bullets: b });
                    }}
                    placeholder="Subtitle"
                  />
                </div>
              )}

              {(cur.layout === 'content' || !cur.layout) && (
                <>
                  <input
                    className="canvas-title"
                    style={{ color: isDark(cur.bg ?? '#FFFFFF') ? '#FFFFFF' : '#1B1B1B' }}
                    value={cur.title}
                    onChange={(e) => editSel({ title: e.target.value })}
                    placeholder="Slide title"
                  />
                  <div
                    className="title-rule"
                    style={{ background: cur.accent ?? 'var(--ppt)' }}
                  />
                  <div className="canvas-bullets">
                    {cur.bullets.map((b, i) => (
                      <div key={i} className="canvas-bullet">
                        <span className="dot" style={{ color: cur.accent ?? 'var(--ppt)' }}>•</span>
                        <input
                          value={b}
                          placeholder="Bullet"
                          style={{ color: isDark(cur.bg ?? '#FFFFFF') ? '#E8E8E8' : '#2B2B2B' }}
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
                          ✕
                        </button>
                      </div>
                    ))}
                    <button className="add-bullet" onClick={() => editSel({ bullets: [...cur.bullets, ''] })}>
                      + Add bullet
                    </button>
                  </div>
                </>
              )}

              {cur.image && <img className="canvas-img" src={cur.image} alt="" />}

              {cur.layout === 'blank' && !cur.image && (
                <p className="blank-hint" style={{ color: isDark(cur.bg ?? '#FFFFFF') ? 'rgba(255,255,255,0.5)' : '#B0B0B0' }}>
                  Blank slide — add an image from the Insert tab.
                </p>
              )}
            </div>
          )
        )}
      </div>

      <div className="ribbon">
        <div className="ribbon-tabs">
          {(['home', 'insert', 'design', 'ai'] as RibbonTab[]).map((t) => (
            <button key={t} className={`ribbon-tab${rTab === t ? ' active' : ''}`} onClick={() => { setRTab(t); setBgPalette(false); }}>
              {t === 'home' ? 'Home' : t === 'insert' ? 'Insert' : t === 'design' ? 'Design' : 'AI'}
            </button>
          ))}
        </div>

        {bgPalette && (
          <Palette
            onPick={(c) => {
              editSel({ bg: c });
              setBgPalette(false);
            }}
          />
        )}

        {rTab === 'home' && (
          <div className="ribbon-row">
            <RGroup label="Slides">
              <RBtn icon="plus" label="New" onRun={() => addSlide('content')} />
              <RBtn icon="copy" label="Duplicate" disabled={!cur} onRun={duplicateSlide} />
              <RBtn icon="trash" label="Delete" disabled={!cur || slides.length === 0} onRun={() => removeSlide(sel)} />
            </RGroup>
            <RGroup label="Reorder">
              <RBtn icon="chevronLeft" label="Earlier" disabled={!cur || sel === 0} onRun={() => move(sel, -1)} />
              <RBtn icon="chevronRight" label="Later" disabled={!cur || sel === slides.length - 1} onRun={() => move(sel, 1)} />
            </RGroup>
            <RGroup label="Layouts">
              {LAYOUTS.map((l) => (
                <RBtn
                  key={l.id}
                  icon={l.icon}
                  label={l.label}
                  active={cur?.layout === l.id || (l.id === 'content' && !cur?.layout)}
                  disabled={!cur}
                  onRun={() => editSel({ layout: l.id })}
                />
              ))}
            </RGroup>
          </div>
        )}

        {rTab === 'insert' && (
          <div className="ribbon-row">
            <RGroup label="Media">
              <RBtn icon="image" label="Picture" disabled={!cur} onRun={() => void insertImage()} />
              <RBtn icon="close" label="Remove" disabled={!cur?.image} onRun={() => editSel({ image: undefined })} />
            </RGroup>
            <RGroup label="Bullets">
              <RBtn icon="plus" label="Add bullet" disabled={!cur || cur.layout !== 'content'} onRun={() => editSel({ bullets: [...(cur?.bullets ?? []), ''] })} />
            </RGroup>
          </div>
        )}

        {rTab === 'design' && (
          <div className="ribbon-row">
            <RGroup label="Themes (this slide)">
              <div className="theme-chips">
                {THEMES.map((t) => (
                  <button
                    key={t.name}
                    className={`theme-chip${curTheme?.name === t.name ? ' active' : ''}`}
                    title={`${t.name} theme`}
                    aria-label={`${t.name} theme`}
                    onClick={() => applyTheme(t)}
                  >
                    <span className="chip-bg" style={{ background: t.bg }}>
                      <span className="chip-accent" style={{ background: t.accent }} />
                    </span>
                  </button>
                ))}
              </div>
            </RGroup>
            <RGroup label="Apply">
              <RWide icon="theme" label={curTheme ? `All slides: ${curTheme.name}` : 'Apply to all slides'} disabled={!cur} onRun={() => curTheme && applyThemeAll(curTheme)} />
            </RGroup>
            <RGroup label="Background">
              <RBtn icon="fill" label="Pick color" disabled={!cur} onRun={() => setBgPalette(!bgPalette)} />
            </RGroup>
          </div>
        )}

        {rTab === 'ai' && (
          <div className="ribbon-row">
            <RGroup label="AI deck builder">
              <RWide icon="sparkle" label="Generate an outline" disabled={aiBusy} onRun={() => setAiOpen(true)} />
            </RGroup>
          </div>
        )}
      </div>

      <div className="filmstrip">
        {slides.map((s, i) => (
          <button
            key={i}
            className={`slide-thumb${i === sel ? ' selected' : ''}`}
            style={{ background: s.bg ?? '#FFFFFF' }}
            onClick={() => setSel(i)}
          >
            {s.shapes && s.shapes.length > 0 ? (
              <span className="thumb-sv" aria-hidden="true">
                <SlideView slide={s} width={96} />
              </span>
            ) : (
              <span className="slide-thumb-title" style={{ color: isDark(s.bg ?? '#FFFFFF') ? '#F0F0F0' : 'inherit' }}>
                {s.title || 'Untitled'}
              </span>
            )}
            <span className="slide-num" style={{ color: isDark(s.bg ?? '#FFFFFF') ? 'rgba(255,255,255,0.65)' : '#9a9a9a' }}>
              {i + 1}
            </span>
          </button>
        ))}
        <button className="slide-thumb add" aria-label="Add slide" onClick={() => addSlide('content')}>
          +
        </button>
      </div>

      {shapeEdit !== null && (
        <div className="modal" onClick={() => setShapeEdit(null)}>
          <div className="modal-body" onClick={(e) => e.stopPropagation()}>
            <h3>Edit text box</h3>
            <p className="hint">One line per paragraph. Font, color and position are kept.</p>
            <textarea
              className="input shape-input"
              value={shapeText}
              rows={5}
              onChange={(e) => setShapeText(e.target.value)}
            />
            <div className="btn-row">
              <button className="btn primary" onClick={saveShapeEdit}>
                Apply
              </button>
              <button className="btn" onClick={() => setShapeEdit(null)}>
                Cancel
              </button>
            </div>
          </div>
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
