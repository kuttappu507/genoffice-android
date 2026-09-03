import { useState } from 'react';
import { FileTypeIcon, Icon } from '../components/Icon';
import { getPrefs, savePrefs } from '../lib/storage';
import { tap } from '../lib/native';

const PAGES: { icon: JSX.Element; title: string; text: string; accent: string }[] = [
  {
    icon: <span className="ob-logo">G</span>,
    title: 'Welcome to GenOffice',
    text: 'Word, Excel and PowerPoint-style editors in one small app. Everything is stored on your phone — no account, no cloud, no telemetry.',
    accent: '#0F6CBD',
  },
  {
    icon: <FileTypeIcon kind="doc" size={72} />,
    title: 'Documents',
    text: 'Styles, fonts, tables, images, comments, headers & footers, find & replace. Open and save real .docx files or export to PDF.',
    accent: '#185ABD',
  },
  {
    icon: <FileTypeIcon kind="sheet" size={72} />,
    title: 'Spreadsheets',
    text: '120+ formulas with autocomplete, charts, freeze panes, filters, sorting, number formats and multi-sheet .xlsx / CSV files.',
    accent: '#107C41',
  },
  {
    icon: <FileTypeIcon kind="deck" size={72} />,
    title: 'Presentations',
    text: 'Themes, layouts, shapes, notes, transitions and a presenter view with timer and laser pointer. Open and save .pptx or export PDF handouts.',
    accent: '#C43E1C',
  },
  {
    icon: <span className="ob-ai"><Icon name="sparkles" size={56} /></span>,
    title: 'AI when you want it',
    text: 'Add your own OpenRouter or NVIDIA key in Settings to draft, rewrite, summarise, fill tables and outline decks. Everything else works offline.',
    accent: '#5B5FC7',
  },
];

export default function Onboarding({ onDone }: { onDone: () => void }) {
  const [i, setI] = useState(0);
  const p = PAGES[i];
  const last = i === PAGES.length - 1;
  const finish = () => {
    savePrefs({ ...getPrefs(), onboarded: true });
    onDone();
  };
  return (
    <div className="onboard" style={{ ['--accent' as string]: p.accent }}>
      <button className="ob-skip" onClick={finish}>Skip</button>
      <div className="ob-art">{p.icon}</div>
      <h1 className="ob-title">{p.title}</h1>
      <p className="ob-text">{p.text}</p>
      <div className="ob-dots">
        {PAGES.map((_, k) => <span key={k} className={`ob-dot${k === i ? ' on' : ''}`} onClick={() => setI(k)} />)}
      </div>
      <div className="ob-actions">
        {i > 0 && <button className="btn" onClick={() => setI(i - 1)}>Back</button>}
        <button className="btn primary" onClick={() => { void tap(); last ? finish() : setI(i + 1); }}>
          {last ? 'Get started' : 'Next'} <Icon name={last ? 'check' : 'chevronRight'} size={16} />
        </button>
      </div>
    </div>
  );
}
