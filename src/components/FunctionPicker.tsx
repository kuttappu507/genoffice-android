import { useMemo, useState } from 'react';
import { FUNCTION_HELP } from '../lib/formulas';
import { BottomSheet } from './Sheet';
import { Icon } from './Icon';

const CATS = ['All', 'Math', 'Conditional', 'Logic', 'Text', 'Lookup', 'Date', 'Finance'];

/** Excel-style "Insert Function" browser: search + categories + signature help. */
export function FunctionPicker({ open, onClose, onPick, initialCat }: { open: boolean; onClose: () => void; onPick: (name: string, sig: string) => void; initialCat?: string }) {
  const [q, setQ] = useState('');
  const [cat, setCat] = useState(initialCat ?? 'All');
  const list = useMemo(() => {
    const qq = q.trim().toUpperCase();
    return FUNCTION_HELP.filter((f) => (cat === 'All' || f.cat === cat) && (!qq || f.name.includes(qq) || f.desc.toUpperCase().includes(qq)));
  }, [q, cat]);
  return (
    <BottomSheet open={open} onClose={onClose} title="Insert function" tall>
      <div className="fp-search">
        <Icon name="search" size={16} className="dim" />
        <input className="input" placeholder="Search functions (e.g. VLOOKUP, IF, PMT)" value={q} onChange={(e) => setQ(e.target.value)} autoFocus />
      </div>
      <div className="chip-row">
        {CATS.map((c) => (
          <button key={c} className={`chip${cat === c ? ' on' : ''}`} onClick={() => setCat(c)}>{c}</button>
        ))}
      </div>
      <div className="fp-list">
        {list.map((f) => (
          <button key={f.name} className="fp-item" onClick={() => { onPick(f.name, f.sig); onClose(); }}>
            <span className="fp-name">{f.name}</span>
            <span className="fp-sig">{f.sig}</span>
            <span className="fp-desc">{f.desc}</span>
          </button>
        ))}
        {list.length === 0 && <p className="empty">No functions match “{q}”.</p>}
      </div>
    </BottomSheet>
  );
}
