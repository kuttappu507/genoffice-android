import { useState } from 'react';
import type { DocKind } from './types';
import Home from './screens/Home';
import Chat from './screens/Chat';
import Docs from './screens/Docs';
import Sheets from './screens/Sheets';
import Slides from './screens/Slides';
import Settings from './screens/Settings';

type Tab = 'home' | 'chat' | 'docs' | 'sheets' | 'slides' | 'settings';

const TABS: { id: Tab; label: string }[] = [
  { id: 'home', label: 'Home' },
  { id: 'chat', label: 'AI Chat' },
  { id: 'docs', label: 'Docs' },
  { id: 'sheets', label: 'Sheets' },
  { id: 'slides', label: 'Slides' },
  { id: 'settings', label: 'Settings' },
];

const TAB_FOR_KIND: Record<DocKind, Tab> = {
  chat: 'chat',
  doc: 'docs',
  sheet: 'sheets',
  deck: 'slides',
};

export default function App() {
  const [tab, setTab] = useState<Tab>('home');
  const [target, setTarget] = useState<{ kind: DocKind; id: string } | null>(null);

  const openDoc = (kind: DocKind, id: string) => {
    setTarget({ kind, id });
    setTab(TAB_FOR_KIND[kind]);
  };

  const keyFor = (kind: DocKind): string => (target?.kind === kind ? target.id : 'new');

  return (
    <div className="app">
      <main className="screen-area">
        {tab === 'home' && <Home onOpen={openDoc} onGo={setTab} key="home" />}
        {tab === 'chat' && <Chat initialId={target?.kind === 'chat' ? target.id : undefined} key={keyFor('chat')} />}
        {tab === 'docs' && <Docs initialId={target?.kind === 'doc' ? target.id : undefined} key={keyFor('doc')} />}
        {tab === 'sheets' && <Sheets initialId={target?.kind === 'sheet' ? target.id : undefined} key={keyFor('sheet')} />}
        {tab === 'slides' && <Slides initialId={target?.kind === 'deck' ? target.id : undefined} key={keyFor('deck')} />}
        {tab === 'settings' && <Settings key="settings" />}
      </main>
      <nav className="tabbar">
        {TABS.map((t) => (
          <button key={t.id} className={tab === t.id ? 'tab active' : 'tab'} onClick={() => setTab(t.id)}>
            {t.label}
          </button>
        ))}
      </nav>
    </div>
  );
}
