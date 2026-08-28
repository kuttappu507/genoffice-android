import { useState } from 'react';
import type { DocKind } from './types';
import { Icon } from './components/Icon';
import Home from './screens/Home';
import Chat from './screens/Chat';
import Docs from './screens/Docs';
import Sheets from './screens/Sheets';
import Slides from './screens/Slides';
import Settings from './screens/Settings';

type Tab = 'home' | 'chat' | 'docs' | 'sheets' | 'slides' | 'settings';

const TABS: { id: Tab; label: string; icon: string }[] = [
  { id: 'home', label: 'Home', icon: 'home' },
  { id: 'docs', label: 'Docs', icon: 'fileText' },
  { id: 'sheets', label: 'Sheets', icon: 'grid' },
  { id: 'slides', label: 'Slides', icon: 'monitor' },
  { id: 'chat', label: 'AI', icon: 'chat' },
  { id: 'settings', label: 'Settings', icon: 'settings' },
];

/** Office brand accent per tab (Word blue / Excel green / PPT orange / Teams purple). */
const ACCENT: Record<Tab, string> = {
  home: '#0F6CBD',
  docs: '#185ABD',
  sheets: '#107C41',
  slides: '#C43E1C',
  chat: '#5B5FC7',
  settings: '#0F6CBD',
};

const EDITOR_TABS: Tab[] = ['docs', 'sheets', 'slides'];

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
  const isEditor = EDITOR_TABS.includes(tab);
  const exit = () => setTab('home');

  return (
    <div className="app" style={{ ['--accent' as string]: ACCENT[tab] }}>
      <main className={isEditor ? 'screen-area flush' : 'screen-area'}>
        {tab === 'home' && <Home onOpen={openDoc} onGo={setTab} key="home" />}
        {tab === 'chat' && <Chat initialId={target?.kind === 'chat' ? target.id : undefined} key={keyFor('chat')} />}
        {tab === 'docs' && <Docs initialId={target?.kind === 'doc' ? target.id : undefined} key={keyFor('doc')} onExit={exit} />}
        {tab === 'sheets' && <Sheets initialId={target?.kind === 'sheet' ? target.id : undefined} key={keyFor('sheet')} onExit={exit} />}
        {tab === 'slides' && <Slides initialId={target?.kind === 'deck' ? target.id : undefined} key={keyFor('deck')} onExit={exit} />}
        {tab === 'settings' && <Settings key="settings" />}
      </main>
      {!isEditor && (
        <nav className="tabbar">
          {TABS.map((t) => (
            <button key={t.id} className={tab === t.id ? 'tab active' : 'tab'} onClick={() => setTab(t.id)}>
              <Icon name={t.icon} size={21} strokeWidth={tab === t.id ? 2.3 : 2} />
              <span>{t.label}</span>
            </button>
          ))}
        </nav>
      )}
    </div>
  );
}
