import { useEffect, useState } from 'react';
import type { DocKind } from './types';
import { Icon } from './components/Icon';
import Home from './screens/Home';
import Chat from './screens/Chat';
import Docs from './screens/Docs';
import Sheets from './screens/Sheets';
import Slides from './screens/Slides';
import Settings from './screens/Settings';
import Onboarding from './screens/Onboarding';
import { applyTheme, getPrefs, onPrefsChange } from './lib/storage';
import { onBack, setStatusBar } from './lib/native';

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

function isDarkNow(): boolean {
  return document.documentElement.dataset.theme === 'dark';
}

export default function App() {
  const [tab, setTab] = useState<Tab>('home');
  const [target, setTarget] = useState<{ kind: DocKind; id: string } | null>(null);
  const [onboarded, setOnboarded] = useState(() => getPrefs().onboarded);
  const [, setThemeTick] = useState(0);

  // Theme: apply on boot, follow prefs + OS changes.
  useEffect(() => {
    applyTheme(getPrefs().theme);
    const off = onPrefsChange(() => {
      applyTheme(getPrefs().theme);
      setOnboarded(getPrefs().onboarded);
      setThemeTick((t) => t + 1);
    });
    const mq = window.matchMedia('(prefers-color-scheme: dark)');
    const onMq = () => {
      applyTheme(getPrefs().theme);
      setThemeTick((t) => t + 1);
    };
    mq.addEventListener('change', onMq);
    return () => {
      off();
      mq.removeEventListener('change', onMq);
    };
  }, []);

  // Status bar colour follows the active app (editors paint their own bar).
  const isEditor = EDITOR_TABS.includes(tab);
  useEffect(() => {
    if (isEditor) void setStatusBar(ACCENT[tab], true);
    else void setStatusBar(isDarkNow() ? '#1b1b1f' : '#f5f5f5', isDarkNow());
  });

  // Hardware back: any non-home tab returns home; on home let Android minimise the app.
  useEffect(() => {
    return onBack(() => {
      if (tab !== 'home') {
        setTab('home');
        return true;
      }
      return false;
    });
  }, [tab]);

  const openDoc = (kind: DocKind, id: string) => {
    setTarget({ kind, id });
    setTab(TAB_FOR_KIND[kind]);
  };

  const keyFor = (kind: DocKind): string => (target?.kind === kind ? target.id : 'new');
  const exit = () => {
    setTarget(null);
    setTab('home');
  };
  const go = (t: Tab) => {
    if (EDITOR_TABS.includes(t) || t === 'chat') setTarget(null);
    setTab(t);
  };

  if (!onboarded) return <Onboarding onDone={() => setOnboarded(true)} />;

  return (
    <div className="app" style={{ ['--accent' as string]: ACCENT[tab] }}>
      <main className={isEditor ? 'screen-area flush' : 'screen-area'}>
        {tab === 'home' && <Home onOpen={openDoc} onGo={go} key="home" />}
        {tab === 'chat' && <Chat initialId={target?.kind === 'chat' ? target.id : undefined} key={keyFor('chat')} />}
        {tab === 'docs' && <Docs initialId={target?.kind === 'doc' ? target.id : undefined} key={keyFor('doc')} onExit={exit} />}
        {tab === 'sheets' && <Sheets initialId={target?.kind === 'sheet' ? target.id : undefined} key={keyFor('sheet')} onExit={exit} />}
        {tab === 'slides' && <Slides initialId={target?.kind === 'deck' ? target.id : undefined} key={keyFor('deck')} onExit={exit} />}
        {tab === 'settings' && <Settings key="settings" />}
      </main>
      {!isEditor && (
        <nav className="tabbar">
          {TABS.map((t) => (
            <button key={t.id} className={tab === t.id ? 'tab active' : 'tab'} onClick={() => go(t.id)}>
              <Icon name={t.icon} size={21} strokeWidth={tab === t.id ? 2.3 : 2} />
              <span>{t.label}</span>
            </button>
          ))}
        </nav>
      )}
    </div>
  );
}
