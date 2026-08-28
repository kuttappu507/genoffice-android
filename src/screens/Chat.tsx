import { useEffect, useRef, useState } from 'react';
import type { ChatMsg } from '../types';
import { chatStream, errMsg, trimMessages } from '../lib/ai-client';
import { contextBudget } from '../lib/models';
import { debounce, getDoc, getSettings, listDocs, putDoc, removeDoc, uid } from '../lib/storage';
import { Markdown } from '../lib/markdown';

interface Conv {
  msgs: ChatMsg[];
}

const SYSTEM_PROMPT =
  'You are the GenOffice mobile assistant. Answer in the user\'s language. ' +
  'Be concise and dense: no preamble, no restating the question, no filler. ' +
  'Use markdown (lists, bold, code fences) when it helps. If the request is ' +
  'ambiguous, state your assumption in one line and answer anyway.';

export default function Chat({ initialId }: { initialId?: string }) {
  const [convId, setConvId] = useState<string>(initialId ?? uid());
  const [title, setTitle] = useState(initialId ? 'Conversation' : 'New chat');
  const [msgs, setMsgs] = useState<ChatMsg[]>([]);
  const [input, setInput] = useState('');
  const [busy, setBusy] = useState(false);
  const [showList, setShowList] = useState(false);
  const abortRef = useRef<AbortController | null>(null);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!initialId) return;
    const d = getDoc<Conv>(initialId); // load saved conversation
    if (d) {
      setMsgs(d.msgs);
      const meta = listDocs('chat').find((c) => c.id === initialId);
      if (meta) setTitle(meta.title);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const save = useRef(
    debounce((id: string, t: string, m: ChatMsg[]) => {
      if (m.length > 0) putDoc<Conv>('chat', id, t, { msgs: m });
    }, 700),
  ).current;

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [msgs]);

  const send = async () => {
    const text = input.trim();
    if (!text || busy) return;
    const s = getSettings();
    if (!s.apiKey) {
      setMsgs((m) => [...m, { role: 'assistant', content: 'Add your API key in Settings first (OpenRouter or NVIDIA free tier works).' }]);
      return;
    }
    setInput('');
    const userMsg: ChatMsg = { role: 'user', content: text };
    const history = [...msgs, userMsg];
    const aiMsg: ChatMsg = { role: 'assistant', content: '' };
    setMsgs([...history, aiMsg]);
    setBusy(true);
    setTitle(text.slice(0, 40));

    const ac = new AbortController();
    abortRef.current = ac;
    let acc = '';
    let reasoning = '';
    try {
      const budget = contextBudget(s.provider, s.model);
      const payload = trimMessages([{ role: 'system', content: SYSTEM_PROMPT }, ...history], budget);
      await chatStream(s, payload, {
        signal: ac.signal,
        onDelta: (d) => {
          acc += d;
          setMsgs([...history, { ...aiMsg, content: acc }]);
        },
        onReasoning: (d) => {
          reasoning += d;
          setMsgs([...history, { ...aiMsg, content: acc, reasoning }]);
        },
      });
    } catch (e) {
      const aborted = ac.signal.aborted;
      if (!aborted) acc = acc || `Error: ${errMsg(e)}`;
      setMsgs([...history, { ...aiMsg, content: acc, reasoning: reasoning || undefined }]);
    } finally {
      setBusy(false);
      abortRef.current = null;
      const finalMsgs = [...history, { ...aiMsg, content: acc, reasoning: reasoning || undefined }];
      setMsgs(finalMsgs);
      save(convId, title, finalMsgs);
    }
  };

  const stop = () => abortRef.current?.abort();

  const newChat = () => {
    stop();
    setConvId(uid());
    setTitle('New chat');
    setMsgs([]);
    setInput('');
  };

  const chats = listDocs('chat');

  return (
    <div className="screen chat-screen">
      <header className="screen-head">
        <input className="title-input" value={title} onChange={(e) => { setTitle(e.target.value); }} placeholder="Chat title" />
        <button className="btn small" onClick={() => setShowList(!showList)}>
          {showList ? 'Close' : 'Chats'}
        </button>
        <button className="btn small primary" onClick={newChat}>
          New
        </button>
      </header>

      {showList && (
        <div className="list panel">
          {chats.length === 0 && <p className="empty">No saved chats yet.</p>}
          {chats.map((c) => (
            <div key={c.id} className="list-row">
              <button
                className="list-title as-btn"
                onClick={() => {
                  const d = getDoc<Conv>(c.id);
                  setConvId(c.id);
                  setTitle(c.title);
                  setMsgs(d?.msgs ?? []);
                  setShowList(false);
                }}
              >
                {c.title}
              </button>
              <button className="btn small danger" onClick={() => { removeDoc(c.id); if (c.id === convId) newChat(); }}>
                Del
              </button>
            </div>
          ))}
        </div>
      )}

      <div className="msgs">
        {msgs.length === 0 && (
          <p className="empty">
            Ask anything. Streaming responses, your own API key, history trimmed automatically to fit
            free-tier context windows.
          </p>
        )}
        {msgs.map((m, i) => (
          <div key={i} className={`msg ${m.role}`}>
            {m.role === 'assistant' ? (
              <>
                {m.reasoning && (
                  <details className="reasoning">
                    <summary>Thinking (not sent back to the model)</summary>
                    <div className="reasoning-body">{m.reasoning}</div>
                  </details>
                )}
                <Markdown text={m.content || '...'} />
              </>
            ) : (
              m.content
            )}
          </div>
        ))}
        <div ref={bottomRef} />
      </div>

      <div className="composer">
        <textarea
          className="input"
          rows={2}
          value={input}
          placeholder={busy ? 'Generating...' : 'Message'}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              void send();
            }
          }}
        />
        {busy ? (
          <button className="btn danger" onClick={stop}>
            Stop
          </button>
        ) : (
          <button className="btn primary" onClick={() => void send()} disabled={!input.trim()}>
            Send
          </button>
        )}
      </div>
    </div>
  );
}
