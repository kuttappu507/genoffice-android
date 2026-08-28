import { useState } from 'react';
import type { Provider } from '../types';
import { Icon } from '../components/Icon';
import { listModels, testConnection, errMsg, type RemoteModel, type TestResult } from '../lib/ai-client';
import { PROVIDER_PRESETS, contextBudget } from '../lib/models';
import { downloadText, exportAll, getSettings, importAll, saveSettings } from '../lib/storage';

/** 1048576 -> "1M", 131072 -> "131K" - compact context-length for option labels. */
function fmtCtx(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(n % 1_000_000 === 0 ? 0 : 1)}M`;
  if (n >= 1000) return `${Math.round(n / 1000)}K`;
  return String(n);
}

export default function Settings() {
  const [s, setS] = useState(getSettings());
  const [showKey, setShowKey] = useState(false);
  const [testing, setTesting] = useState(false);
  const [result, setResult] = useState<TestResult | null>(null);
  const [error, setError] = useState('');
  const [toast, setToast] = useState('');
  const [remote, setRemote] = useState<RemoteModel[] | null>(null);
  const [loadingModels, setLoadingModels] = useState(false);
  const [freeOnly, setFreeOnly] = useState(false);
  const [modelErr, setModelErr] = useState('');

  const patch = (p: Partial<typeof s>) => {
    const next = { ...s, ...p };
    setS(next);
    saveSettings(next);
  };

  const pickProvider = (p: Provider) => {
    const preset = PROVIDER_PRESETS[p];
    patch({ provider: p, baseUrl: preset.baseUrl, model: preset.models[0]?.id ?? s.model });
    setResult(null);
    setError('');
    setRemote(null);
    setModelErr('');
    setFreeOnly(false);
  };

  const loadModels = async () => {
    setLoadingModels(true);
    setModelErr('');
    try {
      const list = await listModels(s);
      setRemote(list);
      if (list.length === 0) setModelErr('Provider returned an empty model list - type the id manually below.');
    } catch (e) {
      setModelErr(errMsg(e));
    } finally {
      setLoadingModels(false);
    }
  };

  const runTest = async () => {
    setTesting(true);
    setResult(null);
    setError('');
    try {
      setResult(await testConnection(s));
    } catch (e) {
      setError(errMsg(e));
    } finally {
      setTesting(false);
    }
  };

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  };

  const preset = PROVIDER_PRESETS[s.provider];

  return (
    <div className="screen">
      <header className="screen-head settings-head">
        <h2 className="screen-title">Settings</h2>
      </header>

      <section className="card">
        <h3><Icon name="cpu" size={17} /> AI provider (BYOK)</h3>
        <p className="hint">Your key is stored only on this device and sent only to the provider you choose.</p>
        <div className="btn-row">
          {(Object.keys(PROVIDER_PRESETS) as Provider[]).map((p) => (
            <button key={p} className={`btn small${s.provider === p ? ' primary' : ''}`} onClick={() => pickProvider(p)}>
              {PROVIDER_PRESETS[p].label}
            </button>
          ))}
        </div>

        <label className="field">
          <span>Base URL</span>
          <input
            className="input"
            value={s.baseUrl}
            onChange={(e) => {
              patch({ baseUrl: e.target.value });
              setRemote(null);
              setModelErr('');
            }}
          />
        </label>

        <label className="field">
          <span>API key ({preset.keyHint})</span>
          <div className="key-row">
            <input
              className="input"
              type={showKey ? 'text' : 'password'}
              value={s.apiKey}
              placeholder={preset.keyHint}
              onChange={(e) => patch({ apiKey: e.target.value })}
              autoComplete="off"
            />
            <button className="btn small" onClick={() => setShowKey(!showKey)}>
              {showKey ? 'Hide' : 'Show'}
            </button>
          </div>
          {preset.keysUrl && (
            <a className="hint link" href={preset.keysUrl} target="_blank" rel="noreferrer">
              Get a key at {preset.keysUrl.replace('https://', '')}
            </a>
          )}
        </label>

        <label className="field">
          <span>Model</span>
          {preset.models.length > 0 && (
            <select
              className="input"
              value={preset.models.some((m) => m.id === s.model) ? s.model : ''}
              onChange={(e) => e.target.value && patch({ model: e.target.value })}
            >
              <option value="" disabled>
                Custom / typed below
              </option>
              {preset.models.map((m) => (
                <option key={m.id} value={m.id}>
                  {m.label}
                </option>
              ))}
            </select>
          )}
          {remote && remote.length > 0 && (
            <select
              className="input"
              value={remote.some((m) => m.id === s.model) ? s.model : ''}
              onChange={(e) => e.target.value && patch({ model: e.target.value })}
            >
              <option value="" disabled>
                {freeOnly ? remote.filter((m) => m.free).length : remote.length} models loaded - pick one
              </option>
              {(freeOnly ? remote.filter((m) => m.free) : remote).map((m) => (
                <option key={m.id} value={m.id}>
                  {m.free ? '[free] ' : ''}
                  {m.label}
                  {m.context >= 1000 ? ` · ${fmtCtx(m.context)} ctx` : ''}
                </option>
              ))}
            </select>
          )}
          <div className="btn-row">
            <button type="button" className="btn small" disabled={loadingModels} onClick={() => void loadModels()}>
              {loadingModels ? 'Loading models...' : remote ? 'Refresh model list' : 'List models from my key'}
            </button>
            {remote && remote.some((m) => m.free) && (
              <button type="button" className={`btn small${freeOnly ? ' primary' : ''}`} onClick={() => setFreeOnly(!freeOnly)}>
                {freeOnly ? 'Free only: on' : 'Free only: off'}
              </button>
            )}
          </div>
          {remote && remote.length > 0 && (
            <p className="hint">
              {remote.length} models · {remote.filter((m) => m.free).length} free. Free ones are listed first.
            </p>
          )}
          {modelErr && <p className="err">{modelErr}</p>}
          <input className="input" value={s.model} onChange={(e) => patch({ model: e.target.value })} placeholder="model id (or type manually)" />
        </label>

        <label className="field">
          <span>Temperature: {s.temperature.toFixed(1)}</span>
          <input
            type="range"
            min={0}
            max={1.5}
            step={0.1}
            value={s.temperature}
            onChange={(e) => patch({ temperature: parseFloat(e.target.value) })}
          />
        </label>

        {s.provider === 'openrouter' && (
          <label className="field row">
            <input
              type="checkbox"
              checked={s.excludeReasoning}
              onChange={(e) => patch({ excludeReasoning: e.target.checked })}
            />
            <span>Exclude reasoning tokens (OpenRouter only - saves output tokens on free models)</span>
          </label>
        )}

        <div className="btn-row">
          <button className="btn primary" disabled={testing || !s.apiKey} onClick={() => void runTest()}>
            {testing ? 'Testing...' : 'Test connection'}
          </button>
        </div>
        {result && (
          <p className="ok">
            Connected. {result.latencyMs} ms - model replied: {result.reply || '(empty)'}
          </p>
        )}
        {error && <p className="err">{error}</p>}

        <p className="hint">
          History sent per chat is capped at about {contextBudget(s.provider, s.model).toLocaleString()} characters
          (free-tier friendly). Reasoning text is never sent back to the model.
        </p>
      </section>

      <section className="card">
        <h3><Icon name="database" size={17} /> Data</h3>
        <p className="hint">Everything lives in on-device storage. Export a JSON backup to move devices.</p>
        <div className="btn-row">
          <button
            className="btn"
            onClick={() => void downloadText(`genoffice-backup-${Date.now()}.json`, exportAll(), 'application/json').then(flash)}
          >
            Export backup
          </button>
          <label className="btn file-btn">
            Import backup
            <input
              type="file"
              accept="application/json"
              style={{ display: 'none' }}
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                const reader = new FileReader();
                reader.onload = () => {
                  try {
                    const n = importAll(String(reader.result));
                    flash(`Imported ${n} documents. Reload to see them.`);
                  } catch (err) {
                    flash(`Import failed: ${errMsg(err)}`);
                  }
                };
                reader.readAsText(f);
              }}
            />
          </label>
        </div>
      </section>

      <section className="card">
        <h3><Icon name="info" size={17} /> About</h3>
        <p className="hint">
          GenOffice Mobile - the Android port of the GenOffice AI office suite. BYOK (bring your own
          key) via OpenRouter or NVIDIA NIM. No account, no login, no telemetry. Documents stay on
          your device.
        </p>
      </section>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
