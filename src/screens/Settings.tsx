import { useState } from 'react';
import type { Provider } from '../types';
import { testConnection, errMsg, TestResult } from '../lib/ai-client';
import { PROVIDER_PRESETS, contextBudget } from '../lib/models';
import { downloadText, exportAll, getSettings, importAll, saveSettings } from '../lib/storage';

export default function Settings() {
  const [s, setS] = useState(getSettings());
  const [showKey, setShowKey] = useState(false);
  const [testing, setTesting] = useState(false);
  const [result, setResult] = useState<TestResult | null>(null);
  const [error, setError] = useState('');
  const [toast, setToast] = useState('');

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
      <header className="screen-head">
        <h2 className="screen-title">Settings</h2>
      </header>

      <section className="card">
        <h3>AI provider (BYOK)</h3>
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
          <input className="input" value={s.baseUrl} onChange={(e) => patch({ baseUrl: e.target.value })} />
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
          <input className="input" value={s.model} onChange={(e) => patch({ model: e.target.value })} placeholder="model id" />
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
        <h3>Data</h3>
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
        <h3>About</h3>
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
