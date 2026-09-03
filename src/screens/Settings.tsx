import { useMemo, useState } from 'react';
import type { AppPrefs, Provider, ThemeMode } from '../types';
import { Icon } from '../components/Icon';
import { ConfirmSheet, Toast, useToast } from '../components/Sheet';
import { RSeg } from '../components/Ribbon';
import { listModels, testConnection, errMsg, type RemoteModel, type TestResult } from '../lib/ai-client';
import { PROVIDER_PRESETS, contextBudget } from '../lib/models';
import { applyTheme, clearAllData, downloadText, exportAll, getPrefs, getSettings, importAll, savePrefs, saveSettings, storageUsage } from '../lib/storage';
import { isNative, tap } from '../lib/native';

const CURRENCIES = ['$', '€', '£', '₹', '¥', '₩', 'R$', 'CHF ', 'A$', 'C$'];
const DOC_FONTS = ['Calibri', 'Segoe UI', 'Arial', 'Times New Roman', 'Georgia', 'Cambria', 'Verdana', 'Roboto'];

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
  const [toast, flash] = useToast();
  const [prefs, setPrefs] = useState<AppPrefs>(() => getPrefs());
  const [confirmClear, setConfirmClear] = useState(false);
  const [dataTick, setDataTick] = useState(0);
  const usage = useMemo(() => storageUsage(), [dataTick]); // eslint-disable-line react-hooks/exhaustive-deps
  const setPref = <K extends keyof AppPrefs>(k: K, v: AppPrefs[K]) => {
    const next = { ...prefs, [k]: v };
    setPrefs(next);
    savePrefs(next);
    if (k === 'theme') applyTheme(v as ThemeMode);
    if (k === 'haptics' && v) void tap('medium');
  };
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
        <h3><Icon name="theme" size={17} /> Appearance</h3>
        <label className="field">
          <span>Theme</span>
          <RSeg value={prefs.theme} options={[{ v: 'system', t: 'System', icon: 'phone' }, { v: 'light', t: 'Light', icon: 'sun' }, { v: 'dark', t: 'Dark', icon: 'moon' }]} onChange={(v) => setPref('theme', v)} />
        </label>
        <label className="field">
          <span>Default document view</span>
          <RSeg value={prefs.docView} options={[{ v: 'mobile', t: 'Mobile (reflow)', icon: 'phone' }, { v: 'print', t: 'Print layout', icon: 'pageSize' }]} onChange={(v) => setPref('docView', v)} />
        </label>
        <label className="field">
          <span>Default document font</span>
          <select className="input" value={prefs.docFont} onChange={(e) => setPref('docFont', e.target.value)}>
            {DOC_FONTS.map((f) => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
          </select>
        </label>
        <label className="field">
          <span>Currency symbol (spreadsheets)</span>
          <div className="chip-row">
            {CURRENCIES.map((c) => <button key={c} className={`chip${prefs.currency === c ? ' on' : ''}`} onClick={() => setPref('currency', c)}>{c.trim()}</button>)}
          </div>
        </label>
        <label className="field row">
          <input type="checkbox" checked={prefs.haptics} onChange={(e) => setPref('haptics', e.target.checked)} />
          <span>Haptic feedback on taps{isNative() ? '' : ' (Android app only)'}</span>
        </label>
      </section>

      <section className="card">
        <h3><Icon name="database" size={17} /> Data</h3>
        <p className="hint">Everything lives in on-device storage ({usage.docs} file{usage.docs === 1 ? '' : 's'}, {(usage.bytes / 1024).toFixed(0)} KB). Export a JSON backup to move devices — it includes your files, preferences and provider settings.</p>
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
                    setDataTick((t) => t + 1);
                    setPrefs(getPrefs());
                    setS(getSettings());
                    flash(`Imported ${n} document${n === 1 ? '' : 's'}.`);
                  } catch (err) {
                    flash(`Import failed: ${errMsg(err)}`);
                  }
                };
                reader.readAsText(f);
              }}
            />
          </label>
          <button className="btn danger" onClick={() => setConfirmClear(true)}>Delete all files</button>
        </div>
      </section>

      <section className="card">
        <h3><Icon name="info" size={17} /> About</h3>
        <p className="hint">
          GenOffice Mobile — the Android port of the GenOffice AI office suite: Word-, Excel- and PowerPoint-style
          editors with real .docx / .xlsx / .pptx / PDF import & export, a 120-function formula engine and BYOK
          (bring your own key) AI via OpenRouter or NVIDIA NIM. No account, no login, no telemetry. Documents stay on
          your device.
        </p>
        <button className="btn small" onClick={() => setPref('onboarded', false)}>Show welcome tour again</button>
      </section>

      <ConfirmSheet open={confirmClear} title="Delete all files on this device?" message="Documents, spreadsheets, presentations and chats will be removed. Your API key and preferences are kept. Export a backup first if you need one." confirmLabel="Delete everything" onConfirm={() => { clearAllData(); setDataTick((t) => t + 1); flash('All files deleted.'); }} onClose={() => setConfirmClear(false)} />
      <Toast msg={toast} />
    </div>
  );
}
