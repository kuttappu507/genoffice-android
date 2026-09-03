import { Capacitor } from '@capacitor/core';
import { Directory, Encoding, Filesystem } from '@capacitor/filesystem';
import type { AISettings, AppPrefs, DocKind, DocMeta } from '../types';
import { DEFAULT_SETTINGS } from './models';

const IDX = 'genoffice:index';
const SETTINGS_KEY = 'genoffice:settings';
const PREFS_KEY = 'genoffice:prefs';

export function uid(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function loadIndex(): DocMeta[] {
  try {
    return JSON.parse(localStorage.getItem(IDX) ?? '[]') as DocMeta[];
  } catch {
    return [];
  }
}

function saveIndex(list: DocMeta[]): void {
  localStorage.setItem(IDX, JSON.stringify(list));
}

/** Pinned first, then most recently updated. */
export function listDocs(kind?: DocKind): DocMeta[] {
  return loadIndex()
    .filter((d) => !kind || d.kind === kind)
    .sort((a, b) => Number(!!b.pinned) - Number(!!a.pinned) || b.updated - a.updated);
}

export function getMeta(id: string): DocMeta | undefined {
  return loadIndex().find((d) => d.id === id);
}

export function getDoc<T>(id: string): T | null {
  try {
    const raw = localStorage.getItem(`genoffice:doc:${id}`);
    return raw ? (JSON.parse(raw) as T) : null;
  } catch {
    return null;
  }
}

export function putDoc<T>(kind: DocKind, id: string, title: string, data: T): DocMeta {
  const now = Date.now();
  const list = loadIndex();
  const prev = list.find((d) => d.id === id);
  const meta: DocMeta = {
    id,
    kind,
    title: title.trim() || 'Untitled',
    created: prev?.created ?? now,
    updated: now,
    pinned: prev?.pinned,
  };
  const i = list.findIndex((d) => d.id === id);
  if (i >= 0) list[i] = meta;
  else list.unshift(meta);
  saveIndex(list);
  localStorage.setItem(`genoffice:doc:${id}`, JSON.stringify(data));
  return meta;
}

export function removeDoc(id: string): void {
  saveIndex(loadIndex().filter((d) => d.id !== id));
  localStorage.removeItem(`genoffice:doc:${id}`);
}

export function renameDoc(id: string, title: string): void {
  const list = loadIndex();
  const m = list.find((d) => d.id === id);
  if (!m) return;
  m.title = title.trim() || 'Untitled';
  saveIndex(list);
}

export function togglePin(id: string): boolean {
  const list = loadIndex();
  const m = list.find((d) => d.id === id);
  if (!m) return false;
  m.pinned = !m.pinned;
  saveIndex(list);
  return !!m.pinned;
}

/** Copy a document (same kind) under a new id; returns the new meta. */
export function duplicateDoc(id: string): DocMeta | null {
  const m = getMeta(id);
  const data = getDoc<unknown>(id);
  if (!m || data === null) return null;
  return putDoc(m.kind, uid(), `${m.title} (copy)`, data);
}

/** Approximate bytes used by app data in localStorage. */
export function storageUsage(): { bytes: number; docs: number } {
  let bytes = 0;
  let docs = 0;
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k || !k.startsWith('genoffice:')) continue;
    const v = localStorage.getItem(k) ?? '';
    bytes += (k.length + v.length) * 2;
    if (k.startsWith('genoffice:doc:')) docs++;
  }
  return { bytes, docs };
}

export interface Backup {
  version: number;
  exportedAt: string;
  index: DocMeta[];
  docs: unknown[];
  settings?: AISettings;
  prefs?: AppPrefs;
}

export function exportAll(): string {
  const index = loadIndex();
  const backup: Backup = {
    version: 2,
    exportedAt: new Date().toISOString(),
    index,
    docs: index.map((d) => getDoc(d.id)),
    settings: getSettings(),
    prefs: getPrefs(),
  };
  return JSON.stringify(backup, null, 2);
}

export function importAll(json: string): number {
  const j = JSON.parse(json) as Partial<Backup>;
  if (!j.index || !j.docs) throw new Error('Not a GenOffice backup file');
  // merge: imported docs win on id collisions, existing ones are kept
  const existing = loadIndex().filter((d) => !j.index!.some((n) => n.id === d.id));
  saveIndex([...j.index, ...existing]);
  j.index.forEach((m, i) => {
    localStorage.setItem(`genoffice:doc:${m.id}`, JSON.stringify(j.docs?.[i] ?? null));
  });
  if (j.settings && !getSettings().apiKey) saveSettings({ ...DEFAULT_SETTINGS, ...j.settings });
  if (j.prefs) savePrefs({ ...DEFAULT_PREFS, ...j.prefs });
  return j.index.length;
}

export function clearAllData(): void {
  const keys: string[] = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.startsWith('genoffice:') && k !== SETTINGS_KEY && k !== PREFS_KEY) keys.push(k);
  }
  keys.forEach((k) => localStorage.removeItem(k));
}

export async function downloadText(name: string, text: string, mime = 'text/plain'): Promise<string> {
  if (Capacitor.isNativePlatform()) {
    await Filesystem.writeFile({
      path: name,
      data: text,
      directory: Directory.Documents,
      encoding: Encoding.UTF8,
      recursive: true,
    });
    return `Saved to device Documents/${name}`;
  }
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  return `Downloaded ${name}`;
}

export function debounce<A extends unknown[]>(fn: (...a: A) => void, ms: number): (...a: A) => void {
  let t: ReturnType<typeof setTimeout> | undefined;
  return (...a: A) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn(...a), ms);
  };
}

export function getSettings(): AISettings {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (!raw) return DEFAULT_SETTINGS;
    return { ...DEFAULT_SETTINGS, ...(JSON.parse(raw) as Partial<AISettings>) };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

export function saveSettings(s: AISettings): void {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
}

// ---------------------------------------------------------------------------
// App preferences (theme, haptics, editor defaults)
// ---------------------------------------------------------------------------

export const DEFAULT_PREFS: AppPrefs = {
  theme: 'system',
  haptics: true,
  currency: '$',
  docFont: 'Calibri',
  docView: 'mobile',
  onboarded: false,
};

const PREFS_EVENT = 'genoffice:prefs-changed';

export function getPrefs(): AppPrefs {
  try {
    const raw = localStorage.getItem(PREFS_KEY);
    if (!raw) return DEFAULT_PREFS;
    return { ...DEFAULT_PREFS, ...(JSON.parse(raw) as Partial<AppPrefs>) };
  } catch {
    return DEFAULT_PREFS;
  }
}

export function savePrefs(p: AppPrefs): void {
  localStorage.setItem(PREFS_KEY, JSON.stringify(p));
  window.dispatchEvent(new CustomEvent(PREFS_EVENT));
}

export function onPrefsChange(fn: () => void): () => void {
  window.addEventListener(PREFS_EVENT, fn);
  return () => window.removeEventListener(PREFS_EVENT, fn);
}

/** Apply the light/dark theme to <html> so CSS variables switch. */
export function applyTheme(mode: AppPrefs['theme']): void {
  const prefersDark = window.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false;
  const dark = mode === 'dark' || (mode === 'system' && prefersDark);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.setAttribute('content', dark ? '#1b1b1f' : '#f5f5f5');
}
