import { Capacitor } from '@capacitor/core';
import { Directory, Encoding, Filesystem } from '@capacitor/filesystem';
import type { AISettings, DocKind, DocMeta } from '../types';
import { DEFAULT_SETTINGS } from './models';

const IDX = 'genoffice:index';
const SETTINGS_KEY = 'genoffice:settings';

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

export function listDocs(kind?: DocKind): DocMeta[] {
  return loadIndex()
    .filter((d) => !kind || d.kind === kind)
    .sort((a, b) => b.updated - a.updated);
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

export interface Backup {
  version: number;
  exportedAt: string;
  index: DocMeta[];
  docs: unknown[];
}

export function exportAll(): string {
  const index = loadIndex();
  const backup: Backup = {
    version: 1,
    exportedAt: new Date().toISOString(),
    index,
    docs: index.map((d) => getDoc(d.id)),
  };
  return JSON.stringify(backup, null, 2);
}

export function importAll(json: string): number {
  const j = JSON.parse(json) as Partial<Backup>;
  if (!j.index || !j.docs) throw new Error('Not a GenOffice backup file');
  saveIndex(j.index);
  j.index.forEach((m, i) => {
    localStorage.setItem(`genoffice:doc:${m.id}`, JSON.stringify(j.docs?.[i] ?? null));
  });
  return j.index.length;
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
