export type Provider = 'openrouter' | 'nvidia' | 'custom';

export interface AISettings {
  provider: Provider;
  baseUrl: string;
  apiKey: string;
  model: string;
  temperature: number;
  /** OpenRouter only: ask the provider not to return reasoning tokens. */
  excludeReasoning: boolean;
}

export type ThemeMode = 'system' | 'light' | 'dark';

/** Device-level preferences (appearance, feedback, editor defaults). */
export interface AppPrefs {
  theme: ThemeMode;
  haptics: boolean;
  /** currency symbol used by the Sheets "Currency" number format */
  currency: string;
  /** default font family for new documents */
  docFont: string;
  /** Docs view: reflowed mobile view or page-sized print layout */
  docView: 'mobile' | 'print';
  /** first-run coach marks dismissed */
  onboarded: boolean;
}

export type DocKind = 'doc' | 'sheet' | 'deck' | 'chat';

export interface DocMeta {
  id: string;
  kind: DocKind;
  title: string;
  created: number;
  updated: number;
  /** pinned to the top of Home */
  pinned?: boolean;
}

export interface ChatMsg {
  role: 'system' | 'user' | 'assistant';
  content: string;
  /** Display-only; never sent back to the API (saves tokens). */
  reasoning?: string;
}
