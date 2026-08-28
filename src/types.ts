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

export type DocKind = 'doc' | 'sheet' | 'deck' | 'chat';

export interface DocMeta {
  id: string;
  kind: DocKind;
  title: string;
  created: number;
  updated: number;
}

export interface ChatMsg {
  role: 'system' | 'user' | 'assistant';
  content: string;
  /** Display-only; never sent back to the API (saves tokens). */
  reasoning?: string;
}
