import type { AISettings, Provider } from '../types';

export interface ModelPreset {
  id: string;
  label: string;
}

export interface ProviderPreset {
  label: string;
  baseUrl: string;
  keyHint: string;
  keysUrl: string;
  models: ModelPreset[];
}

export const PROVIDER_PRESETS: Record<Provider, ProviderPreset> = {
  openrouter: {
    label: 'OpenRouter',
    baseUrl: 'https://openrouter.ai/api/v1',
    keyHint: 'sk-or-v1-...',
    keysUrl: 'https://openrouter.ai/keys',
    models: [
      { id: 'deepseek/deepseek-chat-v3-0324:free', label: 'DeepSeek V3 0324 (free)' },
      { id: 'meta-llama/llama-3.3-70b-instruct:free', label: 'Llama 3.3 70B (free)' },
      { id: 'google/gemini-2.0-flash-exp:free', label: 'Gemini 2.0 Flash (free)' },
      { id: 'qwen/qwen-2.5-72b-instruct:free', label: 'Qwen 2.5 72B (free)' },
      { id: 'nvidia/nemotron-nano-9b-v2:free', label: 'Nemotron Nano 9B (free)' },
    ],
  },
  nvidia: {
    label: 'NVIDIA NIM',
    baseUrl: 'https://integrate.api.nvidia.com/v1',
    keyHint: 'nvapi-...',
    keysUrl: 'https://build.nvidia.com',
    models: [
      { id: 'meta/llama-3.3-70b-instruct', label: 'Llama 3.3 70B' },
      { id: 'nvidia/llama-3.1-nemotron-70b-instruct', label: 'Nemotron 70B' },
      { id: 'deepseek-ai/deepseek-r1', label: 'DeepSeek R1' },
      { id: 'mistralai/mistral-small-24b-instruct', label: 'Mistral Small 24B' },
    ],
  },
  custom: {
    label: 'Custom (OpenAI-compatible)',
    baseUrl: 'http://localhost:1234/v1',
    keyHint: 'optional',
    keysUrl: '',
    models: [],
  },
};

export const DEFAULT_SETTINGS: AISettings = {
  provider: 'openrouter',
  baseUrl: PROVIDER_PRESETS.openrouter.baseUrl,
  apiKey: '',
  model: 'deepseek/deepseek-chat-v3-0324:free',
  temperature: 0.7,
  excludeReasoning: true,
};

/**
 * Character budget used to trim conversation history before sending.
 * Free tiers have small context windows, so we keep history short there.
 * This is a key part of "less tokens, same useful output".
 */
export function contextBudget(provider: Provider, model: string): number {
  const m = model.toLowerCase();
  if (m.includes(':free')) return 60_000;
  if (provider === 'nvidia') return 60_000;
  return 200_000;
}
