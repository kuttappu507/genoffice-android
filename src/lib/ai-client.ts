import type { AISettings, ChatMsg } from '../types';

export class ApiError extends Error {
  status: number;
  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

const THINK_OPEN = '<think>';
const THINK_CLOSE = '</think>';

function partialTagLen(s: string, tag: string): number {
  const max = Math.min(tag.length - 1, s.length);
  for (let n = max; n > 0; n--) {
    if (s.endsWith(tag.slice(0, n))) return n;
  }
  return 0;
}

/**
 * Stateful filter that strips <think>...</think> blocks from model output,
 * even when the tags are split across streaming chunk boundaries.
 * Thinking text is captured separately so it can be shown in a collapsible
 * section instead of being sent back to the API (saves tokens).
 */
export class ThinkFilter {
  private buf = '';
  private inside = false;
  private reasoning = '';

  feed(chunk: string): string {
    this.buf += chunk;
    let out = '';
    for (;;) {
      if (!this.inside) {
        const i = this.buf.indexOf(THINK_OPEN);
        if (i === -1) {
          const keep = partialTagLen(this.buf, THINK_OPEN);
          out += this.buf.slice(0, this.buf.length - keep);
          this.buf = this.buf.slice(this.buf.length - keep);
          return out;
        }
        out += this.buf.slice(0, i);
        this.buf = this.buf.slice(i + THINK_OPEN.length);
        this.inside = true;
      } else {
        const j = this.buf.indexOf(THINK_CLOSE);
        if (j === -1) {
          const keep = partialTagLen(this.buf, THINK_CLOSE);
          this.reasoning += this.buf.slice(0, this.buf.length - keep);
          this.buf = this.buf.slice(this.buf.length - keep);
          return out;
        }
        this.reasoning += this.buf.slice(0, j);
        this.buf = this.buf.slice(j + THINK_CLOSE.length);
        this.inside = false;
      }
    }
  }

  flush(): string {
    const rest = this.inside ? '' : this.buf;
    this.buf = '';
    return rest;
  }

  drainedReasoning(): string {
    const r = this.reasoning;
    this.reasoning = '';
    return r;
  }
}

function endpoint(s: AISettings, path: string): string {
  return `${s.baseUrl.replace(/\/+$/, '')}${path}`;
}

function buildHeaders(s: AISettings): Record<string, string> {
  const h: Record<string, string> = { 'Content-Type': 'application/json' };
  if (s.apiKey) h['Authorization'] = `Bearer ${s.apiKey}`;
  if (s.provider === 'openrouter') {
    h['HTTP-Referer'] = 'https://github.com/kuttappu507/genoffice-android';
    h['X-Title'] = 'GenOffice Mobile';
  }
  return h;
}

function buildBody(s: AISettings, messages: ChatMsg[], stream: boolean, maxTokens?: number): string {
  const body: Record<string, unknown> = {
    model: s.model,
    messages: messages.map((m) => ({ role: m.role, content: m.content })),
    stream,
    temperature: s.temperature,
  };
  if (maxTokens) body['max_tokens'] = maxTokens;
  if (s.provider === 'openrouter' && s.excludeReasoning) body['reasoning'] = { exclude: true };
  return JSON.stringify(body);
}

function friendly(status: number, raw: string): string {
  if (status === 401) return 'Invalid or missing API key (401). Check Settings.';
  if (status === 402) return 'This model needs credits (402). Pick a ":free" model to stay on the free tier.';
  if (status === 404) return 'Model not found (404). Check the model id in Settings.';
  if (status === 429) return `Rate limited (429). Free tier limits hit; retrying later helps. ${raw.slice(0, 120)}`;
  return `HTTP ${status}: ${raw.slice(0, 180)}`;
}

const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

export interface StreamHandlers {
  onDelta?: (text: string) => void;
  onReasoning?: (text: string) => void;
  signal?: AbortSignal;
}

/**
 * Stream a chat completion with:
 * - 429/5xx backoff: 4 attempts, honors Retry-After (capped at 20s), else 1/2/4s
 * - reasoning.exclude for OpenRouter (fewer tokens billed/returned)
 * - think-tag stripping for models like Nemotron that leak <think> blocks
 * Returns the full visible text.
 */
export async function chatStream(s: AISettings, messages: ChatMsg[], h: StreamHandlers = {}): Promise<string> {
  let full = '';
  let lastErr: Error = new ApiError(0, 'request failed');
  for (let attempt = 0; attempt < 4; attempt++) {
    let res: Response;
    try {
      res = await fetch(endpoint(s, '/chat/completions'), {
        method: 'POST',
        headers: buildHeaders(s),
        body: buildBody(s, messages, true),
        signal: h.signal,
      });
    } catch (e) {
      if (h.signal?.aborted) throw e;
      lastErr = new ApiError(0, e instanceof Error ? e.message : 'network error');
      if (attempt < 3) {
        await sleep(1000 * 2 ** attempt);
        continue;
      }
      throw lastErr;
    }
    if (!res.ok) {
      const raw = await res.text().catch(() => '');
      const retryable = res.status === 429 || res.status >= 500;
      if (retryable && attempt < 3) {
        const ra = parseFloat(res.headers.get('retry-after') ?? '');
        const ms = Number.isFinite(ra) ? Math.min(ra * 1000, 20_000) : 1000 * 2 ** attempt;
        await sleep(ms);
        continue;
      }
      throw new ApiError(res.status, friendly(res.status, raw));
    }
    if (!res.body) throw new ApiError(0, 'Empty response body');
    const reader = res.body.getReader();
    const dec = new TextDecoder();
    const tf = new ThinkFilter();
    let buf = '';
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });
      const lines = buf.split('\n');
      buf = lines.pop() ?? '';
      for (const line of lines) {
        const t = line.trim();
        if (!t.startsWith('data:')) continue;
        const data = t.slice(5).trim();
        if (data === '[DONE]') continue;
        try {
          const j = JSON.parse(data) as {
            choices?: { delta?: { content?: string | null; reasoning_content?: string | null } }[];
            error?: { message?: string };
          };
          if (j.error?.message) throw new ApiError(0, j.error.message);
          const d = j.choices?.[0]?.delta;
          if (d?.reasoning_content) h.onReasoning?.(d.reasoning_content);
          if (d?.content) {
            const vis = tf.feed(d.content);
            if (vis) {
              full += vis;
              h.onDelta?.(vis);
            }
          }
        } catch (e) {
          if (e instanceof ApiError) throw e;
          // ignore malformed chunks
        }
      }
    }
    const tail = tf.flush();
    if (tail) {
      full += tail;
      h.onDelta?.(tail);
    }
    const r = tf.drainedReasoning();
    if (r) h.onReasoning?.(r);
    return full;
  }
  throw lastErr;
}

export interface TestResult {
  ok: boolean;
  latencyMs: number;
  reply: string;
}

export interface RemoteModel {
  id: string;
  label: string;
  free: boolean;
  context: number;
}

/**
 * Fetch the provider's live model catalogue (OpenAI-compatible GET /models).
 * OpenRouter works without a key (381+ models, `:free` ones detected via id
 * suffix or zero pricing); NVIDIA NIM and OpenAI-compatible local servers use
 * the same endpoint shape with Bearer auth.
 */
export async function listModels(s: AISettings): Promise<RemoteModel[]> {
  const res = await fetch(endpoint(s, '/models'), {
    method: 'GET',
    headers: buildHeaders(s),
  });
  if (!res.ok) {
    throw new ApiError(res.status, friendly(res.status, await res.text().catch(() => '')));
  }
  const j = (await res.json()) as {
    data?: {
      id?: string;
      name?: string;
      context_length?: number;
      pricing?: { prompt?: string; completion?: string };
    }[];
  };
  const arr = Array.isArray(j.data) ? j.data : [];
  const out: RemoteModel[] = [];
  for (const m of arr) {
    if (!m || typeof m.id !== 'string' || !m.id) continue;
    const prompt = Number(m.pricing?.prompt ?? '1');
    const comp = Number(m.pricing?.completion ?? '1');
    const zero = Number.isFinite(prompt) && Number.isFinite(comp) && prompt === 0 && comp === 0;
    out.push({
      id: m.id,
      label: typeof m.name === 'string' && m.name ? m.name : m.id,
      free: m.id.endsWith(':free') || zero,
      context: typeof m.context_length === 'number' ? m.context_length : 0,
    });
  }
  // free tier first, then alphabetical
  out.sort((a, b) => (a.free === b.free ? a.id.localeCompare(b.id) : a.free ? -1 : 1));
  return out;
}

/** Small non-streaming ping used by Settings "Test connection". */
export async function testConnection(s: AISettings): Promise<TestResult> {
  const t0 = Date.now();
  const res = await fetch(endpoint(s, '/chat/completions'), {
    method: 'POST',
    headers: buildHeaders(s),
    body: buildBody(s, [{ role: 'user', content: 'Reply with exactly: ok' }], false, 24),
  });
  if (!res.ok) {
    throw new ApiError(res.status, friendly(res.status, await res.text().catch(() => '')));
  }
  const j = (await res.json()) as { choices?: { message?: { content?: string } }[] };
  const raw = j.choices?.[0]?.message?.content ?? '';
  const tf = new ThinkFilter();
  const reply = (tf.feed(raw) + tf.flush()).trim().slice(0, 80);
  return { ok: true, latencyMs: Date.now() - t0, reply };
}

/** Keep the system prompt plus as many recent messages as fit the char budget. */
export function trimMessages(msgs: ChatMsg[], budget: number): ChatMsg[] {
  if (msgs.length === 0) return msgs;
  const sys = msgs[0]?.role === 'system' ? [msgs[0]] : [];
  const rest = msgs.slice(sys.length);
  let used = sys.reduce((n, m) => n + m.content.length, 0);
  const kept: ChatMsg[] = [];
  for (let i = rest.length - 1; i >= 0; i--) {
    const len = rest[i].content.length;
    if (used + len > budget && kept.length >= 2) break;
    used += len;
    kept.unshift(rest[i]);
  }
  return [...sys, ...kept];
}

export function errMsg(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
