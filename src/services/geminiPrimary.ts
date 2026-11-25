import type { AiProviderId } from '../types/mail';
import { recordLlmUsageEvent } from './llmUsageTracker';

export const DEFAULT_GEMINI_MODEL = 'gemini-4-hydrogen';
export const GEMINI_PRIMARY_TIMEOUT_MS = 20_000;
const GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta/models';

export interface GeminiUsageMetadata {
  promptTokenCount?: number;
  candidatesTokenCount?: number;
  totalTokenCount?: number;
  thoughtsTokenCount?: number;
}

export interface GeminiPrimaryParams {
  prompt: string;
  systemInstruction?: string;
  model?: string;
  apiKey?: string;
  temperature?: number;
  thinkingLevel?: 'minimal' | 'low' | 'medium' | 'high';
  responseMimeType?: string;
  timeoutMs?: number;
  fetchImpl?: typeof fetch;
}

export interface GeminiPrimaryResult {
  text: string;
  model: string;
  modelId: string;
  provider: 'gemini-primary';
  usageMetadata?: GeminiUsageMetadata;
}

export interface SecondaryProviderResponse {
  text: string;
  model: string;
  providerUsed?: AiProviderId;
  usageMetadata?: GeminiUsageMetadata;
}

export interface PriorityAnalysisResult {
  text: string;
  model: string;
  providerUsed: AiProviderId;
  usedPrimary: boolean;
  usageMetadata?: GeminiUsageMetadata;
  fallbackReason?: string;
}

/**
 * Strips any accidental API key strings (AIza...) or raw key values from error messages/logs.
 */
export function sanitizeGeminiErrorText(message: string, apiKey?: string): string {
  let clean = String(message || 'Unknown error');
  if (apiKey && apiKey.trim().length >= 6) {
    clean = clean.split(apiKey.trim()).join('[REDACTED_KEY]');
  }
  clean = clean.replace(/AIza[0-9A-Za-z_-]{8,}/g, '[REDACTED_KEY]');
  clean = clean.replace(/([?&]key=)[^&\s]+/gi, '$1[REDACTED_KEY]');
  return clean.slice(0, 240);
}

export function isValidGeminiApiKey(raw?: string): boolean {
  const key = (raw ?? '').trim();
  return key.length > 0 && key !== 'MY_GEMINI_API_KEY';
}

export function resolveGeminiModel(override?: string): string {
  const candidate = (override || process.env.GEMINI_MODEL || DEFAULT_GEMINI_MODEL).trim();
  return candidate || DEFAULT_GEMINI_MODEL;
}

/**
 * Priority-1 Gemini 4 Hydrogen client built strictly on standard fetch + AbortController (20s timeout).
 * Passes GEMINI_API_KEY exclusively via the `x-goog-api-key` HTTP header (never URL query params).
 */
export async function geminiPrimary(params: GeminiPrimaryParams): Promise<GeminiPrimaryResult> {
  const apiKey = (params.apiKey ?? process.env.GEMINI_API_KEY ?? '').trim();
  if (!isValidGeminiApiKey(apiKey)) {
    throw new Error('GEMINI_API_KEY 환경변수가 설정되지 않았거나 유효하지 않습니다.');
  }

  const promptText = String(params.prompt || '').trim();
  if (!promptText) {
    throw new Error('Gemini 요청 프롬프트가 비어 있습니다.');
  }

  const modelId = resolveGeminiModel(params.model);
  const endpoint = `${GEMINI_API_BASE}/${encodeURIComponent(modelId)}:generateContent`;
  const timeoutMs =
    typeof params.timeoutMs === 'number' && params.timeoutMs > 0
      ? params.timeoutMs
      : GEMINI_PRIMARY_TIMEOUT_MS;
  const fetchFn = params.fetchImpl ?? fetch;

  const generationConfig: Record<string, unknown> = {
    temperature: typeof params.temperature === 'number' ? params.temperature : 0.3,
    thinkingConfig: {
      thinkingLevel: params.thinkingLevel ?? 'low',
    },
  };
  if (params.responseMimeType) {
    generationConfig.responseMimeType = params.responseMimeType;
  }

  const requestBody: Record<string, unknown> = {
    contents: [
      {
        role: 'user',
        parts: [{ text: promptText }],
      },
    ],
    generationConfig,
  };

  if (params.systemInstruction && params.systemInstruction.trim()) {
    requestBody.systemInstruction = {
      parts: [{ text: params.systemInstruction.trim() }],
    };
  }

  const controller = new AbortController();
  const startedAt = Date.now();
  const promptChars = promptText.length + (params.systemInstruction?.length || 0);
  const timer = setTimeout(() => {
    controller.abort();
  }, timeoutMs);

  let response: Response;
  try {
    response = await fetchFn(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey,
      },
      body: JSON.stringify(requestBody),
      signal: controller.signal,
    });
  } catch (err) {
    const isAbort =
      err instanceof Error && (err.name === 'AbortError' || /abort/i.test(err.message));
    const errMsg = isAbort
      ? `Gemini Primary 요청이 ${Math.round(timeoutMs / 1000)}초 타임아웃을 초과했습니다.`
      : `Gemini Primary 네트워크 오류: ${sanitizeGeminiErrorText(err instanceof Error ? err.message : String(err), apiKey)}`;
    recordLlmUsageEvent({
      provider: 'gemini',
      modelId,
      feature: 'gemini-primary',
      promptChars,
      latencyMs: Date.now() - startedAt,
      status: 'error',
      errorMessage: errMsg,
    });
    throw new Error(errMsg);
  } finally {
    clearTimeout(timer);
  }

  if (!response.ok) {
    let rawErr = '';
    try {
      rawErr = await response.text();
    } catch {
      rawErr = response.statusText || '';
    }
    const safeErr = sanitizeGeminiErrorText(rawErr, apiKey);
    recordLlmUsageEvent({
      provider: 'gemini',
      modelId,
      feature: 'gemini-primary',
      promptChars,
      latencyMs: Date.now() - startedAt,
      status: 'error',
      errorMessage: `HTTP ${response.status}: ${safeErr}`,
    });
    throw new Error(`Gemini Primary HTTP ${response.status} 오류: ${safeErr}`);
  }

  let data: {
    candidates?: Array<{
      content?: {
        parts?: Array<{ text?: string }>;
      };
    }>;
    usageMetadata?: GeminiUsageMetadata;
  };

  try {
    data = (await response.json()) as typeof data;
  } catch (err) {
    const safeMsg = sanitizeGeminiErrorText(err instanceof Error ? err.message : String(err), apiKey);
    throw new Error(`Gemini Primary JSON 파싱 오류: ${safeMsg}`);
  }

  const parts = data.candidates?.[0]?.content?.parts ?? [];
  const text = parts
    .map((p) => (typeof p?.text === 'string' ? p.text : ''))
    .join('')
    .trim();

  if (!text) {
    throw new Error('Gemini Primary가 빈 응답을 반환했습니다.');
  }

  const usageMetadata: GeminiUsageMetadata | undefined = data.usageMetadata
    ? {
        promptTokenCount: data.usageMetadata.promptTokenCount,
        candidatesTokenCount: data.usageMetadata.candidatesTokenCount,
        totalTokenCount: data.usageMetadata.totalTokenCount,
        thoughtsTokenCount: data.usageMetadata.thoughtsTokenCount,
      }
    : undefined;

  recordLlmUsageEvent({
    provider: 'gemini',
    modelId,
    feature: 'gemini-primary',
    inputTokens: usageMetadata?.promptTokenCount,
    outputTokens:
      typeof usageMetadata?.candidatesTokenCount === 'number'
        ? usageMetadata.candidatesTokenCount + (usageMetadata.thoughtsTokenCount || 0)
        : undefined,
    promptChars,
    responseChars: text.length,
    latencyMs: Date.now() - startedAt,
    status: 'ok',
  });

  return {
    text,
    model: `Gemini Primary · ${modelId}`,
    modelId,
    provider: 'gemini-primary',
    usageMetadata,
  };
}

/**
 * Priority router: attempts Gemini 4 Hydrogen (1st priority) first, and seamlessly falls back
 * to the configured secondary provider on missing/invalid key, HTTP 429/5xx, timeout, network, or empty response.
 */
export async function analyzeWithPriority(params: {
  primary: GeminiPrimaryParams;
  fallback: () => Promise<SecondaryProviderResponse>;
  skipPrimary?: boolean;
  onFallback?: (reason: string) => void;
}): Promise<PriorityAnalysisResult> {
  if (!params.skipPrimary) {
    try {
      const primaryResult = await geminiPrimary(params.primary);
      return {
        text: primaryResult.text,
        model: primaryResult.model,
        providerUsed: 'builtin',
        usedPrimary: true,
        usageMetadata: primaryResult.usageMetadata,
      };
    } catch (err) {
      const reason = sanitizeGeminiErrorText(
        err instanceof Error ? err.message : String(err),
        params.primary.apiKey ?? process.env.GEMINI_API_KEY
      );
      params.onFallback?.(reason);
      console.warn(`[analyzeWithPriority] Gemini Primary fallback triggered: ${reason}`);

      const secondary = await params.fallback();
      return {
        text: secondary.text,
        model: secondary.model,
        providerUsed: secondary.providerUsed ?? 'groq',
        usedPrimary: false,
        usageMetadata: secondary.usageMetadata,
        fallbackReason: reason,
      };
    }
  }

  const secondary = await params.fallback();
  return {
    text: secondary.text,
    model: secondary.model,
    providerUsed: secondary.providerUsed ?? 'groq',
    usedPrimary: false,
    usageMetadata: secondary.usageMetadata,
  };
}
