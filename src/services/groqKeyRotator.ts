import fs from 'node:fs';
import path from 'node:path';
import type { AiProviderId } from '../types/mail';
import { selectClaudeModelByComplexity, type ClaudeRoutingInput } from './claudeModelRouter';
import { getLlmUsageLimits, recordLlmUsageEvent } from './llmUsageTracker';

export const GROQ_CHAT_COMPLETIONS_URL = 'https://api.groq.com/openai/v1/chat/completions';
export const DEFAULT_GROQ_MODEL = 'openai/gpt-oss-120b';
export const MAX_GROQ_KEY_SLOTS = 5;

const GROQ_KEY_SLOT_NAMES = [
  'GROQ_API_KEY',
  'GROQ_API_KEY_1',
  'GROQ_API_KEY_2',
  'GROQ_API_KEY_3',
  'GROQ_API_KEY_4',
  'GROQ_API_KEY_5',
] as const;

const COOLDOWN_RATE_OR_SERVER_MS = 60_000;
const COOLDOWN_AUTH_MS = 300_000;
const COOLDOWN_NETWORK_MS = 30_000;
const DEFAULT_TIMEOUT_MS = 25_000;

export interface GroqKeyPoolStatus {
  configuredCount: number;
  healthyCount: number;
  maxSlots: number;
}

export interface GroqStreamingCompletionParams {
  systemPrompt?: string;
  userPrompt: string;
  model?: string;
  jsonMode?: boolean;
  maxCompletionTokens?: number;
  feature?: string;
  timeoutMs?: number;
  envOverrides?: Record<string, string | undefined>;
  fetchImpl?: typeof fetch;
}

export interface GroqStreamingCompletionResult {
  text: string;
  model: string;
  keySlotUsed: number;
  inputTokens?: number;
  outputTokens?: number;
}

export interface ClaudeEscalationCheckInput {
  explicitProvider?: AiProviderId;
  highPerformance?: boolean;
  routing?: ClaudeRoutingInput;
  promptChars?: number;
}

function isUsableGroqKey(raw: string | undefined): string {
  const clean = String(raw || '').trim();
  if (clean.length < 10) return '';
  if (/^MY_GROQ/i.test(clean)) return '';
  if (/^(placeholder|changeme|example|none|null|undefined)$/i.test(clean)) return '';
  return clean;
}

function splitMultiKeyValue(raw: string | undefined): string[] {
  if (!raw) return [];
  return String(raw)
    .split(/[,;\s]+/)
    .map((part) => isUsableGroqKey(part))
    .filter(Boolean);
}

function parseGroqKeyLinesFromText(text: string): string[] {
  const results: string[] = [];
  const clean = text.replace(/^\uFEFF/, '');
  for (const rawLine of clean.split(/\r?\n/)) {
    const trimmed = rawLine.trim();
    if (!trimmed || trimmed.startsWith('#') || trimmed.startsWith('//')) continue;
    const eqMatch = trimmed.match(/^[A-Za-z0-9_]+\s*=\s*(.+)$/);
    const val = (eqMatch ? eqMatch[1] : trimmed).trim().replace(/^['"]|['"]$/g, '');
    for (const k of splitMultiKeyValue(val)) {
      results.push(k);
    }
  }
  return results;
}

export function readGroqKeysFromFile(filePath: string): string[] {
  try {
    if (!filePath || !fs.existsSync(filePath)) return [];
    const content = fs.readFileSync(filePath, 'utf8');
    return parseGroqKeyLinesFromText(content);
  } catch {
    return [];
  }
}

/**
 * Collects up to 5 deduplicated valid Groq API keys from overrides, process.env, and `groq-keys.txt`.
 * Ignores empty strings and placeholders shorter than 10 chars.
 */
export function collectGroqApiKeys(
  envOverrides?: Record<string, string | undefined>,
  envSource: Record<string, string | undefined> = process.env
): string[] {
  const seen = new Set<string>();
  const collected: string[] = [];

  const pushCandidate = (candidate: string | undefined) => {
    const clean = isUsableGroqKey(candidate);
    if (!clean || seen.has(clean)) return;
    if (collected.length >= MAX_GROQ_KEY_SLOTS) return;
    seen.add(clean);
    collected.push(clean);
  };

  for (const source of [envOverrides, envSource]) {
    if (!source) continue;
    for (const slotName of GROQ_KEY_SLOT_NAMES) {
      pushCandidate(source[slotName]);
    }
    for (const multiKey of splitMultiKeyValue(source.GROQ_API_KEYS)) {
      pushCandidate(multiKey);
    }
  }

  // Also load simple key-per-line/comma file (`groq-keys.txt`) unless envOverrides or BMAIL_GROQ_KEYS_FILE overrides it
  const candidateFiles = envOverrides
    ? [envOverrides.BMAIL_GROQ_KEYS_FILE].filter((p): p is string => Boolean(p))
    : envSource.BMAIL_GROQ_KEYS_FILE
      ? [envSource.BMAIL_GROQ_KEYS_FILE]
      : Array.from(
          new Set(
            [
              envSource.BMAIL_DATA_DIR ? path.join(envSource.BMAIL_DATA_DIR, 'groq-keys.txt') : '',
              path.join(process.cwd(), 'groq-keys.txt'),
              envSource.BMAIL_APP_ROOT ? path.join(envSource.BMAIL_APP_ROOT, 'groq-keys.txt') : '',
            ].filter((p): p is string => Boolean(p))
          )
        );

  for (const filePath of candidateFiles) {
    for (const fileKey of readGroqKeysFromFile(filePath)) {
      pushCandidate(fileKey);
    }
  }

  return collected;
}

export function sanitizeGroqError(message: string, apiKey?: string): string {
  let clean = String(message || 'Unknown Groq API error');
  if (apiKey && apiKey.trim().length >= 6) {
    clean = clean.split(apiKey.trim()).join('[REDACTED_GROQ_KEY]');
  }
  clean = clean.replace(/gsk_[A-Za-z0-9_-]+/gi, '[REDACTED_GROQ_KEY]');
  return clean.slice(0, 240);
}

/**
 * Escalates to Claude for advanced analysis (executive reports, multi-doc analysis, meetings,
 * medium/high reduce volumes, explicit Claude/highPerformance, or large prompts), while keeping
 * fast single-mail & parallel map chunks on the Groq 5-key pool.
 */
export function shouldEscalateToClaude(input: ClaudeEscalationCheckInput = {}): boolean {
  if (input.explicitProvider === 'claude') return true;
  if (input.highPerformance === true) return true;
  const effectiveChars = Number(input.promptChars ?? input.routing?.promptChars ?? 0);
  if (effectiveChars > 18000) return true;

  if (!input.explicitProvider) {
    const mode = getLlmUsageLimits().claudeAdvancedMode;
    if (mode === 'claude-only') return true;
    if (mode === 'groq-only') return false;
  }

  const taskKind = input.routing?.taskKind;
  if (taskKind === 'mail-map' || taskKind === 'quick-summary') {
    // Keep parallel map chunks & quick summaries fast on Groq unless highPerformance/huge prompt
    return false;
  }
  if (
    taskKind === 'executive-report' ||
    taskKind === 'doc-analysis' ||
    taskKind === 'document' ||
    taskKind === 'meeting'
  ) {
    return true;
  }
  if (taskKind === 'mail-reduce' && Number(input.routing?.emailCount ?? 0) >= 6) {
    return true;
  }
  if (input.routing) {
    const decision = selectClaudeModelByComplexity(input.routing);
    if (decision.tier === 'high') return true;
  }
  return false;
}

/**
 * Executes a batch of AI tasks in parallel across up to all configured Groq API key slots
 * (default concurrency = min(items.length, configuredGroqKeys, limits.groqParallelMax, 5)).
 */
export async function runGroqParallelBatch<T, R>(
  items: readonly T[],
  worker: (item: T, index: number) => Promise<R>,
  options?: {
    maxConcurrency?: number;
    envOverrides?: Record<string, string | undefined>;
  }
): Promise<R[]> {
  if (items.length === 0) return [];
  const poolCount = Math.max(1, getGroqKeyPoolStatus(options?.envOverrides).configuredCount);
  const adminMax = getLlmUsageLimits().groqParallelMax;
  const concurrency = Math.max(
    1,
    Math.min(
      items.length,
      options?.maxConcurrency ?? Math.min(poolCount, adminMax),
      MAX_GROQ_KEY_SLOTS
    )
  );
  const results = new Array<R>(items.length);
  let cursor = 0;

  const runners = Array.from({ length: concurrency }, async () => {
    while (true) {
      const idx = cursor;
      cursor += 1;
      if (idx >= items.length) return;
      results[idx] = await worker(items[idx], idx);
    }
  });

  await Promise.all(runners);
  return results;
}

interface ParsedSseResult {
  text: string;
  inputTokens?: number;
  outputTokens?: number;
}

async function parseGroqResponseStream(response: Response): Promise<ParsedSseResult> {
  const contentType = (response.headers.get('content-type') || '').toLowerCase();
  if (contentType.includes('application/json')) {
    const data = (await response.json()) as {
      choices?: Array<{ message?: { content?: string }; delta?: { content?: string } }>;
      usage?: { prompt_tokens?: number; completion_tokens?: number };
    };
    const text =
      data.choices?.[0]?.message?.content ?? data.choices?.[0]?.delta?.content ?? '';
    return {
      text: String(text || ''),
      inputTokens: data.usage?.prompt_tokens,
      outputTokens: data.usage?.completion_tokens,
    };
  }

  if (!response.body) {
    const fallbackText = await response.text();
    if (!fallbackText.trim()) {
      return { text: '' };
    }
    // Attempt SSE line parsing on full text if stream body was unavailable
    return parseSseLines(fallbackText);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let buffer = '';
  let contentOut = '';
  let inputTokens: number | undefined;
  let outputTokens: number | undefined;
  let sawDone = false;

  const consumeLine = (rawLine: string) => {
    const line = rawLine.trim();
    if (!line || line.startsWith(':')) return;
    if (!line.startsWith('data:')) return;
    const payload = line.slice(5).trim();
    if (!payload) return;
    if (payload === '[DONE]') {
      sawDone = true;
      return;
    }
    try {
      const chunk = JSON.parse(payload) as {
        choices?: Array<{ delta?: { content?: string }; message?: { content?: string } }>;
        usage?: { prompt_tokens?: number; completion_tokens?: number };
        x_groq?: { usage?: { prompt_tokens?: number; completion_tokens?: number } };
      };
      const piece =
        chunk.choices?.[0]?.delta?.content ?? chunk.choices?.[0]?.message?.content ?? '';
      if (piece) {
        contentOut += piece;
      }
      const usage = chunk.x_groq?.usage ?? chunk.usage;
      if (usage) {
        if (typeof usage.prompt_tokens === 'number') inputTokens = usage.prompt_tokens;
        if (typeof usage.completion_tokens === 'number') outputTokens = usage.completion_tokens;
      }
    } catch {
      // Ignore malformed partial SSE keepalive/event JSON chunks
    }
  };

  while (!sawDone) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    let newlineIdx = buffer.indexOf('\n');
    while (newlineIdx !== -1) {
      const line = buffer.slice(0, newlineIdx);
      buffer = buffer.slice(newlineIdx + 1);
      consumeLine(line);
      if (sawDone) break;
      newlineIdx = buffer.indexOf('\n');
    }
  }

  buffer += decoder.decode();
  if (!sawDone && buffer.trim()) {
    for (const remaining of buffer.split(/\r?\n/)) {
      consumeLine(remaining);
    }
  }

  if (sawDone) {
    try {
      await reader.cancel();
    } catch {
      // Ignore stream cancel errors after [DONE]
    }
  }

  return { text: contentOut, inputTokens, outputTokens };
}

function parseSseLines(rawText: string): ParsedSseResult {
  let contentOut = '';
  let inputTokens: number | undefined;
  let outputTokens: number | undefined;
  for (const rawLine of rawText.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line.startsWith('data:')) continue;
    const payload = line.slice(5).trim();
    if (!payload || payload === '[DONE]') break;
    try {
      const chunk = JSON.parse(payload) as {
        choices?: Array<{ delta?: { content?: string }; message?: { content?: string } }>;
        usage?: { prompt_tokens?: number; completion_tokens?: number };
        x_groq?: { usage?: { prompt_tokens?: number; completion_tokens?: number } };
      };
      const piece =
        chunk.choices?.[0]?.delta?.content ?? chunk.choices?.[0]?.message?.content ?? '';
      if (piece) contentOut += piece;
      const usage = chunk.x_groq?.usage ?? chunk.usage;
      if (usage) {
        if (typeof usage.prompt_tokens === 'number') inputTokens = usage.prompt_tokens;
        if (typeof usage.completion_tokens === 'number') outputTokens = usage.completion_tokens;
      }
    } catch {
      // Ignore unparseable line
    }
  }
  return { text: contentOut, inputTokens, outputTokens };
}

export class GroqKeyRotator {
  private nextSlotIndex = 0;
  private readonly cooldownUntil = new Map<string, number>();
  private readonly nowFn: () => number;

  constructor(options?: { nowFn?: () => number }) {
    this.nowFn = options?.nowFn ?? (() => Date.now());
  }

  reset(): void {
    this.nextSlotIndex = 0;
    this.cooldownUntil.clear();
  }

  getPoolStatus(envOverrides?: Record<string, string | undefined>): GroqKeyPoolStatus {
    const keys = collectGroqApiKeys(envOverrides);
    const now = this.nowFn();
    const healthyCount = keys.filter((k) => (this.cooldownUntil.get(k) ?? 0) <= now).length;
    return {
      configuredCount: keys.length,
      healthyCount,
      maxSlots: MAX_GROQ_KEY_SLOTS,
    };
  }

  private setCooldown(key: string, durationMs: number): void {
    this.cooldownUntil.set(key, this.nowFn() + durationMs);
  }

  async invokeStreamingCompletion(
    params: GroqStreamingCompletionParams
  ): Promise<GroqStreamingCompletionResult> {
    const keys = collectGroqApiKeys(params.envOverrides);
    if (keys.length === 0) {
      throw new Error(
        '유효한 Groq API 키(GROQ_API_KEY ~ GROQ_API_KEY_5)가 설정되지 않았습니다.'
      );
    }

    const now = this.nowFn();
    const startIdx = this.nextSlotIndex % keys.length;
    this.nextSlotIndex = (startIdx + 1) % keys.length;

    const orderedCandidates: Array<{ key: string; slotNumber: number }> = [];
    for (let offset = 0; offset < keys.length; offset++) {
      const idx = (startIdx + offset) % keys.length;
      const key = keys[idx];
      if ((this.cooldownUntil.get(key) ?? 0) <= now) {
        orderedCandidates.push({ key, slotNumber: idx + 1 });
      }
    }

    if (orderedCandidates.length === 0) {
      throw new Error(
        `모든 Groq API 키(${keys.length}개)가 일시적 쿨다운 상태입니다. 잠시 후 다시 시도하거나 Claude로 자동 전환합니다.`
      );
    }

    const model = params.model?.trim() || process.env.GROQ_MODEL?.trim() || DEFAULT_GROQ_MODEL;
    const systemPrompt = String(params.systemPrompt || '');
    const userPrompt = String(params.userPrompt || '');
    const messages: Array<{ role: 'system' | 'user'; content: string }> = [];
    if (systemPrompt.trim()) {
      messages.push({ role: 'system', content: systemPrompt });
    }
    messages.push({ role: 'user', content: userPrompt });

    const fetchImpl = params.fetchImpl ?? fetch;
    const timeoutMs = Math.max(2000, params.timeoutMs ?? DEFAULT_TIMEOUT_MS);
    const promptChars = systemPrompt.length + userPrompt.length;
    const feature = params.feature || 'groq-stream';
    const errors: string[] = [];

    for (const { key, slotNumber } of orderedCandidates) {
      const startedAt = this.nowFn();
      const attemptRequest = async (includeResponseFormat: boolean) => {
        const bodyPayload: Record<string, unknown> = {
          model,
          messages,
          temperature: 1,
          max_completion_tokens: params.maxCompletionTokens ?? 2048,
          top_p: 1,
          reasoning_effort: 'medium',
          stream: true,
          stop: null,
        };
        if (includeResponseFormat) {
          bodyPayload.response_format = { type: 'json_object' };
        }

        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), timeoutMs);
        try {
          return await fetchImpl(GROQ_CHAT_COMPLETIONS_URL, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              Authorization: `Bearer ${key}`,
            },
            body: JSON.stringify(bodyPayload),
            signal: controller.signal,
          });
        } finally {
          clearTimeout(timer);
        }
      };

      try {
        let response = await attemptRequest(Boolean(params.jsonMode));
        if (!response.ok && response.status === 400 && params.jsonMode) {
          const errBody = await response.text();
          if (/response_format|json_object/i.test(errBody)) {
            response = await attemptRequest(false);
          } else {
            const cleanErr = sanitizeGroqError(errBody, key);
            recordLlmUsageEvent({
              provider: 'groq',
              modelId: model,
              feature,
              promptChars,
              latencyMs: Math.max(0, this.nowFn() - startedAt),
              status: 'error',
              errorMessage: `Slot#${slotNumber} HTTP 400: ${cleanErr.slice(0, 100)}`,
            });
            errors.push(`Slot#${slotNumber}(400): ${cleanErr}`);
            continue;
          }
        }

        if (!response.ok) {
          const errText = await response.text();
          const cleanErr = sanitizeGroqError(errText, key);
          if (response.status === 401 || response.status === 403) {
            this.setCooldown(key, COOLDOWN_AUTH_MS);
          } else if (response.status === 429 || response.status >= 500) {
            this.setCooldown(key, COOLDOWN_RATE_OR_SERVER_MS);
          }
          recordLlmUsageEvent({
            provider: 'groq',
            modelId: model,
            feature,
            promptChars,
            latencyMs: Math.max(0, this.nowFn() - startedAt),
            status: 'error',
            errorMessage: `Slot#${slotNumber} HTTP ${response.status}: ${cleanErr.slice(0, 100)}`,
          });
          errors.push(`Slot#${slotNumber}(${response.status}): ${cleanErr}`);
          continue;
        }

        const parsed = await parseGroqResponseStream(response);
        const text = parsed.text.trim();
        if (!text) {
          this.setCooldown(key, COOLDOWN_RATE_OR_SERVER_MS);
          recordLlmUsageEvent({
            provider: 'groq',
            modelId: model,
            feature,
            promptChars,
            latencyMs: Math.max(0, this.nowFn() - startedAt),
            status: 'error',
            errorMessage: `Slot#${slotNumber}: Empty stream response`,
          });
          errors.push(`Slot#${slotNumber}: 빈 스트림 응답`);
          continue;
        }

        recordLlmUsageEvent({
          provider: 'groq',
          modelId: model,
          feature,
          inputTokens: parsed.inputTokens,
          outputTokens: parsed.outputTokens,
          promptChars,
          responseChars: text.length,
          latencyMs: Math.max(0, this.nowFn() - startedAt),
          status: 'ok',
        });

        return {
          text,
          model: `Groq · ${model} (#${slotNumber})`,
          keySlotUsed: slotNumber,
          inputTokens: parsed.inputTokens,
          outputTokens: parsed.outputTokens,
        };
      } catch (err) {
        const cleanMsg = sanitizeGroqError(
          err instanceof Error ? err.message : String(err),
          key
        );
        this.setCooldown(key, COOLDOWN_NETWORK_MS);
        recordLlmUsageEvent({
          provider: 'groq',
          modelId: model,
          feature,
          promptChars,
          latencyMs: Math.max(0, this.nowFn() - startedAt),
          status: 'error',
          errorMessage: `Slot#${slotNumber}: ${cleanMsg.slice(0, 100)}`,
        });
        errors.push(`Slot#${slotNumber}(net): ${cleanMsg}`);
      }
    }

    throw new Error(`Groq 5-키 풀 호출 실패: ${errors.join(' | ').slice(0, 360)}`);
  }
}

const defaultGroqKeyRotator = new GroqKeyRotator();

export function getGroqKeyPoolStatus(
  envOverrides?: Record<string, string | undefined>
): GroqKeyPoolStatus {
  return defaultGroqKeyRotator.getPoolStatus(envOverrides);
}

export function resetGroqKeyRotatorState(): void {
  defaultGroqKeyRotator.reset();
}

export async function invokeGroqStreamingCompletion(
  params: GroqStreamingCompletionParams
): Promise<GroqStreamingCompletionResult> {
  return defaultGroqKeyRotator.invokeStreamingCompletion(params);
}
