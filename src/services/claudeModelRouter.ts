import { recordLlmUsageEvent } from './llmUsageTracker';

export const CLAUDE_MODEL_HAIKU = 'claude-haiku-5-5';
export const CLAUDE_MODEL_SONNET = 'claude-sonnet-5-5';
export const CLAUDE_MODEL_OPUS = 'claude-opus-5-5';
export const DEFAULT_CLAUDE_MODEL = CLAUDE_MODEL_HAIKU;

const AUTO_MODE_OVERRIDES = new Set([
  '',
  'auto',
  CLAUDE_MODEL_HAIKU,
  'claude-sonnet-4-20250514',
]);

const HIGH_DIFFICULTY_RE =
  /(심층|정밀|감리|감사|아키텍처|보안|계약|법적|분쟁|리스크\s*평가|트레이드오프|근본\s*원인|전략|최종\s*결론|오퍼스|opus|deep|audit|legal|contract|security)/i;

const MEDIUM_DIFFICULTY_RE =
  /(비교|교차|분석|보고서|블로그|제안서|도면|리비전|소넷|sonnet|compare|report|proposal|drawing)/i;

export type ClaudeTaskKind =
  | 'mail-map'
  | 'mail-reduce'
  | 'document'
  | 'doc-analysis'
  | 'executive-report'
  | 'meeting'
  | 'sgw-assist'
  | 'quick-summary'
  | 'general';

export type ClaudeDifficultyTier = 'low' | 'medium' | 'high';

export interface ClaudeRoutingInput {
  taskKind?: ClaudeTaskKind;
  emailCount?: number;
  documentCount?: number;
  promptChars?: number;
  customFocus?: string;
  explicitModelOverride?: string;
  autoTierEnabled?: boolean;
}

export interface ClaudeRoutingDecision {
  modelId: string;
  tier: ClaudeDifficultyTier;
  tierLabel: string;
  effort: 'medium' | 'high';
  useServerFallback: boolean;
  reason: string;
}

export interface ClaudeUsageSummary {
  input_tokens?: number;
  output_tokens?: number;
  cache_read_input_tokens?: number;
}

export interface InvokeClaudeResult {
  text: string;
  model: string;
  decision: ClaudeRoutingDecision;
  usage?: ClaudeUsageSummary;
  stopReason?: string;
}

/**
 * Redacts Anthropic keys (`sk-ant-...` or explicit API key string) from error messages.
 */
export function sanitizeAnthropicError(message: string, apiKey?: string): string {
  let clean = String(message || 'Unknown Anthropic API error');
  if (apiKey && apiKey.trim().length >= 6) {
    clean = clean.split(apiKey.trim()).join('[REDACTED_KEY]');
  }
  clean = clean.replace(/sk-ant-[A-Za-z0-9_-]+/gi, '[REDACTED_KEY]');
  return clean.slice(0, 240);
}

/**
 * Selects Claude 5.5 model tier by task complexity:
 * - Default/low complexity -> `claude-haiku-5-5`
 * - Medium complexity -> `claude-sonnet-5-5`
 * - High complexity -> `claude-opus-5-5`
 */
export function selectClaudeModelByComplexity(input: ClaudeRoutingInput = {}): ClaudeRoutingDecision {
  const override = String(input.explicitModelOverride || '').trim();
  const taskKind: ClaudeTaskKind = input.taskKind ?? 'general';
  const emailCount = Math.max(0, Number(input.emailCount) || 0);
  const documentCount = Math.max(0, Number(input.documentCount) || 0);
  const promptChars = Math.max(0, Number(input.promptChars) || 0);
  const focus = String(input.customFocus || '').trim();

  if (override && !AUTO_MODE_OVERRIDES.has(override)) {
    const useServerFallback = /opus-5|sonnet-5-5|fable-5/i.test(override);
    const isHighModel = /opus-5|fable-5/i.test(override);
    const isMediumModel = /sonnet-5-5/i.test(override);
    return {
      modelId: override,
      tier: isHighModel ? 'high' : isMediumModel ? 'medium' : 'low',
      tierLabel: `수동지정·${override}`,
      effort: isHighModel ? 'high' : 'medium',
      useServerFallback,
      reason: 'explicit-model-override',
    };
  }

  const autoEnvDisabled = String(process.env.ANTHROPIC_AUTO_TIER || '').trim().toLowerCase() === 'false';
  if (input.autoTierEnabled === false || autoEnvDisabled) {
    return {
      modelId: CLAUDE_MODEL_HAIKU,
      tier: 'low',
      tierLabel: '고정·Haiku 5.5',
      effort: 'medium',
      useServerFallback: false,
      reason: 'auto-tier-disabled',
    };
  }

  const isHighComplexity =
    taskKind === 'executive-report' ||
    promptChars >= 18000 ||
    emailCount >= 20 ||
    documentCount >= 4 ||
    ((taskKind === 'document' || taskKind === 'doc-analysis') &&
      documentCount >= 2 &&
      emailCount >= 5) ||
    HIGH_DIFFICULTY_RE.test(focus);

  if (isHighComplexity) {
    return {
      modelId: CLAUDE_MODEL_OPUS,
      tier: 'high',
      tierLabel: '고난도·Opus 5.5',
      effort: 'high',
      useServerFallback: true,
      reason: HIGH_DIFFICULTY_RE.test(focus)
        ? 'high-difficulty-keyword'
        : 'high-volume-or-cross-analysis',
    };
  }

  const isMediumComplexity =
    taskKind === 'meeting' ||
    taskKind === 'document' ||
    taskKind === 'doc-analysis' ||
    emailCount >= 6 ||
    documentCount >= 1 ||
    promptChars >= 5500 ||
    MEDIUM_DIFFICULTY_RE.test(focus);

  if (isMediumComplexity) {
    return {
      modelId: CLAUDE_MODEL_SONNET,
      tier: 'medium',
      tierLabel: '중난도·Sonnet 5.5',
      effort: 'medium',
      useServerFallback: true,
      reason:
        taskKind === 'meeting' || taskKind === 'document' || taskKind === 'doc-analysis'
          ? `task-${taskKind}`
          : MEDIUM_DIFFICULTY_RE.test(focus)
            ? 'medium-difficulty-keyword'
            : 'medium-volume-or-length',
    };
  }

  return {
    modelId: CLAUDE_MODEL_HAIKU,
    tier: 'low',
    tierLabel: '기본·Haiku 5.5',
    effort: 'medium',
    useServerFallback: false,
    reason: 'default-routine-haiku',
  };
}

/**
 * Invokes Anthropic Messages API (`POST /v1/messages`) using Claude 5.5 difficulty routing.
 * Never sends deprecated `temperature`, `top_p`, `top_k`, or `budget_tokens`.
 */
export async function invokeClaudeMessagesApi(params: {
  apiKey?: string;
  systemPrompt: string;
  userPrompt: string;
  maxTokens?: number;
  routing?: ClaudeRoutingInput;
  fetchImpl?: typeof fetch;
  timeoutMs?: number;
}): Promise<InvokeClaudeResult> {
  const apiKey = (params.apiKey ?? process.env.ANTHROPIC_API_KEY ?? '').trim();
  if (!apiKey) {
    throw new Error('ANTHROPIC_API_KEY 환경변수가 .env(Secrets) 또는 EMV 볼트에 설정되지 않았습니다.');
  }

  const decision = selectClaudeModelByComplexity(params.routing);
  const tierMaxTokens =
    decision.tier === 'high' ? 16384 : decision.tier === 'medium' ? 8192 : 6144;
  const timeoutMs =
    typeof params.timeoutMs === 'number' && params.timeoutMs > 0
      ? params.timeoutMs
      : decision.tier === 'high'
        ? 90_000
        : decision.tier === 'medium'
          ? 45_000
          : 30_000;

  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    'x-api-key': apiKey,
    'anthropic-version': '2023-06-01',
  };
  if (decision.useServerFallback) {
    headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
  }

  const body: Record<string, unknown> = {
    model: decision.modelId,
    max_tokens: params.maxTokens ?? tierMaxTokens,
    output_config: { effort: decision.effort },
    system: params.systemPrompt,
    messages: [{ role: 'user', content: params.userPrompt }],
  };
  if (decision.useServerFallback) {
    body.fallbacks = 'default';
  }

  const fetchFn = params.fetchImpl ?? fetch;
  const startedAt = Date.now();
  const promptChars = (params.systemPrompt?.length || 0) + (params.userPrompt?.length || 0);
  let response: Response;
  try {
    response = await fetchFn('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(timeoutMs),
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    const cleanErr = sanitizeAnthropicError(msg, apiKey);
    recordLlmUsageEvent({
      provider: 'claude',
      modelId: decision.modelId,
      tier: decision.tier,
      feature: params.routing?.taskKind || 'claude-messages',
      promptChars,
      latencyMs: Date.now() - startedAt,
      status: 'error',
      errorMessage: cleanErr,
    });
    throw new Error(`Claude API 네트워크 오류: ${cleanErr}`);
  }

  if (!response.ok) {
    const errText = await response.text().catch(() => '');
    const cleanErr = sanitizeAnthropicError(errText, apiKey);
    recordLlmUsageEvent({
      provider: 'claude',
      modelId: decision.modelId,
      tier: decision.tier,
      feature: params.routing?.taskKind || 'claude-messages',
      promptChars,
      latencyMs: Date.now() - startedAt,
      status: 'error',
      errorMessage: `HTTP ${response.status}: ${cleanErr}`,
    });
    throw new Error(
      `Claude API 호출 실패 (${response.status}): ${cleanErr}`
    );
  }

  const data = (await response.json()) as {
    stop_reason?: string;
    stop_details?: { category?: string } | string;
    content?: Array<{ type: string; text?: string }>;
    usage?: ClaudeUsageSummary;
  };

  if (data.stop_reason === 'refusal') {
    const cat =
      typeof data.stop_details === 'object' && data.stop_details?.category
        ? data.stop_details.category
        : typeof data.stop_details === 'string'
          ? data.stop_details
          : '';
    recordLlmUsageEvent({
      provider: 'claude',
      modelId: decision.modelId,
      tier: decision.tier,
      feature: params.routing?.taskKind || 'claude-messages',
      inputTokens: data.usage?.input_tokens,
      outputTokens: data.usage?.output_tokens,
      promptChars,
      latencyMs: Date.now() - startedAt,
      status: 'error',
      errorMessage: `refusal${cat ? `: ${cat}` : ''}`,
    });
    throw new Error(
      `Claude 안전 정책에 의해 요청이 거절되었습니다 (stop_reason: refusal${cat ? `, ${cat}` : ''}).`
    );
  }

  const text = Array.isArray(data.content)
    ? data.content.find((block) => block.type === 'text')?.text
    : undefined;
  if (!text || !text.trim()) {
    throw new Error('Claude API가 빈 응답을 반환했습니다.');
  }

  recordLlmUsageEvent({
    provider: 'claude',
    modelId: decision.modelId,
    tier: decision.tier,
    feature: params.routing?.taskKind || 'claude-messages',
    inputTokens: data.usage?.input_tokens,
    outputTokens: data.usage?.output_tokens,
    promptChars,
    responseChars: text.length,
    latencyMs: Date.now() - startedAt,
    status: 'ok',
  });

  return {
    text,
    model: `Claude · ${decision.modelId} (${decision.tierLabel})`,
    decision,
    usage: data.usage,
    stopReason: data.stop_reason,
  };
}
