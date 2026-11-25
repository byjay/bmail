import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

export type LlmProviderFamily = 'claude' | 'gemini' | 'groq' | 'openai' | 'agent-builder';

export interface LlmUsageEventInput {
  provider: LlmProviderFamily;
  modelId: string;
  tier?: 'low' | 'medium' | 'high' | 'standard';
  feature?: string;
  inputTokens?: number;
  outputTokens?: number;
  promptChars?: number;
  responseChars?: number;
  latencyMs?: number;
  status?: 'ok' | 'error';
  errorMessage?: string;
}

export interface LlmUsageEventRecord {
  id: string;
  timestamp: string;
  date: string;
  provider: LlmProviderFamily;
  modelId: string;
  tier: 'low' | 'medium' | 'high' | 'standard';
  feature: string;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  estimatedTokens: boolean;
  costUsd: number;
  latencyMs: number;
  status: 'ok' | 'error';
  errorMessage?: string;
}

export interface LlmModelUsageRow {
  provider: LlmProviderFamily;
  modelId: string;
  tier: string;
  calls: number;
  errorCalls: number;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  costUsd: number;
  avgLatencyMs: number;
  lastUsedAt: string;
}

export interface LlmDailyUsageBucket {
  date: string;
  calls: number;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  costUsd: number;
}

export interface LlmUsageSummary {
  updatedAt: string;
  totals: {
    totalCalls: number;
    successCalls: number;
    errorCalls: number;
    inputTokens: number;
    outputTokens: number;
    totalTokens: number;
    totalCostUsd: number;
    avgLatencyMs: number;
  };
  today: {
    date: string;
    calls: number;
    totalTokens: number;
    costUsd: number;
  };
  byProvider: Record<
    LlmProviderFamily,
    {
      calls: number;
      totalTokens: number;
      costUsd: number;
    }
  >;
  byModel: LlmModelUsageRow[];
  dailyBuckets: LlmDailyUsageBucket[];
  recentEvents: LlmUsageEventRecord[];
}

interface ModelPricingPerMillion {
  inputUsdPer1M: number;
  outputUsdPer1M: number;
}

/**
 * Standard per-1M token pricing table for transparent admin cost estimation.
 */
const MODEL_PRICING_PER_1M: Array<{ match: RegExp; pricing: ModelPricingPerMillion }> = [
  { match: /haiku-5-5|haiku-4-5/i, pricing: { inputUsdPer1M: 0.8, outputUsdPer1M: 4.0 } },
  { match: /sonnet-5-5|sonnet-5|sonnet-4/i, pricing: { inputUsdPer1M: 3.0, outputUsdPer1M: 15.0 } },
  { match: /opus-5-5|opus-5|opus-4|fable-5/i, pricing: { inputUsdPer1M: 15.0, outputUsdPer1M: 75.0 } },
  { match: /gemini/i, pricing: { inputUsdPer1M: 0.35, outputUsdPer1M: 1.05 } },
  { match: /gpt-4o-mini/i, pricing: { inputUsdPer1M: 0.15, outputUsdPer1M: 0.6 } },
  { match: /gpt-4o|gpt-5/i, pricing: { inputUsdPer1M: 2.5, outputUsdPer1M: 10.0 } },
  { match: /groq|gpt-oss|llama|mixtral|whisper/i, pricing: { inputUsdPer1M: 0.0, outputUsdPer1M: 0.0 } },
];

export function estimateModelCostUsd(
  provider: LlmProviderFamily,
  modelId: string,
  inputTokens: number,
  outputTokens: number
): number {
  if (provider === 'groq' || provider === 'agent-builder') return 0;
  const inTok = Math.max(0, inputTokens);
  const outTok = Math.max(0, outputTokens);
  const entry = MODEL_PRICING_PER_1M.find((row) => row.match.test(modelId));
  const pricing = entry
    ? entry.pricing
    : provider === 'claude'
      ? { inputUsdPer1M: 0.8, outputUsdPer1M: 4.0 }
      : provider === 'gemini'
        ? { inputUsdPer1M: 0.35, outputUsdPer1M: 1.05 }
        : { inputUsdPer1M: 0.15, outputUsdPer1M: 0.6 };

  const cost = (inTok / 1_000_000) * pricing.inputUsdPer1M + (outTok / 1_000_000) * pricing.outputUsdPer1M;
  return Math.round(cost * 1_000_000) / 1_000_000;
}

const MAX_STORED_EVENTS = 500;

interface StoredLedgerPayload {
  version: 1;
  events: LlmUsageEventRecord[];
}

let globalUsageFilePath = path.resolve(
  process.env.BMAIL_DATA_DIR || path.resolve(process.cwd(), 'data'),
  'llm-usage-ledger.json'
);

let cachedEvents: LlmUsageEventRecord[] | null = null;

function ensureLoaded(): LlmUsageEventRecord[] {
  if (cachedEvents !== null) return cachedEvents;
  try {
    if (fs.existsSync(globalUsageFilePath)) {
      const raw = fs.readFileSync(globalUsageFilePath, 'utf8');
      const parsed = JSON.parse(raw) as StoredLedgerPayload;
      if (parsed && Array.isArray(parsed.events)) {
        cachedEvents = parsed.events.slice(-MAX_STORED_EVENTS);
        return cachedEvents;
      }
    }
  } catch {
    // start clean if corrupted
  }
  cachedEvents = [];
  return cachedEvents;
}

function persistEvents(events: LlmUsageEventRecord[]): void {
  try {
    const dir = path.dirname(globalUsageFilePath);
    fs.mkdirSync(dir, { recursive: true });
    const payload: StoredLedgerPayload = {
      version: 1,
      events: events.slice(-MAX_STORED_EVENTS),
    };
    fs.writeFileSync(globalUsageFilePath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
  } catch {
    // Non-fatal in read-only or ephemeral test environments
  }
}

export function configureLlmUsageLedgerPath(customFilePath: string): void {
  globalUsageFilePath = path.resolve(customFilePath);
  cachedEvents = null;
}

export function recordLlmUsageEvent(input: LlmUsageEventInput): LlmUsageEventRecord {
  const events = ensureLoaded();
  const nowIso = new Date().toISOString();
  const dateStr = nowIso.slice(0, 10);

  const hasExactIn = typeof input.inputTokens === 'number' && Number.isFinite(input.inputTokens) && input.inputTokens >= 0;
  const hasExactOut = typeof input.outputTokens === 'number' && Number.isFinite(input.outputTokens) && input.outputTokens >= 0;

  const inputTokens = hasExactIn
    ? Math.round(input.inputTokens!)
    : Math.max(1, Math.round((input.promptChars || 0) / 3.5));
  const outputTokens = hasExactOut
    ? Math.round(input.outputTokens!)
    : Math.max(0, Math.round((input.responseChars || 0) / 3.5));
  const totalTokens = inputTokens + outputTokens;
  const estimatedTokens = !(hasExactIn && hasExactOut);

  const cleanModelId = String(input.modelId || 'unknown').trim() || 'unknown';
  const tier =
    input.tier ||
    (/opus|fable/i.test(cleanModelId)
      ? 'high'
      : /sonnet/i.test(cleanModelId)
        ? 'medium'
        : /haiku/i.test(cleanModelId)
          ? 'low'
          : 'standard');

  const costUsd = estimateModelCostUsd(input.provider, cleanModelId, inputTokens, outputTokens);
  const record: LlmUsageEventRecord = {
    id: crypto.randomUUID(),
    timestamp: nowIso,
    date: dateStr,
    provider: input.provider,
    modelId: cleanModelId,
    tier,
    feature: String(input.feature || 'AI 분석').trim().slice(0, 60),
    inputTokens,
    outputTokens,
    totalTokens,
    estimatedTokens,
    costUsd,
    latencyMs: Math.max(0, Math.round(input.latencyMs || 0)),
    status: input.status === 'error' ? 'error' : 'ok',
    ...(input.errorMessage ? { errorMessage: String(input.errorMessage).slice(0, 140) } : {}),
  };

  events.push(record);
  if (events.length > MAX_STORED_EVENTS) {
    events.splice(0, events.length - MAX_STORED_EVENTS);
  }
  persistEvents(events);
  return record;
}

export function resetLlmUsageLedger(): void {
  cachedEvents = [];
  persistEvents(cachedEvents);
}

export function getLlmUsageSummary(recentLimit = 12): LlmUsageSummary {
  const events = ensureLoaded();
  const todayDate = new Date().toISOString().slice(0, 10);

  let totalCalls = 0;
  let successCalls = 0;
  let errorCalls = 0;
  let inputTokens = 0;
  let outputTokens = 0;
  let totalTokens = 0;
  let totalCostUsd = 0;
  let latencySum = 0;

  let todayCalls = 0;
  let todayTokens = 0;
  let todayCost = 0;

  const byProvider: LlmUsageSummary['byProvider'] = {
    claude: { calls: 0, totalTokens: 0, costUsd: 0 },
    gemini: { calls: 0, totalTokens: 0, costUsd: 0 },
    groq: { calls: 0, totalTokens: 0, costUsd: 0 },
    openai: { calls: 0, totalTokens: 0, costUsd: 0 },
    'agent-builder': { calls: 0, totalTokens: 0, costUsd: 0 },
  };

  const modelMap = new Map<
    string,
    {
      provider: LlmProviderFamily;
      modelId: string;
      tier: string;
      calls: number;
      errorCalls: number;
      inputTokens: number;
      outputTokens: number;
      totalTokens: number;
      costUsd: number;
      latencySum: number;
      lastUsedAt: string;
    }
  >();

  const dailyMap = new Map<string, LlmDailyUsageBucket>();

  for (const ev of events) {
    totalCalls += 1;
    if (ev.status === 'error') {
      errorCalls += 1;
    } else {
      successCalls += 1;
    }
    inputTokens += ev.inputTokens;
    outputTokens += ev.outputTokens;
    totalTokens += ev.totalTokens;
    totalCostUsd += ev.costUsd;
    latencySum += ev.latencyMs;

    if (ev.date === todayDate) {
      todayCalls += 1;
      todayTokens += ev.totalTokens;
      todayCost += ev.costUsd;
    }

    const provBucket = byProvider[ev.provider] || byProvider.claude;
    provBucket.calls += 1;
    provBucket.totalTokens += ev.totalTokens;
    provBucket.costUsd = Math.round((provBucket.costUsd + ev.costUsd) * 1_000_000) / 1_000_000;

    const mKey = `${ev.provider}::${ev.modelId}`;
    const existingModel = modelMap.get(mKey);
    if (!existingModel) {
      modelMap.set(mKey, {
        provider: ev.provider,
        modelId: ev.modelId,
        tier: ev.tier,
        calls: 1,
        errorCalls: ev.status === 'error' ? 1 : 0,
        inputTokens: ev.inputTokens,
        outputTokens: ev.outputTokens,
        totalTokens: ev.totalTokens,
        costUsd: ev.costUsd,
        latencySum: ev.latencyMs,
        lastUsedAt: ev.timestamp,
      });
    } else {
      existingModel.calls += 1;
      if (ev.status === 'error') existingModel.errorCalls += 1;
      existingModel.inputTokens += ev.inputTokens;
      existingModel.outputTokens += ev.outputTokens;
      existingModel.totalTokens += ev.totalTokens;
      existingModel.costUsd += ev.costUsd;
      existingModel.latencySum += ev.latencyMs;
      if (ev.timestamp > existingModel.lastUsedAt) {
        existingModel.lastUsedAt = ev.timestamp;
      }
    }

    const dayBucket = dailyMap.get(ev.date);
    if (!dayBucket) {
      dailyMap.set(ev.date, {
        date: ev.date,
        calls: 1,
        inputTokens: ev.inputTokens,
        outputTokens: ev.outputTokens,
        totalTokens: ev.totalTokens,
        costUsd: ev.costUsd,
      });
    } else {
      dayBucket.calls += 1;
      dayBucket.inputTokens += ev.inputTokens;
      dayBucket.outputTokens += ev.outputTokens;
      dayBucket.totalTokens += ev.totalTokens;
      dayBucket.costUsd = Math.round((dayBucket.costUsd + ev.costUsd) * 1_000_000) / 1_000_000;
    }
  }

  const byModel: LlmModelUsageRow[] = Array.from(modelMap.values())
    .map((m) => ({
      provider: m.provider,
      modelId: m.modelId,
      tier: m.tier,
      calls: m.calls,
      errorCalls: m.errorCalls,
      inputTokens: m.inputTokens,
      outputTokens: m.outputTokens,
      totalTokens: m.totalTokens,
      costUsd: Math.round(m.costUsd * 1_000_000) / 1_000_000,
      avgLatencyMs: m.calls > 0 ? Math.round(m.latencySum / m.calls) : 0,
      lastUsedAt: m.lastUsedAt,
    }))
    .sort((a, b) => b.totalTokens - a.totalTokens || b.calls - a.calls);

  const dailyBuckets = Array.from(dailyMap.values())
    .sort((a, b) => b.date.localeCompare(a.date))
    .slice(0, 7);

  return {
    updatedAt: new Date().toISOString(),
    totals: {
      totalCalls,
      successCalls,
      errorCalls,
      inputTokens,
      outputTokens,
      totalTokens,
      totalCostUsd: Math.round(totalCostUsd * 1_000_000) / 1_000_000,
      avgLatencyMs: totalCalls > 0 ? Math.round(latencySum / totalCalls) : 0,
    },
    today: {
      date: todayDate,
      calls: todayCalls,
      totalTokens: todayTokens,
      costUsd: Math.round(todayCost * 1_000_000) / 1_000_000,
    },
    byProvider,
    byModel,
    dailyBuckets,
    recentEvents: events.slice(-Math.max(1, recentLimit)).reverse(),
  };
}
