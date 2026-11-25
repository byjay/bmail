import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  configureLlmUsageLedgerPath,
  estimateModelCostUsd,
  getLlmUsageSummary,
  recordLlmUsageEvent,
  resetLlmUsageLedger,
} from './llmUsageTracker';

test('llmUsageTracker records exact and estimated token usage, aggregates by model/provider, and persists to disk', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-llm-usage-'));
  const ledgerFile = path.join(tmpDir, 'llm-usage-ledger.json');
  configureLlmUsageLedgerPath(ledgerFile);
  resetLlmUsageLedger();

  try {
    const e1 = recordLlmUsageEvent({
      provider: 'claude',
      modelId: 'claude-haiku-5-5',
      tier: 'low',
      feature: '메일 AI 어시스턴트',
      inputTokens: 1000,
      outputTokens: 500,
      latencyMs: 650,
      status: 'ok',
    });
    assert.equal(e1.totalTokens, 1500);
    assert.equal(e1.estimatedTokens, false);
    assert.ok(e1.costUsd > 0);

    const e2 = recordLlmUsageEvent({
      provider: 'claude',
      modelId: 'claude-opus-5-5',
      tier: 'high',
      feature: '대시보드 종합 분석글',
      inputTokens: 2000,
      outputTokens: 1000,
      latencyMs: 2400,
      status: 'ok',
    });
    assert.equal(e2.totalTokens, 3000);
    assert.ok(e2.costUsd > e1.costUsd);

    const e3 = recordLlmUsageEvent({
      provider: 'groq',
      modelId: 'openai/gpt-oss-120b',
      feature: 'Groq 무료 청킹',
      promptChars: 3500,
      responseChars: 700,
      latencyMs: 320,
      status: 'ok',
    });
    assert.equal(e3.estimatedTokens, true);
    assert.equal(e3.costUsd, 0);

    const summary = getLlmUsageSummary(10);
    assert.equal(summary.totals.totalCalls, 3);
    assert.equal(summary.totals.successCalls, 3);
    assert.equal(summary.totals.errorCalls, 0);
    assert.equal(summary.byProvider.claude.calls, 2);
    assert.equal(summary.byProvider.groq.calls, 1);
    assert.equal(summary.byModel.length, 3);
    assert.equal(summary.recentEvents.length, 3);

    // Reload from disk
    configureLlmUsageLedgerPath(ledgerFile);
    const reloaded = getLlmUsageSummary(10);
    assert.equal(reloaded.totals.totalCalls, 3);

    resetLlmUsageLedger();
    assert.equal(getLlmUsageSummary().totals.totalCalls, 0);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('estimateModelCostUsd prices Haiku < Sonnet < Opus accurately and returns $0 for Groq', () => {
  const haikuCost = estimateModelCostUsd('claude', 'claude-haiku-5-5', 100_000, 20_000);
  const sonnetCost = estimateModelCostUsd('claude', 'claude-sonnet-5-5', 100_000, 20_000);
  const opusCost = estimateModelCostUsd('claude', 'claude-opus-5-5', 100_000, 20_000);
  const groqCost = estimateModelCostUsd('groq', 'openai/gpt-oss-120b', 100_000, 20_000);

  assert.ok(haikuCost > 0);
  assert.ok(sonnetCost > haikuCost);
  assert.ok(opusCost > sonnetCost);
  assert.equal(groqCost, 0);
});
