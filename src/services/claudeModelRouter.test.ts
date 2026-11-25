import test from 'node:test';
import assert from 'node:assert/strict';
import {
  CLAUDE_MODEL_HAIKU,
  CLAUDE_MODEL_OPUS,
  CLAUDE_MODEL_SONNET,
  invokeClaudeMessagesApi,
  selectClaudeModelByComplexity,
} from './claudeModelRouter';

test('selectClaudeModelByComplexity defaults to Haiku 5.5 on routine workloads', () => {
  const decision = selectClaudeModelByComplexity({
    taskKind: 'mail-reduce',
    emailCount: 2,
    documentCount: 0,
    promptChars: 1200,
    customFocus: '',
    explicitModelOverride: 'claude-haiku-5-5',
  });

  assert.equal(decision.modelId, CLAUDE_MODEL_HAIKU);
  assert.equal(decision.tier, 'low');
  assert.equal(decision.effort, 'medium');
  assert.equal(decision.useServerFallback, false);
});

test('selectClaudeModelByComplexity escalates to Sonnet 5.5 for medium tasks (6-19 emails, docs, meeting, keywords)', () => {
  const byEmails = selectClaudeModelByComplexity({
    taskKind: 'mail-reduce',
    emailCount: 8,
    promptChars: 3200,
  });
  assert.equal(byEmails.modelId, CLAUDE_MODEL_SONNET);
  assert.equal(byEmails.tier, 'medium');
  assert.equal(byEmails.useServerFallback, true);

  const byMeeting = selectClaudeModelByComplexity({
    taskKind: 'meeting',
    emailCount: 0,
    promptChars: 800,
  });
  assert.equal(byMeeting.modelId, CLAUDE_MODEL_SONNET);
  assert.equal(byMeeting.tier, 'medium');

  const byDoc = selectClaudeModelByComplexity({
    taskKind: 'document',
    documentCount: 1,
    emailCount: 2,
  });
  assert.equal(byDoc.modelId, CLAUDE_MODEL_SONNET);
  assert.equal(byDoc.tier, 'medium');

  const byKeyword = selectClaudeModelByComplexity({
    taskKind: 'quick-summary',
    emailCount: 1,
    customFocus: '도면 리비전 비교 보고서 작성',
  });
  assert.equal(byKeyword.modelId, CLAUDE_MODEL_SONNET);
});

test('selectClaudeModelByComplexity escalates to Opus 5.5 for high-difficulty or large-volume tasks', () => {
  const byVolume = selectClaudeModelByComplexity({
    taskKind: 'mail-reduce',
    emailCount: 25,
    promptChars: 9000,
  });
  assert.equal(byVolume.modelId, CLAUDE_MODEL_OPUS);
  assert.equal(byVolume.tier, 'high');
  assert.equal(byVolume.effort, 'high');
  assert.equal(byVolume.useServerFallback, true);

  const byPromptChars = selectClaudeModelByComplexity({
    taskKind: 'mail-map',
    emailCount: 3,
    promptChars: 18500,
  });
  assert.equal(byPromptChars.modelId, CLAUDE_MODEL_OPUS);
  assert.equal(byPromptChars.tier, 'high');

  const byKeyword = selectClaudeModelByComplexity({
    taskKind: 'general',
    emailCount: 1,
    customFocus: '계약 보안 리스크 정밀 감리',
  });
  assert.equal(byKeyword.modelId, CLAUDE_MODEL_OPUS);
  assert.equal(byKeyword.effort, 'high');

  const byCrossDoc = selectClaudeModelByComplexity({
    taskKind: 'document',
    documentCount: 2,
    emailCount: 6,
  });
  assert.equal(byCrossDoc.modelId, CLAUDE_MODEL_OPUS);
});

test('selectClaudeModelByComplexity honors explicit overrides and autoTierEnabled=false', () => {
  const pinnedOpus = selectClaudeModelByComplexity({
    taskKind: 'quick-summary',
    emailCount: 1,
    explicitModelOverride: CLAUDE_MODEL_OPUS,
  });
  assert.equal(pinnedOpus.modelId, CLAUDE_MODEL_OPUS);
  assert.equal(pinnedOpus.tier, 'high');
  assert.equal(pinnedOpus.reason, 'explicit-model-override');

  const disabledAuto = selectClaudeModelByComplexity({
    taskKind: 'document',
    emailCount: 30,
    documentCount: 5,
    autoTierEnabled: false,
  });
  assert.equal(disabledAuto.modelId, CLAUDE_MODEL_HAIKU);
  assert.equal(disabledAuto.tier, 'low');
  assert.equal(disabledAuto.useServerFallback, false);
});

test('invokeClaudeMessagesApi sends valid Claude 5.5 headers/payload, omits deprecated sampling params, and handles refusal/redaction', async () => {
  const captured: Array<{ url: string; headers: Record<string, string>; body: Record<string, unknown> }> = [];
  const fakeFetch: typeof fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    captured.push({
      url: String(input),
      headers: (init?.headers || {}) as Record<string, string>,
      body: JSON.parse(String(init?.body || '{}')) as Record<string, unknown>,
    });
    return new Response(
      JSON.stringify({
        stop_reason: 'end_turn',
        content: [{ type: 'text', text: '{"ok":true}' }],
        usage: { input_tokens: 12, output_tokens: 5 },
      }),
      { status: 200, headers: { 'content-type': 'application/json' } }
    );
  }) as typeof fetch;

  const resHaiku = await invokeClaudeMessagesApi({
    apiKey: 'sk-ant-test-secret-key-0001',
    systemPrompt: 'sys',
    userPrompt: 'hello',
    routing: { taskKind: 'quick-summary', emailCount: 1 },
    fetchImpl: fakeFetch,
  });
  assert.equal(resHaiku.decision.modelId, CLAUDE_MODEL_HAIKU);
  assert.equal(captured[0].headers['anthropic-beta'], undefined);
  assert.equal(captured[0].body.max_tokens, 6144);
  assert.equal(captured[0].body.fallbacks, undefined);
  assert.equal(captured[0].body.temperature, undefined);
  assert.equal(captured[0].body.top_p, undefined);
  assert.deepEqual(captured[0].body.output_config, { effort: 'medium' });

  const resOpus = await invokeClaudeMessagesApi({
    apiKey: 'sk-ant-test-secret-key-0001',
    systemPrompt: 'sys',
    userPrompt: 'analyze contract',
    routing: { taskKind: 'mail-reduce', customFocus: '계약 정밀 감리' },
    fetchImpl: fakeFetch,
  });
  assert.equal(resOpus.decision.modelId, CLAUDE_MODEL_OPUS);
  assert.equal(captured[1].headers['anthropic-beta'], 'server-side-fallback-2026-07-01');
  assert.equal(captured[1].body.max_tokens, 16384);
  assert.equal(captured[1].body.fallbacks, 'default');
  assert.deepEqual(captured[1].body.output_config, { effort: 'high' });

  const refusalFetch: typeof fetch = (async () =>
    new Response(
      JSON.stringify({
        stop_reason: 'refusal',
        stop_details: { category: 'policy' },
        content: [],
      }),
      { status: 200, headers: { 'content-type': 'application/json' } }
    )) as typeof fetch;

  await assert.rejects(
    () =>
      invokeClaudeMessagesApi({
        apiKey: 'sk-ant-test-secret-key-0001',
        systemPrompt: 'sys',
        userPrompt: 'test',
        routing: { taskKind: 'quick-summary' },
        fetchImpl: refusalFetch,
      }),
    /stop_reason: refusal/
  );

  const errFetch: typeof fetch = (async () =>
    new Response('invalid x-api-key sk-ant-test-secret-key-0001 rejected', { status: 401 })) as typeof fetch;

  await assert.rejects(
    () =>
      invokeClaudeMessagesApi({
        apiKey: 'sk-ant-test-secret-key-0001',
        systemPrompt: 'sys',
        userPrompt: 'test',
        routing: { taskKind: 'quick-summary' },
        fetchImpl: errFetch,
      }),
    (err: Error) => !err.message.includes('sk-ant-test-secret-key-0001') && err.message.includes('[REDACTED_KEY]')
  );
});
