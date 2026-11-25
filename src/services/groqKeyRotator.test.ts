import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  collectGroqApiKeys,
  DEFAULT_GROQ_MODEL,
  GroqKeyRotator,
  runGroqParallelBatch,
  shouldEscalateToClaude,
} from './groqKeyRotator';

function makeSseStreamResponse(chunks: string[]): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) {
        controller.enqueue(encoder.encode(chunk));
      }
      controller.close();
    },
  });
  return new Response(stream, {
    status: 200,
    headers: { 'Content-Type': 'text/event-stream; charset=utf-8' },
  });
}

test('collectGroqApiKeys gathers up to 5 valid keys, deduplicates, and filters short placeholders', () => {
  const keys = collectGroqApiKeys(
    {
      GROQ_API_KEY: 'x', // placeholder < 10 chars ignored
      GROQ_API_KEY_2: 'gsk_key_two_123456',
      GROQ_API_KEY_3: 'gsk_key_three_1234',
      GROQ_API_KEYS: 'gsk_key_two_123456, gsk_key_four_12345, gsk_key_five_12345',
    },
    {
      GROQ_API_KEY: 'gsk_key_one_123456',
      GROQ_API_KEY_5: 'gsk_key_six_ignored',
    }
  );

  assert.equal(keys.length, 5);
  assert.deepEqual(keys, [
    'gsk_key_two_123456',
    'gsk_key_three_1234',
    'gsk_key_four_12345',
    'gsk_key_five_12345',
    'gsk_key_one_123456',
  ]);
});

test('GroqKeyRotator round-robins across 5 keys and sends exact streaming params', async () => {
  let currentNow = 1_000_000;
  const rotator = new GroqKeyRotator({ nowFn: () => currentNow });
  const envOverrides = {
    GROQ_API_KEY_1: 'gsk_slot_1_abcdef123',
    GROQ_API_KEY_2: 'gsk_slot_2_abcdef123',
    GROQ_API_KEY_3: 'gsk_slot_3_abcdef123',
    GROQ_API_KEY_4: 'gsk_slot_4_abcdef123',
    GROQ_API_KEY_5: 'gsk_slot_5_abcdef123',
  };

  const seenAuthHeaders: string[] = [];
  const seenPayloads: Array<Record<string, unknown>> = [];

  const mockFetch = (async (_url: string | URL | Request, init?: RequestInit) => {
    const auth = String((init?.headers as Record<string, string>)?.Authorization || '');
    seenAuthHeaders.push(auth);
    seenPayloads.push(JSON.parse(String(init?.body || '{}')) as Record<string, unknown>);
    return makeSseStreamResponse([
      'data: {"choices":[{"delta":{"content":"안녕"}}]}\n\n',
      'data: {"choices":[{"delta":{"content":"하세요"}}],"x_groq":{"usage":{"prompt_tokens":10,"completion_tokens":5}}}\n\n',
      'data: [DONE]\n\n',
    ]);
  }) as typeof fetch;

  const slotsUsed: number[] = [];
  for (let i = 0; i < 6; i++) {
    currentNow += 100;
    const res = await rotator.invokeStreamingCompletion({
      systemPrompt: 'sys',
      userPrompt: `hello ${i}`,
      envOverrides,
      fetchImpl: mockFetch,
    });
    assert.equal(res.text, '안녕하세요');
    assert.equal(res.inputTokens, 10);
    assert.equal(res.outputTokens, 5);
    slotsUsed.push(res.keySlotUsed);
  }

  assert.deepEqual(slotsUsed, [1, 2, 3, 4, 5, 1]);
  assert.equal(seenAuthHeaders[0], 'Bearer gsk_slot_1_abcdef123');
  assert.equal(seenAuthHeaders[4], 'Bearer gsk_slot_5_abcdef123');
  assert.equal(seenAuthHeaders[5], 'Bearer gsk_slot_1_abcdef123');

  const firstPayload = seenPayloads[0];
  assert.equal(firstPayload.model, DEFAULT_GROQ_MODEL);
  assert.equal(firstPayload.temperature, 1);
  assert.equal(firstPayload.max_completion_tokens, 2048);
  assert.equal(firstPayload.top_p, 1);
  assert.equal(firstPayload.reasoning_effort, 'medium');
  assert.equal(firstPayload.stream, true);
  assert.equal(firstPayload.stop, null);
});

test('GroqKeyRotator fails over cleanly on 429 and puts failing key on cooldown', async () => {
  let currentNow = 2_000_000;
  const rotator = new GroqKeyRotator({ nowFn: () => currentNow });
  const envOverrides = {
    GROQ_API_KEY_1: 'gsk_fail_1_abcdef123',
    GROQ_API_KEY_2: 'gsk_good_2_abcdef123',
  };

  const calledKeys: string[] = [];
  const mockFetch = (async (_url: string | URL | Request, init?: RequestInit) => {
    const auth = String((init?.headers as Record<string, string>)?.Authorization || '');
    calledKeys.push(auth);
    if (auth.includes('gsk_fail_1')) {
      return new Response('{"error":"rate_limit_exceeded"}', { status: 429 });
    }
    return makeSseStreamResponse([
      'data: {"choices":[{"delta":{"content":"성공 응답"}}]}\n\n',
      'data: [DONE]\n\n',
    ]);
  }) as typeof fetch;

  // Call 1 starts at slot 1 -> gets 429 -> fails over immediately to slot 2
  const res1 = await rotator.invokeStreamingCompletion({
    userPrompt: 'req 1',
    envOverrides,
    fetchImpl: mockFetch,
  });
  assert.equal(res1.keySlotUsed, 2);
  assert.equal(res1.text, '성공 응답');
  assert.equal(rotator.getPoolStatus(envOverrides).healthyCount, 1);

  // Call 2 (1s later): slot 1 is on 60s cooldown so ONLY slot 2 should be called
  currentNow += 1000;
  const res2 = await rotator.invokeStreamingCompletion({
    userPrompt: 'req 2',
    envOverrides,
    fetchImpl: mockFetch,
  });
  assert.equal(res2.keySlotUsed, 2);
  assert.equal(calledKeys.length, 3); // call1(fail1, good2) + call2(good2 only)
});

test('shouldEscalateToClaude escalates only on explicit Claude, highPerformance, high tier, or huge prompts', () => {
  assert.equal(
    shouldEscalateToClaude({
      explicitProvider: 'groq',
      promptChars: 2400,
      routing: { taskKind: 'general', emailCount: 2, promptChars: 2400, customFocus: '오늘 메일 요약' },
    }),
    false
  );

  assert.equal(shouldEscalateToClaude({ explicitProvider: 'claude', promptChars: 500 }), true);
  assert.equal(shouldEscalateToClaude({ highPerformance: true, promptChars: 500 }), true);
  assert.equal(shouldEscalateToClaude({ promptChars: 21000 }), true);
  assert.equal(
    shouldEscalateToClaude({
      explicitProvider: 'groq',
      routing: { customFocus: '보안 감사 및 계약 법적 리스크 정밀 평가' },
    }),
    true
  );
});

test('collectGroqApiKeys loads plain 1-key-per-line groq-keys.txt ignoring comments and blanks', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-groq-keys-'));
  try {
    const keysFile = path.join(tmpDir, 'groq-keys.txt');
    fs.writeFileSync(
      keysFile,
      [
        '# comment line',
        '',
        'gsk_file_one_123456',
        '// another comment',
        'GROQ_API_KEY=gsk_file_two_123456',
        'gsk_file_three_1234',
        'gsk_file_four_12345',
        'gsk_file_five_12345',
        'gsk_file_six_ignored',
      ].join('\n'),
      'utf8'
    );
    const keys = collectGroqApiKeys({ BMAIL_GROQ_KEYS_FILE: keysFile }, {});
    assert.deepEqual(keys, [
      'gsk_file_one_123456',
      'gsk_file_two_123456',
      'gsk_file_three_1234',
      'gsk_file_four_12345',
      'gsk_file_five_12345',
    ]);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('collectGroqApiKeys parses comma-separated 5 keys on a single line', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-groq-csv-'));
  try {
    const keysFile = path.join(tmpDir, 'groq-keys.txt');
    fs.writeFileSync(
      keysFile,
      '# comma list\ngsk_csv_one_123456, gsk_csv_two_123456, gsk_csv_three_1234, gsk_csv_four_12345, gsk_csv_five_12345\n',
      'utf8'
    );
    const keys = collectGroqApiKeys({ BMAIL_GROQ_KEYS_FILE: keysFile }, {});
    assert.deepEqual(keys, [
      'gsk_csv_one_123456',
      'gsk_csv_two_123456',
      'gsk_csv_three_1234',
      'gsk_csv_four_12345',
      'gsk_csv_five_12345',
    ]);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('runGroqParallelBatch executes up to 5 workers concurrently while preserving output order', async () => {
  const envOverrides = {
    GROQ_API_KEY_1: 'gsk_par_1_abcdef12',
    GROQ_API_KEY_2: 'gsk_par_2_abcdef12',
    GROQ_API_KEY_3: 'gsk_par_3_abcdef12',
    GROQ_API_KEY_4: 'gsk_par_4_abcdef12',
    GROQ_API_KEY_5: 'gsk_par_5_abcdef12',
  };
  let active = 0;
  let peakActive = 0;
  const items = [10, 20, 30, 40, 50, 60];

  const results = await runGroqParallelBatch(
    items,
    async (item, index) => {
      active += 1;
      peakActive = Math.max(peakActive, active);
      await new Promise((r) => setTimeout(r, (items.length - index) * 8));
      active -= 1;
      return `${index}:${item * 2}`;
    },
    { envOverrides }
  );

  assert.equal(peakActive >= 3 && peakActive <= 5, true, `Expected parallel concurrency 3..5, got ${peakActive}`);
  assert.deepEqual(results, ['0:20', '1:40', '2:60', '3:80', '4:100', '5:120']);
});

test('shouldEscalateToClaude sends advanced analysis to Claude while keeping fast map-chunks on Groq', () => {
  assert.equal(
    shouldEscalateToClaude({ routing: { taskKind: 'executive-report', promptChars: 1200 } }),
    true
  );
  assert.equal(
    shouldEscalateToClaude({ routing: { taskKind: 'doc-analysis', promptChars: 1200 } }),
    true
  );
  assert.equal(
    shouldEscalateToClaude({ routing: { taskKind: 'mail-map', emailCount: 50, promptChars: 5500 } }),
    false
  );
});
