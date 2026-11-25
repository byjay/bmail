import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {
  analyzeWithPriority,
  DEFAULT_GEMINI_MODEL,
  geminiPrimary,
  sanitizeGeminiErrorText,
} from './geminiPrimary';

test('geminiPrimary sends POST to v1beta generateContent with x-goog-api-key header, generationConfig, and maps usageMetadata', async () => {
  let capturedUrl = '';
  let capturedInit: RequestInit | undefined;

  const mockFetch: typeof fetch = async (input, init) => {
    capturedUrl = String(input);
    capturedInit = init;
    return new Response(
      JSON.stringify({
        candidates: [
          {
            content: {
              parts: [{ text: '{"overview":"정상 분석 완료"}' }],
            },
          },
        ],
        usageMetadata: {
          promptTokenCount: 120,
          candidatesTokenCount: 45,
          totalTokenCount: 165,
        },
      }),
      { status: 200, headers: { 'Content-Type': 'application/json' } }
    );
  };

  const result = await geminiPrimary({
    prompt: '오늘 메일 핵심을 요약해 주세요.',
    systemInstruction: '업무 분석 전문가입니다.',
    apiKey: 'test-valid-secret-value',
    fetchImpl: mockFetch,
  });

  assert.equal(
    capturedUrl,
    `https://generativelanguage.googleapis.com/v1beta/models/${DEFAULT_GEMINI_MODEL}:generateContent`
  );
  assert.equal(capturedUrl.includes('key='), false, 'API key must never appear in URL query params');

  const headers = capturedInit?.headers as Record<string, string>;
  assert.equal(headers['Content-Type'], 'application/json');
  assert.equal(headers['x-goog-api-key'], 'test-valid-secret-value');

  const parsedBody = JSON.parse(String(capturedInit?.body)) as {
    contents: Array<{ parts: Array<{ text: string }> }>;
    systemInstruction?: { parts: Array<{ text: string }> };
    generationConfig: { temperature: number; thinkingConfig: { thinkingLevel: string } };
  };
  assert.equal(parsedBody.contents[0].parts[0].text, '오늘 메일 핵심을 요약해 주세요.');
  assert.equal(parsedBody.systemInstruction?.parts[0].text, '업무 분석 전문가입니다.');
  assert.equal(parsedBody.generationConfig.temperature, 0.3);
  assert.equal(parsedBody.generationConfig.thinkingConfig.thinkingLevel, 'low');

  assert.equal(result.text, '{"overview":"정상 분석 완료"}');
  assert.equal(result.modelId, DEFAULT_GEMINI_MODEL);
  assert.equal(result.usageMetadata?.promptTokenCount, 120);
  assert.equal(result.usageMetadata?.candidatesTokenCount, 45);
  assert.equal(result.usageMetadata?.totalTokenCount, 165);
});

test('analyzeWithPriority runs Gemini Primary 1st and falls back cleanly on missing key, 429, 5xx, empty response, and timeout', async () => {
  // Case 1: Empty / missing API key -> immediate fallback without throwing
  const missingKeyRes = await analyzeWithPriority({
    primary: {
      prompt: '분석 요청',
      apiKey: '',
    },
    fallback: async () => ({
      text: 'secondary-ok-missing-key',
      model: 'Groq · openai/gpt-oss-120b',
      providerUsed: 'groq',
    }),
  });
  assert.equal(missingKeyRes.usedPrimary, false);
  assert.equal(missingKeyRes.providerUsed, 'groq');
  assert.equal(missingKeyRes.text, 'secondary-ok-missing-key');

  // Case 2: HTTP 429 rate limit -> immediate fallback
  const rateLimitedRes = await analyzeWithPriority({
    primary: {
      prompt: '분석 요청',
      apiKey: 'test-valid-secret-value',
      fetchImpl: async () => new Response('Rate limit exceeded', { status: 429 }),
    },
    fallback: async () => ({
      text: 'secondary-ok-429',
      model: 'Claude · claude-sonnet-4-20250514',
      providerUsed: 'claude',
    }),
  });
  assert.equal(rateLimitedRes.usedPrimary, false);
  assert.equal(rateLimitedRes.providerUsed, 'claude');
  assert.equal(rateLimitedRes.text, 'secondary-ok-429');
  assert.match(rateLimitedRes.fallbackReason ?? '', /429/);

  // Case 3: HTTP 503 server error -> immediate fallback
  const serverErrRes = await analyzeWithPriority({
    primary: {
      prompt: '분석 요청',
      apiKey: 'test-valid-secret-value',
      fetchImpl: async () => new Response('Service unavailable', { status: 503 }),
    },
    fallback: async () => ({
      text: 'secondary-ok-503',
      model: 'ChatGPT · gpt-4o-mini',
      providerUsed: 'openai',
    }),
  });
  assert.equal(serverErrRes.usedPrimary, false);
  assert.equal(serverErrRes.text, 'secondary-ok-503');

  // Case 4: Empty candidates response -> immediate fallback
  const emptyRes = await analyzeWithPriority({
    primary: {
      prompt: '분석 요청',
      apiKey: 'test-valid-secret-value',
      fetchImpl: async () =>
        new Response(JSON.stringify({ candidates: [{ content: { parts: [{ text: '   ' }] } }] }), {
          status: 200,
        }),
    },
    fallback: async () => ({
      text: 'secondary-ok-empty',
      model: 'Groq · openai/gpt-oss-120b',
      providerUsed: 'groq',
    }),
  });
  assert.equal(emptyRes.usedPrimary, false);
  assert.equal(emptyRes.text, 'secondary-ok-empty');

  // Case 5: AbortController timeout -> immediate fallback
  const timeoutRes = await analyzeWithPriority({
    primary: {
      prompt: '분석 요청',
      apiKey: 'test-valid-secret-value',
      timeoutMs: 30,
      fetchImpl: (_input, init) =>
        new Promise((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => {
            const abortErr = new Error('The operation was aborted');
            abortErr.name = 'AbortError';
            reject(abortErr);
          });
        }),
    },
    fallback: async () => ({
      text: 'secondary-ok-timeout',
      model: 'Groq · openai/gpt-oss-120b',
      providerUsed: 'groq',
    }),
  });
  assert.equal(timeoutRes.usedPrimary, false);
  assert.equal(timeoutRes.text, 'secondary-ok-timeout');
  assert.match(timeoutRes.fallbackReason ?? '', /타임아웃/);
});

test('sanitizeGeminiErrorText redacts sensitive key patterns and repository contains zero hardcoded AIza keys', () => {
  const sampleSecret = 'FAKE_SECRET_TOKEN_98765';
  const redacted = sanitizeGeminiErrorText(`Error calling endpoint ?key=${sampleSecret} with ${sampleSecret}`, sampleSecret);
  assert.equal(redacted.includes(sampleSecret), false);
  assert.ok(redacted.includes('[REDACTED_KEY]'));

  // Walk src/ and server.ts to verify zero hardcoded AIza keys exist
  const rootDir = process.cwd();
  const keyPattern = new RegExp('AI' + 'za[0-9A-Za-z_-]{10,}');
  const violations: string[] = [];

  const scanFile = (filePath: string) => {
    const content = fs.readFileSync(filePath, 'utf8');
    if (keyPattern.test(content)) {
      violations.push(path.relative(rootDir, filePath));
    }
  };

  const walk = (dir: string) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (entry.name === 'node_modules' || entry.name === 'dist' || entry.name.startsWith('.')) continue;
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(full);
      } else if (/\.(ts|tsx|js|jsx|json|env\.example)$/.test(entry.name)) {
        scanFile(full);
      }
    }
  };

  scanFile(path.join(rootDir, 'server.ts'));
  walk(path.join(rootDir, 'src'));
  assert.deepEqual(violations, []);
});
