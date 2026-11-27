import assert from 'node:assert/strict';
import test from 'node:test';
import {
  maskAnthropicApiKey,
  parseAssistantJsonRobust,
  sanitizeAnthropicError,
  sanitizeClaudeModelOption,
  upsertEnvKeyInText,
  validateAnthropicApiKeyInput,
} from './localClaudeEnv';

test('maskAnthropicApiKey and sanitizeAnthropicError never expose raw secret values', () => {
  const sampleKey = 'sk-ant-api03-testsecretvalue9876543210';
  assert.equal(maskAnthropicApiKey(sampleKey), 'sk-ant-••••••3210');
  assert.equal(maskAnthropicApiKey(''), '');

  const sanitized = sanitizeAnthropicError(`Auth error for ${sampleKey}`, sampleKey);
  assert.ok(!sanitized.includes(sampleKey));
  assert.ok(sanitized.includes('[REDACTED_CLAUDE_KEY]'));
});

test('validateAnthropicApiKeyInput accepts valid keys and rejects whitespace or placeholders', () => {
  assert.equal(validateAnthropicApiKeyInput('sk-ant-api03-validkey1234567890').ok, true);
  assert.equal(validateAnthropicApiKeyInput('').ok, false);
  assert.equal(validateAnthropicApiKeyInput('MY_ANTHROPIC_API_KEY').ok, false);
  assert.equal(validateAnthropicApiKeyInput('sk-ant-has\nnewline1234567890').ok, false);
  assert.equal(validateAnthropicApiKeyInput('short').ok, false);
});

test('upsertEnvKeyInText updates, appends, and removes keys while preserving other lines', () => {
  const initial = '# Bmail env\nPORT=3000\nANTHROPIC_API_KEY=old-key-123\nGROQ_API_KEY=gsk_keep\n';
  const updated = upsertEnvKeyInText(initial, 'ANTHROPIC_API_KEY', 'sk-ant-new-9999');
  assert.ok(updated.includes('ANTHROPIC_API_KEY=sk-ant-new-9999'));
  assert.ok(updated.includes('PORT=3000'));
  assert.ok(updated.includes('GROQ_API_KEY=gsk_keep'));
  assert.ok(!updated.includes('old-key-123'));

  const appended = upsertEnvKeyInText(updated, 'ANTHROPIC_MODEL', 'claude-sonnet-5-5');
  assert.ok(appended.includes('ANTHROPIC_MODEL=claude-sonnet-5-5'));

  const removed = upsertEnvKeyInText(appended, 'ANTHROPIC_API_KEY', '');
  assert.ok(!removed.includes('ANTHROPIC_API_KEY='));
  assert.ok(removed.includes('ANTHROPIC_MODEL=claude-sonnet-5-5'));
});

test('parseAssistantJsonRobust parses fenced JSON and repairs unescaped newlines inside strings', () => {
  const fenced = '```json\n{"summaryLines":["라인1","라인2"],"actionChecklist":["조치1"]}\n```';
  const parsed1 = parseAssistantJsonRobust(fenced);
  assert.ok(parsed1);
  assert.deepEqual(parsed1.summaryLines, ['라인1', '라인2']);

  const unescapedNewline =
    '{"summaryLines":["1050호선 간섭 검토 필요"],"replyDrafts":[{"id":"formal","toneLabel":"공식","subject":"RE: 테스트","body":"안녕하세요.\n확인했습니다.\n감사합니다."}]}';
  const parsed2 = parseAssistantJsonRobust(unescapedNewline);
  assert.ok(parsed2);
  assert.equal((parsed2.summaryLines as string[])[0], '1050호선 간섭 검토 필요');

  assert.equal(sanitizeClaudeModelOption('claude-sonnet-5-5'), 'claude-sonnet-5-5');
  assert.equal(sanitizeClaudeModelOption('unknown-model'), 'auto');
});
