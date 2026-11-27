export const ALLOWED_CLAUDE_MODELS = [
  'auto',
  'claude-haiku-5-5',
  'claude-sonnet-5-5',
  'claude-opus-5-5',
] as const;

export type AllowedClaudeModel = (typeof ALLOWED_CLAUDE_MODELS)[number];

export function maskAnthropicApiKey(rawKey: string | undefined | null): string {
  const clean = String(rawKey || '').trim();
  if (!clean) return '';
  if (clean.length <= 10) return '••••••';
  const prefix = clean.startsWith('sk-ant-') ? 'sk-ant-' : clean.slice(0, 4);
  const suffix = clean.slice(-4);
  return `${prefix}••••••${suffix}`;
}

export function validateAnthropicApiKeyInput(raw: unknown): { ok: true; value: string } | { ok: false; error: string } {
  const clean = String(raw ?? '').trim();
  if (!clean) {
    return { ok: false, error: 'Claude API 키를 입력해 주세요.' };
  }
  if (/[\r\n\0]/.test(clean) || /\s/.test(clean)) {
    return { ok: false, error: 'API 키에 공백이나 줄바꿈이 포함될 수 없습니다.' };
  }
  if (clean.length < 16) {
    return { ok: false, error: '유효한 Anthropic Claude API 키(sk-ant-...) 형식이 아닙니다.' };
  }
  if (/^(MY_ANTHROPIC|placeholder|changeme|example|your[_-]?key)/i.test(clean)) {
    return { ok: false, error: '예시 문자열이 아닌 실제 Claude API 키를 입력해 주세요.' };
  }
  return { ok: true, value: clean };
}

export function sanitizeClaudeModelOption(raw: unknown): AllowedClaudeModel {
  const clean = String(raw ?? '').trim();
  if ((ALLOWED_CLAUDE_MODELS as readonly string[]).includes(clean)) {
    return clean as AllowedClaudeModel;
  }
  return 'auto';
}

export function sanitizeAnthropicError(message: string, apiKey?: string): string {
  let clean = String(message || 'Claude API 호출 중 오류가 발생했습니다.');
  if (apiKey && apiKey.trim().length >= 8) {
    clean = clean.split(apiKey.trim()).join('[REDACTED_CLAUDE_KEY]');
  }
  clean = clean.replace(/sk-ant-[A-Za-z0-9_-]+/gi, '[REDACTED_CLAUDE_KEY]');
  return clean.slice(0, 260);
}

/**
 * Safely updates, appends, or removes a single KEY=VALUE entry in `.env` text
 * while preserving comments and unrelated environment variables.
 */
export function upsertEnvKeyInText(envText: string, keyName: string, nextValue: string): string {
  const safeKey = keyName.trim();
  const safeVal = nextValue.replace(/[\r\n\0]/g, '').trim();
  const lines = envText ? envText.replace(/\r\n/g, '\n').split('\n') : [];
  const keyRe = new RegExp(`^\\s*${safeKey.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*=`);

  let found = false;
  const out: string[] = [];
  for (const line of lines) {
    if (keyRe.test(line)) {
      found = true;
      if (safeVal) {
        out.push(`${safeKey}=${safeVal}`);
      }
      continue;
    }
    out.push(line);
  }

  if (!found && safeVal) {
    if (out.length > 0 && out[out.length - 1].trim() !== '') {
      out.push(`${safeKey}=${safeVal}`);
    } else if (out.length > 0 && out[out.length - 1].trim() === '') {
      out[out.length - 1] = `${safeKey}=${safeVal}`;
      out.push('');
    } else {
      out.push(`${safeKey}=${safeVal}`);
    }
  }

  const joined = out.join('\n').replace(/\n{3,}/g, '\n\n');
  return joined.endsWith('\n') || joined === '' ? joined : `${joined}\n`;
}

/**
 * Escapes literal control characters (newlines/tabs) that LLMs occasionally emit inside JSON string literals.
 */
function escapeLiteralControlCharsInJsonStrings(input: string): string {
  let out = '';
  let inString = false;
  let escaped = false;
  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    if (inString) {
      if (escaped) {
        out += ch;
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        out += ch;
        escaped = true;
        continue;
      }
      if (ch === '"') {
        out += ch;
        inString = false;
        continue;
      }
      if (ch === '\n') {
        out += '\\n';
        continue;
      }
      if (ch === '\r') {
        out += '\\r';
        continue;
      }
      if (ch === '\t') {
        out += '\\t';
        continue;
      }
      out += ch;
    } else {
      if (ch === '"') {
        inString = true;
      }
      out += ch;
    }
  }
  return out;
}

/**
 * Robustly extracts and parses the assistant JSON payload from Claude/Groq text output.
 */
export function parseAssistantJsonRobust(rawText: string): Record<string, unknown> | null {
  const stripped = String(rawText || '')
    .trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();
  const match = stripped.match(/\{[\s\S]*\}/);
  if (!match) return null;

  const candidates = [match[0], escapeLiteralControlCharsInJsonStrings(match[0])];
  for (const candidate of candidates) {
    try {
      const obj = JSON.parse(candidate) as Record<string, unknown>;
      if (
        obj &&
        typeof obj === 'object' &&
        Array.isArray(obj.summaryLines) &&
        obj.summaryLines.some((line) => String(line || '').trim().length > 0)
      ) {
        return obj;
      }
    } catch {
      // try repaired candidate
    }
  }
  return null;
}
