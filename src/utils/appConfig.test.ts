import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildEnvSecretStatuses,
  buildEnvTemplateText,
  DEFAULT_APP_CONFIG,
  sanitizeAppConfig,
  stripSecretKeysFromObject,
} from './appConfig';
import { auditAndMaskEmails } from '../services/agentService';
import type { MailItem } from '../types/mail';

test('stripSecretKeysFromObject recursively removes API keys, secrets, tokens, and passwords', () => {
  const dirty = {
    plugin: {
      displayName: 'MailLog AI',
      adminPassword: 'super-secret-password',
    },
    ai: {
      defaultProvider: 'groq',
      GROQ_API_KEY: 'gsk_leak12345678901234567890',
      nested: {
        clientSecret: 'azure-secret',
        accessToken: 'jwt-token',
        safeFlag: true,
      },
    },
  };

  const cleaned = stripSecretKeysFromObject(dirty) as Record<string, Record<string, unknown>>;
  assert.equal('adminPassword' in cleaned.plugin, false);
  assert.equal('GROQ_API_KEY' in cleaned.ai, false);
  assert.equal('clientSecret' in (cleaned.ai.nested as Record<string, unknown>), false);
  assert.equal('accessToken' in (cleaned.ai.nested as Record<string, unknown>), false);
  assert.equal((cleaned.ai.nested as Record<string, unknown>).safeFlag, true);
  assert.equal(cleaned.plugin.displayName, 'MailLog AI');
});

test('sanitizeAppConfig clamps out-of-range numbers and falls back invalid enums to defaults', () => {
  const sanitized = sanitizeAppConfig({
    plugin: {
      displayName: '   ',
      supportsPinning: false,
    },
    ai: {
      defaultProvider: 'unknown-provider',
      maxEmailsPerRun: 9999,
      models: {
        groq: 'openai/gpt-oss-120b',
      },
    },
    agent: {
      autoRunOnImport: true,
      defaultBlogTone: 'invalid-tone',
      customSystemDirective: '  핵심 성과 위주  ',
    },
    pst: {
      maxExtractedMessages: 5,
      maxUploadSizeMb: 9000,
    },
  });

  assert.equal(sanitized.plugin.displayName, DEFAULT_APP_CONFIG.plugin.displayName);
  assert.equal(sanitized.plugin.supportsPinning, false);
  assert.equal(sanitized.ai.defaultProvider, DEFAULT_APP_CONFIG.ai.defaultProvider);
  assert.equal(sanitized.ai.maxEmailsPerRun, 100);
  assert.equal(sanitized.ai.models.groq, 'openai/gpt-oss-120b');
  assert.equal(sanitized.agent.autoRunOnImport, true);
  assert.equal(sanitized.agent.defaultBlogTone, DEFAULT_APP_CONFIG.agent.defaultBlogTone);
  assert.equal(sanitized.agent.customSystemDirective, '핵심 성과 위주');
  assert.equal(sanitized.pst.maxExtractedMessages, 20);
  assert.equal(sanitized.pst.maxUploadSizeMb, 500);
});

test('buildEnvSecretStatuses returns only booleans and ignores MY_* placeholders', () => {
  const statuses = buildEnvSecretStatuses({
    GROQ_API_KEY: 'gsk_real_secret_value',
    ANTHROPIC_API_KEY: 'MY_ANTHROPIC_API_KEY',
    OPENAI_API_KEY: '',
    MICROSOFT_CLIENT_ID: '1234-5678',
    MICROSOFT_CLIENT_SECRET: undefined,
    CONFIG_ADMIN_PASSWORD: 'admin-pw',
  });

  const byKey = new Map(statuses.map((s) => [s.key, s]));
  assert.equal(byKey.get('GROQ_API_KEY')?.configured, true);
  assert.equal(byKey.get('ANTHROPIC_API_KEY')?.configured, false);
  assert.equal(byKey.get('OPENAI_API_KEY')?.configured, false);
  assert.equal(byKey.get('MICROSOFT_CLIENT_ID')?.configured, true);
  assert.equal(byKey.get('MICROSOFT_CLIENT_SECRET')?.configured, false);
  assert.equal(byKey.get('CONFIG_ADMIN_PASSWORD')?.configured, true);

  const serialized = JSON.stringify(statuses);
  assert.ok(!serialized.includes('gsk_real_secret_value'));
  assert.ok(!serialized.includes('admin-pw'));
});

test('buildEnvTemplateText emits clean empty placeholders and exact callback URL', () => {
  const template = buildEnvTemplateText('https://mail.corp.example.com');
  assert.ok(template.includes('APP_URL="https://mail.corp.example.com"'));
  assert.ok(template.includes('https://mail.corp.example.com/auth/microsoft/callback'));
  assert.ok(template.includes('GROQ_API_KEY=""'));
  assert.ok(template.includes('CONFIG_ADMIN_PASSWORD=""'));
});

test('auditAndMaskEmails counts all redacted sensitive patterns accurately', () => {
  const mails: MailItem[] = [
    {
      id: 'm1',
      folder: 'inbox',
      subject: '점검 요청',
      fromName: '담당자',
      fromEmail: 'ops@corp.kr',
      to: ['me@corp.kr'],
      date: '2026-11-19',
      timeLabel: '10:00',
      timestamp: '2026-11-19T10:00:00+09:00',
      snippet: '요약',
      body: '담당자 이메일 ops@corp.kr, 연락처 010-1234-5678, 서버 192.168.0.10, 키 sk-proj1234567890abcdef, 예산 3,500만원',
      isRead: false,
      hasAttachment: false,
      importance: 'normal',
      source: 'pst-import',
    },
  ];

  const maskedAudit = auditAndMaskEmails(mails, true);
  assert.equal(maskedAudit.redactedCount, 5);

  const skippedAudit = auditAndMaskEmails(mails, false);
  assert.equal(skippedAudit.redactedCount, 0);
});

test('SupportsPinning manifest option toggles XML and Unified JSON pinnable flags cleanly', async () => {
  const { buildOutlookManifestXml, buildUnifiedManifestJson } = await import('./outlookManifest');
  const { loadTaskpanePinnedPreference } = await import('../services/officeBridge');

  const pinnedXml = buildOutlookManifestXml({
    baseUrl: 'https://maillog.example.com',
    supportsPinning: true,
  });
  assert.ok(pinnedXml.includes('<SupportsPinning>true</SupportsPinning>'));

  const unpinnedXml = buildOutlookManifestXml({
    baseUrl: 'https://maillog.example.com',
    supportsPinning: false,
  });
  assert.ok(unpinnedXml.includes('<SupportsPinning>false</SupportsPinning>'));

  const unpinnedUnified = JSON.parse(
    buildUnifiedManifestJson({ baseUrl: 'https://maillog.example.com', supportsPinning: false })
  );
  assert.equal(unpinnedUnified.extensions[0].runtimes[0].actions[0].pinnable, false);

  // In Node (no window), loadTaskpanePinnedPreference respects manifest SupportsPinning default
  assert.equal(loadTaskpanePinnedPreference(true), true);
  assert.equal(loadTaskpanePinnedPreference(false), false);
});
