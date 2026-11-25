import test from 'node:test';
import assert from 'node:assert/strict';
import type { MailItem } from '../types/mail';
import {
  buildAgentBuilderArchiveJsonl,
  callAgentBuilderRag,
  evaluateCreditExpiryGuard,
  getHybridCreditStatus,
  queryHistoricalMailHybrid,
} from './hybridCreditRouter';

const SAMPLE_EMAILS: MailItem[] = [
  {
    id: 'mail-100',
    folder: 'inbox',
    subject: '[PRJ-H101] 지난달 케이블 견적 협의',
    fromName: '김영업',
    fromEmail: 'sales@partner.example',
    to: ['me@company.example'],
    date: '2026-10-15',
    timeLabel: '10:00',
    timestamp: '2026-10-15T10:00:00Z',
    snippet: '케이블 견적서 송부드립니다.',
    body: '지난달 요청하신 케이블 견적 12,000,000원 송부드립니다.',
    isRead: true,
    hasAttachment: true,
    attachments: [
      { id: 'a1', filename: 'quote_h101.pdf', kind: 'pdf', mimeType: 'application/pdf', sizeBytes: 1024 },
    ],
    source: 'pst-import',
  },
];

test('evaluateCreditExpiryGuard allows before expiry and hard-stops after 2027-02-02 to prevent billing', () => {
  const before = evaluateCreditExpiryGuard({
    expiryDate: '2027-02-02',
    hardStopAfterExpiry: true,
    nowDate: new Date('2026-11-20T00:00:00Z'),
  });
  assert.equal(before.isExpired, false);
  assert.equal(before.allowAgentBuilderCall, true);
  assert.ok(before.daysRemaining > 60);

  const near = evaluateCreditExpiryGuard({
    expiryDate: '2027-02-02',
    hardStopAfterExpiry: true,
    nowDate: new Date('2027-01-25T00:00:00Z'),
  });
  assert.equal(near.isNearExpiry, true);
  assert.equal(near.allowAgentBuilderCall, true);

  const after = evaluateCreditExpiryGuard({
    expiryDate: '2027-02-02',
    hardStopAfterExpiry: true,
    nowDate: new Date('2027-02-03T00:00:00Z'),
  });
  assert.equal(after.isExpired, true);
  assert.equal(after.allowAgentBuilderCall, false);
  assert.match(after.statusMessage, /자동 과금 방지/);
});

test('callAgentBuilderRag invokes Cloudflare Worker proxy with Bearer token and parses summary/citations', async () => {
  let capturedUrl = '';
  let capturedAuth = '';
  const mockFetch: typeof fetch = async (input, init) => {
    capturedUrl = String(input);
    const hdrs = init?.headers as Record<string, string>;
    capturedAuth = hdrs.Authorization || '';
    return new Response(
      JSON.stringify({
        summary: { summaryText: '지난달 PRJ-H101 케이블 견적이 10월 15일에 접수되었습니다.' },
        results: [
          {
            id: 'doc-1',
            document: {
              id: 'mail-100',
              structData: {
                subject: '[PRJ-H101] 지난달 케이블 견적 협의',
                date: '2026-10-15',
                snippet: '케이블 견적서 송부',
              },
            },
          },
        ],
      }),
      { status: 200, headers: { 'Content-Type': 'application/json' } }
    );
  };

  const res = await callAgentBuilderRag({
    query: '지난달 케이블 견적 정리해줘',
    workerUrl: 'https://agent-proxy.example.workers.dev/search',
    workerToken: 'secret-worker-token-123',
    fetchImpl: mockFetch,
  });

  assert.equal(capturedUrl, 'https://agent-proxy.example.workers.dev/search');
  assert.equal(capturedAuth, 'Bearer secret-worker-token-123');
  assert.match(res.summary, /PRJ-H101/);
  assert.equal(res.citations.length, 1);
});

test('queryHistoricalMailHybrid prioritizes Agent Builder RAG -> falls back to Gemini Primary -> falls back to Local on failure or expiry', async () => {
  // Case 1: Active Agent Builder -> Priority 1
  const p1 = await queryHistoricalMailHybrid({
    question: '지난달 케이블 견적',
    emails: SAMPLE_EMAILS,
    projects: [],
    calendarEvents: [],
    nowDate: new Date('2026-11-20T00:00:00Z'),
    agentBuilderOverride: { workerUrl: 'https://worker.example.dev/rag' },
    fetchImpl: async () =>
      new Response(JSON.stringify({ summary: { summaryText: 'Agent Builder RAG 응답 완료' } }), {
        status: 200,
      }),
  });
  assert.equal(p1.creditRoute, 'agent-builder-rag');
  assert.equal(p1.answer, 'Agent Builder RAG 응답 완료');

  // Case 2: Expired Agent Builder (2027-02-03) -> Hard-stop skips Agent Builder and uses Priority 2 Gemini Primary!
  let calledUrls: string[] = [];
  const p2 = await queryHistoricalMailHybrid({
    question: '지난달 케이블 견적',
    emails: SAMPLE_EMAILS,
    projects: [],
    calendarEvents: [],
    nowDate: new Date('2027-02-03T00:00:00Z'),
    agentBuilderOverride: { workerUrl: 'https://worker.example.dev/rag' },
    geminiApiKeyOverride: 'test-gemini-key-valid',
    fetchImpl: async (input) => {
      calledUrls.push(String(input));
      return new Response(
        JSON.stringify({
          candidates: [{ content: { parts: [{ text: 'Gemini 4 Hydrogen 폴백 응답' }] } }],
          usageMetadata: { totalTokenCount: 88 },
        }),
        { status: 200 }
      );
    },
  });
  assert.equal(p2.creditRoute, 'gemini-primary');
  assert.equal(p2.answer, 'Gemini 4 Hydrogen 폴백 응답');
  assert.equal(calledUrls.some((u) => u.includes('worker.example.dev')), false);
  assert.equal(p2.usageMetadata?.totalTokenCount, 88);

  // Case 3: Both fail / unconfigured -> Priority 3 Local Deterministic 0-LLM succeeds
  const p3 = await queryHistoricalMailHybrid({
    question: '지난달 케이블 견적',
    emails: SAMPLE_EMAILS,
    projects: [],
    calendarEvents: [],
    geminiApiKeyOverride: '',
  });
  assert.equal(p3.creditRoute, 'deterministic-local');
  assert.ok(p3.answer.length > 0);
});

test('buildAgentBuilderArchiveJsonl outputs valid NDJSON and getHybridCreditStatus never leaks secret values', () => {
  const jsonl = buildAgentBuilderArchiveJsonl(SAMPLE_EMAILS, true);
  const lines = jsonl.split('\n').filter(Boolean);
  assert.equal(lines.length, 1);
  const parsed = JSON.parse(lines[0]) as {
    id: string;
    structData: { subject: string; body: string };
    content: { mimeType: string; rawBytes: string };
  };
  assert.equal(parsed.id, 'mail-100');
  assert.equal(parsed.content.mimeType, 'text/plain');
  assert.equal(parsed.structData.body.includes('12,000,000원'), false);

  const status = getHybridCreditStatus(new Date('2026-11-20T00:00:00Z'));
  assert.equal(status.setupChecklist.length, 4);
  assert.equal(status.agentBuilderRagCredit.expiryGuard.expiryDate, '2027-02-02');
});
