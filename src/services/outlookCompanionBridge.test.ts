import test from 'node:test';
import assert from 'node:assert/strict';
import { verifyOutlookItemSafety } from './outlookCompanionBridge';

test('verifyOutlookItemSafety blocks composing mail with attachment mention and 0 files', () => {
  const report = verifyOutlookItemSafety(
    {
      subject: 'H042 도면 송부 건',
      fromEmail: 'design@seastargo.com',
      to: ['client@yard.com'],
      body: '안녕하세요. 요청하신 도면을 첨부드립니다. 확인 부탁드립니다.',
      attachmentsCount: 0,
      attachments: [],
    },
    'composing'
  );

  assert.equal(report.status, 'block');
  assert.ok(report.findings.some((f) => f.code === 'missing_attachment' && f.severity === 'critical'));
  assert.ok(report.score <= 60);
});

test('verifyOutlookItemSafety flags missing recipient (critical) and missing subject (warning) when composing', () => {
  const report = verifyOutlookItemSafety(
    {
      subject: '',
      fromEmail: 'design@seastargo.com',
      to: [],
      cc: [],
      bcc: [],
      body: '확인 후 회신 부탁드립니다.',
      attachmentsCount: 0,
    },
    'composing'
  );

  assert.equal(report.status, 'block');
  assert.ok(report.findings.some((f) => f.code === 'missing_recipient' && f.severity === 'critical'));
  assert.ok(report.findings.some((f) => f.code === 'missing_subject' && f.severity === 'warning'));
});

test('verifyOutlookItemSafety warns when sending sensitive patterns to external recipients', () => {
  const report = verifyOutlookItemSafety(
    {
      subject: '견적 전달 건',
      fromEmail: 'sales@seastargo.com',
      to: ['partner@gmail.com'],
      body: '담당자 연락처 010-1234-5678 / 견적금액 25,000,000원 기준 검토 바랍니다.',
      attachmentsCount: 1,
      attachments: ['quote.pdf'],
    },
    'composing'
  );

  assert.equal(report.status, 'caution');
  const leak = report.findings.find((f) => f.code === 'external_sensitive_leak');
  assert.ok(leak);
  assert.equal(leak?.severity, 'warning');
});

test('verifyOutlookItemSafety marks clean email with attachments safe and extracts drawing & deadlines', () => {
  const report = verifyOutlookItemSafety(
    {
      subject: 'RE: H042-1000-E1A-0001 케이블 도면 송부',
      fromEmail: 'design@seastargo.com',
      to: ['eng@seastargo.com'],
      body: 'H042-1000-E1A-0001 Rev.A 도면 첨부드립니다. 2026-12-05까지 코멘트 부탁드립니다.',
      attachmentsCount: 1,
      attachments: ['H042-1000-E1A-0001.pdf'],
    },
    'composing'
  );

  assert.equal(report.status, 'safe');
  assert.equal(report.score, 100);
  assert.ok(report.drawings.some((d) => d.includes('H042')));
  assert.ok(report.deadlines.includes('2026-12-05'));
});
