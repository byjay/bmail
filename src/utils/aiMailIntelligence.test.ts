import test from 'node:test';
import assert from 'node:assert/strict';
import type { MailItem } from '../types/mail';
import {
  answerInboxQuestion,
  buildMorningBrief,
  classifyEmailIntent,
  coachEmailDraft,
  DEFAULT_AUTO_LABEL_RULES,
  expandBulletsToDraft,
  generateInstantReplyOptions,
} from './aiMailIntelligence';

function makeMail(overrides: Partial<MailItem>): MailItem {
  return {
    id: 'm-1',
    folder: 'inbox',
    subject: 'Test Subject',
    fromName: 'Test Sender',
    fromEmail: 'sender@example.test',
    to: ['me@example.test'],
    date: '2026-11-18',
    timestamp: '2026-11-18T09:00:00Z',
    timeLabel: '09:00',
    snippet: 'Test body',
    body: 'Test body',
    isRead: false,
    isFlagged: false,
    hasAttachment: false,
    importance: 'normal',
    source: 'pst-import',
    ...overrides,
  };
}

test('classifyEmailIntent classifies English sales quote, drawing review, and newsletter accurately', () => {
  const salesMail = makeMail({
    id: 'm-sales',
    subject: 'URGENT RFQ Quotation for Vessel Cable Package USD 48,500',
    fromName: 'Alex Buyer',
    fromEmail: 'buyer@client.example.test',
    body: 'Dear Team,\nPlease confirm your quotation USD 48,500 and delivery schedule by Nov 25, 2026. Best regards.',
  });
  const salesProfile = classifyEmailIntent(salesMail, DEFAULT_AUTO_LABEL_RULES);
  assert.equal(salesProfile.isForeignLanguage, true);
  assert.equal(salesProfile.detectedLanguage, 'en');
  assert.equal(salesProfile.urgency, 'urgent');
  assert.ok(salesProfile.extractedAmounts.some((a) => a.includes('48,500')));
  assert.ok(salesProfile.matchedLabels.some((l) => l.includes('영업')));
  assert.ok(salesProfile.matchedLabels.some((l) => l.includes('외국어')));

  const drawingMail = makeMail({
    id: 'm-dwg',
    subject: '[PRJ-101] DWG-E1001 개정 도면 검토 요청',
    body: '안녕하세요. 첨부한 DWG-E1001 Rev.2 도면 검토 부탁드립니다. 11월 22일까지 의견 공유 요청드립니다.',
  });
  const drawingProfile = classifyEmailIntent(drawingMail, DEFAULT_AUTO_LABEL_RULES);
  assert.equal(drawingProfile.isForeignLanguage, false);
  assert.ok(drawingProfile.extractedDrawingNos.includes('DWG-E1001'));
  assert.ok(drawingProfile.extractedDeadlines.length >= 1);
  assert.ok(drawingProfile.matchedLabels.some((l) => l.includes('도면')));

  const promoMail = makeMail({
    id: 'm-promo',
    subject: '[광고] 뉴스레터 주간 프로모션 안내',
    fromEmail: 'noreply@newsletter.example.test',
    body: '수신거부 unsubscribe 링크 안내입니다.',
  });
  const promoProfile = classifyEmailIntent(promoMail, DEFAULT_AUTO_LABEL_RULES);
  assert.equal(promoProfile.intent, 'fyi_newsletter');
  assert.equal(promoProfile.urgency, 'low');
});

test('buildMorningBrief separates Action, Decision, Waiting-On SLA, and Auto-Sweep FYI cleanly', () => {
  const emails: MailItem[] = [
    makeMail({
      id: 'in-urgent',
      folder: 'inbox',
      subject: '긴급 승인 요청 - 출도 결재 건',
      body: '금일 오후 5시까지 승인 회신 부탁드립니다.',
      timestamp: '2026-11-19T08:00:00Z',
    }),
    makeMail({
      id: 'in-quote',
      folder: 'inbox',
      subject: '견적서 송부 - USD 12,000',
      body: '단가 및 발주 조건 검토 부탁드립니다.',
      timestamp: '2026-11-19T07:00:00Z',
    }),
    makeMail({
      id: 'sent-unanswered',
      folder: 'sent',
      subject: 'DWG-E1002 납기 확인 요청',
      to: ['vendor@partner.example.test'],
      date: '2026-11-15',
      timestamp: '2026-11-15T09:00:00Z',
      body: '납기 일정 확인 회신 부탁드립니다.',
    }),
    makeMail({
      id: 'sent-answered',
      folder: 'sent',
      subject: '주간 보고 완료 확인',
      to: ['team@example.test'],
      date: '2026-11-16',
      timestamp: '2026-11-16T09:00:00Z',
      body: '확인 부탁드립니다.',
    }),
    makeMail({
      id: 'in-reply-answered',
      folder: 'inbox',
      subject: 'Re: 주간 보고 완료 확인',
      date: '2026-11-17',
      timestamp: '2026-11-17T10:00:00Z',
      body: '확인 완료했습니다.',
    }),
    makeMail({
      id: 'in-promo',
      folder: 'inbox',
      subject: '[광고] 세미나 안내',
      fromEmail: 'noreply@promo.example.test',
      body: 'unsubscribe',
    }),
  ];

  const brief = buildMorningBrief(emails, DEFAULT_AUTO_LABEL_RULES, '2026-11-20T09:00:00Z');
  assert.equal(brief.totalAnalyzed, 6);
  assert.ok(brief.actionRequired.some((x) => x.mail.id === 'in-urgent'));
  assert.ok(brief.decisionNeeded.some((x) => x.mail.id === 'in-quote'));
  assert.equal(brief.waitingOnOthers.length, 1);
  assert.equal(brief.waitingOnOthers[0].emailId, 'sent-unanswered');
  assert.equal(brief.waitingOnOthers[0].daysWaiting, 5);
  assert.ok(brief.fyiSweepable.some((x) => x.mail.id === 'in-promo'));
  assert.ok(brief.commitments.length >= 1);

  // Defensive empty check
  const emptyBrief = buildMorningBrief([]);
  assert.equal(emptyBrief.totalAnalyzed, 0);
  assert.equal(emptyBrief.actionRequired.length, 0);
});

test('generateInstantReplyOptions generates 3 language-matched full-sentence contextual drafts', () => {
  const enMail = makeMail({
    subject: 'RFQ for Control Panel USD 25,000',
    fromName: 'Sarah Lee',
    body: 'Dear Team,\nPlease confirm your quotation and delivery lead time by Nov 28, 2026.\nBest regards,',
  });
  const enReplies = generateInstantReplyOptions(enMail, 'Bmail Engineering');
  assert.equal(enReplies.length, 3);
  assert.ok(enReplies[0].body.includes('Dear Sarah Lee'));
  assert.ok(enReplies[0].body.includes('Bmail Engineering'));
  assert.equal(enReplies[0].subject, 'Re: RFQ for Control Panel USD 25,000');

  const koMail = makeMail({
    subject: 'Re: DWG-E1001 도면 검토 요청',
    fromName: '담당자 A',
    body: '안녕하세요. 도면 검토 후 회신 부탁드립니다.',
  });
  const koReplies = generateInstantReplyOptions(koMail);
  assert.equal(koReplies.length, 3);
  assert.ok(koReplies[0].body.includes('안녕하세요, 담당자 A님'));
  assert.equal(koReplies[0].subject, 'Re: DWG-E1001 도면 검토 요청'); // No double Re: Re:
});

test('answerInboxQuestion retrieves grounded answers with clickable citations', () => {
  const emails: MailItem[] = [
    makeMail({
      id: 'q-1',
      subject: 'RFQ Cable Tray Quotation USD 32,000',
      fromName: 'Vendor Alpha',
      body: 'Attached please find quotation USD 32,000.',
      hasAttachment: true,
      attachments: [{ id: 'att-1', filename: 'quote_32000.pdf', kind: 'pdf', mimeType: 'application/pdf', sizeBytes: 1024 }],
    }),
    makeMail({
      id: 'q-2',
      subject: '사내 공지 - 주차장 보수',
      fromName: '총무팀',
      body: '주차장 보수 일정 안내드립니다.',
    }),
  ];

  const res = answerInboxQuestion('견적 금액 얼마인지 알려줘', emails);
  assert.equal(res.matchedCount >= 1, true);
  assert.equal(res.citations[0].emailId, 'q-1');
  assert.ok(res.bulletFindings[0].includes('USD 32,000'));

  const emptyRes = answerInboxQuestion('테스트 질문', []);
  assert.equal(emptyRes.matchedCount, 0);
  assert.equal(emptyRes.citations.length, 0);
});

test('expandBulletsToDraft and coachEmailDraft detect missing attachment, harsh tone, and sensitive data', () => {
  const formal = expandBulletsToDraft('- DWG-E1001 검토 완료\n- 내일 오전 중 출도 예정', {
    tone: 'formal',
    recipientName: '김팀장',
    signature: '설계팀 드림',
  });
  assert.ok(formal.body.includes('안녕하세요, 김팀장님'));
  assert.ok(formal.body.includes('1. DWG-E1001 검토 완료'));
  assert.ok(formal.body.includes('설계팀 드림'));

  const en = expandBulletsToDraft('- Confirmed price\n- Shipping next week', {
    tone: 'english',
    recipientName: 'John',
  });
  assert.ok(en.body.includes('Dear John,'));

  // Coaching catches forgotten attachment, harsh tone, sensitive phone, missing recipient
  const badDraft = coachEmailDraft({
    to: '',
    subject: '',
    body: '왜 안 보내셨나요? 첨부 파일 확인하시고 010-1234-5678로 당장 연락하세요.',
    hasAttachment: false,
  });
  assert.equal(badDraft.readyToSend, false);
  const codes = badDraft.warnings.map((w) => w.code);
  assert.ok(codes.includes('missing_recipient'));
  assert.ok(codes.includes('missing_subject'));
  assert.ok(codes.includes('missing_attachment'));
  assert.ok(codes.includes('sensitive_data'));
  assert.ok(codes.includes('harsh_tone'));

  // Clean draft passes coaching
  const goodDraft = coachEmailDraft({
    to: 'partner@example.test',
    subject: 'DWG-E1001 검토 결과 회신',
    body: '안녕하세요. 첨부드린 자료 확인 부탁드립니다. 감사합니다.',
    hasAttachment: true,
  });
  assert.equal(goodDraft.readyToSend, true);
  assert.equal(goodDraft.warnings.length, 0);
});
