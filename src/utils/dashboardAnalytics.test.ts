import test from 'node:test';
import assert from 'node:assert/strict';
import type { AnalysisResult, MailItem } from '../types/mail';
import { buildWorkEfficiencyMetrics } from './dashboardAnalytics';

function makeMail(partial: Partial<MailItem> & { id: string }): MailItem {
  return {
    id: partial.id,
    folder: partial.folder ?? 'inbox',
    subject: partial.subject ?? '테스트 메일',
    fromName: partial.fromName ?? '홍길동',
    fromEmail: partial.fromEmail ?? 'hong@company.co.kr',
    to: partial.to ?? ['me@company.co.kr'],
    date: partial.date ?? '2026-11-20',
    timestamp: partial.timestamp ?? '2026-11-20T09:00:00Z',
    timeLabel: partial.timeLabel ?? '09:00',
    snippet: partial.snippet ?? '',
    body: partial.body ?? '',
    isRead: partial.isRead ?? true,
    isFlagged: partial.isFlagged ?? false,
    source: partial.source ?? 'pst-import',
    hasAttachment: partial.hasAttachment,
    attachments: partial.attachments,
    category: partial.category,
    folderName: partial.folderName,
  };
}

test('buildWorkEfficiencyMetrics returns zeroed metrics and empty lists when emails is empty', () => {
  const result = buildWorkEfficiencyMetrics({ emails: [], projects: [] });
  assert.equal(result.totalEmails, 0);
  assert.equal(result.overallEfficiencyPercent, 0);
  assert.equal(result.readRatePercent, 0);
  assert.equal(result.responseRatePercent, 0);
  assert.equal(result.followUpClosurePercent, 0);
  assert.deepEqual(result.projectProgress, []);
  assert.deepEqual(result.dailyTrend, []);
});

test('buildWorkEfficiencyMetrics computes read rate, response ratio, project progress, and chronological trend', () => {
  const emails: MailItem[] = [
    makeMail({
      id: 'm1',
      folder: 'inbox',
      subject: '[H1001] 선체 도면 검토 요청',
      body: 'H1001 도면 검토 부탁드립니다.',
      date: '2026-11-18',
      isRead: true,
      isFlagged: false,
    }),
    makeMail({
      id: 'm2',
      folder: 'sent',
      subject: 'RE: [H1001] 선체 도면 검토 완료 회신',
      body: 'H1001 검토 완료하여 회신합니다.',
      date: '2026-11-18',
      isRead: true,
    }),
    makeMail({
      id: 'm3',
      folder: 'inbox',
      subject: '[P2002] 배관 자재 승인 마감 임박',
      body: '긴급 승인 요청',
      date: '2026-11-19',
      isRead: false,
      isFlagged: true,
    }),
  ];

  const projects = [
    {
      id: 'p1',
      code: 'H1001',
      name: '1호 프로젝트 선체',
      keywords: ['H1001'],
      color: '#0F6CBD',
      createdAt: '2026-11-01T00:00:00Z',
      isFavorite: true,
    },
    {
      id: 'p2',
      code: 'P2002',
      name: '2호 프로젝트 배관',
      keywords: ['P2002'],
      color: '#0B7A75',
      createdAt: '2026-11-01T00:00:00Z',
      isFavorite: false,
    },
  ];

  const metrics = buildWorkEfficiencyMetrics({ emails, projects });

  assert.equal(metrics.totalEmails, 3);
  assert.equal(metrics.totalInboxCount, 2);
  assert.equal(metrics.readInboxCount, 1);
  assert.equal(metrics.readRatePercent, 50);
  assert.equal(metrics.sentCount, 1);
  assert.equal(metrics.responseRatePercent, 50);
  assert.equal(metrics.flaggedPendingCount, 1);
  assert.ok(metrics.overallEfficiencyPercent >= 0 && metrics.overallEfficiencyPercent <= 100);

  // Project progress sorted by total emails descending
  assert.equal(metrics.projectProgress.length, 2);
  assert.equal(metrics.projectProgress[0].projectCode, 'H1001');
  assert.equal(metrics.projectProgress[0].total, 2);
  assert.equal(metrics.projectProgress[1].projectCode, 'P2002');
  assert.equal(metrics.projectProgress[1].status, '후속 조치 필요');

  // Chronological trend ascending by date
  assert.equal(metrics.dailyTrend.length, 2);
  assert.equal(metrics.dailyTrend[0].date, '2026-11-18');
  assert.equal(metrics.dailyTrend[1].date, '2026-11-19');
});

test('buildWorkEfficiencyMetrics respects AI AnalysisResult projectSummaries status weighting', () => {
  const emails: MailItem[] = [
    makeMail({
      id: 'a1',
      folder: 'inbox',
      subject: '[H3000] 최종 시운전 보고',
      date: '2026-11-20',
      isRead: true,
    }),
  ];

  const aiResult: AnalysisResult = {
    id: 'res-1',
    createdAt: '2026-11-20T10:00:00Z',
    providerUsed: 'builtin',
    providerModel: 'gemini-4-hydrogen',
    periodLabel: '2026-11-20',
    scopeLabel: '전체',
    analyzedCount: 1,
    receivedCount: 1,
    sentCount: 0,
    overview: '시운전 완료',
    workPerformed: [],
    incomingHighlights: [],
    projectSummaries: [
      {
        projectCode: 'H3000',
        projectName: '3호 시운전',
        status: '완료',
        myActions: ['최종 보고 승인'],
        keyDecisions: [],
        attachmentsMentioned: [],
        nextAction: '',
      },
    ],
    workLog: {
      title: '일지',
      periodLabel: '2026-11-20',
      executiveSummary: '',
      completedTasks: [],
      inProgressTasks: [],
      keyCommunications: [],
      nextStepPlan: [],
      issuesAndRisks: [],
    },
    blogPost: {
      title: '블로그',
      subtitle: '',
      readingTimeMinutes: 1,
      tags: [],
      intro: '',
      sections: [],
      conclusion: '',
    },
  };

  const metrics = buildWorkEfficiencyMetrics({ emails, analysisResult: aiResult });
  assert.equal(metrics.projectProgress[0].status, '완료');
  assert.ok(metrics.projectProgress[0].progressPercent >= 88);
});

test('buildWorkEfficiencyMetrics computes thread response speed hours and monthly MoM delta accurately', () => {
  const emails: MailItem[] = [
    makeMail({
      id: 'oct-1',
      folder: 'inbox',
      subject: '[H1001] 10월 선체 검토 요청',
      date: '2026-10-15',
      timestamp: '2026-10-15T08:00:00Z',
      isRead: true,
    }),
    makeMail({
      id: 'oct-2',
      folder: 'sent',
      subject: 'RE: [H1001] 10월 선체 검토 요청',
      date: '2026-10-15',
      timestamp: '2026-10-15T10:00:00Z', // 2.0h response time
      isRead: true,
    }),
    makeMail({
      id: 'nov-1',
      folder: 'inbox',
      subject: '[H1001] 11월 도면 승인 요청',
      date: '2026-11-10',
      timestamp: '2026-11-10T09:00:00Z',
      isRead: true,
    }),
    makeMail({
      id: 'nov-2',
      folder: 'sent',
      subject: 'RE: [H1001] 11월 도면 승인 요청',
      date: '2026-11-10',
      timestamp: '2026-11-10T13:00:00Z', // 4.0h response time
      isRead: true,
    }),
    makeMail({
      id: 'nov-3',
      folder: 'inbox',
      subject: '[P2002] 배관 자재 납기 확인',
      date: '2026-11-12',
      timestamp: '2026-11-12T10:00:00Z',
      isRead: false,
    }),
    makeMail({
      id: 'nov-4',
      folder: 'inbox',
      subject: '[P2002] 배관 자재 도면 추가',
      date: '2026-11-14',
      timestamp: '2026-11-14T11:00:00Z',
      isRead: true,
    }),
  ];

  const projects = [
    {
      id: 'p1',
      code: 'H1001',
      name: '1호 프로젝트 선체',
      keywords: ['H1001'],
      color: '#0F6CBD',
      createdAt: '2026-10-01T00:00:00Z',
      isFavorite: true,
    },
    {
      id: 'p2',
      code: 'P2002',
      name: '2호 프로젝트 배관',
      keywords: ['P2002'],
      color: '#0B7A75',
      createdAt: '2026-10-01T00:00:00Z',
      isFavorite: false,
    },
  ];

  const metrics = buildWorkEfficiencyMetrics({ emails, projects });
  // Monthly trend: Oct = 2 mails, Nov = 4 mails => +100% MoM
  assert.equal(metrics.monthlyTrend.currentMonthLabel, '2026-11');
  assert.equal(metrics.monthlyTrend.currentMonthTotal, 4);
  assert.equal(metrics.monthlyTrend.previousMonthLabel, '2026-10');
  assert.equal(metrics.monthlyTrend.previousMonthTotal, 2);
  assert.equal(metrics.monthlyTrend.momDeltaPercent, 100);
  // Two reply pairs: 2.0h and 4.0h => avg 3.0h, fastest 2.0h
  assert.equal(metrics.monthlyTrend.measuredReplyPairs, 2);
  assert.equal(metrics.monthlyTrend.avgResponseHours, 3);
  assert.equal(metrics.monthlyTrend.fastestResponseHours, 2);

  // Project comparison: H1001 has 4 mails (67%) & avgResponseHours = 3
  const h1001 = metrics.projectResponseComparison.find((p) => p.projectCode === 'H1001');
  assert.ok(h1001);
  assert.equal(h1001.mailCount, 4);
  assert.equal(h1001.avgResponseHours, 3);
  assert.equal(h1001.pairedThreads, 2);
  assert.ok(h1001.responseSpeedScore >= 85);
});

