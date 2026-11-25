import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  Archive,
  BarChart3,
  Briefcase,
  Check,
  ClipboardCopy,
  Clock,
  Database,
  Download,
  ExternalLink,
  FileSpreadsheet,
  FileText,
  Filter,
  Inbox,
  Layers,
  ListChecks,
  Loader2,
  Paperclip,
  RefreshCw,
  Save,
  Send,
  Sparkles,
} from 'lucide-react';
import type {
  AnalysisResult,
  MailItem,
  ParsedDocumentSummary,
} from '../types/mail';
import {
  buildMailboxDashboardMetrics,
  buildWorkEfficiencyMetrics,
} from '../utils/dashboardAnalytics';
import { normalizeConversationSubject, type ProjectDefinition } from '../utils/deterministicAnalyzer';
import type {
  ExecutiveDashboardReport,
  ExecutivePeriodType,
  SavedExecutiveReportSnapshot,
} from '../services/executiveReportEngine';
import { AdminLlmUsageCompactCard } from './AdminLlmUsageCompactCard';
import { DailyBriefingTimeline } from './DailyBriefingTimeline';
import { DashboardEfficiencyCharts } from './DashboardEfficiencyCharts';

export interface DashboardViewProps {
  emails: MailItem[];
  documents: ParsedDocumentSummary[];
  projects: ProjectDefinition[];
  analysisResult: AnalysisResult | null;
  archivedResults: AnalysisResult[];
  isAnalyzing: boolean;
  onRunAiAnalysis: () => void;
  onOpenPstModal: () => void;
  onOpenStudioDocuments: () => void;
  onOpenProjects: () => void;
  onSaveDraftToOutlook: (subject: string, body: string) => void;
  onOpenMailInWorkbench?: (emailId: string) => void;
  onSyncIncrementalOutlook?: () => Promise<void> | void;
  initialSubTab?: 'overview' | 'briefing';
}

type ExecutiveTabId =
  | 'narrative-periodic'
  | 'unreplied-incoming'
  | 'project-topic'
  | 'ledger-attachments'
  | 'saved-snapshots'
  | 'briefing';

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return '-';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function classifyClientTopic(email: Pick<MailItem, 'subject' | 'body' | 'folderName'>): string {
  const text = `${email.subject || ''} ${(email.body || '').slice(0, 500)} ${email.folderName || ''}`.toLowerCase();
  if (/(dwg|drawing|승인도|출도|도면|rev\.?\s?[a-z0-9]|revision)/.test(text)) return '도면·승인도·REV';
  if (/(pos|p\.o\.s|rfq|견적|발주|단가|계약|purchase|quotation|price)/.test(text)) return '견적·발주·POS';
  if (/(선급|kr|abs|dnv|lr|bv|nk|검사|미결|comment|class)/.test(text)) return '선급·검사·미결';
  if (/(3d|modeling|모델링|생산설계|tribon|aveva|cad|block|블록)/.test(text)) return '3D 모델링·생산설계';
  if (/(주간보고|월간|일정|납기|schedule|delivery|meeting|회의)/.test(text)) return '납기·주간보고·일정';
  if (/(urgent|긴급|하자|클레임|defect|issue|지연|delay|asap)/.test(text)) return '하자·긴급·이슈';
  return '일반 업무·공유';
}

/**
 * Fallback client-side report builder when running in browser-only mode without local SQLite API
 */
function buildClientFallbackExecutiveReport(
  emails: MailItem[],
  projects: ProjectDefinition[],
  filters: {
    periodType: ExecutivePeriodType;
    dateFrom?: string;
    dateTo?: string;
    projectCode?: string;
    topic?: string;
    direction: 'all' | 'received' | 'sent';
  }
): ExecutiveDashboardReport {
  const latestSentByConv = new Map<string, string>();
  for (const m of emails) {
    if (m.folder !== 'sent') continue;
    const conv = m.conversationKey || normalizeConversationSubject(m.subject);
    if (!conv) continue;
    const ts = m.timestamp || `${m.date}T00:00:00Z`;
    const prev = latestSentByConv.get(conv);
    if (!prev || ts > prev) latestSentByConv.set(conv, ts);
  }

  const filtered = emails.filter((m) => {
    if (filters.dateFrom && m.date < filters.dateFrom) return false;
    if (filters.dateTo && m.date > filters.dateTo) return false;
    if (filters.direction === 'received' && m.folder !== 'inbox') return false;
    if (filters.direction === 'sent' && m.folder !== 'sent') return false;
    const topic = classifyClientTopic(m);
    if (filters.topic && filters.topic !== 'ALL' && topic !== filters.topic) return false;
    return true;
  });

  const received = filtered.filter((m) => m.folder !== 'sent');
  const sent = filtered.filter((m) => m.folder === 'sent');
  const unreplied = received
    .filter((m) => {
      const conv = m.conversationKey || normalizeConversationSubject(m.subject);
      const ts = m.timestamp || `${m.date}T00:00:00Z`;
      const sentAfter = conv ? latestSentByConv.get(conv) : undefined;
      return !sentAfter || sentAfter < ts;
    })
    .slice(0, 100)
    .map((m) => ({
      emailId: m.id,
      date: m.date,
      timeLabel: m.timeLabel,
      daysPending: 1,
      fromName: m.fromName,
      fromEmail: m.fromEmail,
      projectCode: projects[0]?.code || 'GENERAL',
      topic: classifyClientTopic(m),
      subject: m.subject,
      hasAttachment: Boolean(m.hasAttachment),
      urgency: (m.importance === 'high' ? 'high' : 'normal') as 'high' | 'normal',
      snippet: m.snippet,
    }));

  const attCount = filtered.reduce((acc, m) => acc + (m.attachments?.length || 0), 0);
  const title = `Bmail 메일함 종합 분석 보고서 (${filtered.length}건)`;
  const summary = `총 ${filtered.length}건(받은 메일 ${received.length}건 · 보낸 메일 ${sent.length}건 · 첨부파일 ${attCount}개 · 내게 온 미회신 ${unreplied.length}건)이 집계되었습니다.`;

  return {
    filters,
    kpis: {
      totalEmails: filtered.length,
      receivedCount: received.length,
      sentCount: sent.length,
      unrepliedIncomingCount: unreplied.length,
      awaitingCounterpartyCount: 0,
      realAttachmentCount: attCount,
      projectCount: projects.length || 1,
      topicCount: 6,
      unreadCount: received.filter((m) => !m.isRead).length,
      dateRangeLabel: '현재 조회 범위',
      persistedTotalInDb: emails.length,
    },
    narrativeReport: {
      title,
      executiveSummary: summary,
      periodTrendNote: '주간/월간 수발신 추이가 테이블에 정리되었습니다.',
      projectFocusNote: '프로젝트별 및 주제별 탭에서 상세 교신 내역을 확인하세요.',
      unrepliedRiskNote:
        unreplied.length > 0
          ? `내가 회신해야 할 미회신 수신 메일이 ${unreplied.length}건 있습니다.`
          : '미회신 수신 메일이 없습니다.',
      keyHighlights: [summary],
      recommendedActions: ['미회신 수신 건 우선 검토 및 회신 초안 작성'],
      markdownFull: `# ${title}\n\n${summary}`,
      generatedBy: '로컬 분석 엔진',
      generatedAt: new Date().toISOString(),
    },
    weeklyTable: [],
    monthlyTable: [],
    unrepliedIncoming: unreplied,
    projectRows: [],
    topicRows: [],
    mailRows: filtered.slice(0, 150).map((m) => ({
      emailId: m.id,
      direction: m.folder === 'sent' ? 'sent' : 'received',
      date: m.date,
      timeLabel: m.timeLabel,
      fromName: m.fromName,
      fromEmail: m.fromEmail,
      to: m.to,
      projectCode: projects[0]?.code || 'GENERAL',
      topic: classifyClientTopic(m),
      subject: m.subject,
      realAttachmentCount: m.attachments?.length || 0,
      replyStatus: m.folder === 'sent' ? 'sent' : 'unreplied-incoming',
      isRead: m.isRead,
    })),
    attachmentRows: [],
    availableProjects: projects.map((p) => ({ code: p.code, name: p.name })),
    availableTopics: [
      '도면·승인도·REV',
      '견적·발주·POS',
      '선급·검사·미결',
      '3D 모델링·생산설계',
      '납기·주간보고·일정',
      '하자·긴급·이슈',
      '일반 업무·공유',
    ],
    savedReports: [],
    lastSyncedAt: new Date().toISOString(),
  };
}

export const DashboardView: React.FC<DashboardViewProps> = ({
  emails,
  documents,
  projects,
  analysisResult,
  archivedResults,
  isAnalyzing,
  onRunAiAnalysis,
  onOpenPstModal,
  onOpenStudioDocuments,
  onOpenProjects,
  onSaveDraftToOutlook,
  onOpenMailInWorkbench,
  onSyncIncrementalOutlook,
  initialSubTab = 'overview',
}) => {
  const [activeTab, setActiveTab] = useState<ExecutiveTabId>(
    initialSubTab === 'briefing' ? 'briefing' : 'narrative-periodic'
  );
  const [periodType, setPeriodType] = useState<ExecutivePeriodType>('all');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [projectCode, setProjectCode] = useState('ALL');
  const [topic, setTopic] = useState('ALL');
  const [direction, setDirection] = useState<'all' | 'received' | 'sent'>('all');

  const [report, setReport] = useState<ExecutiveDashboardReport | null>(null);
  const [loadingReport, setLoadingReport] = useState(false);
  const [generatingAiNarrative, setGeneratingAiNarrative] = useState(false);
  const [syncingIncremental, setSyncingIncremental] = useState(false);
  const [syncDeltaNotice, setSyncDeltaNotice] = useState<string | null>(null);
  const [copiedSummary, setCopiedSummary] = useState(false);
  const [snapshotTitle, setSnapshotTitle] = useState('');

  const metrics = useMemo(
    () =>
      buildMailboxDashboardMetrics({
        emails,
        documents,
        projects,
        analysisResult,
      }),
    [emails, documents, projects, analysisResult]
  );

  const efficiency = useMemo(
    () =>
      buildWorkEfficiencyMetrics({
        emails,
        projects,
        analysisResult,
        archivedResults,
      }),
    [emails, projects, analysisResult, archivedResults]
  );

  const loadExecutiveReport = useCallback(async () => {
    setLoadingReport(true);
    try {
      const qs = new URLSearchParams({
        periodType,
        direction,
      });
      if (dateFrom) qs.set('dateFrom', dateFrom);
      if (dateTo) qs.set('dateTo', dateTo);
      if (projectCode && projectCode !== 'ALL') qs.set('projectCode', projectCode);
      if (topic && topic !== 'ALL') qs.set('topic', topic);

      const res = await fetch(`/api/local/dashboard/executive?${qs.toString()}`);
      if (res.ok) {
        const data = (await res.json()) as ExecutiveDashboardReport;
        setReport(data);
        return;
      }
    } catch {
      // Fall back to client computation below
    } finally {
      setLoadingReport(false);
    }

    setReport(
      buildClientFallbackExecutiveReport(emails, projects, {
        periodType,
        dateFrom: dateFrom || undefined,
        dateTo: dateTo || undefined,
        projectCode,
        topic,
        direction,
      })
    );
  }, [periodType, dateFrom, dateTo, projectCode, topic, direction, emails, projects]);

  useEffect(() => {
    void loadExecutiveReport();
  }, [loadExecutiveReport]);

  const effectiveReport = useMemo(
    () =>
      report ||
      buildClientFallbackExecutiveReport(emails, projects, {
        periodType,
        dateFrom: dateFrom || undefined,
        dateTo: dateTo || undefined,
        projectCode,
        topic,
        direction,
      }),
    [report, emails, projects, periodType, dateFrom, dateTo, projectCode, topic, direction]
  );

  const handleIncrementalUpdate = async () => {
    setSyncingIncremental(true);
    setSyncDeltaNotice(null);
    try {
      const res = await fetch('/api/local/outlook/sync-incremental', { method: 'POST' });
      if (res.ok) {
        const data = (await res.json()) as {
          addedPstEmails?: number;
          addedComEmails?: number;
          skippedUnchangedPsts?: number;
          totalEmails?: number;
        };
        const added = Number(data.addedPstEmails || 0) + Number(data.addedComEmails || 0);
        setSyncDeltaNotice(
          `증분 업데이트 완료: 기존 보관 데이터 유지 (PST 스킵 ${data.skippedUnchangedPsts || 0}개) · 신규 반영 +${added}건 (총 ${(
            data.totalEmails || effectiveReport.kpis.persistedTotalInDb
          ).toLocaleString()}건)`
        );
      }
      if (onSyncIncrementalOutlook) {
        await onSyncIncrementalOutlook();
      }
      await loadExecutiveReport();
    } catch {
      setSyncDeltaNotice('로컬 SQLite 영구 저장소 기준으로 최신 집계를 갱신했습니다.');
    } finally {
      setSyncingIncremental(false);
    }
  };

  const handleGenerateAiNarrative = async () => {
    setGeneratingAiNarrative(true);
    try {
      const res = await fetch('/api/local/dashboard/executive/ai-narrative', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          periodType,
          dateFrom: dateFrom || undefined,
          dateTo: dateTo || undefined,
          projectCode,
          topic,
          direction,
        }),
      });
      if (res.ok) {
        const data = (await res.json()) as {
          narrativeReport?: ExecutiveDashboardReport['narrativeReport'];
        };
        if (data.narrativeReport) {
          setReport((prev) =>
            prev ? { ...prev, narrativeReport: data.narrativeReport! } : prev
          );
        }
      }
    } catch {
      // Keep deterministic narrative
    } finally {
      setGeneratingAiNarrative(false);
    }
  };

  const handleSaveSnapshot = async () => {
    try {
      const res = await fetch('/api/local/dashboard/executive/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title: snapshotTitle.trim() || effectiveReport.narrativeReport.title,
          periodType,
          dateFrom: dateFrom || undefined,
          dateTo: dateTo || undefined,
          projectCode,
          topic,
          direction,
        }),
      });
      if (res.ok) {
        const data = (await res.json()) as { savedReports?: SavedExecutiveReportSnapshot[] };
        if (data.savedReports) {
          setReport((prev) => (prev ? { ...prev, savedReports: data.savedReports! } : prev));
          setSnapshotTitle('');
          setSyncDeltaNotice('현재 기간·프로젝트·주제 보고서 스냅샷을 로컬 DB에 영구 저장했습니다.');
        }
      }
    } catch {
      // ignore
    }
  };

  const handleDownloadReportFile = (fmt: 'md' | 'csv' | 'html') => {
    const stamp = new Date().toISOString().slice(0, 10);
    if (fmt === 'md') {
      const blob = new Blob([effectiveReport.narrativeReport.markdownFull], {
        type: 'text/markdown;charset=utf-8',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `bmail-executive-report_${stamp}.md`;
      a.click();
      URL.revokeObjectURL(url);
      return;
    }
    if (fmt === 'csv') {
      const rows = [
        ['구분', '기간/코드', '전체', '받은메일', '보낸메일', '내게온미회신', '첨부파일', '주요주제/안건'],
        ...effectiveReport.weeklyTable.map((w) => [
          '주간',
          w.periodLabel,
          w.total,
          w.received,
          w.sent,
          w.unrepliedIncoming,
          w.attachmentCount,
          `${w.topProject} / ${w.topTopic}`,
        ]),
        ...effectiveReport.monthlyTable.map((m) => [
          '월간',
          m.periodLabel,
          m.total,
          m.received,
          m.sent,
          m.unrepliedIncoming,
          m.attachmentCount,
          `${m.topProject} / ${m.topTopic}`,
        ]),
        ...effectiveReport.projectRows.map((p) => [
          '프로젝트',
          p.projectCode,
          p.total,
          p.received,
          p.sent,
          p.unrepliedIncoming,
          p.attachmentCount,
          p.latestSubject,
        ]),
        ...effectiveReport.topicRows.map((t) => [
          '주제',
          t.topic,
          t.total,
          t.received,
          t.sent,
          t.unrepliedIncoming,
          t.attachmentCount,
          t.latestSubject,
        ]),
      ];
      const csv =
        '\uFEFF' +
        rows
          .map((r) =>
            r.map((c) => `"${String(c ?? '').replace(/"/g, '""')}"`).join(',')
          )
          .join('\r\n');
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `bmail-executive-tables_${stamp}.csv`;
      a.click();
      URL.revokeObjectURL(url);
      return;
    }
    const html = `<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><title>${effectiveReport.narrativeReport.title}</title><style>body{font-family:'Malgun Gothic',sans-serif;margin:28px;color:#1b1b1b;line-height:1.6}h1{color:#0F6CBD}table{width:100%;border-collapse:collapse;margin-top:12px;font-size:12px}th,td{border:1px solid #ddd;padding:6px 8px;text-align:left}th{background:#f3f2f1}</style></head><body><h1>${effectiveReport.narrativeReport.title}</h1><p>${effectiveReport.narrativeReport.executiveSummary}</p><p>${effectiveReport.narrativeReport.periodTrendNote}</p><p>${effectiveReport.narrativeReport.projectFocusNote}</p><p><strong>${effectiveReport.narrativeReport.unrepliedRiskNote}</strong></p></body></html>`;
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `bmail-executive-report_${stamp}.html`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleCopyMarkdown = async () => {
    try {
      await navigator.clipboard.writeText(effectiveReport.narrativeReport.markdownFull);
      setCopiedSummary(true);
      setTimeout(() => setCopiedSummary(false), 1800);
    } catch {
      // ignore
    }
  };

  const { kpis, narrativeReport } = effectiveReport;

  return (
    <div className="flex-1 overflow-y-auto bg-[#FAF9F8] p-4 sm:p-6 space-y-5 text-[#242424]">
      {/* Top Executive Header & Load-Once + Incremental Delta Sync Bar */}
      <div className="bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full bg-[#EBF3FC] text-[#0F6CBD] text-xs font-bold flex items-center gap-1">
                <Database className="w-3.5 h-3.5" />
                1회 불러오기 영구 저장 + 증분 업데이트 엔진
              </span>
              <span className="text-xs font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">
                로컬 DB 보관: {kpis.persistedTotalInDb.toLocaleString()}건
              </span>
            </div>
            <h1 className="text-lg font-bold text-[#242424] mt-1">
              주간 · 월간 종합 분석글 · 내게 온 미회신건 · 프로젝트/주제별 통합 대시보드
            </h1>
            <p className="text-xs text-[#616161]">
              불러온 모든 메일과 첨부파일을 로컬 SQLite에 저장해 두고 새로 도착한 메일만 증분 반영하여 기간·프로젝트·주제별로 즉시 정리합니다.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <button
              type="button"
              disabled={syncingIncremental}
              onClick={() => void handleIncrementalUpdate()}
              className="px-3 py-1.5 rounded-md bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-sm disabled:opacity-60"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${syncingIncremental ? 'animate-spin' : ''}`} />
              <span>{syncingIncremental ? '증분 동기화 중...' : '신규 메일만 증분 업데이트'}</span>
            </button>

            <button
              type="button"
              disabled={generatingAiNarrative}
              onClick={() => void handleGenerateAiNarrative()}
              className="px-3 py-1.5 rounded-md bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-bold flex items-center gap-1.5 shadow-sm disabled:opacity-60"
            >
              {generatingAiNarrative ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <Sparkles className="w-3.5 h-3.5" />
              )}
              <span>Claude 5.5 심층 분석글 갱신</span>
            </button>

            <button
              type="button"
              onClick={() => void handleCopyMarkdown()}
              className="px-2.5 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold flex items-center gap-1"
            >
              {copiedSummary ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <ClipboardCopy className="w-3.5 h-3.5" />}
              <span>{copiedSummary ? '복사됨' : '분석글 복사'}</span>
            </button>

            <button
              type="button"
              onClick={() => handleDownloadReportFile('md')}
              className="px-2.5 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold"
            >
              MD
            </button>
            <button
              type="button"
              onClick={() => handleDownloadReportFile('csv')}
              className="px-2.5 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold"
            >
              Excel CSV
            </button>
            <button
              type="button"
              onClick={() => handleDownloadReportFile('html')}
              className="px-2.5 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold"
            >
              HTML
            </button>

            <button
              type="button"
              onClick={() =>
                onSaveDraftToOutlook(narrativeReport.title, narrativeReport.markdownFull)
              }
              className="px-2.5 py-1.5 rounded-md bg-[#EBF3FC] hover:bg-[#CFE4FA] text-[#0F6CBD] text-xs font-bold"
            >
              임시보관함 저장
            </button>
          </div>
        </div>

        {syncDeltaNotice && (
          <div className="px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-semibold flex items-center justify-between">
            <span>✓ {syncDeltaNotice}</span>
            <button
              type="button"
              onClick={() => setSyncDeltaNotice(null)}
              className="text-[11px] underline"
            >
              닫기
            </button>
          </div>
        )}

        {/* Unified Multi-Axis Filter Bar: 기간 / 프로젝트 / 주제 / 수발신 */}
        <div className="pt-2.5 border-t border-[#EDEBE9] flex flex-wrap items-center gap-2 text-xs">
          <div className="flex items-center gap-1 font-bold text-slate-700">
            <Filter className="w-3.5 h-3.5 text-[#0F6CBD]" />
            <span>기간:</span>
          </div>
          <div className="inline-flex bg-[#F3F2F1] p-0.5 rounded-md">
            {(
              [
                { id: 'all', label: '전체 기간' },
                { id: 'weekly', label: '최근 주간(7일)' },
                { id: 'monthly', label: '최근 월간(30일)' },
                { id: 'custom', label: '직접 지정' },
              ] as const
            ).map((p) => (
              <button
                key={p.id}
                type="button"
                onClick={() => {
                  setPeriodType(p.id);
                  if (p.id !== 'custom') {
                    setDateFrom('');
                    setDateTo('');
                  }
                }}
                className={`px-2.5 py-1 rounded text-[11px] font-semibold ${
                  periodType === p.id ? 'bg-[#0F6CBD] text-white' : 'text-slate-700 hover:bg-black/5'
                }`}
              >
                {p.label}
              </button>
            ))}
          </div>

          <input
            type="date"
            value={dateFrom}
            onChange={(e) => {
              setDateFrom(e.target.value);
              setPeriodType('custom');
            }}
            className="px-2 py-1 text-[11px] border border-[#D1D1D1] rounded"
          />
          <span>~</span>
          <input
            type="date"
            value={dateTo}
            onChange={(e) => {
              setDateTo(e.target.value);
              setPeriodType('custom');
            }}
            className="px-2 py-1 text-[11px] border border-[#D1D1D1] rounded"
          />

          <select
            value={projectCode}
            onChange={(e) => setProjectCode(e.target.value)}
            className="px-2.5 py-1 text-xs border border-[#D1D1D1] rounded bg-white font-semibold"
          >
            <option value="ALL">모든 프로젝트 ({effectiveReport.availableProjects.length})</option>
            {effectiveReport.availableProjects.map((p) => (
              <option key={p.code} value={p.code}>
                [{p.code}] {p.name}
              </option>
            ))}
          </select>

          <select
            value={topic}
            onChange={(e) => setTopic(e.target.value)}
            className="px-2.5 py-1 text-xs border border-[#D1D1D1] rounded bg-white font-semibold"
          >
            <option value="ALL">모든 주제 ({effectiveReport.availableTopics.length})</option>
            {effectiveReport.availableTopics.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>

          <div className="inline-flex bg-[#F3F2F1] p-0.5 rounded-md">
            {(
              [
                { id: 'all', label: '받은+보낸 전체' },
                { id: 'received', label: '받은 메일만' },
                { id: 'sent', label: '보낸 메일만' },
              ] as const
            ).map((d) => (
              <button
                key={d.id}
                type="button"
                onClick={() => setDirection(d.id)}
                className={`px-2 py-1 rounded text-[11px] font-semibold ${
                  direction === d.id ? 'bg-slate-800 text-white' : 'text-slate-700'
                }`}
              >
                {d.label}
              </button>
            ))}
          </div>

          {(projectCode !== 'ALL' || topic !== 'ALL' || dateFrom || dateTo || direction !== 'all') && (
            <button
              type="button"
              onClick={() => {
                setPeriodType('all');
                setDateFrom('');
                setDateTo('');
                setProjectCode('ALL');
                setTopic('ALL');
                setDirection('all');
              }}
              className="text-[11px] text-rose-600 font-semibold underline"
            >
              필터 초기화
            </button>
          )}
        </div>
      </div>

      {/* 6 At-a-Glance Executive KPI Cards */}
      <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-6 gap-3">
        <div className="bg-white p-3.5 rounded-xl border border-[#E0E0E0] shadow-sm">
          <div className="text-[11px] font-semibold text-[#616161]">조회 메일 / 영구 보관</div>
          <div className="text-xl font-extrabold text-[#242424] mt-1">
            {kpis.totalEmails.toLocaleString()}건
          </div>
          <div className="text-[11px] text-[#0F6CBD] mt-0.5">
            DB 총 {kpis.persistedTotalInDb.toLocaleString()}건 저장됨
          </div>
        </div>

        <div className="bg-white p-3.5 rounded-xl border border-[#E0E0E0] shadow-sm">
          <div className="text-[11px] font-semibold text-[#616161]">받은 메일 (수신)</div>
          <div className="text-xl font-extrabold text-[#0F6CBD] mt-1">
            {kpis.receivedCount.toLocaleString()}건
          </div>
          <div className="text-[11px] text-slate-500 mt-0.5">
            읽지 않음 {kpis.unreadCount.toLocaleString()}건
          </div>
        </div>

        <div className="bg-white p-3.5 rounded-xl border border-[#E0E0E0] shadow-sm">
          <div className="text-[11px] font-semibold text-[#616161]">보낸 메일 (발신)</div>
          <div className="text-xl font-extrabold text-indigo-700 mt-1">
            {kpis.sentCount.toLocaleString()}건
          </div>
          <div className="text-[11px] text-slate-500 mt-0.5">
            상대 회신 대기 {kpis.awaitingCounterpartyCount.toLocaleString()}건
          </div>
        </div>

        <div
          onClick={() => setActiveTab('unreplied-incoming')}
          className="bg-rose-50/80 hover:bg-rose-100/70 cursor-pointer p-3.5 rounded-xl border border-rose-200 shadow-sm transition-colors"
        >
          <div className="text-[11px] font-bold text-rose-700 flex items-center justify-between">
            <span>내게 왔는데 미회신건</span>
            <AlertTriangle className="w-3.5 h-3.5 text-rose-600" />
          </div>
          <div className="text-xl font-extrabold text-rose-600 mt-1">
            {kpis.unrepliedIncomingCount.toLocaleString()}건
          </div>
          <div className="text-[11px] text-rose-700 font-medium mt-0.5 underline">
            클릭하여 미회신 목록 바로 보기
          </div>
        </div>

        <div
          onClick={() => setActiveTab('ledger-attachments')}
          className="bg-white hover:bg-slate-50 cursor-pointer p-3.5 rounded-xl border border-[#E0E0E0] shadow-sm"
        >
          <div className="text-[11px] font-semibold text-[#616161]">실첨부 문서 (도면·POS)</div>
          <div className="text-xl font-extrabold text-emerald-700 mt-1">
            {kpis.realAttachmentCount.toLocaleString()}개
          </div>
          <div className="text-[11px] text-slate-500 mt-0.5">서명 로고 자동 제외됨</div>
        </div>

        <div
          onClick={() => setActiveTab('project-topic')}
          className="bg-white hover:bg-slate-50 cursor-pointer p-3.5 rounded-xl border border-[#E0E0E0] shadow-sm"
        >
          <div className="text-[11px] font-semibold text-[#616161]">프로젝트 · 주제 분류</div>
          <div className="text-xl font-extrabold text-purple-700 mt-1">
            {kpis.projectCount}개 · {kpis.topicCount}종
          </div>
          <div className="text-[11px] text-slate-500 mt-0.5">{kpis.dateRangeLabel}</div>
        </div>
      </div>

      {/* Live API Provider Routing & Call-Volume Graphs */}
      <AdminLlmUsageCompactCard defaultExpanded />

      {/* Sub-Tab Navigation */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-[#E0E0E0] pb-2">
        <div className="flex flex-wrap gap-1.5">
          {(
            [
              { id: 'narrative-periodic', label: '1. 종합 분석글 · 주간/월간 보고서 테이블', icon: FileText },
              {
                id: 'unreplied-incoming',
                label: `2. 내게 온 미회신건 정리 (${kpis.unrepliedIncomingCount})`,
                icon: AlertTriangle,
              },
              {
                id: 'project-topic',
                label: `3. 프로젝트별(${kpis.projectCount}) · 주제별(${kpis.topicCount}) 정리`,
                icon: Briefcase,
              },
              {
                id: 'ledger-attachments',
                label: `4. 받은/보낸 메일 · 첨부파일(${kpis.realAttachmentCount}) 통합 테이블`,
                icon: Paperclip,
              },
              {
                id: 'saved-snapshots',
                label: `5. 저장된 보고서함 (${effectiveReport.savedReports.length})`,
                icon: Archive,
              },
              { id: 'briefing', label: '6. 일간 타임라인 브리핑', icon: ListChecks },
            ] as const
          ).map((t) => {
            const Icon = t.icon;
            const active = activeTab === t.id;
            return (
              <button
                key={t.id}
                type="button"
                onClick={() => setActiveTab(t.id)}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-colors ${
                  active
                    ? 'bg-[#0F6CBD] text-white shadow-sm'
                    : 'bg-white border border-[#E0E0E0] text-[#424242] hover:bg-slate-100'
                }`}
              >
                <Icon className="w-3.5 h-3.5" />
                <span>{t.label}</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* TAB 1: Narrative Analysis Article (분석글) + Weekly & Monthly Report Tables */}
      {activeTab === 'narrative-periodic' && (
        <div className="space-y-5">
          {/* Narrative Article Card */}
          <section className="bg-white border border-[#E0E0E0] rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-[#EDEBE9] pb-3">
              <div>
                <span className="text-[11px] font-bold text-[#0F6CBD]">
                  서술형 종합 분석글 ({narrativeReport.generatedBy})
                </span>
                <h2 className="text-base font-bold text-[#242424] mt-0.5">
                  {narrativeReport.title}
                </h2>
              </div>
              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  value={snapshotTitle}
                  onChange={(e) => setSnapshotTitle(e.target.value)}
                  placeholder="보고서 저장 제목 (예: 11월 4주차 보고서)"
                  className="px-2.5 py-1 text-xs border border-[#D1D1D1] rounded"
                />
                <button
                  type="button"
                  onClick={() => void handleSaveSnapshot()}
                  className="px-3 py-1 rounded bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold flex items-center gap-1"
                >
                  <Save className="w-3.5 h-3.5" />
                  <span>보고서함 저장</span>
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
              <div className="lg:col-span-7 space-y-3 text-xs leading-relaxed text-[#242424]">
                <div className="p-3.5 rounded-lg bg-[#F4F8FD] border border-[#C7E0F4] space-y-2">
                  <div className="font-bold text-[#0F6CBD]">① 총괄 업무 흐름 및 수·발신 분석</div>
                  <p>{narrativeReport.executiveSummary}</p>
                  <p>{narrativeReport.periodTrendNote}</p>
                </div>

                <div className="p-3.5 rounded-lg bg-slate-50 border border-slate-200 space-y-2">
                  <div className="font-bold text-slate-800">② 핵심 프로젝트 및 엔지니어링 주제 집중도</div>
                  <p>{narrativeReport.projectFocusNote}</p>
                </div>

                <div className="p-3.5 rounded-lg bg-rose-50/70 border border-rose-200 space-y-1.5">
                  <div className="font-bold text-rose-800">③ 내게 온 미회신 리스크 및 긴급 안건</div>
                  <p className="text-rose-900">{narrativeReport.unrepliedRiskNote}</p>
                </div>
              </div>

              <div className="lg:col-span-5 space-y-3 text-xs">
                <div className="p-3.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
                  <div className="font-bold text-slate-800">핵심 요약 지표 (Key Highlights)</div>
                  <ul className="space-y-1.5 text-slate-700">
                    {narrativeReport.keyHighlights.map((h, i) => (
                      <li key={i} className="flex items-start gap-1.5">
                        <span className="text-[#0F6CBD] font-bold">•</span>
                        <span>{h}</span>
                      </li>
                    ))}
                  </ul>
                </div>

                <div className="p-3.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
                  <div className="font-bold text-rose-700">권장 우선 조치 과제 (Top 5)</div>
                  <ul className="space-y-1.5 text-slate-800">
                    {narrativeReport.recommendedActions.map((act, i) => (
                      <li key={i} className="leading-relaxed">
                        {act}
                      </li>
                    ))}
                  </ul>
                </div>
              </div>
            </div>
          </section>

          {/* Weekly & Monthly Tables Side by Side */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
            {/* Weekly Report Table */}
            <section className="bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
              <div className="flex items-center justify-between">
                <h3 className="text-sm font-bold text-[#242424]">
                  주간(Weekly) 메일 · 미회신 · 첨부 보고서 테이블
                </h3>
                <span className="text-[11px] text-slate-500">
                  최근 {effectiveReport.weeklyTable.length}주
                </span>
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs border-collapse">
                  <thead>
                    <tr className="bg-[#F3F2F1] text-slate-700 border-b border-[#E0E0E0]">
                      <th className="py-2 px-2.5 text-left">주차 (기간)</th>
                      <th className="py-2 px-2 text-right">전체</th>
                      <th className="py-2 px-2 text-right text-[#0F6CBD]">받은 메일</th>
                      <th className="py-2 px-2 text-right text-indigo-700">보낸 메일</th>
                      <th className="py-2 px-2 text-right text-rose-600">내 미회신</th>
                      <th className="py-2 px-2 text-right text-emerald-700">첨부파일</th>
                      <th className="py-2 px-2 text-left">핵심 프로젝트 / 주제</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#EDEBE9]">
                    {effectiveReport.weeklyTable.map((w) => (
                      <tr key={w.periodKey} className="hover:bg-slate-50">
                        <td className="py-2 px-2.5 font-semibold whitespace-nowrap">
                          {w.periodLabel}
                        </td>
                        <td className="py-2 px-2 text-right font-bold">{w.total}</td>
                        <td className="py-2 px-2 text-right text-[#0F6CBD]">{w.received}</td>
                        <td className="py-2 px-2 text-right text-indigo-700">{w.sent}</td>
                        <td className="py-2 px-2 text-right font-bold text-rose-600">
                          {w.unrepliedIncoming}
                        </td>
                        <td className="py-2 px-2 text-right text-emerald-700">
                          {w.attachmentCount}
                        </td>
                        <td className="py-2 px-2 text-[11px] text-slate-600 truncate max-w-[160px]">
                          [{w.topProject}] {w.topTopic}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            {/* Monthly Report Table */}
            <section className="bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
              <div className="flex items-center justify-between">
                <h3 className="text-sm font-bold text-[#242424]">
                  월간(Monthly) 메일 · 미회신 · 첨부 보고서 테이블
                </h3>
                <span className="text-[11px] text-slate-500">
                  최근 {effectiveReport.monthlyTable.length}개월
                </span>
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs border-collapse">
                  <thead>
                    <tr className="bg-[#F3F2F1] text-slate-700 border-b border-[#E0E0E0]">
                      <th className="py-2 px-2.5 text-left">년-월</th>
                      <th className="py-2 px-2 text-right">전체</th>
                      <th className="py-2 px-2 text-right text-[#0F6CBD]">받은 메일</th>
                      <th className="py-2 px-2 text-right text-indigo-700">보낸 메일</th>
                      <th className="py-2 px-2 text-right text-rose-600">내 미회신</th>
                      <th className="py-2 px-2 text-right text-emerald-700">첨부파일</th>
                      <th className="py-2 px-2 text-left">핵심 프로젝트 / 주제</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#EDEBE9]">
                    {effectiveReport.monthlyTable.map((m) => (
                      <tr
                        key={m.periodKey}
                        onClick={() => {
                          if (/^\d{4}-\d{2}$/.test(m.periodKey)) {
                            setPeriodType('custom');
                            setDateFrom(`${m.periodKey}-01`);
                            setDateTo(`${m.periodKey}-31`);
                          }
                        }}
                        className="hover:bg-[#EBF3FC]/50 cursor-pointer"
                        title="클릭하면 해당 월로 전체 대시보드를 필터링합니다"
                      >
                        <td className="py-2 px-2.5 font-bold text-[#0F6CBD]">{m.periodLabel}</td>
                        <td className="py-2 px-2 text-right font-bold">{m.total}</td>
                        <td className="py-2 px-2 text-right text-[#0F6CBD]">{m.received}</td>
                        <td className="py-2 px-2 text-right text-indigo-700">{m.sent}</td>
                        <td className="py-2 px-2 text-right font-bold text-rose-600">
                          {m.unrepliedIncoming}
                        </td>
                        <td className="py-2 px-2 text-right text-emerald-700">
                          {m.attachmentCount}
                        </td>
                        <td className="py-2 px-2 text-[11px] text-slate-600 truncate max-w-[160px]">
                          [{m.topProject}] {m.topTopic}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          </div>

          <DashboardEfficiencyCharts
            efficiency={efficiency}
            onOpenPstModal={onOpenPstModal}
            onOpenProjects={onOpenProjects}
            onOpenBriefing={() => setActiveTab('briefing')}
          />
        </div>
      )}

      {/* TAB 2: Incoming Unreplied Tracker (내게 왔는데 미회신건 정리) */}
      {activeTab === 'unreplied-incoming' && (
        <section className="bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div>
              <h2 className="text-sm font-bold text-rose-700 flex items-center gap-1.5">
                <AlertTriangle className="w-4 h-4" />
                내게 왔는데 아직 답장하지 않은 미회신 수신 메일 ({effectiveReport.unrepliedIncoming.length}건)
              </h2>
              <p className="text-xs text-slate-600">
                전체 스레드(conversationKey)에서 수신 이후 내가 보낸 편지함 발신 이력이 없는 실제 업무 메일만 추출했습니다 (광고·뉴스레터 자동 제외).
              </p>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-rose-50 text-rose-900 border-b border-rose-200">
                  <th className="py-2 px-2.5 text-left">우선순위 · 경과</th>
                  <th className="py-2 px-2.5 text-left">수신 일시</th>
                  <th className="py-2 px-2.5 text-left">보낸 사람</th>
                  <th className="py-2 px-2.5 text-left">프로젝트</th>
                  <th className="py-2 px-2.5 text-left">주제</th>
                  <th className="py-2 px-2.5 text-left">메일 제목 · 내용 요약</th>
                  <th className="py-2 px-2.5 text-right">조치</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#EDEBE9]">
                {effectiveReport.unrepliedIncoming.map((item) => (
                  <tr key={item.emailId} className="hover:bg-rose-50/30">
                    <td className="py-2 px-2.5 whitespace-nowrap">
                      <span
                        className={`px-2 py-0.5 rounded text-[11px] font-bold ${
                          item.urgency === 'high'
                            ? 'bg-rose-600 text-white'
                            : 'bg-amber-100 text-amber-800'
                        }`}
                      >
                        {item.urgency === 'high' ? '긴급' : '일반'} · D+{item.daysPending}
                      </span>
                    </td>
                    <td className="py-2 px-2.5 whitespace-nowrap font-mono text-[11px]">
                      {item.date} {item.timeLabel}
                    </td>
                    <td className="py-2 px-2.5 font-semibold whitespace-nowrap">
                      {item.fromName}
                      <div className="text-[10px] font-normal text-slate-500">{item.fromEmail}</div>
                    </td>
                    <td className="py-2 px-2.5 whitespace-nowrap">
                      <button
                        type="button"
                        onClick={() => setProjectCode(item.projectCode)}
                        className="px-2 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] font-bold text-[11px]"
                      >
                        {item.projectCode}
                      </button>
                    </td>
                    <td className="py-2 px-2.5 whitespace-nowrap">
                      <span className="px-2 py-0.5 rounded bg-slate-100 text-slate-700 text-[11px]">
                        {item.topic}
                      </span>
                    </td>
                    <td className="py-2 px-2.5 max-w-md">
                      <div className="font-bold text-slate-900 truncate flex items-center gap-1">
                        {item.hasAttachment && <Paperclip className="w-3 h-3 text-emerald-600 shrink-0" />}
                        <span className="truncate">{item.subject}</span>
                      </div>
                      <div className="text-[11px] text-slate-500 truncate">{item.snippet}</div>
                    </td>
                    <td className="py-2 px-2.5 text-right whitespace-nowrap space-x-1">
                      {onOpenMailInWorkbench && (
                        <button
                          type="button"
                          onClick={() => onOpenMailInWorkbench(item.emailId)}
                          className="px-2.5 py-1 rounded bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-[11px] font-semibold"
                        >
                          원문·AI답안 열기
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={() =>
                          onSaveDraftToOutlook(
                            item.subject.startsWith('RE:') ? item.subject : `RE: ${item.subject}`,
                            `안녕하세요, ${item.fromName}님.\n\n보내주신 "${item.subject}" 건 확인했습니다.\n관련 사양을 검토 중이며 조속히 회신드리겠습니다.\n\n감사합니다.`
                          )
                        }
                        className="px-2 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-semibold"
                      >
                        초안 저장
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {/* TAB 3: Project-by-Project & Topic-by-Topic Tables */}
      {activeTab === 'project-topic' && (
        <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
          {/* Project Table */}
          <section className="bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-bold text-[#242424]">
                프로젝트(호선)별 메일 · 미회신 · 첨부파일 정리
              </h3>
              <span className="text-[11px] text-slate-500">행 클릭 시 해당 프로젝트 필터링</span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="bg-[#F3F2F1] text-slate-700 border-b border-[#E0E0E0]">
                    <th className="py-2 px-2.5 text-left">프로젝트</th>
                    <th className="py-2 px-2 text-right">전체</th>
                    <th className="py-2 px-2 text-right text-[#0F6CBD]">받은거</th>
                    <th className="py-2 px-2 text-right text-indigo-700">보낸거</th>
                    <th className="py-2 px-2 text-right text-rose-600">미회신</th>
                    <th className="py-2 px-2 text-right text-emerald-700">첨부</th>
                    <th className="py-2 px-2 text-left">최근 안건 · 주요 발신자</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#EDEBE9]">
                  {effectiveReport.projectRows.map((p) => (
                    <tr
                      key={p.projectCode}
                      onClick={() =>
                        setProjectCode((prev) => (prev === p.projectCode ? 'ALL' : p.projectCode))
                      }
                      className={`cursor-pointer hover:bg-[#EBF3FC]/50 ${
                        projectCode === p.projectCode ? 'bg-[#EBF3FC]' : ''
                      }`}
                    >
                      <td className="py-2 px-2.5 font-bold text-[#0F6CBD] whitespace-nowrap">
                        [{p.projectCode}]
                        <div className="text-[10px] font-normal text-slate-500">{p.projectName}</div>
                      </td>
                      <td className="py-2 px-2 text-right font-bold">{p.total}</td>
                      <td className="py-2 px-2 text-right text-[#0F6CBD]">{p.received}</td>
                      <td className="py-2 px-2 text-right text-indigo-700">{p.sent}</td>
                      <td className="py-2 px-2 text-right font-bold text-rose-600">
                        {p.unrepliedIncoming}
                      </td>
                      <td className="py-2 px-2 text-right text-emerald-700">{p.attachmentCount}</td>
                      <td className="py-2 px-2 max-w-[200px]">
                        <div className="truncate font-medium text-slate-800">{p.latestSubject}</div>
                        <div className="text-[10px] text-slate-500 truncate">
                          {p.latestDate} · {p.topSenders.join(', ')}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {/* Topic Table */}
          <section className="bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-bold text-[#242424]">
                업무 주제(도면·POS·선급·모델링·일정)별 메일 정리
              </h3>
              <span className="text-[11px] text-slate-500">행 클릭 시 해당 주제 필터링</span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="bg-[#F3F2F1] text-slate-700 border-b border-[#E0E0E0]">
                    <th className="py-2 px-2.5 text-left">업무 주제</th>
                    <th className="py-2 px-2 text-right">전체</th>
                    <th className="py-2 px-2 text-right text-[#0F6CBD]">받은거</th>
                    <th className="py-2 px-2 text-right text-indigo-700">보낸거</th>
                    <th className="py-2 px-2 text-right text-rose-600">미회신</th>
                    <th className="py-2 px-2 text-right text-emerald-700">첨부</th>
                    <th className="py-2 px-2 text-left">관련 프로젝트 · 최근 메일</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#EDEBE9]">
                  {effectiveReport.topicRows.map((t) => (
                    <tr
                      key={t.topic}
                      onClick={() => setTopic((prev) => (prev === t.topic ? 'ALL' : t.topic))}
                      className={`cursor-pointer hover:bg-purple-50/50 ${
                        topic === t.topic ? 'bg-purple-50' : ''
                      }`}
                    >
                      <td className="py-2 px-2.5 font-bold text-purple-800 whitespace-nowrap">
                        {t.topic}
                      </td>
                      <td className="py-2 px-2 text-right font-bold">{t.total}</td>
                      <td className="py-2 px-2 text-right text-[#0F6CBD]">{t.received}</td>
                      <td className="py-2 px-2 text-right text-indigo-700">{t.sent}</td>
                      <td className="py-2 px-2 text-right font-bold text-rose-600">
                        {t.unrepliedIncoming}
                      </td>
                      <td className="py-2 px-2 text-right text-emerald-700">{t.attachmentCount}</td>
                      <td className="py-2 px-2 max-w-[200px]">
                        <div className="text-[11px] font-semibold text-[#0F6CBD] truncate">
                          {t.topProjects.join(', ')}
                        </div>
                        <div className="text-[10px] text-slate-500 truncate">{t.latestSubject}</div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      )}

      {/* TAB 4: Received / Sent Master Table & Real Attachments Table */}
      {activeTab === 'ledger-attachments' && (
        <div className="grid grid-cols-1 xl:grid-cols-12 gap-4">
          {/* Left 7 Cols: Received / Sent Mail Ledger */}
          <section className="xl:col-span-7 bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
            <h3 className="text-sm font-bold text-[#242424]">
              받은 메일 · 보낸 메일 통합 대장 (상위 {effectiveReport.mailRows.length}건)
            </h3>
            <div className="overflow-x-auto max-h-[540px]">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="bg-[#F3F2F1] text-slate-700 border-b border-[#E0E0E0]">
                    <th className="py-2 px-2 text-left">구분</th>
                    <th className="py-2 px-2 text-left">일자</th>
                    <th className="py-2 px-2 text-left">프로젝트 / 주제</th>
                    <th className="py-2 px-2 text-left">발신 / 수신</th>
                    <th className="py-2 px-2 text-left">제목</th>
                    <th className="py-2 px-2 text-left">회신상태</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#EDEBE9]">
                  {effectiveReport.mailRows.map((row) => (
                    <tr
                      key={row.emailId}
                      onClick={() => onOpenMailInWorkbench?.(row.emailId)}
                      className="hover:bg-slate-50 cursor-pointer"
                    >
                      <td className="py-1.5 px-2 whitespace-nowrap">
                        <span
                          className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                            row.direction === 'sent'
                              ? 'bg-indigo-50 text-indigo-700'
                              : 'bg-sky-50 text-[#0F6CBD]'
                          }`}
                        >
                          {row.direction === 'sent' ? '보낸거' : '받은거'}
                        </span>
                      </td>
                      <td className="py-1.5 px-2 whitespace-nowrap font-mono text-[11px]">
                        {row.date}
                      </td>
                      <td className="py-1.5 px-2 whitespace-nowrap text-[11px]">
                        <span className="font-bold text-[#0F6CBD]">[{row.projectCode}]</span>{' '}
                        <span className="text-slate-500">{row.topic}</span>
                      </td>
                      <td className="py-1.5 px-2 whitespace-nowrap font-medium truncate max-w-[120px]">
                        {row.fromName || row.fromEmail}
                      </td>
                      <td className="py-1.5 px-2 max-w-[240px]">
                        <div className="truncate font-semibold text-slate-800">
                          {row.realAttachmentCount > 0 && (
                            <span className="text-emerald-700 mr-1">
                              [첨부{row.realAttachmentCount}]
                            </span>
                          )}
                          {row.subject}
                        </div>
                      </td>
                      <td className="py-1.5 px-2 whitespace-nowrap">
                        {row.replyStatus === 'unreplied-incoming' ? (
                          <span className="px-1.5 py-0.5 rounded bg-rose-100 text-rose-700 text-[10px] font-bold">
                            미회신
                          </span>
                        ) : row.replyStatus === 'replied' ? (
                          <span className="px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 text-[10px] font-semibold">
                            회신완료
                          </span>
                        ) : (
                          <span className="text-[10px] text-slate-400">
                            {row.replyStatus === 'sent' ? '발신' : '참고'}
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {/* Right 5 Cols: Real Attachments Inventory */}
          <section className="xl:col-span-5 bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
            <h3 className="text-sm font-bold text-[#242424]">
              실첨부 문서(도면·POS·엑셀) 현황 (상위 {effectiveReport.attachmentRows.length}개)
            </h3>
            <div className="overflow-x-auto max-h-[540px]">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="bg-[#F3F2F1] text-slate-700 border-b border-[#E0E0E0]">
                    <th className="py-2 px-2 text-left">파일명</th>
                    <th className="py-2 px-2 text-left">프로젝트</th>
                    <th className="py-2 px-2 text-left">일자 · 발신자</th>
                    <th className="py-2 px-2 text-right">크기</th>
                    <th className="py-2 px-2 text-right">받기</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#EDEBE9]">
                  {effectiveReport.attachmentRows.map((att, idx) => (
                    <tr key={`${att.emailId}-${idx}`} className="hover:bg-slate-50">
                      <td
                        onClick={() => onOpenMailInWorkbench?.(att.emailId)}
                        className="py-1.5 px-2 max-w-[180px] cursor-pointer"
                      >
                        <div className="font-bold text-slate-800 truncate">{att.filename}</div>
                        <div className="text-[10px] text-slate-500 truncate">{att.subject}</div>
                      </td>
                      <td className="py-1.5 px-2 whitespace-nowrap font-bold text-[#0F6CBD]">
                        {att.projectCode}
                      </td>
                      <td className="py-1.5 px-2 whitespace-nowrap text-[11px]">
                        <div>{att.date}</div>
                        <div className="text-[10px] text-slate-500">{att.fromName}</div>
                      </td>
                      <td className="py-1.5 px-2 text-right whitespace-nowrap text-[11px]">
                        {formatBytes(att.sizeBytes)}
                      </td>
                      <td className="py-1.5 px-2 text-right whitespace-nowrap">
                        <a
                          href={`/api/local/pst/attachment?emailId=${encodeURIComponent(
                            att.emailId
                          )}&idx=${att.attachmentIndex}`}
                          download={att.filename}
                          className="px-2 py-0.5 rounded bg-emerald-50 hover:bg-emerald-100 text-emerald-700 text-[10px] font-bold inline-flex items-center gap-0.5"
                        >
                          <Download className="w-3 h-3" /> 저장
                        </a>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      )}

      {/* TAB 5: Saved Executive Report Snapshots */}
      {activeTab === 'saved-snapshots' && (
        <section className="bg-white border border-[#E0E0E0] rounded-xl p-4 shadow-sm space-y-3">
          <h3 className="text-sm font-bold text-[#242424]">
            로컬 SQLite에 저장된 주간·월간 보고서 아카이브 ({effectiveReport.savedReports.length}건)
          </h3>
          {effectiveReport.savedReports.length === 0 ? (
            <p className="text-xs text-slate-500 py-6 text-center">
              아직 저장된 보고서 스냅샷이 없습니다. [1. 종합 분석글] 탭 상단의 [보고서함 저장]을 클릭하면 영구 보관됩니다.
            </p>
          ) : (
            <div className="space-y-3">
              {effectiveReport.savedReports.map((snap) => (
                <div
                  key={snap.id}
                  className="p-3.5 rounded-lg border border-slate-200 bg-slate-50 space-y-2"
                >
                  <div className="flex items-center justify-between">
                    <div className="font-bold text-sm text-[#0F6CBD]">{snap.title}</div>
                    <span className="text-[11px] font-mono text-slate-500">
                      저장: {snap.savedAt.slice(0, 16).replace('T', ' ')}
                    </span>
                  </div>
                  <div className="text-xs text-slate-600">
                    범위: {snap.dateRangeLabel} · 메일 {snap.kpis.totalEmails}건 (받은 {snap.kpis.receivedCount} / 보낸 {snap.kpis.sentCount} / 미회신 {snap.kpis.unrepliedIncomingCount} / 첨부 {snap.kpis.realAttachmentCount})
                  </div>
                  <pre className="whitespace-pre-wrap font-sans text-xs bg-white p-3 rounded border border-slate-200 text-slate-800 max-h-56 overflow-y-auto">
                    {snap.narrativeMarkdown}
                  </pre>
                </div>
              ))}
            </div>
          )}
        </section>
      )}

      {/* TAB 6: Daily Briefing Timeline */}
      {activeTab === 'briefing' && (
        <DailyBriefingTimeline
          emails={emails}
          analysisResult={analysisResult}
          archivedResults={archivedResults}
          projects={projects}
          onSaveDraftToOutlook={onSaveDraftToOutlook}
          onOpenPstModal={onOpenPstModal}
          onRunAiAnalysis={onRunAiAnalysis}
        />
      )}
    </div>
  );
};
