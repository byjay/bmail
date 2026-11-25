import React, { useState } from 'react';
import {
  AlertCircle,
  Bot,
  Check,
  CheckCircle2,
  ClipboardCopy,
  Download,
  FileText,
  ListChecks,
  Loader2,
  MailPlus,
  Newspaper,
  SearchCheck,
} from 'lucide-react';
import type { AgentExecutionStep, AnalysisResult, MailItem } from '../types/mail';
import { exportDraftToOutlook } from '../services/officeBridge';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import { formatBlogPostMarkdown, formatWorkLogMarkdown } from '../utils/mailFilter';
import { DailyBriefingTimeline } from './DailyBriefingTimeline';

interface AnalysisResultPanelProps {
  result: AnalysisResult | null;
  isAnalyzing: boolean;
  agentSteps?: AgentExecutionStep[];
  emails?: MailItem[];
  projects?: ProjectDefinition[];
  onSaveDraftToOutlook?: (subject: string, body: string) => void;
  onOpenPstModal?: () => void;
}

type ResultTab = 'analysis' | 'briefing' | 'worklog' | 'blog';

function triggerMarkdownDownload(filename: string, content: string) {
  const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export const AnalysisResultPanel: React.FC<AnalysisResultPanelProps> = ({
  result,
  isAnalyzing,
  agentSteps = [],
  emails = [],
  projects = [],
  onSaveDraftToOutlook,
  onOpenPstModal,
}) => {
  const [activeTab, setActiveTab] = useState<ResultTab>('analysis');
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const [exportNotice, setExportNotice] = useState<string | null>(null);

  const handleCopy = async (key: string, text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedKey(key);
      window.setTimeout(() => setCopiedKey((prev) => (prev === key ? null : prev)), 2000);
    } catch {
      setExportNotice('클립보드 복사가 브라우저 정책으로 차단되었습니다. 다운로드 버튼을 이용해 주세요.');
    }
  };

  const handleSendToOutlook = (subject: string, body: string) => {
    if (onSaveDraftToOutlook) {
      onSaveDraftToOutlook(subject, body);
      setExportNotice('Outlook 임시 보관함(Drafts)에 메일 초안을 저장하고 작성 창을 열었습니다.');
      window.setTimeout(() => setExportNotice(null), 3000);
      return;
    }
    const res = exportDraftToOutlook(subject, body);
    setExportNotice(
      res.deliveredVia === 'outlook-compose'
        ? 'Outlook 새 메일 작성 창으로 초안을 전달했습니다.'
        : '기본 메일 클라이언트로 초안을 열었습니다.'
    );
    window.setTimeout(() => setExportNotice(null), 3000);
  };

  if (isAnalyzing) {
    return (
      <div className="p-6 max-w-2xl space-y-6" aria-busy="true" aria-label="AI 에이전트 자동 실행 중">
        <div className="flex items-center gap-3 pb-4 border-b border-slate-200">
          <div className="w-9 h-9 rounded-lg bg-[#0F6CBD]/10 text-[#0F6CBD] flex items-center justify-center shrink-0">
            <Bot className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-sm font-bold text-slate-900">
              AI 에이전트가 실제 메일을 자동으로 분석·작성하고 있습니다
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              대상 선별 → 민감정보 감사 → 업무 분석 → 업무 일지·블로그 초안 생성
            </p>
          </div>
        </div>

        <ol className="space-y-3" data-testid="agent-live-steps">
          {(agentSteps.length > 0
            ? agentSteps
            : [
                {
                  id: 'init',
                  title: '1. 대상 메일 자동 선별 및 스레드 분류',
                  status: 'running' as const,
                  detail: '연동된 실제 메일함에서 조건에 맞는 메일을 읽고 있습니다...',
                },
              ]
          ).map((step) => (
            <li
              key={step.id}
              className="flex items-start gap-3 p-3.5 rounded-lg border border-slate-200 bg-slate-50/70"
            >
              <div className="mt-0.5 shrink-0">
                {step.status === 'completed' ? (
                  <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                ) : step.status === 'running' ? (
                  <Loader2 className="w-4 h-4 text-[#0F6CBD] animate-spin" />
                ) : step.status === 'error' ? (
                  <AlertCircle className="w-4 h-4 text-red-600" />
                ) : (
                  <div className="w-4 h-4 rounded-full border-2 border-slate-300" />
                )}
              </div>
              <div className="min-w-0 space-y-0.5">
                <div className="text-xs font-semibold text-slate-900">{step.title}</div>
                <div className="text-xs text-slate-600 leading-relaxed">{step.detail}</div>
              </div>
            </li>
          ))}
        </ol>
      </div>
    );
  }

  if (!result) {
    return (
      <div className="h-full flex flex-col items-center justify-center text-center px-6 py-12">
        <img
          src="/assets/icon-64.png"
          alt="MailLog AI"
          width={48}
          height={48}
          className="w-12 h-12 rounded-xl mb-4"
        />
        <h3 className="text-base font-semibold text-slate-900">
          생성된 AI 업무 일지 · 블로그 리포트가 없습니다
        </h3>
        <p className="mt-2 max-w-md text-sm text-slate-600 leading-relaxed">
          실제 Outlook 원본 <code className="font-mono">.pst/.eml</code> 파일이나 Microsoft 365 계정을 연동한 뒤, 우측 Outlook 플러그인 작업 창에서 <strong>AI 에이전트 자동 실행</strong>을 누르면 업무 분석·업무 일지·블로그 초안이 자동으로 생성됩니다.
        </p>
        {onOpenPstModal && (
          <button
            type="button"
            onClick={onOpenPstModal}
            className="mt-5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg transition-colors"
          >
            Outlook 원본 .pst / .eml 파일 불러오기
          </button>
        )}
      </div>
    );
  }

  const workLogMarkdown = formatWorkLogMarkdown(result);
  const blogMarkdown = formatBlogPostMarkdown(result);

  return (
    <div className="flex flex-col h-full bg-white">
      {/* Summary Metadata Bar (Zero-Pill clean typography) */}
      <div className="px-6 py-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
        <div className="space-y-1 min-w-0">
          <h3 className="text-base font-bold text-slate-900 truncate">
            {result.workLog.title}
          </h3>
          <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
            <span className="font-mono tabular-nums text-slate-700">{result.periodLabel}</span>
            <span aria-hidden="true">·</span>
            <span>{result.scopeLabel}</span>
            <span aria-hidden="true">·</span>
            <span className="font-mono tabular-nums">
              총 {result.analyzedCount}건 (수신 {result.receivedCount} / 발신 {result.sentCount})
            </span>
            <span aria-hidden="true">·</span>
            <span className="text-[#0F6CBD] font-medium">{result.providerModel}</span>
          </div>
        </div>

        {/* Result View Tabs */}
        <div
          role="tablist"
          aria-label="생성 결과 보기 전환"
          className="inline-flex items-center p-1 bg-slate-100 rounded-lg shrink-0"
        >
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'analysis'}
            onClick={() => setActiveTab('analysis')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
              activeTab === 'analysis'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <SearchCheck className="w-3.5 h-3.5" />
            업무 분석
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'briefing'}
            onClick={() => setActiveTab('briefing')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
              activeTab === 'briefing'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <ListChecks className="w-3.5 h-3.5" />
            일간 브리핑
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'worklog'}
            onClick={() => setActiveTab('worklog')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
              activeTab === 'worklog'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <FileText className="w-3.5 h-3.5" />
            업무 일지
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'blog'}
            onClick={() => setActiveTab('blog')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
              activeTab === 'blog'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Newspaper className="w-3.5 h-3.5" />
            블로그 초안
          </button>
        </div>
      </div>

      {exportNotice && (
        <div className="px-6 py-2 bg-emerald-50 border-b border-emerald-200 text-xs text-emerald-800">
          {exportNotice}
        </div>
      )}

      {/* Tab Content */}
      <div className="flex-1 overflow-y-auto px-6 py-6">
        {activeTab === 'analysis' && (
          <div className="max-w-3xl space-y-8">
            {result.agentSteps && result.agentSteps.length > 0 && (
              <section className="p-4 rounded-lg border border-slate-200 bg-slate-50/70 space-y-2">
                <div className="flex items-center justify-between text-xs">
                  <span className="inline-flex items-center gap-1.5 font-semibold text-slate-900">
                    <Bot className="w-3.5 h-3.5 text-[#0F6CBD]" />
                    AI 에이전트 자동 실행 완료 ({result.agentSteps.length}단계)
                  </span>
                  {typeof result.redactedCount === 'number' && (
                    <span className="font-mono text-[11px] text-slate-500">
                      민감정보 마스킹 {result.redactedCount}건
                    </span>
                  )}
                </div>
                <ul className="grid grid-cols-1 sm:grid-cols-2 gap-1.5 text-[11px] text-slate-600">
                  {result.agentSteps.map((st) => (
                    <li key={st.id} className="flex items-start gap-1.5">
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0 mt-0.5" />
                      <span className="leading-snug">{st.detail}</span>
                    </li>
                  ))}
                </ul>
              </section>
            )}

            {/* Overview */}
            <section className="space-y-2">
              <h4 className="text-xs font-semibold text-slate-500">전체 메일 흐름 요약</h4>
              <p className="text-sm text-slate-800 leading-relaxed bg-slate-50 border border-slate-200 rounded-lg p-4">
                {result.overview}
              </p>
            </section>

            {/* Project-by-Project Rollup */}
            {result.projectSummaries && result.projectSummaries.length > 0 && (
              <section className="space-y-3">
                <div className="flex items-center justify-between">
                  <h4 className="text-sm font-semibold text-slate-900">
                    프로젝트별 전체 정리 ({result.projectSummaries.length}개 프로젝트)
                  </h4>
                  <span className="text-[11px] text-slate-500">
                    주고받은 메일 · 첨부파일 · 폴더 기준 자동 분류
                  </span>
                </div>
                <div className="space-y-3">
                  {result.projectSummaries.map((prj, idx) => (
                    <div
                      key={`${prj.projectCode}-${idx}`}
                      className="p-4 rounded-lg border border-slate-200 bg-white space-y-2.5"
                    >
                      <div className="flex items-center justify-between gap-3 pb-2 border-b border-slate-100">
                        <div className="flex items-center gap-2 min-w-0">
                          <span className="font-mono text-xs font-bold text-[#0F6CBD] bg-[#EBF3FC] px-2 py-0.5 rounded">
                            {prj.projectCode}
                          </span>
                          <h5 className="text-sm font-semibold text-slate-900 truncate">
                            {prj.projectName}
                          </h5>
                        </div>
                        <span
                          className={`text-xs font-medium shrink-0 ${
                            prj.status === '완료'
                              ? 'text-emerald-700'
                              : prj.status === '후속 조치 필요'
                                ? 'text-amber-700'
                                : 'text-[#0F6CBD]'
                          }`}
                        >
                          {prj.status}
                        </span>
                      </div>

                      {prj.myActions.length > 0 && (
                        <div className="space-y-1">
                          <div className="text-[11px] font-semibold text-slate-500">내가 수행한 일</div>
                          <ul className="list-disc pl-4 space-y-0.5 text-xs text-slate-800">
                            {prj.myActions.map((act, aIdx) => (
                              <li key={aIdx}>{act}</li>
                            ))}
                          </ul>
                        </div>
                      )}

                      {prj.keyDecisions.length > 0 && (
                        <div className="space-y-1">
                          <div className="text-[11px] font-semibold text-slate-500">주요 협의 · 결정</div>
                          <ul className="list-disc pl-4 space-y-0.5 text-xs text-slate-700">
                            {prj.keyDecisions.map((dec, dIdx) => (
                              <li key={dIdx}>{dec}</li>
                            ))}
                          </ul>
                        </div>
                      )}

                      <div className="flex flex-wrap items-center justify-between gap-2 pt-1 text-[11px] text-slate-600">
                        <span>
                          <strong className="text-slate-700">다음 조치:</strong> {prj.nextAction}
                        </span>
                        {prj.attachmentsMentioned.length > 0 && (
                          <span className="font-mono text-slate-500">
                            첨부: {prj.attachmentsMentioned.join(', ')}
                          </span>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </section>
            )}

            {/* Work Performed (My Accomplished Work) */}
            <section className="space-y-3">
              <div className="flex items-center justify-between">
                <h4 className="text-sm font-semibold text-slate-900">
                  내가 한 일 상세 분석 ({result.workPerformed.length}건)
                </h4>
              </div>
              <div className="divide-y divide-slate-200 border border-slate-200 rounded-lg overflow-hidden">
                {result.workPerformed.map((item, idx) => (
                  <div key={`${item.title}-${idx}`} className="p-4 bg-white space-y-1.5">
                    <div className="flex items-start justify-between gap-3">
                      <h5 className="text-sm font-semibold text-slate-900">{item.title}</h5>
                      <span
                        className={`text-xs font-medium shrink-0 ${
                          item.status === '완료'
                            ? 'text-emerald-700'
                            : item.status === '후속 조치 필요'
                              ? 'text-amber-700'
                              : 'text-[#0F6CBD]'
                        }`}
                      >
                        {item.status}
                      </span>
                    </div>
                    <p className="text-xs text-slate-600 leading-relaxed">{item.detail}</p>
                    {item.relatedEmails.length > 0 && (
                      <p className="text-[11px] text-slate-400">
                        근거 메일: {item.relatedEmails.join(' · ')}
                      </p>
                    )}
                  </div>
                ))}
              </div>
            </section>

            {/* Incoming Highlights */}
            <section className="space-y-3">
              <h4 className="text-sm font-semibold text-slate-900">
                수신된 주요 메일 및 대응 사항 ({result.incomingHighlights.length}건)
              </h4>
              {result.incomingHighlights.length === 0 ? (
                <p className="text-xs text-slate-500">
                  분석 대상에 수신 메일이 없거나 별도 대응 항목이 없습니다.
                </p>
              ) : (
                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-slate-600">
                        <th className="py-2.5 px-3 font-semibold">발신자 / 제목</th>
                        <th className="py-2.5 px-3 font-semibold">핵심 목적</th>
                        <th className="py-2.5 px-3 font-semibold">중요도</th>
                        <th className="py-2.5 px-3 font-semibold">필요 조치</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {result.incomingHighlights.map((inc, idx) => (
                        <tr key={`${inc.subject}-${idx}`} className="hover:bg-slate-50/70">
                          <td className="py-3 px-3 align-top">
                            <div className="font-medium text-slate-900">{inc.sender}</div>
                            <div className="text-slate-500 mt-0.5">{inc.subject}</div>
                          </td>
                          <td className="py-3 px-3 align-top text-slate-700">{inc.intent}</td>
                          <td className="py-3 px-3 align-top whitespace-nowrap">
                            <span
                              className={`font-medium ${
                                inc.priority === '높음'
                                  ? 'text-red-700'
                                  : inc.priority === '보통'
                                    ? 'text-amber-700'
                                    : 'text-slate-500'
                              }`}
                            >
                              {inc.priority}
                            </span>
                          </td>
                          <td className="py-3 px-3 align-top text-slate-800">{inc.actionNeeded}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </div>
        )}

        {activeTab === 'worklog' && (
          <div className="max-w-3xl space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-2 pb-4 border-b border-slate-200">
              <div>
                <h4 className="text-base font-bold text-slate-900">{result.workLog.title}</h4>
                <p className="text-xs text-slate-500 mt-0.5">
                  팀 보고 · Notion · Confluence · 이메일 공유용 구조화 업무 일지
                </p>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => handleCopy('worklog', workLogMarkdown)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap"
                >
                  {copiedKey === 'worklog' ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                      복사됨
                    </>
                  ) : (
                    <>
                      <ClipboardCopy className="w-3.5 h-3.5" />
                      Markdown 복사
                    </>
                  )}
                </button>
                <button
                  type="button"
                  onClick={() =>
                    triggerMarkdownDownload(
                      `worklog-${result.periodLabel.replace(/\s+/g, '')}.md`,
                      workLogMarkdown
                    )
                  }
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap"
                >
                  <Download className="w-3.5 h-3.5" />
                  .md 저장
                </button>
                <button
                  type="button"
                  onClick={() => handleSendToOutlook(result.workLog.title, workLogMarkdown)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-md transition-colors whitespace-nowrap"
                >
                  <MailPlus className="w-3.5 h-3.5" />
                  Outlook 메일 작성
                </button>
              </div>
            </div>

            <div className="space-y-5 text-sm text-slate-800">
              <div>
                <h5 className="text-xs font-semibold text-slate-500 mb-1.5">핵심 요약</h5>
                <p className="leading-relaxed bg-slate-50 border border-slate-200 rounded-lg p-4">
                  {result.workLog.executiveSummary}
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div className="border border-slate-200 rounded-lg p-4 space-y-2">
                  <h5 className="text-xs font-semibold text-emerald-800">완료한 주요 업무</h5>
                  <ul className="space-y-1.5 text-xs text-slate-700 list-disc pl-4">
                    {result.workLog.completedTasks.map((t, i) => (
                      <li key={i} className="leading-relaxed">{t}</li>
                    ))}
                  </ul>
                </div>

                <div className="border border-slate-200 rounded-lg p-4 space-y-2">
                  <h5 className="text-xs font-semibold text-[#0F6CBD]">진행 중인 업무</h5>
                  <ul className="space-y-1.5 text-xs text-slate-700 list-disc pl-4">
                    {result.workLog.inProgressTasks.map((t, i) => (
                      <li key={i} className="leading-relaxed">{t}</li>
                    ))}
                  </ul>
                </div>
              </div>

              <div className="border border-slate-200 rounded-lg p-4 space-y-2">
                <h5 className="text-xs font-semibold text-slate-900">주요 메일 협의 및 결정 사항</h5>
                <ul className="space-y-1.5 text-xs text-slate-700 list-disc pl-4">
                  {result.workLog.keyCommunications.map((t, i) => (
                    <li key={i} className="leading-relaxed">{t}</li>
                  ))}
                </ul>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div className="border border-slate-200 rounded-lg p-4 space-y-2">
                  <h5 className="text-xs font-semibold text-amber-800">이슈 및 리스크</h5>
                  <ul className="space-y-1.5 text-xs text-slate-700 list-disc pl-4">
                    {result.workLog.issuesAndRisks.length > 0 ? (
                      result.workLog.issuesAndRisks.map((t, i) => (
                        <li key={i} className="leading-relaxed">{t}</li>
                      ))
                    ) : (
                      <li>특이 리스크 없음</li>
                    )}
                  </ul>
                </div>

                <div className="border border-slate-200 rounded-lg p-4 space-y-2">
                  <h5 className="text-xs font-semibold text-slate-900">익일 / 후속 실행 계획</h5>
                  <ul className="space-y-1.5 text-xs text-slate-700 list-disc pl-4">
                    {result.workLog.nextStepPlan.map((t, i) => (
                      <li key={i} className="leading-relaxed">{t}</li>
                    ))}
                  </ul>
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'blog' && (
          <article className="max-w-2xl space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-2 pb-4 border-b border-slate-200">
              <div className="text-xs text-slate-500 flex items-center gap-2">
                <span className="font-mono tabular-nums">
                  약 {result.blogPost.readingTimeMinutes}분 분량
                </span>
                <span aria-hidden="true">·</span>
                <span>
                  {result.blogPost.tags.map((t) => `#${t.replace(/^#/, '')}`).join('  ')}
                </span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => handleCopy('blog', blogMarkdown)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap"
                >
                  {copiedKey === 'blog' ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                      복사됨
                    </>
                  ) : (
                    <>
                      <ClipboardCopy className="w-3.5 h-3.5" />
                      블로그 Markdown 복사
                    </>
                  )}
                </button>
                <button
                  type="button"
                  onClick={() =>
                    triggerMarkdownDownload(
                      `blog-draft-${result.periodLabel.replace(/\s+/g, '')}.md`,
                      blogMarkdown
                    )
                  }
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-md transition-colors whitespace-nowrap"
                >
                  <Download className="w-3.5 h-3.5" />
                  .md 다운로드
                </button>
              </div>
            </div>

            <header className="space-y-2">
              <h1 className="text-2xl font-bold text-slate-900 leading-snug tracking-tight">
                {result.blogPost.title}
              </h1>
              <p className="text-sm text-slate-600 leading-relaxed">
                {result.blogPost.subtitle}
              </p>
            </header>

            <div className="space-y-6 text-[15px] text-slate-800 leading-[1.75]">
              <p className="whitespace-pre-line">{result.blogPost.intro}</p>

              {result.blogPost.sections.map((sec, idx) => (
                <section key={`${sec.heading}-${idx}`} className="space-y-2.5">
                  <h2 className="text-lg font-bold text-slate-900 pt-2">
                    {sec.heading}
                  </h2>
                  <p className="whitespace-pre-line text-slate-700">{sec.body}</p>
                </section>
              ))}

              <section className="space-y-2 pt-3 border-t border-slate-200">
                <h2 className="text-base font-bold text-slate-900">마치며</h2>
                <p className="whitespace-pre-line text-slate-700">{result.blogPost.conclusion}</p>
              </section>
            </div>
          </article>
        )}

        {activeTab === 'briefing' && (
          <div className="max-w-4xl">
            <DailyBriefingTimeline
              compact
              emails={emails}
              analysisResult={result}
              projects={projects}
              onSaveDraftToOutlook={onSaveDraftToOutlook}
              onOpenPstModal={onOpenPstModal}
            />
          </div>
        )}
      </div>
    </div>
  );
};
