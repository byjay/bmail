import React, { useMemo, useState } from 'react';
import {
  AlertTriangle,
  ArrowRight,
  CalendarPlus,
  CheckCircle2,
  Clock,
  FileEdit,
  Inbox,
  MessageSquare,
  Plus,
  Search,
  Send,
  ShieldCheck,
  Sparkles,
  Tag,
  Trash2,
} from 'lucide-react';
import type { MailItem } from '../types/mail';
import type { CalendarScheduleEvent, ProjectDefinition } from '../utils/deterministicAnalyzer';
import {
  answerInboxQuestion,
  type AskInboxResponse,
  buildMorningBrief,
  coachEmailDraft,
  type CustomAutoLabelRule,
  DEFAULT_AUTO_LABEL_RULES,
  expandBulletsToDraft,
  generateInstantReplyOptions,
  type InstantReplyOption,
} from '../utils/aiMailIntelligence';
import type { OutlookDraftItem } from './OutlookWorkbench';

interface AiInboxIntelligenceViewProps {
  emails: MailItem[];
  projects: ProjectDefinition[];
  calendarEvents: CalendarScheduleEvent[];
  onOpenPstModal: () => void;
  onSaveDraft: (draft: Omit<OutlookDraftItem, 'id' | 'updatedAt'>) => void;
  onAddCalendarEvent: (ev: Omit<CalendarScheduleEvent, 'id'>) => void;
  onDeleteEmail: (id: string) => void;
  onOpenMailInWorkbench: (emailId: string) => void;
  onOpenExecutiveDashboard?: () => void;
}

type HubTab = 'brief' | 'ask' | 'labels' | 'write';

const QUICK_ASK_PROMPTS = [
  '회신해야 할 긴급 메일 알려줘',
  '견적·금액이 포함된 영업 메일 정리해줘',
  '첨부파일이나 도면이 있는 메일 찾아줘',
  '일정·마감이 언급된 메일 요약해줘',
];

export const AiInboxIntelligenceView: React.FC<AiInboxIntelligenceViewProps> = ({
  emails,
  projects,
  calendarEvents,
  onOpenPstModal,
  onSaveDraft,
  onAddCalendarEvent,
  onDeleteEmail,
  onOpenMailInWorkbench,
  onOpenExecutiveDashboard,
}) => {
  const [activeTab, setActiveTab] = useState<HubTab>('brief');
  const [notice, setNotice] = useState<string | null>(null);

  // Auto-Label rules (Superhuman / Notion Mail style)
  const [labelRules, setLabelRules] = useState<CustomAutoLabelRule[]>(DEFAULT_AUTO_LABEL_RULES);
  const [selectedLabelName, setSelectedLabelName] = useState<string>('ALL');
  const [newLabelName, setNewLabelName] = useState('');
  const [newLabelKeywords, setNewLabelKeywords] = useState('');
  const [newLabelPrompt, setNewLabelPrompt] = useState('');

  // Selected mail for Instant 3-Option Reply preview
  const [instantReplyTargetId, setInstantReplyTargetId] = useState<string>('');

  // Ask AI state (Superhuman / Shortwave Ask AI + Hybrid GCP Credit Priority)
  type AskHistoryEntry = AskInboxResponse & {
    creditRoute?: 'agent-builder-rag' | 'gemini-primary' | 'deterministic-local';
    creditLabel?: string;
    fallbackReason?: string;
  };
  const [askInput, setAskInput] = useState('');
  const [askHistory, setAskHistory] = useState<AskHistoryEntry[]>([]);
  const [isAsking, setIsAsking] = useState(false);

  // Write with AI + Copilot Coaching state
  const [writeTone, setWriteTone] = useState<'formal' | 'concise' | 'english'>('formal');
  const [writeTo, setWriteTo] = useState('');
  const [writeRecipientName, setWriteRecipientName] = useState('');
  const [writeSubject, setWriteSubject] = useState('');
  const [writeBullets, setWriteBullets] = useState('');
  const [writeBody, setWriteBody] = useState('');
  const [writeHasAttachment, setWriteHasAttachment] = useState(false);

  const brief = useMemo(
    () => buildMorningBrief(emails, labelRules),
    [emails, labelRules]
  );

  const emailById = useMemo(() => {
    const map = new Map<string, MailItem>();
    for (const m of emails) map.set(m.id, m);
    return map;
  }, [emails]);

  const instantReplyTarget = useMemo(() => {
    if (instantReplyTargetId && emailById.has(instantReplyTargetId)) {
      return emailById.get(instantReplyTargetId) || null;
    }
    return brief.actionRequired[0]?.mail || brief.decisionNeeded[0]?.mail || emails[0] || null;
  }, [instantReplyTargetId, emailById, brief, emails]);

  const instantReplies: InstantReplyOption[] = useMemo(() => {
    if (!instantReplyTarget) return [];
    return generateInstantReplyOptions(instantReplyTarget);
  }, [instantReplyTarget]);

  const coaching = useMemo(
    () =>
      coachEmailDraft({
        to: writeTo,
        subject: writeSubject,
        body: writeBody,
        hasAttachment: writeHasAttachment,
      }),
    [writeTo, writeSubject, writeBody, writeHasAttachment]
  );

  const handleSaveInstantReply = (opt: InstantReplyOption, targetMail: MailItem) => {
    onSaveDraft({
      to: targetMail.fromEmail || targetMail.fromName,
      cc: '',
      subject: opt.subject,
      body: opt.body,
      kind: 'reply',
    });
    setNotice(`"${opt.label}" 회신 초안을 임시보관함에 저장했습니다. (자동 발송되지 않음)`);
  };

  const handleSaveReminder = (item: MorningBriefReportWaitingItem) => {
    onSaveDraft({
      to: item.recipientLabel,
      cc: '',
      subject: item.reminderSubject,
      body: item.reminderBody,
      kind: 'reply',
    });
    setNotice(`"${item.subject}" 미회신 리마인더 초안을 임시보관함에 저장했습니다.`);
  };

  const handleAutoSweepFyi = () => {
    if (brief.fyiSweepable.length === 0) return;
    for (const entry of brief.fyiSweepable) {
      onDeleteEmail(entry.mail.id);
    }
    setNotice(`참고·뉴스레터 ${brief.fyiSweepable.length}건을 휴지통으로 일괄 정리했습니다.`);
  };

  const handleAddCommitmentToCalendar = (task: string, deadline: string, sourceSubject: string) => {
    const isoMatch = deadline.match(/(20\d{2})[-./](\d{1,2})[-./](\d{1,2})/);
    const date = isoMatch
      ? `${isoMatch[1]}-${isoMatch[2].padStart(2, '0')}-${isoMatch[3].padStart(2, '0')}`
      : new Date().toISOString().slice(0, 10);
    onAddCalendarEvent({
      title: `[AI 약속] ${task.slice(0, 60)}`,
      date,
      timeLabel: '10:00',
      category: 'action-item',
      notes: `출처 메일: ${sourceSubject}`,
    });
    setNotice(`약속 항목을 Bmail 일정 달력(${date})에 등록했습니다.`);
  };

  const handleRunAsk = async (questionOverride?: string) => {
    const q = (questionOverride ?? askInput).trim();
    if (!q) return;
    setIsAsking(true);
    try {
      const res = await fetch('/api/ai-inbox/ask', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          question: q,
          emails,
          projects,
          calendarEvents,
        }),
      });
      if (res.ok) {
        const payload = (await res.json()) as { response?: AskHistoryEntry } | AskHistoryEntry;
        const resolved =
          payload && typeof payload === 'object' && 'response' in payload && payload.response
            ? payload.response
            : (payload as AskHistoryEntry);
        setAskHistory((prev) => [resolved, ...prev].slice(0, 10));
      } else {
        const local = answerInboxQuestion(q, emails, projects, calendarEvents);
        setAskHistory((prev) => [
          { ...local, creditRoute: 'deterministic-local' as const, creditLabel: '로컬 0-LLM 결정적 검색 엔진' },
          ...prev,
        ].slice(0, 10));
      }
    } catch {
      const local = answerInboxQuestion(q, emails, projects, calendarEvents);
      setAskHistory((prev) => [
        { ...local, creditRoute: 'deterministic-local' as const, creditLabel: '로컬 0-LLM 결정적 검색 엔진' },
        ...prev,
      ].slice(0, 10));
    } finally {
      setAskInput('');
      setIsAsking(false);
    }
  };

  const handleAddLabelRule = () => {
    const name = newLabelName.trim();
    const keywords = newLabelKeywords
      .split(/[,，]/)
      .map((k) => k.trim())
      .filter(Boolean);
    if (!name || keywords.length === 0) {
      setNotice('라벨 이름과 쉼표(,)로 구분된 키워드를 1개 이상 입력하세요.');
      return;
    }
    const created: CustomAutoLabelRule = {
      id: `lbl-custom-${Date.now()}`,
      name,
      promptDescription: newLabelPrompt.trim() || `${keywords.join(', ')} 관련 자동 분류`,
      keywords,
      enabled: true,
    };
    setLabelRules((prev) => [created, ...prev]);
    setNewLabelName('');
    setNewLabelKeywords('');
    setNewLabelPrompt('');
    setNotice(`사용자 정의 Auto-Label "${name}" 규칙을 추가했습니다.`);
  };

  const handleExpandBullets = () => {
    const generated = expandBulletsToDraft(writeBullets, {
      tone: writeTone,
      recipientName: writeRecipientName,
      subjectHint: writeSubject,
    });
    if (!writeSubject.trim()) setWriteSubject(generated.subject);
    setWriteBody(generated.body);
    setNotice('핵심 불릿 항목을 비즈니스 메일 초안으로 확장했습니다.');
  };

  const handleSaveComposedDraft = () => {
    if (!writeBody.trim()) {
      setNotice('본문을 작성하거나 불릿 확장을 먼저 실행하세요.');
      return;
    }
    onSaveDraft({
      to: writeTo.trim() || writeRecipientName.trim(),
      cc: '',
      subject: writeSubject.trim() || '(제목 없음)',
      body: writeBody,
      kind: 'new',
    });
    setNotice('작성한 메일 초안을 Bmail 임시보관함에 저장했습니다.');
  };

  const splitLabelEmails = useMemo(() => {
    if (selectedLabelName === 'ALL') return emails;
    const found = brief.labelCounts.find((item) => item.label === selectedLabelName);
    if (!found) return [];
    const idSet = new Set(found.emailIds);
    return emails.filter((m) => idSet.has(m.id));
  }, [selectedLabelName, brief.labelCounts, emails]);

  return (
    <div className="flex-1 min-h-0 flex flex-col bg-[#F5F5F5] text-[#242424] overflow-hidden">
      {/* Top Workspace Header (Single-line controls, clean Fluent 2 bar) */}
      <div className="shrink-0 bg-white border-b border-[#E0E0E0] px-6 py-3 flex items-center justify-between gap-4">
        <div className="min-w-0">
          <div className="flex items-center gap-2 text-xs text-[#616161]">
            <span>Bmail</span>
            <span aria-hidden="true">/</span>
            <span>AI 인텔리전스 스위트</span>
            <span aria-hidden="true">·</span>
            <span className="tabular-nums">분석 대상 {emails.length}건</span>
          </div>
          <h1 className="text-[17px] font-semibold text-[#242424] mt-0.5 truncate">
            모닝 브리프 · Ask AI · Auto-Label · 즉시 회신 · 발송 코칭
          </h1>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          {onOpenExecutiveDashboard && (
            <button
              type="button"
              onClick={onOpenExecutiveDashboard}
              className="px-3 py-1.5 rounded bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-bold flex items-center gap-1.5 shadow-xs"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>주간 · 월간 분석글 · 내게 온 미회신 대시보드 열기</span>
            </button>
          )}
          <div className="flex items-center gap-1 bg-[#F3F2F1] p-1 rounded shrink-0">
          {(
            [
              { id: 'brief', label: '모닝 브리프' },
              { id: 'ask', label: 'Ask AI 질의' },
              { id: 'labels', label: 'Auto-Label 스플릿' },
              { id: 'write', label: 'Write & 톤 코칭' },
            ] as const
          ).map((tab) => (
            <button
              key={tab.id}
              type="button"
              onClick={() => setActiveTab(tab.id)}
              className={`px-3 py-1.5 text-xs font-medium rounded whitespace-nowrap shrink-0 transition-colors ${
                activeTab === tab.id
                  ? 'bg-white text-[#0F6CBD] shadow-xs font-semibold'
                  : 'text-[#424242] hover:text-[#242424]'
              }`}
            >
              {tab.label}
            </button>
          ))}
          </div>
        </div>
      </div>

      {notice && (
        <div className="shrink-0 bg-[#EBF3FC] border-b border-[#CFE4FA] px-6 py-2 flex items-center justify-between gap-3 text-xs text-[#0F6CBD]">
          <span className="truncate">{notice}</span>
          <button
            type="button"
            onClick={() => setNotice(null)}
            className="text-xs font-medium hover:underline whitespace-nowrap shrink-0"
          >
            닫기
          </button>
        </div>
      )}

      {/* Main Content Viewport */}
      <div className="flex-1 min-h-0 overflow-y-auto p-6">
        <div className="max-w-[1280px] mx-auto space-y-5">
          {/* Empty State if zero real emails loaded */}
          {emails.length === 0 && activeTab !== 'write' && (
            <div className="bg-white border border-[#E0E0E0] rounded p-8 flex flex-col items-start gap-3">
              <div className="text-sm font-semibold text-[#242424]">
                연동된 실제 메일이 아직 없습니다
              </div>
              <p className="text-xs text-[#616161] leading-relaxed max-w-2xl">
                Bmail은 샘플/시뮬레이션 데이터를 생성하지 않습니다. 실제 .pst/.eml/.mbox 파일을 열거나
                Microsoft 365 계정을 동기화하면 Genspark형 모닝 브리프, Superhuman형 Ask AI 및 3종 즉시 회신,
                Auto-Label 스플릿 인박스, 미회신 SLA 추적이 즉시 작동합니다.
              </p>
              <button
                type="button"
                onClick={onOpenPstModal}
                className="px-4 py-2 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0"
              >
                실제 PST · EML 메일 불러오기
              </button>
            </div>
          )}

          {/* TAB 1: MORNING EXECUTIVE BRIEF + TRIAGE + WAITING-ON SLA + INSTANT REPLIES */}
          {activeTab === 'brief' && emails.length > 0 && (
            <>
              {/* Clean 4-Metric Strip (No nested cards, tabular numerals) */}
              <div className="bg-white border border-[#E0E0E0] rounded grid grid-cols-2 lg:grid-cols-4 divide-y lg:divide-y-0 lg:divide-x divide-[#EDEBE9]">
                <div className="p-4">
                  <div className="text-xs text-[#616161]">즉시 회신 · 조치 필요</div>
                  <div className="text-2xl font-semibold text-[#C50F1F] tabular-nums mt-1">
                    {brief.actionRequired.length}건
                  </div>
                  <div className="text-[11px] text-[#616161] mt-1">우선 처리 대상</div>
                </div>
                <div className="p-4">
                  <div className="text-xs text-[#616161]">견적 · 의사결정 검토</div>
                  <div className="text-2xl font-semibold text-[#0F6CBD] tabular-nums mt-1">
                    {brief.decisionNeeded.length}건
                  </div>
                  <div className="text-[11px] text-[#616161] mt-1">영업 · 발주 · 승인 건</div>
                </div>
                <div className="p-4">
                  <div className="text-xs text-[#616161]">상대 회신 대기 (SLA)</div>
                  <div className="text-2xl font-semibold text-[#9A6700] tabular-nums mt-1">
                    {brief.waitingOnOthers.length}건
                  </div>
                  <div className="text-[11px] text-[#616161] mt-1">보낸 메일 미회신 추적</div>
                </div>
                <div className="p-4 flex items-start justify-between gap-2">
                  <div>
                    <div className="text-xs text-[#616161]">참고 · Auto-Sweep 대상</div>
                    <div className="text-2xl font-semibold text-[#242424] tabular-nums mt-1">
                      {brief.fyiSweepable.length}건
                    </div>
                    <div className="text-[11px] text-[#616161] mt-1">뉴스레터 · 알림 분리</div>
                  </div>
                  {brief.fyiSweepable.length > 0 && (
                    <button
                      type="button"
                      onClick={handleAutoSweepFyi}
                      className="px-2.5 py-1.5 text-[11px] font-medium border border-[#D1D1D1] rounded hover:bg-[#F5F5F5] text-[#424242] whitespace-nowrap shrink-0 flex items-center gap-1"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      일괄 정리
                    </button>
                  )}
                </div>
              </div>

              {/* Headline Brief Bar */}
              <div className="bg-white border border-[#E0E0E0] rounded px-5 py-3.5 flex flex-wrap items-center justify-between gap-3">
                <div className="text-xs text-[#242424] leading-relaxed">{brief.headlineSummary}</div>
                {brief.keyHighlights.length > 0 && (
                  <div className="text-[11px] text-[#616161]">
                    핵심 추출: {brief.keyHighlights.slice(0, 2).join(' · ')}
                  </div>
                )}
              </div>

              {/* 2-Column Main Brief Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
                {/* Left 7 cols: Action Required + Decision Needed */}
                <div className="lg:col-span-7 space-y-5">
                  <section className="bg-white border border-[#E0E0E0] rounded overflow-hidden">
                    <div className="px-5 py-3 border-b border-[#EDEBE9] flex items-center justify-between">
                      <h2 className="text-sm font-semibold text-[#242424]">
                        즉시 회신 · 조치 필요 ({brief.actionRequired.length})
                      </h2>
                      <span className="text-xs text-[#616161]">긴급도 우선 정렬</span>
                    </div>
                    {brief.actionRequired.length === 0 ? (
                      <div className="p-5 text-xs text-[#616161]">즉시 회신이 필요한 메일이 없습니다.</div>
                    ) : (
                      <div className="divide-y divide-[#EDEBE9]">
                        {brief.actionRequired.slice(0, 8).map(({ mail, profile }) => (
                          <div
                            key={mail.id}
                            className={`px-5 py-3.5 hover:bg-[#FAF9F8] transition-colors ${
                              instantReplyTarget?.id === mail.id ? 'bg-[#F3F8FD]' : ''
                            }`}
                          >
                            <div className="flex items-start justify-between gap-3">
                              <div className="min-w-0">
                                <div className="text-xs text-[#616161] flex flex-wrap items-center gap-1.5">
                                  <span className="font-semibold text-[#C50F1F]">{profile.urgencyLabel}</span>
                                  <span aria-hidden="true">·</span>
                                  <span>{mail.fromName}</span>
                                  <span aria-hidden="true">·</span>
                                  <span className="tabular-nums">{mail.date}</span>
                                  {profile.isForeignLanguage && (
                                    <>
                                      <span aria-hidden="true">·</span>
                                      <span className="text-[#0F6CBD]">외국어({profile.detectedLanguage.toUpperCase()})</span>
                                    </>
                                  )}
                                </div>
                                <div className="text-sm font-semibold text-[#242424] mt-0.5 truncate">
                                  {mail.subject}
                                </div>
                                <div className="text-xs text-[#424242] mt-1">{profile.oneLineSummary}</div>
                                {profile.riskFlags.length > 0 && (
                                  <div className="text-[11px] text-[#9A6700] mt-1">
                                    리스크 신호: {profile.riskFlags.join(' · ')}
                                  </div>
                                )}
                              </div>
                              <div className="flex items-center gap-1.5 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => setInstantReplyTargetId(mail.id)}
                                  className="px-2.5 py-1.5 text-[11px] font-medium bg-[#0F6CBD] hover:bg-[#115EA3] text-white rounded whitespace-nowrap"
                                >
                                  3종 회신
                                </button>
                                <button
                                  type="button"
                                  onClick={() => onOpenMailInWorkbench(mail.id)}
                                  className="px-2.5 py-1.5 text-[11px] font-medium border border-[#D1D1D1] hover:bg-[#F5F5F5] text-[#242424] rounded whitespace-nowrap"
                                >
                                  원문
                                </button>
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </section>

                  <section className="bg-white border border-[#E0E0E0] rounded overflow-hidden">
                    <div className="px-5 py-3 border-b border-[#EDEBE9] flex items-center justify-between">
                      <h2 className="text-sm font-semibold text-[#242424]">
                        영업 견적 · 의사결정 검토 ({brief.decisionNeeded.length})
                      </h2>
                      <span className="text-xs text-[#616161]">금액 · 발주 · 승인</span>
                    </div>
                    {brief.decisionNeeded.length === 0 ? (
                      <div className="p-5 text-xs text-[#616161]">의사결정 검토 대기 메일이 없습니다.</div>
                    ) : (
                      <div className="divide-y divide-[#EDEBE9]">
                        {brief.decisionNeeded.slice(0, 6).map(({ mail, profile }) => (
                          <div key={mail.id} className="px-5 py-3.5 hover:bg-[#FAF9F8] flex items-start justify-between gap-3">
                            <div className="min-w-0">
                              <div className="text-xs text-[#616161] flex flex-wrap items-center gap-1.5">
                                <span className="font-semibold text-[#0F6CBD]">{profile.intentLabel}</span>
                                <span aria-hidden="true">·</span>
                                <span>{mail.fromName}</span>
                                <span aria-hidden="true">·</span>
                                <span className="tabular-nums">{mail.date}</span>
                                {profile.extractedAmounts.length > 0 && (
                                  <>
                                    <span aria-hidden="true">·</span>
                                    <span className="font-mono tabular-nums text-[#242424]">
                                      {profile.extractedAmounts.join(', ')}
                                    </span>
                                  </>
                                )}
                              </div>
                              <div className="text-sm font-semibold text-[#242424] mt-0.5 truncate">
                                {mail.subject}
                              </div>
                              <div className="text-xs text-[#424242] mt-1">{profile.oneLineSummary}</div>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0">
                              <button
                                type="button"
                                onClick={() => setInstantReplyTargetId(mail.id)}
                                className="px-2.5 py-1.5 text-[11px] font-medium border border-[#0F6CBD] text-[#0F6CBD] hover:bg-[#EBF3FC] rounded whitespace-nowrap"
                              >
                                3종 회신
                              </button>
                              <button
                                type="button"
                                onClick={() => onOpenMailInWorkbench(mail.id)}
                                className="px-2.5 py-1.5 text-[11px] font-medium border border-[#D1D1D1] hover:bg-[#F5F5F5] text-[#242424] rounded whitespace-nowrap"
                              >
                                원문
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </section>
                </div>

                {/* Right 5 cols: Instant 3-Option Reply + Waiting-On SLA Radar + Commitments */}
                <div className="lg:col-span-5 space-y-5">
                  {/* Superhuman-style Instant 3-Option Reply Cards */}
                  <section className="bg-white border border-[#E0E0E0] rounded overflow-hidden">
                    <div className="px-5 py-3 border-b border-[#EDEBE9]">
                      <div className="text-sm font-semibold text-[#242424] flex items-center gap-1.5">
                        <Sparkles className="w-4 h-4 text-[#0F6CBD]" />
                        즉시 3종 회신 초안 (Instant Reply)
                      </div>
                      {instantReplyTarget && (
                        <div className="text-xs text-[#616161] mt-0.5 truncate">
                          대상: {instantReplyTarget.fromName} · {instantReplyTarget.subject}
                        </div>
                      )}
                    </div>
                    {instantReplyTarget ? (
                      <div className="divide-y divide-[#EDEBE9]">
                        {instantReplies.map((opt) => (
                          <div key={opt.id} className="p-4 space-y-2">
                            <div className="flex items-center justify-between gap-2">
                              <div>
                                <div className="text-xs font-semibold text-[#242424]">{opt.label}</div>
                                <div className="text-[11px] text-[#616161]">{opt.intentHint}</div>
                              </div>
                              <button
                                type="button"
                                onClick={() => handleSaveInstantReply(opt, instantReplyTarget)}
                                className="px-2.5 py-1.5 text-[11px] font-medium bg-[#0F6CBD] hover:bg-[#115EA3] text-white rounded whitespace-nowrap shrink-0"
                              >
                                임시보관함 저장
                              </button>
                            </div>
                            <pre className="text-[11px] leading-relaxed text-[#424242] whitespace-pre-wrap font-sans bg-[#FAF9F8] border border-[#EDEBE9] rounded p-2.5 max-h-32 overflow-y-auto">
                              {opt.body}
                            </pre>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="p-4 text-xs text-[#616161]">회신할 메일을 선택하세요.</div>
                    )}
                  </section>

                  {/* Waiting-On SLA Radar */}
                  <section className="bg-white border border-[#E0E0E0] rounded overflow-hidden">
                    <div className="px-5 py-3 border-b border-[#EDEBE9] flex items-center justify-between">
                      <h2 className="text-sm font-semibold text-[#242424] flex items-center gap-1.5">
                        <Clock className="w-4 h-4 text-[#9A6700]" />
                        상대 회신 대기 SLA 추적 ({brief.waitingOnOthers.length})
                      </h2>
                    </div>
                    {brief.waitingOnOthers.length === 0 ? (
                      <div className="p-4 text-xs text-[#616161]">
                        미회신 상태로 대기 중인 발신 메일이 없습니다.
                      </div>
                    ) : (
                      <div className="divide-y divide-[#EDEBE9]">
                        {brief.waitingOnOthers.slice(0, 5).map((item) => (
                          <div key={item.emailId} className="p-4 flex items-start justify-between gap-3">
                            <div className="min-w-0">
                              <div className="text-xs text-[#9A6700] font-semibold tabular-nums">
                                {item.daysWaiting}일 경과 · 수신 {item.recipientLabel}
                              </div>
                              <div className="text-xs font-semibold text-[#242424] mt-0.5 truncate">
                                {item.subject}
                              </div>
                              <div className="text-[11px] text-[#616161] tabular-nums">발송일 {item.sentDate}</div>
                            </div>
                            <button
                              type="button"
                              onClick={() => handleSaveReminder(item)}
                              className="px-2.5 py-1.5 text-[11px] font-medium border border-[#D1D1D1] hover:bg-[#F5F5F5] text-[#242424] rounded whitespace-nowrap shrink-0"
                            >
                              리마인더 초안
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </section>

                  {/* Commitments -> 1-Click Calendar */}
                  <section className="bg-white border border-[#E0E0E0] rounded overflow-hidden">
                    <div className="px-5 py-3 border-b border-[#EDEBE9]">
                      <h2 className="text-sm font-semibold text-[#242424]">
                        추출된 약속 · 실행 항목 ({brief.commitments.length})
                      </h2>
                    </div>
                    {brief.commitments.length === 0 ? (
                      <div className="p-4 text-xs text-[#616161]">추출된 실행 약속이 없습니다.</div>
                    ) : (
                      <div className="divide-y divide-[#EDEBE9]">
                        {brief.commitments.slice(0, 6).map((cmt) => (
                          <div key={cmt.id} className="p-3.5 flex items-start justify-between gap-2">
                            <div className="min-w-0">
                              <div className="text-xs font-medium text-[#242424]">{cmt.task}</div>
                              <div className="text-[11px] text-[#616161] mt-0.5 tabular-nums">
                                기한 {cmt.deadline} · {cmt.sourceSubject}
                              </div>
                            </div>
                            <button
                              type="button"
                              onClick={() =>
                                handleAddCommitmentToCalendar(cmt.task, cmt.deadline, cmt.sourceSubject)
                              }
                              className="px-2 py-1 text-[11px] font-medium border border-[#D1D1D1] hover:bg-[#EBF3FC] hover:text-[#0F6CBD] rounded whitespace-nowrap shrink-0 flex items-center gap-1"
                            >
                              <CalendarPlus className="w-3.5 h-3.5" />
                              일정
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </section>
                </div>
              </div>
            </>
          )}

          {/* TAB 2: ASK AI INBOX CONVERSATIONAL Q&A (SUPERHUMAN / SHORTWAVE STYLE) */}
          {activeTab === 'ask' && (
            <div className="space-y-4">
              <div className="bg-white border border-[#E0E0E0] rounded p-5 space-y-3">
                <div className="text-sm font-semibold text-[#242424] flex items-center gap-2">
                  <MessageSquare className="w-4 h-4 text-[#0F6CBD]" />
                  Ask AI — 전체 메일함 · 첨부파일 · 일정 자연어 질의
                </div>
                <p className="text-xs text-[#616161]">
                  로드된 실제 메일과 프로젝트·일정 전체에서 근거를 검색하고 출처 메일 번호([1], [2])를 클릭해 원문으로 바로 이동합니다.
                </p>

                <div className="flex items-center gap-2">
                  <div className="relative flex-1">
                    <Search className="w-4 h-4 text-[#616161] absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={askInput}
                      onChange={(e) => setAskInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') void handleRunAsk();
                      }}
                      placeholder="예: 견적 금액이 포함된 영업 메일과 마감일을 모두 찾아줘"
                      aria-label="Ask AI 메일함 질문 입력"
                      className="w-full pl-9 pr-3 py-2 text-xs bg-[#FAF9F8] border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                    />
                  </div>
                  <button
                    type="button"
                    disabled={isAsking || !askInput.trim()}
                    onClick={() => void handleRunAsk()}
                    className="px-4 py-2 bg-[#0F6CBD] hover:bg-[#115EA3] disabled:opacity-50 text-white text-xs font-medium rounded whitespace-nowrap shrink-0 flex items-center gap-1.5"
                  >
                    <Send className="w-3.5 h-3.5" />
                    질의 실행
                  </button>
                </div>

                <div className="flex flex-wrap items-center gap-2 pt-1">
                  {QUICK_ASK_PROMPTS.map((preset) => (
                    <button
                      key={preset}
                      type="button"
                      onClick={() => void handleRunAsk(preset)}
                      className="px-2.5 py-1 text-[11px] text-[#424242] bg-[#F3F2F1] hover:bg-[#EBF3FC] hover:text-[#0F6CBD] rounded transition-colors whitespace-nowrap"
                    >
                      {preset}
                    </button>
                  ))}
                </div>
              </div>

              {askHistory.length === 0 ? (
                <div className="bg-white border border-[#E0E0E0] rounded p-6 text-xs text-[#616161]">
                  위 질문 예시를 클릭하거나 자연어로 메일함에 질문해 보세요.
                </div>
              ) : (
                <div className="space-y-4">
                  {askHistory.map((item, idx) => (
                    <div key={`${item.question}-${idx}`} className="bg-white border border-[#E0E0E0] rounded p-5 space-y-3">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div className="text-xs font-semibold text-[#0F6CBD]">Q. {item.question}</div>
                        {item.creditLabel && (
                          <span
                            className={`text-[10px] font-medium px-2 py-0.5 rounded border whitespace-nowrap ${
                              item.creditRoute === 'agent-builder-rag'
                                ? 'bg-[#EBF3FC] border-[#B4D6FA] text-[#0F6CBD]'
                                : item.creditRoute === 'gemini-primary'
                                  ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                                  : 'bg-[#FAF9F8] border-[#E0E0E0] text-[#616161]'
                            }`}
                          >
                            {item.creditLabel}
                          </span>
                        )}
                      </div>
                      <div className="text-sm text-[#242424] leading-relaxed">{item.answer}</div>
                      {item.fallbackReason && (
                        <div className="text-[11px] text-[#616161] bg-[#FAF9F8] border border-[#EDEBE9] rounded px-2.5 py-1.5">
                          자동 폴백 안내: {item.fallbackReason}
                        </div>
                      )}
                      {item.bulletFindings.length > 0 && (
                        <ul className="space-y-1.5 text-xs text-[#424242] list-disc pl-4">
                          {item.bulletFindings.map((line, i) => (
                            <li key={i}>{line}</li>
                          ))}
                        </ul>
                      )}
                      {item.citations.length > 0 && (
                        <div className="pt-2 border-t border-[#EDEBE9] space-y-1.5">
                          <div className="text-[11px] font-semibold text-[#616161]">인용된 출처 메일</div>
                          <div className="divide-y divide-[#F3F2F1] border border-[#EDEBE9] rounded">
                            {item.citations.map((cit) => (
                              <div
                                key={cit.emailId}
                                className="px-3 py-2 flex items-center justify-between gap-3 text-xs hover:bg-[#FAF9F8]"
                              >
                                <div className="min-w-0">
                                  <span className="font-mono font-semibold text-[#0F6CBD] mr-2">[{cit.index}]</span>
                                  <span className="font-medium text-[#242424]">{cit.subject}</span>
                                  <span className="text-[#616161] ml-2">
                                    {cit.fromName} · <span className="tabular-nums">{cit.date}</span>
                                  </span>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => onOpenMailInWorkbench(cit.emailId)}
                                  className="text-[11px] font-medium text-[#0F6CBD] hover:underline whitespace-nowrap shrink-0 flex items-center gap-0.5"
                                >
                                  메일 열기
                                  <ArrowRight className="w-3 h-3" />
                                </button>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* TAB 3: AUTO-LABEL SPLIT INBOX (SUPERHUMAN / NOTION MAIL) */}
          {activeTab === 'labels' && (
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
              {/* Left 4 cols: Label Rules & Creator */}
              <div className="lg:col-span-4 space-y-4">
                <section className="bg-white border border-[#E0E0E0] rounded overflow-hidden">
                  <div className="px-4 py-3 border-b border-[#EDEBE9] flex items-center justify-between">
                    <h2 className="text-xs font-semibold text-[#242424] flex items-center gap-1.5">
                      <Tag className="w-3.5 h-3.5 text-[#0F6CBD]" />
                      Auto-Label 스플릿 규칙
                    </h2>
                    <span className="text-[11px] text-[#616161] tabular-nums">{labelRules.length}개</span>
                  </div>
                  <div className="divide-y divide-[#EDEBE9]">
                    <button
                      type="button"
                      onClick={() => setSelectedLabelName('ALL')}
                      className={`w-full px-4 py-2.5 text-left text-xs flex items-center justify-between ${
                        selectedLabelName === 'ALL'
                          ? 'bg-[#EBF3FC] text-[#0F6CBD] font-semibold'
                          : 'hover:bg-[#FAF9F8] text-[#242424]'
                      }`}
                    >
                      <span>전체 메일</span>
                      <span className="tabular-nums">{emails.length}</span>
                    </button>
                    {labelRules.map((rule) => {
                      const cnt = brief.labelCounts.find((l) => l.label === rule.name)?.count ?? 0;
                      return (
                        <div
                          key={rule.id}
                          className={`px-4 py-2.5 flex items-center justify-between gap-2 ${
                            selectedLabelName === rule.name ? 'bg-[#EBF3FC]' : 'hover:bg-[#FAF9F8]'
                          }`}
                        >
                          <button
                            type="button"
                            onClick={() => setSelectedLabelName(rule.name)}
                            className="flex-1 min-w-0 text-left"
                          >
                            <div
                              className={`text-xs truncate ${
                                selectedLabelName === rule.name
                                  ? 'text-[#0F6CBD] font-semibold'
                                  : 'text-[#242424] font-medium'
                              }`}
                            >
                              {rule.name}
                            </div>
                            <div className="text-[11px] text-[#616161] truncate">{rule.promptDescription}</div>
                          </button>
                          <span className="text-xs font-mono tabular-nums text-[#424242] shrink-0">{cnt}</span>
                          <input
                            type="checkbox"
                            checked={rule.enabled}
                            aria-label={`${rule.name} 규칙 활성화`}
                            onChange={() =>
                              setLabelRules((prev) =>
                                prev.map((r) => (r.id === rule.id ? { ...r, enabled: !r.enabled } : r))
                              )
                            }
                            className="accent-[#0F6CBD] shrink-0"
                          />
                        </div>
                      );
                    })}
                  </div>
                </section>

                <section className="bg-white border border-[#E0E0E0] rounded p-4 space-y-2.5">
                  <div className="text-xs font-semibold text-[#242424] flex items-center gap-1.5">
                    <Plus className="w-3.5 h-3.5 text-[#0F6CBD]" />
                    새 Auto-Label 규칙 만들기
                  </div>
                  <input
                    type="text"
                    value={newLabelName}
                    onChange={(e) => setNewLabelName(e.target.value)}
                    placeholder="라벨 이름 (예: 선급 검사 · 승인)"
                    className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                  <input
                    type="text"
                    value={newLabelKeywords}
                    onChange={(e) => setNewLabelKeywords(e.target.value)}
                    placeholder="키워드 쉼표 구분 (예: KR, DNV, ABS, 선급, class)"
                    className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                  <input
                    type="text"
                    value={newLabelPrompt}
                    onChange={(e) => setNewLabelPrompt(e.target.value)}
                    placeholder="분류 설명 (선택)"
                    className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                  <button
                    type="button"
                    onClick={handleAddLabelRule}
                    className="w-full py-2 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-medium rounded transition-colors"
                  >
                    스플릿 라벨 추가
                  </button>
                </section>
              </div>

              {/* Right 8 cols: Split Inbox Stream */}
              <div className="lg:col-span-8 bg-white border border-[#E0E0E0] rounded overflow-hidden">
                <div className="px-5 py-3 border-b border-[#EDEBE9] flex items-center justify-between">
                  <div className="text-sm font-semibold text-[#242424] flex items-center gap-2">
                    <Inbox className="w-4 h-4 text-[#0F6CBD]" />
                    {selectedLabelName === 'ALL' ? '전체 메일 스트림' : selectedLabelName}
                  </div>
                  <span className="text-xs text-[#616161] tabular-nums">{splitLabelEmails.length}건</span>
                </div>
                {splitLabelEmails.length === 0 ? (
                  <div className="p-6 text-xs text-[#616161]">해당 라벨에 일치하는 메일이 없습니다.</div>
                ) : (
                  <div className="divide-y divide-[#EDEBE9]">
                    {splitLabelEmails.map((mail) => (
                      <div key={mail.id} className="px-5 py-3.5 hover:bg-[#FAF9F8] flex items-center justify-between gap-4">
                        <div className="min-w-0">
                          <div className="text-xs text-[#616161]">
                            <span>{mail.fromName}</span>
                            <span aria-hidden="true"> · </span>
                            <span className="tabular-nums">{mail.date}</span>
                            {mail.hasAttachment && (
                              <>
                                <span aria-hidden="true"> · </span>
                                <span>첨부 {mail.attachments?.length ?? 1}개</span>
                              </>
                            )}
                          </div>
                          <div className="text-sm font-semibold text-[#242424] truncate mt-0.5">{mail.subject}</div>
                          <div className="text-xs text-[#616161] truncate mt-0.5">{mail.body.slice(0, 110)}</div>
                        </div>
                        <button
                          type="button"
                          onClick={() => onOpenMailInWorkbench(mail.id)}
                          className="px-3 py-1.5 text-xs font-medium border border-[#D1D1D1] hover:bg-[#EBF3FC] hover:text-[#0F6CBD] rounded whitespace-nowrap shrink-0"
                        >
                          읽기·회신
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 4: WRITE WITH AI + COPILOT PRE-SEND COACHING */}
          {activeTab === 'write' && (
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
              {/* Left 7 cols: Bullet Expander + Composer */}
              <div className="lg:col-span-7 bg-white border border-[#E0E0E0] rounded p-5 space-y-4">
                <div className="flex items-center justify-between gap-3">
                  <div>
                    <h2 className="text-sm font-semibold text-[#242424] flex items-center gap-1.5">
                      <FileEdit className="w-4 h-4 text-[#0F6CBD]" />
                      Write with AI — 불릿 메모 → 완성 메일 작성기
                    </h2>
                    <p className="text-xs text-[#616161] mt-0.5">
                      핵심 메모 2~3줄만 입력하면 어조에 맞춘 완성 메일 초안을 생성합니다.
                    </p>
                  </div>
                  <div className="flex items-center gap-1 bg-[#F3F2F1] p-1 rounded shrink-0">
                    {(
                      [
                        { id: 'formal', label: '국문 격식' },
                        { id: 'concise', label: '국문 간결' },
                        { id: 'english', label: 'English' },
                      ] as const
                    ).map((t) => (
                      <button
                        key={t.id}
                        type="button"
                        onClick={() => setWriteTone(t.id)}
                        className={`px-2.5 py-1 text-[11px] font-medium rounded whitespace-nowrap ${
                          writeTone === t.id ? 'bg-white text-[#0F6CBD] shadow-xs' : 'text-[#424242]'
                        }`}
                      >
                        {t.label}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <input
                    type="text"
                    value={writeTo}
                    onChange={(e) => setWriteTo(e.target.value)}
                    placeholder="받는 사람 이메일 (예: buyer@partner.example.test)"
                    className="px-3 py-2 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                  <input
                    type="text"
                    value={writeRecipientName}
                    onChange={(e) => setWriteRecipientName(e.target.value)}
                    placeholder="상대방 호칭 (예: 김팀장 / Sarah)"
                    className="px-3 py-2 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-medium text-[#424242]" htmlFor="write-bullets-input">
                    핵심 요점 불릿 (한 줄에 하나씩 입력)
                  </label>
                  <textarea
                    id="write-bullets-input"
                    rows={3}
                    value={writeBullets}
                    onChange={(e) => setWriteBullets(e.target.value)}
                    placeholder={'- 견적서 USD 48,500 검토 완료, 단가 수락 가능\n- 납기 12월 10일까지 준수 여부 확인 요청\n- 수정 도면 DWG-E1001 첨부'}
                    className="w-full px-3 py-2 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                  <div className="flex items-center justify-between">
                    <label className="inline-flex items-center gap-1.5 text-xs text-[#424242]">
                      <input
                        type="checkbox"
                        checked={writeHasAttachment}
                        onChange={(e) => setWriteHasAttachment(e.target.checked)}
                        className="accent-[#0F6CBD]"
                      />
                      첨부파일 포함 예정
                    </label>
                    <button
                      type="button"
                      onClick={handleExpandBullets}
                      className="px-3.5 py-1.5 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-medium rounded whitespace-nowrap"
                    >
                      불릿 → 완성 초안 생성
                    </button>
                  </div>
                </div>

                <div className="space-y-2 pt-2 border-t border-[#EDEBE9]">
                  <input
                    type="text"
                    value={writeSubject}
                    onChange={(e) => setWriteSubject(e.target.value)}
                    placeholder="메일 제목"
                    className="w-full px-3 py-2 text-xs font-medium border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                  <textarea
                    rows={8}
                    value={writeBody}
                    onChange={(e) => setWriteBody(e.target.value)}
                    placeholder="생성된 초안을 자유롭게 수정하세요..."
                    className="w-full px-3 py-2 text-xs leading-relaxed border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  />
                  <div className="flex justify-end">
                    <button
                      type="button"
                      onClick={handleSaveComposedDraft}
                      className="px-4 py-2 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-medium rounded whitespace-nowrap"
                    >
                      Bmail 임시보관함에 저장
                    </button>
                  </div>
                </div>
              </div>

              {/* Right 5 cols: Real-time Copilot Pre-Send Coaching */}
              <div className="lg:col-span-5 bg-white border border-[#E0E0E0] rounded p-5 space-y-4">
                <div className="flex items-center justify-between">
                  <h2 className="text-sm font-semibold text-[#242424] flex items-center gap-1.5">
                    <ShieldCheck className="w-4 h-4 text-[#0F6CBD]" />
                    발송 전 실시간 톤 · 리스크 코칭
                  </h2>
                  <span
                    className={`text-xs font-semibold ${
                      coaching.readyToSend ? 'text-[#107C10]' : 'text-[#9A6700]'
                    }`}
                  >
                    {coaching.readyToSend ? '발송 준비 완료' : '검토 권장'}
                  </span>
                </div>

                <div className="grid grid-cols-3 divide-x divide-[#EDEBE9] border border-[#EDEBE9] rounded text-center py-2.5">
                  <div>
                    <div className="text-[11px] text-[#616161]">단어 수</div>
                    <div className="text-sm font-semibold tabular-nums text-[#242424]">{coaching.wordCount}</div>
                  </div>
                  <div>
                    <div className="text-[11px] text-[#616161]">예상 읽기</div>
                    <div className="text-sm font-semibold tabular-nums text-[#242424]">{coaching.estimatedReadSec}초</div>
                  </div>
                  <div>
                    <div className="text-[11px] text-[#616161]">어조 진단</div>
                    <div className="text-xs font-semibold text-[#242424] truncate px-1">{coaching.toneSummary}</div>
                  </div>
                </div>

                {coaching.warnings.length === 0 ? (
                  <div className="p-4 bg-[#F3FAF3] border border-[#C6E5C6] rounded text-xs text-[#107C10] flex items-start gap-2">
                    <CheckCircle2 className="w-4 h-4 shrink-0 mt-0.5" />
                    <div>
                      첨부 누락, 민감 정보 노출, 공격적 어조 등 리스크가 감지되지 않았습니다.
                    </div>
                  </div>
                ) : (
                  <div className="divide-y divide-[#EDEBE9] border border-[#EDEBE9] rounded">
                    {coaching.warnings.map((w) => (
                      <div key={w.code} className="p-3.5 space-y-1">
                        <div className="text-xs font-semibold flex items-center gap-1.5 text-[#242424]">
                          <AlertTriangle
                            className={`w-3.5 h-3.5 shrink-0 ${
                              w.severity === 'high'
                                ? 'text-[#C50F1F]'
                                : w.severity === 'medium'
                                  ? 'text-[#9A6700]'
                                  : 'text-[#0F6CBD]'
                            }`}
                          />
                          <span>{w.title}</span>
                        </div>
                        <div className="text-[11px] text-[#616161] pl-5">{w.suggestion}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

interface MorningBriefReportWaitingItem {
  emailId: string;
  subject: string;
  recipientLabel: string;
  sentDate: string;
  daysWaiting: number;
  reminderSubject: string;
  reminderBody: string;
}
