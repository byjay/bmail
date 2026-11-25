import React, { useMemo } from 'react';
import {
  AlertCircle,
  Calendar,
  CheckCircle2,
  Cpu,
  Download,
  FileText,
  Filter,
  Loader2,
  Pin,
  PinOff,
  Play,
  Shield,
  Sparkles,
  X,
} from 'lucide-react';
import type {
  AgentExecutionStep,
  AiProviderId,
  AnalysisOptions,
  BlogTone,
  DateMode,
  MailItem,
  MailScope,
  OutlookHostInfo,
  ProviderStatus,
} from '../types/mail';
import { getAvailableDates, validateAnalysisFilter } from '../utils/mailFilter';

interface OutlookTaskpaneProps {
  emails: MailItem[];
  options: AnalysisOptions;
  providers: ProviderStatus[];
  isAnalyzing: boolean;
  agentSteps: AgentExecutionStep[];
  error: string | null;
  hostInfo: OutlookHostInfo;
  standalone?: boolean;
  isPinned?: boolean;
  supportsPinning?: boolean;
  accentColor?: string;
  onTogglePin?: () => void;
  onClose?: () => void;
  onChangeOptions: (patch: Partial<AnalysisOptions>) => void;
  onRunAnalysis: (overrideProvider?: AiProviderId) => void;
  onCaptureCurrentOutlookItem?: () => void;
}

const BLOG_TONE_OPTIONS: Array<{ id: BlogTone; label: string; desc: string }> = [
  { id: 'insight', label: '비즈니스 인사이트', desc: '의사결정·핵심 성과 중심 요약' },
  { id: 'tech-dev', label: '기술·엔지니어링', desc: '도면·스펙·트러블슈팅 정리' },
  { id: 'essay', label: '업무 회고 에세이', desc: '하루 흐름과 배운 점 서술' },
  { id: 'newsletter', label: '팀 주간 뉴스레터', desc: '부서 공유용 스캔형 포맷' },
];

const FLUENT_INPUT_CLASS =
  'w-full px-2.5 py-1.5 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] focus:outline-none focus:border-[#D1D1D1] focus:border-b-2 focus:border-b-[#0F6CBD] transition-colors duration-150';

export const OutlookTaskpane: React.FC<OutlookTaskpaneProps> = ({
  emails,
  options,
  providers,
  isAnalyzing,
  agentSteps,
  error,
  hostInfo,
  standalone = false,
  isPinned = true,
  supportsPinning = true,
  accentColor = '#0F6CBD',
  onTogglePin,
  onClose,
  onChangeOptions,
  onRunAnalysis,
  onCaptureCurrentOutlookItem,
}) => {
  const filterStatus = useMemo(
    () => validateAnalysisFilter(emails, options),
    [emails, options]
  );

  const availableDates = useMemo(() => getAvailableDates(emails), [emails]);

  const activeProvider = useMemo(
    () => providers.find((p) => p.id === options.provider) ?? providers[0] ?? null,
    [providers, options.provider]
  );

  const configuredCount = useMemo(
    () => providers.filter((p) => p.configured && p.id !== 'builtin').length,
    [providers]
  );

  const handleDateModeChange = (mode: DateMode) => {
    onChangeOptions({ dateMode: mode });
  };

  const handleScopeChange = (scope: MailScope) => {
    onChangeOptions({ scope });
  };

  return (
    <aside
      aria-label="Outlook 우측 플러그인 작업 창"
      className={`${
        standalone
          ? 'w-full max-w-[440px] mx-auto bg-white border-x border-[#E0E0E0] flex flex-col min-h-0 h-full overflow-hidden'
          : 'w-full lg:w-[344px] xl:w-[368px] shrink-0 bg-white border border-[#E0E0E0] rounded-[8px] shadow-[0_1px_2px_rgba(0,0,0,0.06),0_0_2px_rgba(0,0,0,0.04)] flex flex-col min-h-0 overflow-hidden'
      }`}
    >
      {/* Fluent 2 Taskpane Header (48px) */}
      <div className="h-12 px-3.5 bg-[#FAF9F8] border-b border-[#EDEBE9] flex items-center justify-between gap-2 shrink-0 select-none">
        <div className="flex items-center gap-2.5 min-w-0">
          <img
            src="/assets/icon-32.png"
            alt="Bmail AI 작업창"
            width={22}
            height={22}
            className="w-[22px] h-[22px] rounded-[4px] shrink-0"
          />
          <div className="min-w-0">
            <div className="flex items-center gap-1.5">
              <h2 className="text-[13px] font-semibold text-[#242424] tracking-[-0.01em] truncate">
                Bmail AI 작업창
              </h2>
              {isPinned && supportsPinning && (
                <span className="text-[10px] font-medium text-[#0F6CBD] shrink-0">
                  · 고정됨
                </span>
              )}
            </div>
            <p className="text-[11px] text-[#616161] leading-tight truncate">
              {hostInfo.isInsideOutlook
                ? `${hostInfo.hostName} · 연결됨`
                : 'Outlook 리본 추가 기능 · SupportsPinning'}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-0.5 shrink-0">
          {supportsPinning && onTogglePin && (
            <button
              type="button"
              onClick={onTogglePin}
              aria-pressed={isPinned}
              title={
                isPinned
                  ? '작업 창 고정 해제 (SupportsPinning)'
                  : '작업 창 항상 고정 (SupportsPinning)'
              }
              aria-label={isPinned ? '작업 창 고정 해제' : '작업 창 고정'}
              className={`w-8 h-8 flex items-center justify-center rounded-[4px] border transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                isPinned
                  ? 'bg-[#EBF3FC] text-[#0F6CBD] border-[#B4D6FA]'
                  : 'bg-transparent border-transparent text-[#616161] hover:bg-black/[0.05] hover:text-[#242424] active:bg-black/[0.08]'
              }`}
            >
              {isPinned ? (
                <Pin className="w-3.5 h-3.5" strokeWidth={1.75} />
              ) : (
                <PinOff className="w-3.5 h-3.5" strokeWidth={1.75} />
              )}
            </button>
          )}
          {!standalone && onClose && (
            <button
              type="button"
              onClick={onClose}
              title="작업 창 닫기"
              aria-label="작업 창 닫기"
              className="w-8 h-8 flex items-center justify-center rounded-[4px] text-[#616161] hover:bg-black/[0.05] hover:text-[#242424] active:bg-black/[0.08] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
            >
              <X className="w-4 h-4" strokeWidth={1.75} />
            </button>
          )}
        </div>
      </div>

      {/* Scrollable Fluent 2 Taskpane Body */}
      <div className="flex-1 overflow-y-auto px-3.5 py-3.5 space-y-4">
        {/* Live Office.js Mail Capture Banner */}
        {hostInfo.isInsideOutlook && onCaptureCurrentOutlookItem && (
          <div className="rounded-[6px] border border-[#B4D6FA] bg-[#EBF3FC]/70 p-2.5 space-y-2">
            <div className="flex items-center justify-between text-[11px] gap-2">
              <span className="font-semibold text-[#0F6CBD] shrink-0">Outlook 실시간 연결</span>
              <span className="text-[#424242] truncate">
                {hostInfo.userEmail || hostInfo.currentItemSubject || hostInfo.hostName}
              </span>
            </div>
            <button
              type="button"
              onClick={onCaptureCurrentOutlookItem}
              className="w-full py-1.5 px-2.5 text-xs font-medium text-[#0F6CBD] bg-white hover:bg-[#FAF9F8] active:bg-[#F0F0F0] border border-[#B4D6FA] rounded-[4px] inline-flex items-center justify-center gap-1.5 transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
            >
              <Download className="w-3.5 h-3.5" strokeWidth={1.75} />
              현재 열려있는 Outlook 메일 가져오기
            </button>
          </div>
        )}

        {/* Section 1: Date Selection Mode & Date Inputs */}
        <section className="space-y-2.5">
          <div className="flex items-center justify-between">
            <span className="inline-flex items-center gap-1.5 text-[12px] font-semibold text-[#242424]">
              <Calendar className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
              분석 기간
            </span>
            {availableDates.length > 0 && (
              <div className="inline-flex items-center gap-1">
                <button
                  type="button"
                  onClick={() =>
                    onChangeOptions({
                      dateMode: 'single',
                      singleDate: availableDates[0].date,
                    })
                  }
                  className="px-1.5 py-0.5 text-[11px] font-medium text-[#0F6CBD] hover:bg-[#EBF3FC] active:bg-[#DCEBFA] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/50"
                >
                  최근일
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onChangeOptions({
                      dateMode: 'range',
                      startDate: availableDates[availableDates.length - 1].date,
                      endDate: availableDates[0].date,
                    })
                  }
                  className="px-1.5 py-0.5 text-[11px] font-medium text-[#616161] hover:bg-black/[0.05] hover:text-[#242424] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/50"
                >
                  전체
                </button>
              </div>
            )}
          </div>

          {/* Fluent 2 Segmented Control: Single vs Range */}
          <div
            role="group"
            aria-label="날짜 선택 방식"
            className="grid grid-cols-2 gap-0.5 p-0.5 bg-[#F0F0F0] rounded-[6px]"
          >
            {(
              [
                { id: 'single', label: '특정 하루' },
                { id: 'range', label: '기간 범위' },
              ] as const
            ).map((item) => {
              const active = options.dateMode === item.id;
              return (
                <button
                  key={item.id}
                  type="button"
                  aria-pressed={active}
                  onClick={() => handleDateModeChange(item.id)}
                  className={`py-1.5 text-[11px] rounded-[4px] transition-all duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                    active
                      ? 'bg-white text-[#0F6CBD] font-semibold shadow-[0_1px_2px_rgba(0,0,0,0.08)]'
                      : 'text-[#424242] hover:text-[#242424] hover:bg-black/[0.03]'
                  }`}
                >
                  {item.label}
                </button>
              );
            })}
          </div>

          {options.dateMode === 'single' ? (
            <div>
              <label htmlFor="tp-single-date" className="block text-[11px] text-[#616161] mb-1">
                기준 일자
              </label>
              <input
                id="tp-single-date"
                type="date"
                value={options.singleDate}
                onChange={(e) => onChangeOptions({ singleDate: e.target.value })}
                className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
              />
            </div>
          ) : (
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label htmlFor="tp-start-date" className="block text-[11px] text-[#616161] mb-1">
                  시작일
                </label>
                <input
                  id="tp-start-date"
                  type="date"
                  value={options.startDate}
                  onChange={(e) => onChangeOptions({ startDate: e.target.value })}
                  className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                />
              </div>
              <div>
                <label htmlFor="tp-end-date" className="block text-[11px] text-[#616161] mb-1">
                  종료일
                </label>
                <input
                  id="tp-end-date"
                  type="date"
                  value={options.endDate}
                  onChange={(e) => onChangeOptions({ endDate: e.target.value })}
                  className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                />
              </div>
            </div>
          )}

          {availableDates.length > 0 && (
            <div className="flex flex-wrap gap-1 pt-0.5">
              {availableDates.slice(0, 4).map((d) => {
                const isSelected =
                  options.dateMode === 'single' && options.singleDate === d.date;
                return (
                  <button
                    key={d.date}
                    type="button"
                    onClick={() =>
                      onChangeOptions({ dateMode: 'single', singleDate: d.date })
                    }
                    className={`px-2 py-0.5 text-[10px] font-mono tabular-nums rounded-[4px] border transition-colors duration-150 ${
                      isSelected
                        ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-semibold'
                        : 'bg-white border-[#E0E0E0] text-[#424242] hover:bg-[#FAF9F8]'
                    }`}
                  >
                    {d.date.slice(5)} ({d.total})
                  </button>
                );
              })}
            </div>
          )}
        </section>

        {/* Section 2: Mail Scope & Target Count */}
        <section className="space-y-2.5 pt-3 border-t border-[#EDEBE9]">
          <div className="flex items-center justify-between">
            <span className="inline-flex items-center gap-1.5 text-[12px] font-semibold text-[#242424]">
              <Filter className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
              분석 대상 메일
            </span>
            <span className="text-[11px] font-mono tabular-nums text-[#616161]">
              <strong className="text-[#0F6CBD] font-semibold">
                {filterStatus.matchedEmails.length}
              </strong>{' '}
              / {emails.length}건
            </span>
          </div>

          <div
            role="group"
            aria-label="메일함 범위 선택"
            className="grid grid-cols-2 gap-1 p-0.5 bg-[#F0F0F0] rounded-[6px]"
          >
            {(
              [
                { id: 'all', label: '받은·보낸 전체' },
                { id: 'received', label: '받은 메일만' },
                { id: 'sent', label: '보낸 메일만' },
                {
                  id: 'selected',
                  label: `체크 선택 (${options.selectedIds.length})`,
                },
              ] as const
            ).map((item) => {
              const active = options.scope === item.id;
              return (
                <button
                  key={item.id}
                  type="button"
                  aria-pressed={active}
                  onClick={() => handleScopeChange(item.id)}
                  className={`py-1.5 px-2 text-[11px] rounded-[4px] transition-all duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                    active
                      ? 'bg-white text-[#0F6CBD] font-semibold shadow-[0_1px_2px_rgba(0,0,0,0.08)]'
                      : 'text-[#424242] hover:text-[#242424] hover:bg-black/[0.03]'
                  }`}
                >
                  {item.label}
                </button>
              );
            })}
          </div>

          {/* Live matched summary bar */}
          <div className="px-2.5 py-2 rounded-[6px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center justify-between text-[11px]">
            <span className="text-[#424242] truncate">{filterStatus.periodLabel}</span>
            <span className="font-mono tabular-nums text-[#616161] shrink-0">
              수신 {filterStatus.receivedCount} · 발신 {filterStatus.sentCount}
            </span>
          </div>
        </section>

        {/* Section 3: AI Engine Selection */}
        <section className="space-y-2 pt-3 border-t border-[#EDEBE9]">
          <div className="flex items-center justify-between">
            <span className="inline-flex items-center gap-1.5 text-[12px] font-semibold text-[#242424]">
              <Cpu className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
              AI 분석 엔진
            </span>
            <span className="text-[11px] text-[#616161] tabular-nums">
              {configuredCount > 0 ? `${configuredCount}개 API 연결됨` : '로컬 엔진 대기'}
            </span>
          </div>

          <div className="space-y-1.5" role="radiogroup" aria-label="AI 분석 엔진 선택">
            {providers.map((provider) => {
              const selected = options.provider === provider.id;
              const isReady = provider.configured || provider.id === 'builtin';
              return (
                <label
                  key={provider.id}
                  className={`flex items-start gap-2.5 px-2.5 py-2 rounded-[6px] border cursor-pointer transition-colors duration-150 ${
                    selected
                      ? 'bg-[#EBF3FC]/75 border-[#0F6CBD]'
                      : 'bg-white border-[#E0E0E0] hover:bg-[#FAF9F8] hover:border-[#BDBDBD]'
                  }`}
                >
                  <input
                    type="radio"
                    name="taskpane-ai-provider"
                    value={provider.id}
                    checked={selected}
                    onChange={() => onChangeOptions({ provider: provider.id })}
                    className="mt-0.5 accent-[#0F6CBD] shrink-0"
                  />
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center justify-between gap-2">
                      <span
                        className={`text-xs truncate ${
                          selected ? 'font-semibold text-[#0F6CBD]' : 'font-medium text-[#242424]'
                        }`}
                      >
                        {provider.name}
                      </span>
                      <span
                        className={`inline-flex items-center gap-1 text-[10px] font-medium shrink-0 ${
                          isReady ? 'text-[#107C10]' : 'text-[#616161]'
                        }`}
                      >
                        <span
                          aria-hidden="true"
                          className={`w-1.5 h-1.5 rounded-full ${
                            isReady ? 'bg-[#107C10]' : 'bg-[#BDBDBD]'
                          }`}
                        />
                        {provider.id === 'builtin'
                          ? '기본 활성'
                          : isReady
                            ? '연결됨'
                            : '키 필요'}
                      </span>
                    </div>
                    <p className="text-[11px] text-[#616161] mt-0.5 leading-snug truncate">
                      {provider.model} · {provider.tierLabel}
                    </p>
                  </div>
                </label>
              );
            })}
          </div>
        </section>

        {/* Section 4: Blog Tone, Focus & Privacy Options */}
        <section className="space-y-2.5 pt-3 border-t border-[#EDEBE9]">
          <span className="inline-flex items-center gap-1.5 text-[12px] font-semibold text-[#242424]">
            <FileText className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
            일지 · 블로그 작성 옵션
          </span>

          <div>
            <label htmlFor="tp-blog-tone" className="block text-[11px] text-[#616161] mb-1">
              문체 스타일
            </label>
            <select
              id="tp-blog-tone"
              value={options.blogTone}
              onChange={(e) =>
                onChangeOptions({ blogTone: e.target.value as BlogTone })
              }
              className={FLUENT_INPUT_CLASS}
            >
              {BLOG_TONE_OPTIONS.map((tone) => (
                <option key={tone.id} value={tone.id}>
                  {tone.label} — {tone.desc}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label htmlFor="tp-custom-focus" className="block text-[11px] text-[#616161] mb-1">
              중점 분석 지시 (선택)
            </label>
            <input
              id="tp-custom-focus"
              type="text"
              value={options.customFocus}
              onChange={(e) => onChangeOptions({ customFocus: e.target.value })}
              placeholder="예: 선주 코멘트·도면 납기 일정 위주로 정리"
              className={FLUENT_INPUT_CLASS}
            />
          </div>

          <div className="space-y-1">
            <label className="flex items-center gap-2 px-2 py-1.5 rounded-[4px] hover:bg-[#FAF9F8] cursor-pointer transition-colors duration-150">
              <input
                type="checkbox"
                checked={options.includeActionItems}
                onChange={(e) =>
                  onChangeOptions({ includeActionItems: e.target.checked })
                }
                className="accent-[#0F6CBD] w-3.5 h-3.5 rounded-[3px]"
              />
              <span className="text-xs text-[#242424]">후속 조치(Action Items) 자동 추출</span>
            </label>

            <label className="flex items-start gap-2 px-2 py-1.5 rounded-[4px] hover:bg-[#FAF9F8] cursor-pointer transition-colors duration-150">
              <input
                type="checkbox"
                checked={options.maskSensitiveInfo}
                onChange={(e) =>
                  onChangeOptions({ maskSensitiveInfo: e.target.checked })
                }
                className="mt-0.5 accent-[#0F6CBD] w-3.5 h-3.5 rounded-[3px] shrink-0"
              />
              <span className="min-w-0">
                <span className="inline-flex items-center gap-1 text-xs text-[#242424]">
                  <Shield className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
                  민감 정보 자동 마스킹
                </span>
                <span className="block text-[10px] text-[#616161] leading-snug">
                  외부 AI 전송 전 연락처·금액·개인정보 비식별화
                </span>
              </span>
            </label>
          </div>
        </section>

        {/* Agent Step Progress */}
        {agentSteps.length > 0 && (
          <section className="rounded-[6px] border border-[#EDEBE9] bg-[#FAF9F8] p-2.5 space-y-1.5">
            <div className="text-[11px] font-semibold text-[#424242]">AI 에이전트 진행 상태</div>
            <ul className="space-y-1.5">
              {agentSteps.map((step) => (
                <li
                  key={step.id}
                  className="flex items-start gap-2 text-[11px] text-[#424242]"
                >
                  {step.status === 'completed' && (
                    <CheckCircle2
                      className="w-3.5 h-3.5 text-[#107C10] shrink-0 mt-0.5"
                      strokeWidth={1.75}
                    />
                  )}
                  {step.status === 'running' && (
                    <Loader2
                      className="w-3.5 h-3.5 text-[#0F6CBD] animate-spin shrink-0 mt-0.5"
                      strokeWidth={1.75}
                    />
                  )}
                  {step.status === 'error' && (
                    <AlertCircle
                      className="w-3.5 h-3.5 text-[#C50F1F] shrink-0 mt-0.5"
                      strokeWidth={1.75}
                    />
                  )}
                  {(step.status === 'pending' || step.status === 'skipped') && (
                    <span
                      aria-hidden="true"
                      className="w-3.5 h-3.5 flex items-center justify-center shrink-0 mt-0.5"
                    >
                      <span className="w-1.5 h-1.5 rounded-full bg-[#BDBDBD]" />
                    </span>
                  )}
                  <div className="min-w-0">
                    <span className="font-medium text-[#242424]">{step.title}</span>
                    {step.detail && (
                      <span className="block text-[10px] text-[#616161] leading-snug">
                        {step.detail}
                      </span>
                    )}
                  </div>
                </li>
              ))}
            </ul>
          </section>
        )}

        {/* Filter Warning or Runtime Error (Fluent 2 MessageBar) */}
        {(error || (!filterStatus.valid && filterStatus.error)) && (
          <div
            role="alert"
            className={`rounded-[6px] border px-3 py-2 text-xs flex items-start gap-2 ${
              error
                ? 'border-[#F1BBBC] bg-[#FDE7E9] text-[#A4262C]'
                : 'border-[#EDEBE9] bg-[#FAF9F8] text-[#616161]'
            }`}
          >
            <AlertCircle
              className={`w-4 h-4 shrink-0 mt-0.5 ${
                error ? 'text-[#C50F1F]' : 'text-[#616161]'
              }`}
              strokeWidth={1.75}
            />
            <span className="leading-relaxed break-words">
              {error || filterStatus.error}
            </span>
          </div>
        )}
      </div>

      {/* Fluent 2 Sticky Action Footer */}
      <div className="px-3.5 py-3 bg-[#FAF9F8] border-t border-[#EDEBE9] shrink-0 space-y-2">
        <button
          type="button"
          disabled={isAnalyzing || !filterStatus.valid}
          onClick={() => onRunAnalysis()}
          style={
            !isAnalyzing && filterStatus.valid
              ? { backgroundColor: accentColor }
              : undefined
          }
          className={`w-full h-9 px-3 rounded-[4px] text-xs font-semibold inline-flex items-center justify-center gap-2 transition-all duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-1 focus-visible:ring-[#0F6CBD] ${
            isAnalyzing || !filterStatus.valid
              ? 'bg-[#F0F0F0] text-[#A19F9D] border border-[#E0E0E0] cursor-not-allowed'
              : 'text-white hover:brightness-95 active:brightness-90 shadow-[0_1px_2px_rgba(0,0,0,0.12)]'
          }`}
        >
          {isAnalyzing ? (
            <>
              <Loader2 className="w-4 h-4 animate-spin" strokeWidth={1.75} />
              AI 분석 진행 중...
            </>
          ) : (
            <>
              <Sparkles className="w-4 h-4" strokeWidth={1.75} />
              <span>
                {filterStatus.valid
                  ? `${filterStatus.matchedEmails.length}건 메일 일지·블로그 생성`
                  : '분석할 메일이 없습니다'}
              </span>
            </>
          )}
        </button>

        {!isAnalyzing && filterStatus.valid && activeProvider?.id !== 'builtin' && (
          <button
            type="button"
            onClick={() => onRunAnalysis('builtin')}
            className="w-full h-8 px-3 rounded-[4px] text-[11px] font-medium text-[#424242] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] active:bg-[#EBEBEB] inline-flex items-center justify-center gap-1.5 transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
          >
            <Play className="w-3 h-3 text-[#0F6CBD]" strokeWidth={1.75} />
            API 없이 로컬 엔진으로 즉시 생성
          </button>
        )}
      </div>
    </aside>
  );
};
