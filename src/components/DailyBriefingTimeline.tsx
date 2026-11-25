import React, { useMemo, useState } from 'react';
import {
  CalendarClock,
  Check,
  CheckCircle2,
  ClipboardCopy,
  Download,
  ListChecks,
  MailPlus,
  Sparkles,
} from 'lucide-react';
import type { AnalysisResult, MailItem } from '../types/mail';
import {
  buildDailyBriefingTimeline,
  formatDailyBriefingMarkdown,
  type DailyBriefingDay,
} from '../utils/dashboardAnalytics';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';

const CHECKED_STORAGE_KEY = 'maillog.dailybriefing.checked.v1';

interface DailyBriefingTimelineProps {
  emails: MailItem[];
  analysisResult: AnalysisResult | null;
  archivedResults?: AnalysisResult[];
  projects?: ProjectDefinition[];
  compact?: boolean;
  onSaveDraftToOutlook?: (subject: string, body: string) => void;
  onOpenPstModal?: () => void;
  onRunAiAnalysis?: () => void;
}

function loadCheckedIds(): Record<string, boolean> {
  try {
    const raw = window.localStorage.getItem(CHECKED_STORAGE_KEY);
    if (!raw) return {};
    const parsed: unknown = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, boolean>) : {};
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    console.warn(`[DailyBriefingTimeline] Failed to load checked state: ${reason}`);
    return {};
  }
}

export const DailyBriefingTimeline: React.FC<DailyBriefingTimelineProps> = ({
  emails,
  analysisResult,
  archivedResults = [],
  projects = [],
  compact = false,
  onSaveDraftToOutlook,
  onOpenPstModal,
  onRunAiAnalysis,
}) => {
  const [selectedProject, setSelectedProject] = useState<string>('ALL');
  const [onlyFollowUps, setOnlyFollowUps] = useState(false);
  const [checkedMap, setCheckedMap] = useState<Record<string, boolean>>(() => loadCheckedIds());
  const [copied, setCopied] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);

  const allDays = useMemo(
    () =>
      buildDailyBriefingTimeline({
        emails,
        analysisResult,
        archivedResults,
        projects,
      }),
    [emails, analysisResult, archivedResults, projects]
  );

  const projectOptions = useMemo(() => {
    const set = new Set<string>();
    for (const day of allDays) {
      day.projectsActive.forEach((p) => set.add(p));
    }
    return Array.from(set).sort();
  }, [allDays]);

  const filteredDays: DailyBriefingDay[] = useMemo(() => {
    return allDays
      .map((day) => {
        const matchProject = (code?: string) =>
          selectedProject === 'ALL' || (code && code.toLowerCase() === selectedProject.toLowerCase());
        const myWorkItems = day.myWorkItems.filter((item) => matchProject(item.projectCode));
        const followUpItems = day.followUpItems.filter((item) => matchProject(item.projectCode));
        return {
          ...day,
          myWorkItems: onlyFollowUps ? [] : myWorkItems,
          followUpItems,
        };
      })
      .filter((day) => day.myWorkItems.length > 0 || day.followUpItems.length > 0);
  }, [allDays, selectedProject, onlyFollowUps]);

  const totalMyWork = useMemo(
    () => filteredDays.reduce((sum, d) => sum + d.myWorkItems.length, 0),
    [filteredDays]
  );
  const totalFollowUps = useMemo(
    () => filteredDays.reduce((sum, d) => sum + d.followUpItems.length, 0),
    [filteredDays]
  );
  const completedFollowUps = useMemo(
    () =>
      filteredDays.reduce(
        (sum, d) => sum + d.followUpItems.filter((item) => Boolean(checkedMap[item.id])).length,
        0
      ),
    [filteredDays, checkedMap]
  );

  const toggleFollowUpDone = (id: string) => {
    setCheckedMap((prev) => {
      const next = { ...prev, [id]: !prev[id] };
      try {
        window.localStorage.setItem(CHECKED_STORAGE_KEY, JSON.stringify(next));
      } catch (err) {
        const reason = err instanceof Error ? err.message : String(err);
        console.warn(`[DailyBriefingTimeline] Failed to persist checked state: ${reason}`);
      }
      return next;
    });
  };

  const handleCopyMarkdown = async () => {
    const md = formatDailyBriefingMarkdown(filteredDays);
    try {
      await navigator.clipboard.writeText(md);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 2000);
    } catch {
      setNotice('클립보드 복사가 차단되었습니다. Markdown 다운로드를 이용해 주세요.');
    }
  };

  const handleDownloadMarkdown = () => {
    const md = formatDailyBriefingMarkdown(filteredDays);
    const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `daily-briefing-${new Date().toISOString().slice(0, 10)}.md`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleSendDraft = () => {
    if (!onSaveDraftToOutlook) return;
    const md = formatDailyBriefingMarkdown(filteredDays);
    onSaveDraftToOutlook(`[일간 업무 브리핑] ${filteredDays[0]?.date || '최신'} 요약`, md);
    setNotice('Outlook 임시 보관함에 일간 업무 브리핑 초안을 저장했습니다.');
    window.setTimeout(() => setNotice(null), 3000);
  };

  if (allDays.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center text-center py-14 px-6 bg-white border border-[#E0E0E0] rounded-lg">
        <CalendarClock className="w-8 h-8 text-[#0F6CBD] mb-3" />
        <h3 className="text-sm font-bold text-[#242424]">일간 업무 브리핑 타임라인이 비어 있습니다</h3>
        <p className="mt-1.5 text-xs text-[#616161] max-w-md leading-relaxed">
          실제 Outlook 원본 <code className="font-mono">.pst/.eml</code> 파일을 연동하거나 AI 에이전트 분석을 실행하면
          날짜별로 <strong>내가 한 일</strong>과 <strong>후속 조치</strong>가 자동으로 타임라인에 정리됩니다.
        </p>
        <div className="mt-4 flex flex-wrap items-center justify-center gap-2">
          {onOpenPstModal && (
            <button
              type="button"
              onClick={onOpenPstModal}
              className="px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-md transition-colors"
            >
              원본 PST · EML 불러오기
            </button>
          )}
          {onRunAiAnalysis && emails.length > 0 && (
            <button
              type="button"
              onClick={onRunAiAnalysis}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] rounded-md transition-colors"
            >
              <Sparkles className="w-3.5 h-3.5" />
              AI 업무 분석 실행
            </button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className={compact ? 'space-y-4' : 'space-y-5'}>
      {/* Briefing Control & Summary Header */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-[#E0E0E0]">
        <div className="space-y-0.5">
          <h3 className="text-sm font-bold text-[#242424] flex items-center gap-2">
            <ListChecks className="w-4 h-4 text-[#0F6CBD]" />
            일간 업무 브리핑 타임라인
          </h3>
          <div className="flex flex-wrap items-center gap-2 text-xs text-[#616161] tabular-nums">
            <span>활동 일수 {filteredDays.length}일</span>
            <span aria-hidden="true">·</span>
            <span>내가 한 일 {totalMyWork}건</span>
            <span aria-hidden="true">·</span>
            <span>
              후속 조치 {completedFollowUps}/{totalFollowUps}건 완료
            </span>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {projectOptions.length > 0 && (
            <select
              aria-label="프로젝트별 브리핑 필터"
              value={selectedProject}
              onChange={(e) => setSelectedProject(e.target.value)}
              className="h-8 px-2.5 text-xs bg-white border border-[#D1D1D1] rounded-md text-[#242424]"
            >
              <option value="ALL">전체 프로젝트</option>
              {projectOptions.map((code) => (
                <option key={code} value={code}>
                  {code}
                </option>
              ))}
            </select>
          )}

          <div className="inline-flex items-center p-0.5 bg-[#F0F0F0] rounded-md">
            <button
              type="button"
              onClick={() => setOnlyFollowUps(false)}
              className={`px-2.5 py-1 text-xs font-medium rounded transition-colors ${
                !onlyFollowUps ? 'bg-white text-[#242424] shadow-xs' : 'text-[#616161]'
              }`}
            >
              전체 보기
            </button>
            <button
              type="button"
              onClick={() => setOnlyFollowUps(true)}
              className={`px-2.5 py-1 text-xs font-medium rounded transition-colors ${
                onlyFollowUps ? 'bg-white text-[#242424] shadow-xs' : 'text-[#616161]'
              }`}
            >
              후속 조치만
            </button>
          </div>

          <button
            type="button"
            onClick={handleCopyMarkdown}
            className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-md transition-colors"
          >
            {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <ClipboardCopy className="w-3.5 h-3.5" />}
            {copied ? '복사됨' : '복사'}
          </button>

          <button
            type="button"
            onClick={handleDownloadMarkdown}
            className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-md transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            .md
          </button>

          {onSaveDraftToOutlook && (
            <button
              type="button"
              onClick={handleSendDraft}
              className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-md transition-colors"
            >
              <MailPlus className="w-3.5 h-3.5" />
              Outlook 초안
            </button>
          )}
        </div>
      </div>

      {notice && (
        <div className="px-3.5 py-2 rounded-md bg-emerald-50 border border-emerald-200 text-xs text-emerald-800">
          {notice}
        </div>
      )}

      {/* Chronological Timeline */}
      <ol className="relative border-l-2 border-[#D6E6F7] ml-3 space-y-5" aria-label="날짜별 업무 브리핑 타임라인">
        {filteredDays.map((day) => (
          <li key={day.date} className="pl-5 relative">
            <span
              aria-hidden="true"
              className="absolute -left-[7px] top-1.5 w-3 h-3 rounded-full bg-[#0F6CBD] ring-4 ring-white"
            />
            <div className="bg-white border border-[#E0E0E0] rounded-lg overflow-hidden">
              {/* Day Header (Zero-Pill clean typographic metadata) */}
              <div className="px-4 py-2.5 bg-[#FAF9F8] border-b border-[#EDEBE9] flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <time dateTime={day.date} className="font-mono text-sm font-bold text-[#242424] tabular-nums">
                    {day.date}
                  </time>
                  {day.weekdayLabel && (
                    <span className="text-xs font-medium text-[#616161]">({day.weekdayLabel})</span>
                  )}
                  {day.projectsActive.length > 0 && (
                    <>
                      <span aria-hidden="true" className="text-[#BDBDBD]">·</span>
                      <span className="font-mono text-xs font-semibold text-[#0F6CBD]">
                        {day.projectsActive.join(', ')}
                      </span>
                    </>
                  )}
                </div>
                <div className="flex items-center gap-2 text-xs text-[#616161] font-mono tabular-nums">
                  <span>수신 {day.receivedCount}</span>
                  <span aria-hidden="true">/</span>
                  <span>발신 {day.sentCount}</span>
                  {day.attachmentCount > 0 && (
                    <>
                      <span aria-hidden="true">·</span>
                      <span>첨부 {day.attachmentCount}</span>
                    </>
                  )}
                </div>
              </div>

              {/* Two-Column Briefing Grid: What I Did vs Follow-up Actions */}
              <div className="grid grid-cols-1 md:grid-cols-2 divide-y md:divide-y-0 md:divide-x divide-[#EDEBE9]">
                {/* Left: 내가 한 일 */}
                <div className="p-4 space-y-2">
                  <div className="flex items-center justify-between text-xs font-semibold text-[#242424]">
                    <span>내가 한 일</span>
                    <span className="font-mono text-[#0F6CBD] tabular-nums">{day.myWorkItems.length}건</span>
                  </div>
                  {day.myWorkItems.length === 0 ? (
                    <p className="text-xs text-[#8A8886]">해당 일자에 기록된 발신·수행 항목이 없습니다.</p>
                  ) : (
                    <ul className="space-y-2">
                      {day.myWorkItems.map((item) => (
                        <li key={item.id} className="flex items-start gap-2 text-xs text-[#242424] leading-relaxed">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0 mt-0.5" />
                          <div className="min-w-0">
                            <span>{item.text}</span>
                            <div className="text-[11px] text-[#707070] mt-0.5">
                              {item.source === 'ai'
                                ? 'AI 분석 추출'
                                : item.source === 'project-summary'
                                  ? '프로젝트 요약'
                                  : '발신 메일 근거'}
                              {item.projectCode ? ` · ${item.projectCode}` : ''}
                            </div>
                          </div>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>

                {/* Right: 후속 조치 */}
                <div className="p-4 space-y-2 bg-[#FCFDFD]">
                  <div className="flex items-center justify-between text-xs font-semibold text-[#242424]">
                    <span>후속 조치</span>
                    <span className="font-mono text-amber-700 tabular-nums">{day.followUpItems.length}건</span>
                  </div>
                  {day.followUpItems.length === 0 ? (
                    <p className="text-xs text-[#8A8886]">대기 중인 후속 조치가 없습니다.</p>
                  ) : (
                    <ul className="space-y-2">
                      {day.followUpItems.map((item) => {
                        const isDone = Boolean(checkedMap[item.id]);
                        return (
                          <li key={item.id} className="flex items-start gap-2 text-xs leading-relaxed">
                            <input
                              type="checkbox"
                              aria-label={`후속 조치 완료: ${item.text}`}
                              checked={isDone}
                              onChange={() => toggleFollowUpDone(item.id)}
                              className="mt-0.5 rounded border-[#BDBDBD] text-[#0F6CBD] focus:ring-[#0F6CBD]"
                            />
                            <div className="min-w-0">
                              <span className={isDone ? 'line-through text-[#8A8886]' : 'text-[#242424]'}>
                                {item.text}
                              </span>
                              <div className="text-[11px] text-[#707070] mt-0.5">
                                <span
                                  className={
                                    item.priority === '높음'
                                      ? 'font-semibold text-red-700'
                                      : 'text-[#707070]'
                                  }
                                >
                                  {item.priority ? `중요도 ${item.priority}` : '확인 필요'}
                                </span>
                                {item.projectCode ? ` · ${item.projectCode}` : ''}
                              </div>
                            </div>
                          </li>
                        );
                      })}
                    </ul>
                  )}
                </div>
              </div>
            </div>
          </li>
        ))}
      </ol>
    </div>
  );
};
