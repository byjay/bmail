import React, { useMemo, useState } from 'react';
import {
  AlertTriangle,
  ChevronDown,
  ChevronUp,
  CornerUpLeft,
  Eye,
  FolderKanban,
  Sparkles,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import type { ProjectMailOverrides } from '../utils/outlookWorkspaceModel';
import {
  buildUnreadPriorityBriefing,
  type UnreadPriorityBriefing,
} from '../services/unreadBriefingEngine';

export interface UnreadBriefingBannerProps {
  emails: MailItem[];
  projects: ProjectDefinition[];
  projectMailOverrides?: ProjectMailOverrides;
  serverBriefing?: UnreadPriorityBriefing | null;
  onOpenMail: (emailId: string) => void;
  onQuickReplyMail: (emailId: string) => void;
  onDismiss?: () => void;
}

export const UnreadBriefingBanner: React.FC<UnreadBriefingBannerProps> = ({
  emails,
  projects,
  projectMailOverrides,
  serverBriefing,
  onOpenMail,
  onQuickReplyMail,
  onDismiss,
}) => {
  const [expanded, setExpanded] = useState(true);
  const [selectedProjectCode, setSelectedProjectCode] = useState<string>('ALL');

  const briefing = useMemo(() => {
    const clientComputed = buildUnreadPriorityBriefing(emails, projects, projectMailOverrides);
    if (serverBriefing && serverBriefing.totalUnread >= clientComputed.totalUnread) {
      return serverBriefing;
    }
    return clientComputed;
  }, [emails, projects, projectMailOverrides, serverBriefing]);

  if (briefing.totalUnread === 0) {
    return null;
  }

  const visibleGroups =
    selectedProjectCode === 'ALL'
      ? briefing.projectGroups.slice(0, 6)
      : briefing.projectGroups.filter((g) => g.projectCode === selectedProjectCode);

  return (
    <section
      aria-label="미확인 메일 읽지않음 상태 프로젝트·중요도 브리핑"
      className="mx-2 mt-1.5 mb-1 rounded-[6px] border border-[#CFE4FA] bg-white shadow-[0_1px_3px_rgba(0,0,0,0.06)] overflow-hidden shrink-0"
    >
      {/* Top Briefing Summary Bar */}
      <div className="px-3 py-1.5 bg-[#F3F8FE] border-b border-[#DCEBFC] flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 min-w-0 flex-wrap">
          <span className="inline-flex items-center gap-1 text-xs font-bold text-[#0F6CBD]">
            <Sparkles className="w-3.5 h-3.5 shrink-0" />
            미확인 메일 브리핑 (안 읽음 유지)
          </span>
          <span className="px-2 py-0.5 rounded-full bg-white border border-[#CFE4FA] text-[11px] font-semibold text-[#242424] tabular-nums">
            미열람 {briefing.totalUnread}건
          </span>
          {briefing.urgentReplyCount > 0 && (
            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-rose-600 text-white text-[11px] font-bold animate-pulse">
              <AlertTriangle className="w-3 h-3" />
              반드시 긴급회신 {briefing.urgentReplyCount}건
            </span>
          )}
          <span className="text-[11px] text-[#424242] hidden md:inline">
            · 중요 {briefing.importantCount}건 · {briefing.projectCount}개 프로젝트별 정리
          </span>
        </div>

        <div className="flex items-center gap-1.5 shrink-0">
          <button
            type="button"
            onClick={() => setExpanded((v) => !v)}
            className="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold text-[#0F6CBD] bg-white hover:bg-[#EBF3FC] border border-[#CFE4FA] rounded-[4px]"
          >
            {expanded ? (
              <>
                <ChevronUp className="w-3.5 h-3.5" />
                접기
              </>
            ) : (
              <>
                <ChevronDown className="w-3.5 h-3.5" />
                전체 브리핑 펼치기
              </>
            )}
          </button>
          {onDismiss && (
            <button
              type="button"
              onClick={onDismiss}
              aria-label="브리핑 배너 닫기"
              className="p-0.5 text-[#616161] hover:text-[#242424] rounded"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>

      {/* Urgent Mandatory Reply Notification Strip (Always visible when urgent replies exist) */}
      {briefing.urgentReplyAlerts.length > 0 && (
        <div className="px-3 py-1.5 bg-rose-50/90 border-b border-rose-200 flex items-center justify-between gap-2 flex-wrap">
          <div className="flex items-center gap-2 min-w-0 flex-1">
            <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
            <div className="text-[11px] text-rose-950 truncate">
              <strong className="font-bold">[긴급회신 필수 알림]</strong>{' '}
              <span className="font-semibold text-rose-700">
                [{briefing.urgentReplyAlerts[0].projectCode}]
              </span>{' '}
              {briefing.urgentReplyAlerts[0].fromName} — {briefing.urgentReplyAlerts[0].subject}
              <span className="text-rose-800/80 hidden lg:inline">
                {' '}
                · {briefing.urgentReplyAlerts[0].summary}
              </span>
            </div>
          </div>
          <div className="flex items-center gap-1 shrink-0">
            <button
              type="button"
              onClick={() => onQuickReplyMail(briefing.urgentReplyAlerts[0].emailId)}
              className="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold text-white bg-rose-600 hover:bg-rose-700 rounded-[4px]"
            >
              <CornerUpLeft className="w-3 h-3" />
              즉시 회신
            </button>
            <button
              type="button"
              onClick={() => onOpenMail(briefing.urgentReplyAlerts[0].emailId)}
              className="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold text-rose-800 bg-white hover:bg-rose-100 border border-rose-300 rounded-[4px]"
            >
              <Eye className="w-3 h-3" />
              열기
            </button>
          </div>
        </div>
      )}

      {/* Expanded Project-by-Project Importance Briefing */}
      {expanded && (
        <div className="p-2.5 space-y-2 max-h-[215px] overflow-y-auto bg-[#FAF9F8]/60">
          {/* Project Filter Pills sorted by Importance */}
          <div className="flex items-center gap-1 flex-wrap">
            <button
              type="button"
              onClick={() => setSelectedProjectCode('ALL')}
              className={`px-2 py-0.5 rounded-full text-[10.5px] font-semibold border transition-colors ${
                selectedProjectCode === 'ALL'
                  ? 'bg-[#0F6CBD] text-white border-[#0F6CBD]'
                  : 'bg-white text-[#424242] border-[#D1D1D1] hover:bg-[#F5F5F5]'
              }`}
            >
              전체 프로젝트 ({briefing.projectGroups.length})
            </button>
            {briefing.projectGroups.slice(0, 8).map((grp) => (
              <button
                key={grp.projectCode}
                type="button"
                onClick={() => setSelectedProjectCode(grp.projectCode)}
                className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10.5px] font-semibold border transition-colors ${
                  selectedProjectCode === grp.projectCode
                    ? 'bg-[#0F6CBD] text-white border-[#0F6CBD]'
                    : grp.urgentReplyCount > 0
                      ? 'bg-rose-50 text-rose-800 border-rose-200 hover:bg-rose-100'
                      : 'bg-white text-[#242424] border-[#D1D1D1] hover:bg-[#F5F5F5]'
                }`}
              >
                <FolderKanban className="w-3 h-3" />
                <span>{grp.projectCode}</span>
                <span className="opacity-80">
                  {grp.urgentReplyCount > 0
                    ? `긴급${grp.urgentReplyCount}/${grp.totalCount}`
                    : `${grp.totalCount}건`}
                </span>
              </button>
            ))}
          </div>

          <div className="grid grid-cols-1 xl:grid-cols-2 gap-2">
            {visibleGroups.map((grp) => (
              <div
                key={grp.projectCode}
                className="rounded-[5px] border border-[#E0E0E0] bg-white p-2 space-y-1.5"
              >
                <div className="flex items-center justify-between gap-2 pb-1 border-b border-[#F0F0F0]">
                  <div className="flex items-center gap-1.5 min-w-0">
                    <span className="px-1.5 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[10px] font-bold">
                      {grp.projectCode}
                    </span>
                    <span className="text-[11px] font-bold text-[#242424] truncate">
                      {grp.projectName}
                    </span>
                  </div>
                  <div className="flex items-center gap-1 shrink-0 text-[10px] font-semibold">
                    {grp.urgentReplyCount > 0 && (
                      <span className="px-1.5 py-0.5 rounded bg-rose-100 text-rose-800">
                        긴급회신 {grp.urgentReplyCount}
                      </span>
                    )}
                    {grp.importantCount > 0 && (
                      <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-800">
                        중요 {grp.importantCount}
                      </span>
                    )}
                    <span className="text-[#616161]">총 {grp.totalCount}</span>
                  </div>
                </div>

                <ul className="space-y-1">
                  {grp.items.slice(0, 3).map((item) => (
                    <li
                      key={item.emailId}
                      className="flex items-start justify-between gap-2 text-[11px] rounded px-1.5 py-1 hover:bg-[#F5F8FD]"
                    >
                      <div className="min-w-0 flex-1 space-y-0.5">
                        <div className="flex items-center gap-1.5 min-w-0">
                          <span
                            className={`px-1.5 py-0.2 rounded text-[9.5px] font-bold shrink-0 ${
                              item.urgencyTier === 'urgent-reply'
                                ? 'bg-rose-600 text-white'
                                : item.urgencyTier === 'important'
                                  ? 'bg-amber-500 text-white'
                                  : 'bg-slate-200 text-slate-700'
                            }`}
                          >
                            {item.urgencyLabel}
                          </span>
                          <span className="font-semibold text-[#242424] truncate">
                            {item.fromName}: {item.subject}
                          </span>
                        </div>
                        <div className="text-[10.5px] text-[#616161] truncate">{item.summary}</div>
                      </div>
                      <div className="flex items-center gap-1 shrink-0 pt-0.5">
                        <button
                          type="button"
                          onClick={() => onQuickReplyMail(item.emailId)}
                          title="원문 읽음 변경 없이 바로 회신 창 열기"
                          className="px-1.5 py-0.5 rounded border border-[#D1D1D1] bg-white hover:bg-[#EBF3FC] text-[10px] font-semibold text-[#0F6CBD]"
                        >
                          회신
                        </button>
                        <button
                          type="button"
                          onClick={() => onOpenMail(item.emailId)}
                          className="px-1.5 py-0.5 rounded border border-[#D1D1D1] bg-white hover:bg-[#F5F5F5] text-[10px] text-[#424242]"
                        >
                          보기
                        </button>
                      </div>
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>
      )}
    </section>
  );
};
