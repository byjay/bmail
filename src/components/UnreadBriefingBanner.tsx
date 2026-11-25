import React, { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
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
  DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS,
  buildUnreadPriorityBriefing,
  type UnreadPriorityBriefing,
} from '../services/unreadBriefingEngine';

export interface UnreadBriefingBannerProps {
  emails: MailItem[];
  projects: ProjectDefinition[];
  projectMailOverrides?: ProjectMailOverrides;
  serverBriefing?: UnreadPriorityBriefing | null;
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
  onOpenMail: (emailId: string) => void;
  onQuickReplyMail: (emailId: string) => void;
  onDismiss?: () => void;
}

export const UnreadBriefingBanner: React.FC<UnreadBriefingBannerProps> = ({
  emails,
  projects,
  projectMailOverrides,
  serverBriefing,
  open,
  onOpenChange,
  onOpenMail,
  onQuickReplyMail,
  onDismiss,
}) => {
  const [internalOpen, setInternalOpen] = useState(false);
  const [selectedProjectCode, setSelectedProjectCode] = useState<string>('ALL');

  const isOpen = open !== undefined ? open : internalOpen;
  const setOpen = (next: boolean) => {
    if (open === undefined) {
      setInternalOpen(next);
    }
    onOpenChange?.(next);
    if (!next && !onOpenChange) {
      onDismiss?.();
    }
  };

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (open === undefined) {
          setInternalOpen(false);
        }
        onOpenChange?.(false);
        if (!onOpenChange) {
          onDismiss?.();
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, open, onOpenChange, onDismiss]);

  const closedSummary = useMemo(() => {
    if (isOpen) return null;
    if (serverBriefing) {
      return {
        totalUnread: serverBriefing.totalUnread,
        urgentReplyCount: serverBriefing.urgentReplyCount,
        maxAgeDays: serverBriefing.maxAgeDays ?? DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS,
      };
    }
    const cutoffMs = Date.now() - DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS * 86_400_000;
    let totalUnread = 0;
    for (const m of emails) {
      if (!m || m.folder !== 'inbox' || m.isRead) {
        continue;
      }
      const ms = Date.parse(String(m.timestamp || m.date || '').trim());
      if (Number.isFinite(ms) && ms >= cutoffMs) {
        totalUnread += 1;
      }
    }
    return {
      totalUnread,
      urgentReplyCount: 0,
      maxAgeDays: DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS,
    };
  }, [isOpen, emails, serverBriefing]);

  const briefing = useMemo(() => {
    if (!isOpen) return null;
    const clientComputed = buildUnreadPriorityBriefing(emails, projects, projectMailOverrides);
    if (serverBriefing && serverBriefing.totalUnread >= clientComputed.totalUnread) {
      return serverBriefing;
    }
    return clientComputed;
  }, [isOpen, emails, projects, projectMailOverrides, serverBriefing]);

  if (!isOpen) {
    if (!closedSummary || closedSummary.totalUnread === 0) {
      return null;
    }
    return (
      <button
        type="button"
        data-testid="briefing-floating-pill"
        onClick={() => setOpen(true)}
        title="최근 14일 미확인 메일 브리핑 팝업 열기"
        className="fixed bottom-14 right-3 z-40 inline-flex items-center gap-2 px-3 py-2 rounded-full bg-white hover:bg-[#F3F8FE] border border-[#C7E0F4] text-[#242424] text-xs font-semibold shadow-lg transition-all"
      >
        <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />
        <span className="tabular-nums">
          미확인 브리핑 (최근 {closedSummary.maxAgeDays}일) · {closedSummary.totalUnread}건
        </span>
        {closedSummary.urgentReplyCount > 0 && (
          <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-full bg-rose-600 text-white text-[10.5px] font-bold tabular-nums">
            긴급 {closedSummary.urgentReplyCount}
          </span>
        )}
      </button>
    );
  }

  if (!briefing) {
    return null;
  }

  const maxAgeDays = briefing.maxAgeDays ?? DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS;
  const visibleGroups =
    selectedProjectCode === 'ALL'
      ? briefing.projectGroups.slice(0, 6)
      : briefing.projectGroups.filter((g) => g.projectCode === selectedProjectCode);

  return (
    <section
      role="dialog"
      aria-label="미확인 메일 최근 14일 프로젝트·중요도 브리핑"
      data-testid="briefing-floating-popup"
      className="fixed bottom-14 right-3 z-50 w-[min(780px,94vw)] max-h-[72vh] rounded-xl border border-[#0F6CBD]/30 bg-white shadow-2xl flex flex-col overflow-hidden"
    >
      {/* Popup Header */}
      <div className="px-3.5 py-2.5 bg-[#F3F8FE] border-b border-[#DCEBFC] flex items-center justify-between gap-2 flex-wrap shrink-0">
        <div className="flex items-center gap-2 min-w-0 flex-wrap">
          <span className="inline-flex items-center gap-1.5 text-xs font-bold text-[#0F6CBD]">
            <Sparkles className="w-3.5 h-3.5 shrink-0" />
            미확인 메일 브리핑 (최근 {maxAgeDays}일)
          </span>
          <span className="px-2 py-0.5 rounded-full bg-white border border-[#CFE4FA] text-[11px] font-semibold text-[#242424] tabular-nums">
            미열람 {briefing.totalUnread}건
          </span>
          {briefing.urgentReplyCount > 0 && (
            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-rose-600 text-white text-[11px] font-bold tabular-nums">
              <AlertTriangle className="w-3 h-3" />
              반드시 긴급회신 {briefing.urgentReplyCount}건
            </span>
          )}
          <span className="text-[11px] text-[#424242] hidden md:inline">
            · 중요 {briefing.importantCount}건 · {briefing.projectCount}개 프로젝트별 정리
          </span>
          {briefing.excludedOldUnreadCount > 0 && (
            <span className="px-1.5 py-0.5 rounded bg-white/80 border border-[#DCEBFC] text-[10.5px] text-[#616161] tabular-nums">
              이전 {briefing.excludedOldUnreadCount}건 제외
            </span>
          )}
        </div>

        <button
          type="button"
          data-testid="briefing-popup-close"
          onClick={() => setOpen(false)}
          aria-label="브리핑 팝업 닫기"
          className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-semibold text-[#424242] bg-white hover:bg-[#EBF3FC] border border-[#C7E0F4] rounded-[5px] shrink-0"
        >
          <X className="w-3.5 h-3.5" />
          <span>닫기</span>
        </button>
      </div>

      {/* Urgent Mandatory Reply Notification Strip */}
      {briefing.urgentReplyAlerts.length > 0 && (
        <div className="px-3.5 py-2 bg-rose-50/90 border-b border-rose-200 flex items-center justify-between gap-2 flex-wrap shrink-0">
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

      {/* Scrollable Project-by-Project Importance Briefing */}
      <div className="p-3 space-y-2.5 max-h-[46vh] overflow-y-auto bg-[#FAF9F8]">
        {briefing.totalUnread === 0 ? (
          <div className="py-8 text-center text-xs text-[#616161] bg-white rounded-lg border border-[#E1DFDD]">
            최근 {maxAgeDays}일 이내 미확인 메일이 없습니다.
            {briefing.excludedOldUnreadCount > 0 && (
              <span className="block mt-1 text-[11px] text-[#757575]">
                ({maxAgeDays}일 이전 과거 미확인 메일 {briefing.excludedOldUnreadCount}건 자동 제외됨)
              </span>
            )}
          </div>
        ) : (
          <>
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
          </>
        )}
      </div>
    </section>
  );
};
