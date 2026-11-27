import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertOctagon,
  AlertTriangle,
  ArrowLeft,
  BarChart3,
  CheckCircle2,
  ExternalLink,
  Info,
  Loader2,
  PanelRight,
  Paperclip,
  Pause,
  Pin,
  Play,
  RefreshCw,
  ShieldCheck,
  Sparkles,
} from 'lucide-react';
import type { MailItem } from '../types/mail';
import {
  verifyOutlookItemSafety,
  type OutlookActiveItemSnapshot,
  type OutlookCompanionMode,
  type OutlookCompanionPollResult,
  type OutlookVerificationReport,
} from '../utils/outlookCompanionVerifier';

export interface OutlookCompanionPanelProps {
  fallbackEmail?: MailItem | null;
  onSwitchToWorkbench: () => void;
  onSwitchToDashboard: () => void;
  onSendReceiveRefresh?: () => void;
  onOpenAiSettings?: () => void;
  isSyncingLiveOutlook?: boolean;
}

interface AssistantAiResult {
  ok?: boolean;
  model?: string;
  modelLabel?: string;
  usedFallback?: boolean;
  fallbackReason?: string;
  urgency?: string;
  summaryLines?: string[];
  actionChecklist?: string[];
  replyDrafts?: Array<{
    id: string;
    toneLabel: string;
    subject: string;
    body: string;
  }>;
}

function getCompanionEngineBadge(res: AssistantAiResult | null): { label: string; className: string } {
  if (!res) return { label: '분석 대기', className: 'bg-slate-100 text-slate-600' };
  const m = (res.model || '').toLowerCase();
  if (res.usedFallback || m === 'deterministic-local') {
    return { label: '오프라인 발췌', className: 'bg-amber-100 text-amber-900' };
  }
  if (m.includes('opus')) return { label: 'Claude Opus', className: 'bg-rose-100 text-rose-800' };
  if (m.includes('sonnet')) return { label: 'Claude Sonnet', className: 'bg-purple-100 text-purple-800' };
  if (m.includes('haiku')) return { label: 'Claude Haiku', className: 'bg-indigo-100 text-indigo-800' };
  if (/gpt-oss|llama|groq|qwen|mixtral/i.test(m)) {
    return { label: 'Groq AI', className: 'bg-emerald-100 text-emerald-800' };
  }
  return { label: 'AI 분석', className: 'bg-sky-100 text-[#0F6CBD]' };
}

function getBmailDesktopDock():
  | ((opts: { enabled: boolean; width?: number; alwaysOnTop?: boolean }) => Promise<{
      ok?: boolean;
      docked?: boolean;
      alwaysOnTop?: boolean;
    }>)
  | null {
  if (typeof window === 'undefined') return null;
  const desktop = (
    window as unknown as {
      bmailDesktop?: {
        setCompanionDock?: (opts: {
          enabled: boolean;
          width?: number;
          alwaysOnTop?: boolean;
        }) => Promise<{ ok?: boolean; docked?: boolean; alwaysOnTop?: boolean }>;
      };
    }
  ).bmailDesktop;
  return typeof desktop?.setCompanionDock === 'function' ? desktop.setCompanionDock : null;
}

function toSnapshotFromMailItem(mail: MailItem): OutlookActiveItemSnapshot {
  const attNames = (mail.attachments || []).map((a) => a.filename).filter(Boolean);
  return {
    entryId: mail.id,
    sourceType: 'local-fallback',
    subject: mail.subject || '(제목 없음)',
    fromName: mail.fromName || '',
    fromEmail: mail.fromEmail || '',
    to: mail.to || [],
    cc: mail.cc || [],
    bcc: [],
    body: mail.body || mail.snippet || '',
    attachments: attNames,
    attachmentsCount: attNames.length,
    receivedAt: mail.timestamp || mail.date || '',
    sent: mail.folder === 'sent',
    importance: mail.importance || 'normal',
  };
}

export const OutlookCompanionPanel: React.FC<OutlookCompanionPanelProps> = ({
  fallbackEmail,
  onSwitchToWorkbench,
  onSwitchToDashboard,
  onSendReceiveRefresh,
  onOpenAiSettings,
  isSyncingLiveOutlook = false,
}) => {
  const [pollResult, setPollResult] = useState<OutlookCompanionPollResult | null>(null);
  const [isPolling, setIsPolling] = useState(false);
  const [autoPoll, setAutoPoll] = useState(true);
  const [isDocked, setIsDocked] = useState(false);
  const [alwaysOnTop, setAlwaysOnTop] = useState(true);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);
  const [aiResult, setAiResult] = useState<AssistantAiResult | null>(null);
  const aiCacheRef = useRef<Map<string, AssistantAiResult>>(new Map());
  const aiAbortRef = useRef<AbortController | null>(null);
  const [replyTo, setReplyTo] = useState('');
  const [replySubject, setReplySubject] = useState('');
  const [replyBody, setReplyBody] = useState('');
  const [openingOutlook, setOpeningOutlook] = useState(false);
  const [comNotice, setComNotice] = useState<{ ok: boolean; text: string } | null>(null);

  const setCompanionDockFn = useMemo(() => getBmailDesktopDock(), []);

  const fetchActiveOutlookItem = useCallback(async () => {
    setIsPolling(true);
    try {
      const res = await fetch('/api/local/outlook/active-item');
      if (res.ok) {
        const data = (await res.json()) as OutlookCompanionPollResult;
        setPollResult(data);
      }
    } catch {
      // ignore transient network errors during poll
    } finally {
      setIsPolling(false);
    }
  }, []);

  useEffect(() => {
    void fetchActiveOutlookItem();
  }, [fetchActiveOutlookItem]);

  useEffect(() => {
    if (!autoPoll) return;
    const timer = window.setInterval(() => {
      if (typeof document !== 'undefined' && document.visibilityState !== 'visible') {
        return;
      }
      void fetchActiveOutlookItem();
    }, 2500);
    return () => window.clearInterval(timer);
  }, [autoPoll, fetchActiveOutlookItem]);

  const effectiveItem: OutlookActiveItemSnapshot | null = useMemo(() => {
    if (pollResult?.outlookRunning && pollResult.item) {
      return pollResult.item;
    }
    return fallbackEmail ? toSnapshotFromMailItem(fallbackEmail) : null;
  }, [pollResult, fallbackEmail]);

  const effectiveMode: OutlookCompanionMode = useMemo(() => {
    if (pollResult?.outlookRunning && pollResult.item) {
      return pollResult.mode;
    }
    return effectiveItem ? 'reading' : 'idle';
  }, [pollResult, effectiveItem]);

  const effectiveVerification: OutlookVerificationReport | null = useMemo(() => {
    if (pollResult?.outlookRunning && pollResult.verification) {
      return pollResult.verification;
    }
    if (!effectiveItem) return null;
    return verifyOutlookItemSafety(effectiveItem, effectiveMode);
  }, [pollResult, effectiveItem, effectiveMode]);

  // Pre-fill reply editor whenever active subject/sender changes
  useEffect(() => {
    if (!effectiveItem) return;
    setAiResult(null);
    setAiError(null);
    setComNotice(null);
    const subj = effectiveItem.subject || '(제목 없음)';
    setReplyTo(
      effectiveMode === 'composing'
        ? effectiveItem.to.join('; ')
        : effectiveItem.fromEmail || effectiveItem.fromName || ''
    );
    setReplySubject(subj.startsWith('RE:') ? subj : `RE: ${subj}`);
    const sender = effectiveItem.fromName || effectiveItem.fromEmail || '담당자';
    setReplyBody(
      effectiveMode === 'composing'
        ? effectiveItem.body
        : `안녕하세요, ${sender}님.\n\n"${subj}" 건 잘 수신하여 검토 중입니다.\n확인 완료 후 회신드리겠습니다.\n\n감사합니다.`
    );
  }, [effectiveItem?.entryId, effectiveItem?.subject, effectiveMode]);

  // Live pre-send verification of the reply draft itself
  const draftPreSendReport = useMemo(() => {
    if (!replySubject.trim() && !replyBody.trim()) return null;
    return verifyOutlookItemSafety(
      {
        subject: replySubject,
        fromEmail: effectiveItem?.fromEmail || '',
        to: replyTo
          .split(/[;,]/)
          .map((s) => s.trim())
          .filter(Boolean),
        body: replyBody,
        attachmentsCount: 0,
      },
      'composing'
    );
  }, [replyTo, replySubject, replyBody, effectiveItem?.fromEmail]);

  const handleToggleDock = async () => {
    if (!setCompanionDockFn) return;
    const nextDocked = !isDocked;
    const res = await setCompanionDockFn({
      enabled: nextDocked,
      width: 460,
      alwaysOnTop,
    });
    if (res?.ok) {
      setIsDocked(Boolean(res.docked));
    }
  };

  const handleTogglePinTop = async () => {
    if (!setCompanionDockFn) return;
    const nextPin = !alwaysOnTop;
    setAlwaysOnTop(nextPin);
    if (isDocked) {
      await setCompanionDockFn({
        enabled: true,
        width: 460,
        alwaysOnTop: nextPin,
      });
    }
  };

  const handleRunDeepAi = useCallback(
    async (
      highPerformance = false,
      itemOverride?: OutlookActiveItemSnapshot | null,
      bypassCache = false
    ) => {
      const target = itemOverride || effectiveItem;
      if (!target) return;
      const cacheKey = `${target.entryId}:${target.subject}`;
      if (!highPerformance && !bypassCache) {
        const cached = aiCacheRef.current.get(cacheKey);
        if (cached && !cached.usedFallback) {
          setAiResult(cached);
          setAiError(null);
          return;
        }
      }
      if (aiAbortRef.current) {
        aiAbortRef.current.abort();
      }
      const ctrl = new AbortController();
      aiAbortRef.current = ctrl;
      setAiLoading(true);
      setAiError(null);
      try {
        const res = await fetch('/api/local/mail/assistant', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            mode: 'analyze',
            highPerformance,
            email: {
              id: target.entryId,
              folder: target.sent ? 'sent' : 'inbox',
              subject: target.subject,
              fromName: target.fromName,
              fromEmail: target.fromEmail,
              to: target.to,
              body: target.body,
              attachments: target.attachments.map((name, i) => ({
                id: `att-${i}`,
                filename: name,
                mimeType: 'application/octet-stream',
                kind: 'other' as const,
              })),
            },
          }),
          signal: ctrl.signal,
        });
        if (!res.ok) {
          const errBody = (await res.json().catch(() => ({}))) as { error?: string };
          throw new Error(errBody.error || `HTTP ${res.status}`);
        }
        const data = (await res.json()) as AssistantAiResult;
        if (!data.usedFallback) {
          aiCacheRef.current.set(cacheKey, data);
        }
        setAiResult(data);
        if (
          Array.isArray(data.replyDrafts) &&
          data.replyDrafts.length > 0 &&
          effectiveMode !== 'composing'
        ) {
          setReplySubject(data.replyDrafts[0].subject);
          setReplyBody(data.replyDrafts[0].body);
        }
      } catch (err) {
        if (err instanceof DOMException && err.name === 'AbortError') {
          return;
        }
        setAiError(err instanceof Error ? err.message : 'AI 정밀 분석 중 오류가 발생했습니다.');
      } finally {
        if (aiAbortRef.current === ctrl) {
          aiAbortRef.current = null;
        }
        setAiLoading(false);
      }
    },
    [effectiveItem, effectiveMode]
  );

  // Automatically analyze active Outlook item with real Groq/Claude LLM
  useEffect(() => {
    if (!effectiveItem) return;
    const cacheKey = `${effectiveItem.entryId}:${effectiveItem.subject}`;
    const cached = aiCacheRef.current.get(cacheKey);
    if (cached && !cached.usedFallback) {
      setAiResult(cached);
      return;
    }
    const timer = window.setTimeout(() => {
      void handleRunDeepAi(false, effectiveItem);
    }, 350);
    return () => {
      window.clearTimeout(timer);
      if (aiAbortRef.current) {
        aiAbortRef.current.abort();
      }
    };
  }, [effectiveItem?.entryId, effectiveItem?.subject, handleRunDeepAi]);

  const applyTemplateDraft = (kind: 'confirm' | 'schedule' | 'drawing') => {
    if (!effectiveItem) return;
    const sender = effectiveItem.fromName || effectiveItem.fromEmail || '담당자';
    const subj = effectiveItem.subject || '(제목 없음)';
    const reSubj = subj.startsWith('RE:') ? subj : `RE: ${subj}`;
    setReplySubject(reSubj);
    if (kind === 'confirm') {
      setReplyBody(
        `안녕하세요, ${sender}님.\n\n"${subj}" 건 요청 내용 확인했습니다.\n내부 점검 후 진행 사항 다시 안내드리겠습니다.\n\n감사합니다.`
      );
    } else if (kind === 'schedule') {
      const deadlineHint =
        effectiveVerification?.deadlines[0] || '요청하신 기한';
      setReplyBody(
        `안녕하세요, ${sender}님.\n\n"${subj}" 관련 일정(${deadlineHint}) 확인했습니다.\n회신 가능 일시 조율 후 말씀드리겠습니다.\n\n감사합니다.`
      );
    } else {
      const drwHint =
        effectiveVerification?.drawings.slice(0, 3).join(', ') || '송부 도면';
      setReplyBody(
        `안녕하세요, ${sender}님.\n\n"${subj}" 관련 자료(${drwHint}) 접수하여 사양 및 리비전 검토 중입니다.\n코멘트 완료 즉시 회신드리겠습니다.\n\n감사합니다.`
      );
    }
  };

  const handleOpenInOriginalOutlook = async () => {
    setOpeningOutlook(true);
    setComNotice(null);
    try {
      const res = await fetch('/api/local/outlook/send-via-com', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          to: replyTo,
          subject: replySubject || '(제목 없음)',
          body: replyBody,
          action: 'display',
          importance: effectiveItem?.importance || 'normal',
        }),
      });
      const data = (await res.json().catch(() => ({}))) as { ok?: boolean; error?: string };
      if (!res.ok) {
        throw new Error(data.error || `HTTP ${res.status}`);
      }
      setComNotice({
        ok: true,
        text: '오리지널 Outlook 새 작성창에 초안을 열었습니다. Outlook에서 첨부·수신자 최종 확인 후 발송하세요.',
      });
    } catch (err) {
      setComNotice({
        ok: false,
        text:
          err instanceof Error
            ? `Outlook 창 열기 실패: ${err.message}`
            : '오리지널 Outlook 작성창을 열지 못했습니다.',
      });
    } finally {
      setOpeningOutlook(false);
    }
  };

  const isLiveOutlook = Boolean(pollResult?.outlookRunning && pollResult.item);

  return (
    <div className="flex-1 flex flex-col bg-[#FAF9F8] text-[#111827] min-h-0 overflow-hidden">
      {/* Top Companion Status & Dock Control Bar */}
      <div className="px-4 py-2.5 bg-white border-b border-[#E1DFDD] flex items-center justify-between gap-2 flex-wrap shrink-0">
        <div className="flex items-center gap-2 flex-wrap min-w-0">
          <span className="inline-flex items-center gap-1.5 text-sm font-bold text-[#0F6CBD]">
            <ShieldCheck className="w-4 h-4 shrink-0" />
            <span>Outlook 동반 분석 · 발송 전 검증</span>
          </span>
          <span
            data-testid="companion-connection-badge"
            className={`px-2 py-0.5 rounded-full text-[11px] font-bold border ${
              isLiveOutlook
                ? effectiveMode === 'composing'
                  ? 'bg-amber-50 text-amber-900 border-amber-300'
                  : 'bg-emerald-50 text-emerald-900 border-emerald-300'
                : 'bg-slate-100 text-[#374151] border-[#D1D1D1]'
            }`}
          >
            {isLiveOutlook
              ? effectiveMode === 'composing'
                ? '클래식 Outlook 연동 · 작성 중 발송 전 검증'
                : '클래식 Outlook 연동 · 읽기 분석'
              : 'Outlook 미실행 — 로컬 DB 최근 메일 검증'}
          </span>
        </div>

        <div className="flex items-center gap-1.5 flex-wrap shrink-0">
          {onSendReceiveRefresh && (
            <button
              type="button"
              data-testid="companion-send-receive-btn"
              onClick={onSendReceiveRefresh}
              disabled={isSyncingLiveOutlook}
              title="클래식 Outlook 보내기/받기 실행 후 새 메일을 즉시 가져옵니다"
              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-[4px] border border-[#0F6CBD] bg-[#0F6CBD] hover:bg-[#115EA3] text-xs font-semibold text-white disabled:opacity-60"
            >
              <RefreshCw
                className={`w-3.5 h-3.5 ${isSyncingLiveOutlook ? 'animate-spin' : ''}`}
              />
              <span>{isSyncingLiveOutlook ? '새 메일 확인 중...' : '보내기/받기'}</span>
            </button>
          )}
          <button
            type="button"
            onClick={() => void fetchActiveOutlookItem()}
            disabled={isPolling}
            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-[4px] border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-xs font-semibold text-[#111827] disabled:opacity-60"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-[#0F6CBD] ${isPolling ? 'animate-spin' : ''}`} />
            <span>지금 감지</span>
          </button>

          <button
            type="button"
            onClick={() => setAutoPoll((v) => !v)}
            className={`inline-flex items-center gap-1 px-2.5 py-1.5 rounded-[4px] border text-xs font-semibold ${
              autoPoll
                ? 'bg-[#EBF3FC] border-[#B4D6FA] text-[#0F6CBD]'
                : 'bg-white border-[#D1D1D1] text-[#374151]'
            }`}
          >
            {autoPoll ? <Pause className="w-3.5 h-3.5" /> : <Play className="w-3.5 h-3.5" />}
            <span>{autoPoll ? '실시간 감시 ON' : '감시 일시정지'}</span>
          </button>

          {setCompanionDockFn && (
            <>
              <button
                type="button"
                data-testid="companion-dock-toggle"
                onClick={() => void handleToggleDock()}
                className={`inline-flex items-center gap-1 px-2.5 py-1.5 rounded-[4px] border text-xs font-semibold ${
                  isDocked
                    ? 'bg-[#0F6CBD] border-[#0F6CBD] text-white'
                    : 'bg-white border-[#D1D1D1] text-[#111827] hover:bg-[#F3F2F1]'
                }`}
              >
                <PanelRight className="w-3.5 h-3.5" />
                <span>{isDocked ? '우측 도킹 해제' : '우측 사이드카 도킹'}</span>
              </button>
              <button
                type="button"
                aria-pressed={alwaysOnTop}
                onClick={() => void handleTogglePinTop()}
                className={`inline-flex items-center gap-1 px-2 py-1.5 rounded-[4px] border text-xs font-semibold ${
                  alwaysOnTop
                    ? 'bg-[#EDEBE9] border-[#D1D1D1] text-[#0F6CBD]'
                    : 'bg-white border-[#D1D1D1] text-[#374151]'
                }`}
              >
                <Pin className="w-3.5 h-3.5" />
                <span>항상 위</span>
              </button>
            </>
          )}

          <button
            type="button"
            onClick={onSwitchToDashboard}
            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-[4px] border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-xs font-semibold text-[#111827]"
          >
            <BarChart3 className="w-3.5 h-3.5 text-[#0F6CBD]" />
            <span>분석 대시보드</span>
          </button>

          <button
            type="button"
            onClick={onSwitchToWorkbench}
            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-[4px] border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-xs font-semibold text-[#111827]"
          >
            <ArrowLeft className="w-3.5 h-3.5" />
            <span>전체 메일함</span>
          </button>
        </div>
      </div>

      {/* Main Scrollable Companion Stack (Responsive for 460px Sidecar OR Full Width) */}
      <div className="flex-1 overflow-y-auto p-4">
        {!effectiveItem ? (
          <div className="max-w-xl mx-auto mt-10 p-6 bg-white border border-[#E1DFDD] rounded-lg text-center space-y-2">
            <ShieldCheck className="w-8 h-8 text-[#0F6CBD] mx-auto" />
            <p className="text-sm font-bold text-[#111827]">
              클래식 Outlook에서 메일을 선택하거나 새 메일을 작성하면 즉시 분석합니다.
            </p>
            <p className="text-xs text-[#4B5563]">
              Bmail은 아웃룩의 원본 PST와 메일 수발신을 대체하지 않고 옆에서 검증·초안만 지원합니다.
            </p>
          </div>
        ) : (
          <div className="max-w-6xl mx-auto grid grid-cols-1 xl:grid-cols-12 gap-4">
            {/* Left / Top Column (7 cols): Live Target Header + Pre-Send Verification + Instant AI Analysis */}
            <div className="xl:col-span-7 space-y-4">
              {/* Target Snapshot Header */}
              <div className="bg-white border border-[#E1DFDD] rounded-lg p-4 space-y-2">
                <div className="flex items-center justify-between gap-2 flex-wrap">
                  <span className="text-[11px] font-bold px-2 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD]">
                    {effectiveItem.sourceType === 'inspector'
                      ? 'Outlook 작성/열람 창 감지'
                      : effectiveItem.sourceType === 'explorer'
                        ? 'Outlook 선택 메일 감지'
                        : '로컬 보관 최근 메일 검증'}
                  </span>
                  <span className="inline-flex items-center gap-1 text-[11px] font-medium text-[#374151]">
                    <Paperclip className="w-3.5 h-3.5 text-[#4B5563]" />
                    첨부 {effectiveItem.attachmentsCount}건
                  </span>
                </div>

                <h2 className="text-base font-bold text-[#111827] leading-snug">
                  {effectiveItem.subject || '(제목 없음)'}
                </h2>
                <div className="text-xs text-[#374151] space-y-0.5">
                  <div>
                    <strong className="text-[#111827]">보낸 사람:</strong>{' '}
                    {effectiveItem.fromName || '-'} {effectiveItem.fromEmail ? `<${effectiveItem.fromEmail}>` : ''}
                  </div>
                  {effectiveItem.to.length > 0 && (
                    <div className="truncate">
                      <strong className="text-[#111827]">받는 사람:</strong>{' '}
                      {effectiveItem.to.join(', ')}
                    </div>
                  )}
                </div>
                {effectiveItem.attachments.length > 0 && (
                  <div className="pt-1 flex flex-wrap gap-1.5">
                    {effectiveItem.attachments.map((att, idx) => (
                      <span
                        key={`${att}-${idx}`}
                        className="px-2 py-0.5 rounded border border-[#D1D1D1] bg-[#FAF9F8] text-[11px] font-medium text-[#111827]"
                      >
                        {att}
                      </span>
                    ))}
                  </div>
                )}
              </div>

              {/* Section A: Real-Time Verification Report Card */}
              {effectiveVerification && (
                <section
                  aria-label="실시간 메일 검증 리포트"
                  data-testid="companion-verification-card"
                  className="bg-white border border-[#E1DFDD] rounded-lg overflow-hidden"
                >
                  <div
                    className={`px-4 py-2.5 border-b flex items-center justify-between gap-2 ${
                      effectiveVerification.status === 'block'
                        ? 'bg-rose-50 border-rose-200'
                        : effectiveVerification.status === 'caution'
                          ? 'bg-amber-50 border-amber-200'
                          : 'bg-emerald-50 border-emerald-200'
                    }`}
                  >
                    <div className="flex items-center gap-2 min-w-0">
                      {effectiveVerification.status === 'block' ? (
                        <AlertOctagon className="w-4 h-4 text-rose-700 shrink-0" />
                      ) : effectiveVerification.status === 'caution' ? (
                        <AlertTriangle className="w-4 h-4 text-amber-800 shrink-0" />
                      ) : (
                        <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                      )}
                      <span className="text-xs font-bold text-[#111827] truncate">
                        {effectiveVerification.summary}
                      </span>
                    </div>
                    <span className="px-2 py-0.5 rounded bg-white border border-[#D1D1D1] text-xs font-bold text-[#111827] tabular-nums shrink-0">
                      {effectiveVerification.score}점
                    </span>
                  </div>

                  <div className="p-4 space-y-2.5">
                    {effectiveVerification.findings.length === 0 ? (
                      <p className="text-xs text-[#374151]">
                        첨부파일 누락·수신자 누락·외부 민감정보 유출 위험이 감지되지 않았습니다.
                      </p>
                    ) : (
                      effectiveVerification.findings.map((f, idx) => (
                        <div
                          key={`${f.code}-${idx}`}
                          className={`p-2.5 rounded border text-xs space-y-0.5 ${
                            f.severity === 'critical'
                              ? 'bg-rose-50 border-rose-200 text-rose-950'
                              : f.severity === 'warning'
                                ? 'bg-amber-50 border-amber-200 text-amber-950'
                                : 'bg-[#FAF9F8] border-[#E1DFDD] text-[#111827]'
                          }`}
                        >
                          <div className="font-bold flex items-center gap-1.5">
                            {f.severity === 'info' && <Info className="w-3.5 h-3.5 text-[#0F6CBD]" />}
                            <span>{f.title}</span>
                          </div>
                          <p className="leading-relaxed opacity-95">{f.detail}</p>
                        </div>
                      ))
                    )}

                    {(effectiveVerification.drawings.length > 0 ||
                      effectiveVerification.deadlines.length > 0) && (
                      <div className="pt-2 border-t border-[#EDEBE9] flex flex-wrap gap-1.5 text-[11px]">
                        {effectiveVerification.drawings.map((drw) => (
                          <span
                            key={drw}
                            className="px-2 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] font-mono font-bold"
                          >
                            도면 {drw}
                          </span>
                        ))}
                        {effectiveVerification.deadlines.map((dl) => (
                          <span
                            key={dl}
                            className="px-2 py-0.5 rounded bg-slate-100 text-[#111827] font-mono font-semibold"
                          >
                            일정 {dl}
                          </span>
                        ))}
                      </div>
                    )}
                  </div>
                </section>
              )}

              {/* Section B: Instant Summary + Explicit Groq/Claude Deep Analysis */}
              <section
                aria-label="메일 AI 핵심 분석"
                className="bg-white border border-[#E1DFDD] rounded-lg p-4 space-y-3"
              >
                <div className="flex items-center justify-between gap-2 flex-wrap">
                  <div className="flex items-center gap-2 text-xs font-bold text-[#111827]">
                    <Sparkles className="w-4 h-4 text-[#0F6CBD]" />
                    <span>메일 AI 핵심 분석</span>
                    {(() => {
                      const badge = getCompanionEngineBadge(aiResult);
                      return (
                        <span
                          data-testid="companion-ai-engine-badge"
                          className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${badge.className}`}
                        >
                          {aiLoading ? 'AI 분석 중...' : badge.label}
                        </span>
                      );
                    })()}
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      disabled={aiLoading}
                      onClick={() => void handleRunDeepAi(false, undefined, true)}
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-[4px] border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-[#111827] text-xs font-semibold disabled:opacity-60"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${aiLoading ? 'animate-spin' : ''}`} />
                      <span>다시 분석</span>
                    </button>
                    <button
                      type="button"
                      data-testid="companion-claude-deep-btn"
                      disabled={aiLoading}
                      onClick={() => void handleRunDeepAi(true)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-[4px] bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-semibold disabled:opacity-60"
                    >
                      {aiLoading ? (
                        <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      ) : (
                        <Sparkles className="w-3.5 h-3.5" />
                      )}
                      <span>{aiLoading ? 'AI 분석 중...' : 'Claude 정밀 재분석'}</span>
                    </button>
                  </div>
                </div>

                {aiError && (
                  <div className="p-2.5 rounded border border-rose-200 bg-rose-50 text-xs text-rose-900">
                    {aiError}
                  </div>
                )}

                {aiResult?.usedFallback && aiResult.fallbackReason && (
                  <div className="p-2.5 rounded border border-amber-200 bg-amber-50 text-xs text-amber-900 flex flex-wrap items-center justify-between gap-2">
                    <span>
                      <strong>오프라인 발췌 표시 중:</strong> {aiResult.fallbackReason}
                    </span>
                    {onOpenAiSettings && (
                      <button
                        type="button"
                        onClick={onOpenAiSettings}
                        className="px-2.5 py-1 rounded bg-white border border-amber-300 hover:bg-amber-100 text-[11px] font-bold text-amber-950"
                      >
                        Claude API 설정 열기
                      </button>
                    )}
                  </div>
                )}

                {aiResult ? (
                  <div className="space-y-2.5 text-xs text-[#111827]">
                    {aiResult.modelLabel && (
                      <div className="text-[11px] font-semibold text-[#0F6CBD]">
                        엔진: {aiResult.modelLabel}
                      </div>
                    )}
                    {Array.isArray(aiResult.summaryLines) && aiResult.summaryLines.length > 0 && (
                      <ul className="list-disc pl-4 space-y-1">
                        {aiResult.summaryLines.map((line, i) => (
                          <li key={i} className="leading-relaxed">
                            {line}
                          </li>
                        ))}
                      </ul>
                    )}
                    {Array.isArray(aiResult.actionChecklist) && aiResult.actionChecklist.length > 0 && (
                      <div className="p-2.5 rounded bg-[#FAF9F8] border border-[#E1DFDD] space-y-1">
                        <div className="font-bold text-[#111827]">권장 후속 조치</div>
                        {aiResult.actionChecklist.map((act, i) => (
                          <div key={i}>• {act}</div>
                        ))}
                      </div>
                    )}
                  </div>
                ) : (
                  <div className="text-xs text-[#374151] leading-relaxed whitespace-pre-wrap bg-[#FAF9F8] p-3 rounded border border-[#E1DFDD] max-h-44 overflow-y-auto">
                    {effectiveItem.body.slice(0, 1000) || '(본문 텍스트가 비어 있습니다.)'}
                  </div>
                )}
              </section>
            </div>

            {/* Right / Bottom Column (5 cols): Delegate Safe Reply Draft Back to Real Classic Outlook */}
            <div className="xl:col-span-5">
              <section
                aria-label="오리지널 Outlook 회신 초안 위임"
                className="bg-white border border-[#E1DFDD] rounded-lg p-4 space-y-3"
              >
                <div className="flex items-center justify-between gap-2">
                  <h3 className="text-xs font-bold text-[#111827]">
                    오리지널 Outlook 작성창으로 회신 넘기기
                  </h3>
                  <span className="text-[11px] font-semibold text-[#0F6CBD]">MailItem.Display()</span>
                </div>
                <p className="text-[11px] text-[#4B5563]">
                  Bmail이 직접 메일을 발송하지 않고, 오리지널 Outlook 새 창을 열어 초안을 채워 드립니다.
                </p>

                <div className="flex flex-wrap gap-1.5">
                  <button
                    type="button"
                    onClick={() => applyTemplateDraft('confirm')}
                    className="px-2 py-1 rounded border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-[11px] font-semibold text-[#111827]"
                  >
                    확인·진행 회신
                  </button>
                  <button
                    type="button"
                    onClick={() => applyTemplateDraft('schedule')}
                    className="px-2 py-1 rounded border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-[11px] font-semibold text-[#111827]"
                  >
                    일정 조율 회신
                  </button>
                  <button
                    type="button"
                    onClick={() => applyTemplateDraft('drawing')}
                    className="px-2 py-1 rounded border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-[11px] font-semibold text-[#111827]"
                  >
                    도면·자료 검토 회신
                  </button>
                </div>

                <div className="space-y-2 text-xs">
                  <div>
                    <label className="block text-[11px] font-bold text-[#111827] mb-1">
                      받는 사람
                    </label>
                    <input
                      type="text"
                      value={replyTo}
                      onChange={(e) => setReplyTo(e.target.value)}
                      placeholder="recipient@company.com"
                      className="w-full px-2.5 py-1.5 rounded border border-[#D1D1D1] bg-white text-[#111827] focus:outline-none focus:border-[#0F6CBD]"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-[#111827] mb-1">
                      제목
                    </label>
                    <input
                      type="text"
                      value={replySubject}
                      onChange={(e) => setReplySubject(e.target.value)}
                      className="w-full px-2.5 py-1.5 rounded border border-[#D1D1D1] bg-white text-[#111827] focus:outline-none focus:border-[#0F6CBD]"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-[#111827] mb-1">
                      초안 본문
                    </label>
                    <textarea
                      rows={7}
                      value={replyBody}
                      onChange={(e) => setReplyBody(e.target.value)}
                      className="w-full px-2.5 py-1.5 rounded border border-[#D1D1D1] bg-white text-[#111827] leading-relaxed focus:outline-none focus:border-[#0F6CBD]"
                    />
                  </div>
                </div>

                {draftPreSendReport && draftPreSendReport.findings.length > 0 && (
                  <div
                    className={`p-2 rounded border text-[11px] ${
                      draftPreSendReport.status === 'block'
                        ? 'bg-rose-50 border-rose-200 text-rose-950'
                        : 'bg-amber-50 border-amber-200 text-amber-950'
                    }`}
                  >
                    <strong>초안 사전 점검:</strong> {draftPreSendReport.findings[0].title} —{' '}
                    {draftPreSendReport.findings[0].detail}
                  </div>
                )}

                {comNotice && (
                  <div
                    className={`p-2.5 rounded border text-xs ${
                      comNotice.ok
                        ? 'bg-emerald-50 border-emerald-200 text-emerald-950'
                        : 'bg-rose-50 border-rose-200 text-rose-950'
                    }`}
                  >
                    {comNotice.text}
                  </div>
                )}

                <button
                  type="button"
                  data-testid="companion-open-in-outlook-btn"
                  disabled={openingOutlook}
                  onClick={() => void handleOpenInOriginalOutlook()}
                  className="w-full py-2.5 px-3 rounded-[5px] bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-bold flex items-center justify-center gap-1.5 disabled:opacity-60"
                >
                  {openingOutlook ? (
                    <Loader2 className="w-4 h-4 animate-spin" />
                  ) : (
                    <ExternalLink className="w-4 h-4" />
                  )}
                  <span>
                    {openingOutlook ? 'Outlook 창 여는 중...' : '오리지널 Outlook 작성창으로 열기'}
                  </span>
                </button>
              </section>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
