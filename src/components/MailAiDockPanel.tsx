import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  CalendarPlus,
  Check,
  CheckCircle2,
  ClipboardCopy,
  Clock,
  CornerUpLeft,
  Languages,
  ListChecks,
  MessageSquare,
  RefreshCw,
  Send,
  Sparkles,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';
import type { CalendarScheduleEvent } from '../utils/deterministicAnalyzer';
import { extractShipbuildingDrawingRefs } from '../utils/outlookWorkspaceModel';

export type AiDockTab = 'analyze' | 'reply' | 'chat' | 'translate' | 'actions';
export type MailActionStatus = 'pending' | 'in_progress' | 'replied' | 'hold';

export interface MailActionRecord {
  status: MailActionStatus;
  memo: string;
  dueDate: string;
  updatedAt: string;
}

const AI_DOCK_LRU_CACHE = new Map<string, AssistantPayload>();
const MAX_DOCK_CACHE_ENTRIES = 50;

function readDockCache(cacheKey: string): AssistantPayload | null {
  const hit = AI_DOCK_LRU_CACHE.get(cacheKey);
  if (!hit) return null;
  AI_DOCK_LRU_CACHE.delete(cacheKey);
  AI_DOCK_LRU_CACHE.set(cacheKey, hit);
  return hit;
}

function writeDockCache(cacheKey: string, val: AssistantPayload): void {
  if (AI_DOCK_LRU_CACHE.has(cacheKey)) AI_DOCK_LRU_CACHE.delete(cacheKey);
  AI_DOCK_LRU_CACHE.set(cacheKey, val);
  if (AI_DOCK_LRU_CACHE.size > MAX_DOCK_CACHE_ENTRIES) {
    const oldest = AI_DOCK_LRU_CACHE.keys().next().value;
    if (oldest) AI_DOCK_LRU_CACHE.delete(oldest);
  }
}

export interface MailAiDockPanelProps {
  email: MailItem | null;
  onInsertReplyDraft?: (draft: { to: string; subject: string; body: string }) => void;
  onAddCalendarEvent?: (ev: Omit<CalendarScheduleEvent, 'id'>) => void;
  onClose?: () => void;
}

interface AssistantReplyDraft {
  id: string;
  toneLabel: string;
  subject: string;
  body: string;
}

interface AssistantPayload {
  model: string;
  modelLabel: string;
  tier: string;
  usedFallback: boolean;
  urgency: 'high' | 'normal';
  drawings: string[];
  deadlines: string[];
  summaryLines: string[];
  actionChecklist: string[];
  replyDrafts: AssistantReplyDraft[];
  translatedText: string;
  chatAnswer: string;
}

function buildClientFallbackPayload(
  email: MailItem,
  customPrompt?: string,
  targetLang = 'ko'
): AssistantPayload {
  const subject = email.subject || '(제목 없음)';
  const sender = email.fromName || email.fromEmail || '발신자';
  const body = email.body || email.snippet || '';
  const drawings = extractShipbuildingDrawingRefs(`${subject}\n${body}`).map(
    (d) => `${d.raw}${d.revision ? ` (${d.revision})` : ''}`
  );
  const deadlines = Array.from(
    new Set([
      ...(body.match(/\b20\d{2}[-./]\d{1,2}[-./]\d{1,2}\b/g) || []),
      ...(body.match(/\d{1,2}\s*월\s*\d{1,2}\s*일/g) || []),
    ])
  ).slice(0, 4);
  const isUrgent =
    email.importance === 'high' || /긴급|urgent|asap|미결|금일|즉시/i.test(`${subject} ${body}`);

  const summaryLines = [
    `[${isUrgent ? '긴급·우선 검토' : '일반 업무'}] ${sender} 측에서 "${subject}" 건을 발송했습니다.`,
    drawings.length > 0
      ? `식별된 도면·문서 코드: ${drawings.join(', ')}`
      : '식별된 도면 코드 없음 (사양/업무 협의)',
    deadlines.length > 0
      ? `언급된 기한·일정: ${deadlines.join(', ')}`
      : '명시적 마감 기한 미검출',
  ];

  const actionChecklist = [
    '첨부 문서(POS / 도면 / 단가표) 개정 내역 및 기술 사양 일치 여부 확인',
    deadlines.length > 0
      ? `기한(${deadlines[0]}) 내 검토 코멘트 및 승인 여부 회신`
      : '담당 파트 검토 후 회신 일정 안내',
  ];

  const reSubject = subject.startsWith('RE:') ? subject : `RE: ${subject}`;
  const replyDrafts: AssistantReplyDraft[] = [
    {
      id: 'formal',
      toneLabel: '정중·공식 회신',
      subject: reSubject,
      body: `안녕하세요, ${sender}님.\n\n보내주신 "${subject}" 관련 메일 및 첨부 자료 잘 받았습니다.\n${
        customPrompt ? `${customPrompt}\n` : '요청하신 사항을 내부 검토한 후 조속히 확정 회신드리겠습니다.\n'
      }\n감사합니다.`,
    },
    {
      id: 'technical',
      toneLabel: '기술·도면 검토 회신',
      subject: reSubject,
      body: `안녕하세요, ${sender}님.\n\n송부해 주신 "${subject}"${
        drawings.length > 0 ? ` (${drawings.join(', ')})` : ''
      } 자료 접수하여 사양 및 개정 사항을 확인 중입니다.\n검토 완료 즉시 코멘트 반영하여 회신드리겠습니다.\n\n감사합니다.`,
    },
    {
      id: 'concise',
      toneLabel: '간결·일정 조율 회신',
      subject: reSubject,
      body: `${sender}님, 안녕하세요.\n\n"${subject}" 건 확인했습니다. 내부 검토를 거쳐 금주 내로 회신드리겠습니다.\n\n감사합니다.`,
    },
  ];

  const translatedText =
    targetLang === 'en'
      ? `[English Summary & Translation]\nSubject: ${subject}\nFrom: ${sender}\n\n${body.slice(0, 1200)}`
      : `[한국어 요약·번역 해설]\n제목: ${subject}\n발신: ${sender}\n\n${summaryLines.join('\n')}\n\n--- 원문 주요 내용 ---\n${body.slice(0, 1200)}`;

  return {
    model: 'deterministic-local',
    modelLabel: '로컬 결정론적 엔진',
    tier: 'low',
    usedFallback: true,
    urgency: isUrgent ? 'high' : 'normal',
    drawings,
    deadlines,
    summaryLines,
    actionChecklist,
    replyDrafts,
    translatedText,
    chatAnswer: customPrompt
      ? `"${subject}" 메일 기준 "${customPrompt}" 질의 결과:\n- 발신: ${sender}\n- 핵심: ${summaryLines[0]}`
      : summaryLines.join('\n'),
  };
}

export const MailAiDockPanel: React.FC<MailAiDockPanelProps> = ({
  email,
  onInsertReplyDraft,
  onAddCalendarEvent,
  onClose,
}) => {
  const [activeTab, setActiveTab] = useState<AiDockTab>('analyze');
  const [loading, setLoading] = useState(false);
  const [payload, setPayload] = useState<AssistantPayload | null>(null);
  const [customReplyDirective, setCustomReplyDirective] = useState('');
  const [targetLang, setTargetLang] = useState<'ko' | 'en' | 'ja' | 'zh'>('ko');
  const [copiedId, setCopiedId] = useState<string | null>(null);

  // Per-mail persistent Chat & Action states
  const [chatMap, setChatMap] = useState<Record<string, Array<{ role: 'user' | 'assistant'; content: string }>>>(() => {
    try {
      const raw = localStorage.getItem('bmail_mail_chat_histories');
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });
  const [chatInput, setChatInput] = useState('');

  const [actionMap, setActionMap] = useState<Record<string, MailActionRecord>>(() => {
    try {
      const raw = localStorage.getItem('bmail_mail_action_states');
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });
  const [calendarSavedNotice, setCalendarSavedNotice] = useState<string | null>(null);

  const persistActionMap = (next: Record<string, MailActionRecord>) => {
    setActionMap(next);
    try {
      localStorage.setItem('bmail_mail_action_states', JSON.stringify(next));
    } catch {
      // ignore
    }
    void fetch('/api/local/state/mail_action_states', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ value: next }),
    }).catch(() => {});
  };

  const persistChatMap = (next: Record<string, Array<{ role: 'user' | 'assistant'; content: string }>>) => {
    setChatMap(next);
    try {
      localStorage.setItem('bmail_mail_chat_histories', JSON.stringify(next));
    } catch {
      // ignore
    }
    void fetch('/api/local/state/mail_chat_histories', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ value: next }),
    }).catch(() => {});
  };

  const abortControllerRef = useRef<AbortController | null>(null);

  const runAssistant = useCallback(
    async (
      targetEmail: MailItem,
      mode: AiDockTab,
      options?: { prompt?: string; lang?: string; bypassCache?: boolean }
    ) => {
      const effectiveLang = options?.lang || targetLang;
      const cacheKey =
        mode === 'analyze' && !options?.prompt ? `${targetEmail.id}:${effectiveLang}` : '';
      if (cacheKey && !options?.bypassCache) {
        const cached = readDockCache(cacheKey);
        if (cached) {
          setPayload(cached);
          return cached;
        }
      }

      const immediateFallback = buildClientFallbackPayload(
        targetEmail,
        options?.prompt,
        effectiveLang
      );
      setPayload((prev) => prev || immediateFallback);
      if (abortControllerRef.current) {
        abortControllerRef.current.abort();
      }
      const ctrl = new AbortController();
      abortControllerRef.current = ctrl;
      setLoading(true);
      try {
        const res = await fetch('/api/local/mail/assistant', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email: targetEmail,
            mode,
            prompt: options?.prompt,
            targetLang: effectiveLang,
          }),
          signal: ctrl.signal,
        });
        if (!res.ok) {
          setPayload(immediateFallback);
          return immediateFallback;
        }
        const data = (await res.json()) as AssistantPayload;
        if (cacheKey) {
          writeDockCache(cacheKey, data);
        }
        setPayload(data);
        return data;
      } catch {
        setPayload(immediateFallback);
        return immediateFallback;
      } finally {
        if (abortControllerRef.current === ctrl) {
          abortControllerRef.current = null;
        }
        setLoading(false);
      }
    },
    [targetLang]
  );

  useEffect(() => {
    if (!email) {
      setPayload(null);
      return;
    }
    const cacheKey = `${email.id}:${targetLang}`;
    const cached = readDockCache(cacheKey);
    if (cached) {
      setPayload(cached);
      return;
    }
    setPayload(buildClientFallbackPayload(email, undefined, targetLang));
    const timer = window.setTimeout(() => {
      void runAssistant(email, 'analyze', { lang: targetLang });
    }, 350);
    return () => {
      window.clearTimeout(timer);
      if (abortControllerRef.current) {
        abortControllerRef.current.abort();
      }
    };
  }, [email?.id, targetLang]);

  const currentAction: MailActionRecord = useMemo(() => {
    if (!email) {
      return { status: 'pending', memo: '', dueDate: '', updatedAt: '' };
    }
    return (
      actionMap[email.id] || {
        status: 'pending',
        memo: '',
        dueDate: payload?.deadlines?.[0] || email.date || '',
        updatedAt: '',
      }
    );
  }, [email, actionMap, payload?.deadlines]);

  const currentChatHistory = useMemo(() => {
    if (!email) return [];
    return chatMap[email.id] || [];
  }, [email, chatMap]);

  const handleCopy = async (id: string, text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedId(id);
      setTimeout(() => setCopiedId(null), 1500);
    } catch {
      // ignore
    }
  };

  const handleSendChatQuestion = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email || !chatInput.trim()) return;
    const question = chatInput.trim();
    setChatInput('');
    const nextHist: Array<{ role: 'user' | 'assistant'; content: string }> = [
      ...currentChatHistory,
      { role: 'user', content: question },
    ];
    persistChatMap({ ...chatMap, [email.id]: nextHist });

    const res = await runAssistant(email, 'chat', { prompt: question });
    const withAnswer: Array<{ role: 'user' | 'assistant'; content: string }> = [
      ...nextHist,
      { role: 'assistant', content: res.chatAnswer || res.summaryLines.join('\n') },
    ];
    persistChatMap({ ...chatMap, [email.id]: withAnswer });
  };

  if (!email) {
    return (
      <aside
        aria-label="세로 1분할 AI 분석 및 답안 도크"
        className="w-[272px] xl:w-[296px] bg-[#FAF9F8] border-l border-[#E0E0E0] p-4 flex flex-col justify-center items-center text-center text-xs text-slate-500 shrink-0"
      >
        <Sparkles className="w-6 h-6 text-[#0F6CBD] mb-2 opacity-60" />
        <p className="font-semibold text-slate-700">AI 분석 · 답안 도크</p>
        <p className="mt-1">메일을 선택하면 분석·답안·대화·번역·액션 메뉴가 표시됩니다.</p>
      </aside>
    );
  }

  const effective = payload || buildClientFallbackPayload(email, undefined, targetLang);

  return (
    <aside
      aria-label="세로 1분할 AI 분석 및 답안 도크"
      className="w-[272px] xl:w-[296px] bg-[#FAF9F8] border-l border-[#E0E0E0] flex flex-col min-h-0 shrink-0 select-none"
    >
      {/* Top Header */}
      <div className="px-3 py-2 bg-white border-b border-[#E0E0E0] flex items-center justify-between gap-1.5 shrink-0">
        <div className="flex items-center gap-1.5 min-w-0">
          <Sparkles className="w-4 h-4 text-[#0F6CBD] shrink-0" />
          <span className="text-xs font-bold text-[#242424] truncate">AI 분석</span>
          <span className="px-1.5 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] text-[9.5px] font-semibold shrink-0">
            고정
          </span>
          <span
            className={`px-1.5 py-0.5 rounded text-[10px] font-bold truncate ${
              effective.tier === 'high'
                ? 'bg-rose-100 text-rose-800'
                : effective.tier === 'medium'
                  ? 'bg-purple-100 text-purple-800'
                  : 'bg-sky-100 text-[#0F6CBD]'
            }`}
            title={effective.modelLabel}
          >
            {effective.model.includes('opus')
              ? 'Opus 5.5'
              : effective.model.includes('sonnet')
                ? 'Sonnet 5.5'
                : effective.model.includes('haiku')
                  ? 'Haiku 5.5'
                  : '로컬 엔진'}
          </span>
        </div>

        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            disabled={loading}
            onClick={() => void runAssistant(email, activeTab)}
            title="AI 다시 분석"
            className="p-1 rounded hover:bg-slate-100 text-slate-600"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin text-[#0F6CBD]' : ''}`} />
          </button>
          {onClose && (
            <button
              type="button"
              onClick={onClose}
              title="AI 도크 닫기"
              className="p-1 rounded hover:bg-slate-100 text-slate-500"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>

      {/* 5 Menu Tabs: 분석 / 답안 / 대화 / 번역 / 액션 */}
      <div className="grid grid-cols-5 bg-[#F0F0F0] p-1 gap-0.5 border-b border-[#E0E0E0] text-[11px] font-semibold shrink-0">
        {(
          [
            { id: 'analyze', label: '분석', icon: Sparkles },
            { id: 'reply', label: '답안', icon: CornerUpLeft },
            { id: 'chat', label: '대화', icon: MessageSquare },
            { id: 'translate', label: '번역', icon: Languages },
            { id: 'actions', label: '액션', icon: ListChecks },
          ] as const
        ).map((t) => {
          const Icon = t.icon;
          const isActive = activeTab === t.id;
          return (
            <button
              key={t.id}
              type="button"
              onClick={() => setActiveTab(t.id)}
              className={`py-1.5 rounded flex items-center justify-center gap-1 transition-colors ${
                isActive
                  ? 'bg-white text-[#0F6CBD] shadow-sm font-bold'
                  : 'text-[#616161] hover:text-[#242424]'
              }`}
            >
              <Icon className="w-3 h-3 shrink-0" />
              <span>{t.label}</span>
            </button>
          );
        })}
      </div>

      {/* Content Area */}
      <div className="flex-1 overflow-y-auto p-3 space-y-3 text-xs select-text">
        {activeTab === 'analyze' && (
          <>
            <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-[11px] font-bold text-[#0F6CBD]">핵심 브리핑 요약</span>
                <span
                  className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                    effective.urgency === 'high'
                      ? 'bg-rose-50 text-rose-700 border border-rose-200'
                      : 'bg-slate-100 text-slate-700'
                  }`}
                >
                  {effective.urgency === 'high' ? '긴급·우선' : '일반 중요도'}
                </span>
              </div>
              <ul className="space-y-1.5 text-[#242424] leading-relaxed">
                {effective.summaryLines.map((line, i) => (
                  <li key={i} className="flex items-start gap-1.5">
                    <span className="text-[#0F6CBD] font-bold">•</span>
                    <span>{line}</span>
                  </li>
                ))}
              </ul>
            </div>

            {effective.drawings.length > 0 && (
              <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-1.5">
                <div className="text-[11px] font-bold text-slate-700">식별된 도면 · 문서 번호</div>
                <div className="flex flex-wrap gap-1">
                  {effective.drawings.map((d) => (
                    <span
                      key={d}
                      className="px-2 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[11px] font-semibold"
                    >
                      {d}
                    </span>
                  ))}
                </div>
              </div>
            )}

            <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-1.5">
              <div className="text-[11px] font-bold text-slate-700">권장 실행 체크리스트</div>
              <ul className="space-y-1.5 text-slate-700">
                {effective.actionChecklist.map((item, i) => (
                  <li key={i} className="flex items-start gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0 mt-0.5" />
                    <span>{item}</span>
                  </li>
                ))}
              </ul>
            </div>
          </>
        )}

        {activeTab === 'reply' && (
          <>
            <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
              <label className="block text-[11px] font-bold text-slate-700">
                맞춤 답안 지시어 (선택)
              </label>
              <div className="flex gap-1.5">
                <input
                  type="text"
                  value={customReplyDirective}
                  onChange={(e) => setCustomReplyDirective(e.target.value)}
                  placeholder="예: 내일 오전까지 수정 POS 반영 회신 예정"
                  className="flex-1 px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                />
                <button
                  type="button"
                  disabled={loading}
                  onClick={() =>
                    void runAssistant(email, 'reply', { prompt: customReplyDirective })
                  }
                  className="px-2.5 py-1.5 rounded bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-[11px] font-semibold shrink-0"
                >
                  생성
                </button>
              </div>
            </div>

            {effective.replyDrafts.map((draft) => (
              <div
                key={draft.id}
                className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2"
              >
                <div className="flex items-center justify-between">
                  <span className="text-[11px] font-bold text-[#0F6CBD]">{draft.toneLabel}</span>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => void handleCopy(draft.id, draft.body)}
                      className="px-2 py-0.5 rounded bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10px] font-semibold inline-flex items-center gap-1"
                    >
                      {copiedId === draft.id ? (
                        <>
                          <Check className="w-3 h-3 text-emerald-600" /> 복사됨
                        </>
                      ) : (
                        <>
                          <ClipboardCopy className="w-3 h-3" /> 복사
                        </>
                      )}
                    </button>
                    {onInsertReplyDraft && (
                      <button
                        type="button"
                        onClick={() =>
                          onInsertReplyDraft({
                            to: email.fromEmail || email.fromName,
                            subject: draft.subject,
                            body: draft.body,
                          })
                        }
                        className="px-2 py-0.5 rounded bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-[10px] font-semibold"
                      >
                        회신창 삽입
                      </button>
                    )}
                  </div>
                </div>
                <div className="text-[11px] font-semibold text-slate-800 truncate">
                  {draft.subject}
                </div>
                <pre className="whitespace-pre-wrap font-sans text-[11px] text-slate-700 bg-slate-50 p-2 rounded border border-slate-200 leading-relaxed">
                  {draft.body}
                </pre>
              </div>
            ))}
          </>
        )}

        {activeTab === 'chat' && (
          <div className="flex flex-col h-full space-y-2.5">
            <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2 max-h-[360px] overflow-y-auto">
              <div className="text-[11px] font-bold text-[#0F6CBD]">
                이 메일에 대해 AI에게 질문하기
              </div>
              {currentChatHistory.length === 0 ? (
                <p className="text-[11px] text-slate-500">
                  예: "요구하는 납기일과 변경된 POS 핵심이 뭐야?", "영문으로 요약해줘"
                </p>
              ) : (
                <div className="space-y-2">
                  {currentChatHistory.map((msg, idx) => (
                    <div
                      key={idx}
                      className={`p-2 rounded-lg text-[11px] leading-relaxed whitespace-pre-wrap ${
                        msg.role === 'user'
                          ? 'bg-[#EBF3FC] text-[#0F6CBD] font-medium ml-4'
                          : 'bg-slate-100 text-slate-800 mr-2'
                      }`}
                    >
                      {msg.content}
                    </div>
                  ))}
                </div>
              )}
            </div>

            <form onSubmit={handleSendChatQuestion} className="flex gap-1.5">
              <input
                type="text"
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                placeholder="이 메일에 관해 질문 입력..."
                className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
              />
              <button
                type="submit"
                disabled={loading || !chatInput.trim()}
                className="px-2.5 py-1.5 rounded bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 text-white text-xs font-semibold flex items-center gap-1"
              >
                <Send className="w-3.5 h-3.5" />
              </button>
            </form>
          </div>
        )}

        {activeTab === 'translate' && (
          <div className="space-y-2.5">
            <div className="flex items-center justify-between bg-white p-2 rounded-lg border border-[#E0E0E0]">
              <span className="text-[11px] font-bold text-slate-700">목표 번역 언어</span>
              <div className="flex gap-1">
                {(
                  [
                    { code: 'ko', label: '한국어' },
                    { code: 'en', label: 'English' },
                    { code: 'ja', label: '日本語' },
                    { code: 'zh', label: '中文' },
                  ] as const
                ).map((l) => (
                  <button
                    key={l.code}
                    type="button"
                    onClick={() => {
                      setTargetLang(l.code);
                      void runAssistant(email, 'translate', { lang: l.code });
                    }}
                    className={`px-2 py-0.5 rounded text-[11px] font-semibold ${
                      targetLang === l.code
                        ? 'bg-[#0F6CBD] text-white'
                        : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                    }`}
                  >
                    {l.label}
                  </button>
                ))}
              </div>
            </div>

            <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-[11px] font-bold text-[#0F6CBD]">번역 및 비즈니스 해설</span>
                <button
                  type="button"
                  onClick={() => void handleCopy('translate', effective.translatedText)}
                  className="px-2 py-0.5 rounded bg-slate-100 hover:bg-slate-200 text-[10px] font-semibold"
                >
                  {copiedId === 'translate' ? '복사됨' : '복사'}
                </button>
              </div>
              <pre className="whitespace-pre-wrap font-sans text-[11px] text-slate-800 leading-relaxed">
                {effective.translatedText}
              </pre>
            </div>
          </div>
        )}

        {activeTab === 'actions' && (
          <div className="space-y-2.5">
            <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
              <div className="text-[11px] font-bold text-slate-700">후속 조치 상태 관리</div>
              <div className="grid grid-cols-2 gap-1.5">
                {(
                  [
                    { id: 'pending', label: '회신·조치 대기' },
                    { id: 'in_progress', label: '검토 진행 중' },
                    { id: 'replied', label: '회신·처리 완료' },
                    { id: 'hold', label: '보류·참고' },
                  ] as const
                ).map((st) => (
                  <button
                    key={st.id}
                    type="button"
                    onClick={() =>
                      persistActionMap({
                        ...actionMap,
                        [email.id]: {
                          ...currentAction,
                          status: st.id,
                          updatedAt: new Date().toISOString(),
                        },
                      })
                    }
                    className={`py-1.5 px-2 rounded text-[11px] font-semibold border text-left flex items-center justify-between ${
                      currentAction.status === st.id
                        ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD]'
                        : 'bg-slate-50 border-slate-200 text-slate-700'
                    }`}
                  >
                    <span>{st.label}</span>
                    {currentAction.status === st.id && <Check className="w-3.5 h-3.5" />}
                  </button>
                ))}
              </div>
            </div>

            <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
              <label className="block text-[11px] font-bold text-slate-700">
                마감 기한 · 메모 (로컬 DB 영구 저장)
              </label>
              <input
                type="date"
                value={currentAction.dueDate}
                onChange={(e) =>
                  persistActionMap({
                    ...actionMap,
                    [email.id]: {
                      ...currentAction,
                      dueDate: e.target.value,
                      updatedAt: new Date().toISOString(),
                    },
                  })
                }
                className="w-full px-2 py-1 text-xs border border-[#D1D1D1] rounded"
              />
              <textarea
                rows={3}
                value={currentAction.memo}
                onChange={(e) =>
                  persistActionMap({
                    ...actionMap,
                    [email.id]: {
                      ...currentAction,
                      memo: e.target.value,
                      updatedAt: new Date().toISOString(),
                    },
                  })
                }
                placeholder="도면 검토 코멘트, 담당자 확인 메모 입력..."
                className="w-full px-2 py-1.5 text-xs border border-[#D1D1D1] rounded"
              />
              {onAddCalendarEvent && (
                <button
                  type="button"
                  onClick={() => {
                    onAddCalendarEvent({
                      title: `[메일 조치] ${email.subject.slice(0, 42)}`,
                      date: currentAction.dueDate || email.date || new Date().toISOString().slice(0, 10),
                      timeLabel: '10:00',
                      category: 'deadline',
                      sourceEmailId: email.id,
                      notes: currentAction.memo || effective.summaryLines.join(' / '),
                    });
                    setCalendarSavedNotice('Outlook 일정 달력에 등록되었습니다.');
                    setTimeout(() => setCalendarSavedNotice(null), 2500);
                  }}
                  className="w-full py-1.5 rounded bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-[11px] font-semibold flex items-center justify-center gap-1.5"
                >
                  <CalendarPlus className="w-3.5 h-3.5" />
                  <span>Outlook 일정 달력에 등록</span>
                </button>
              )}
              {calendarSavedNotice && (
                <div className="text-[11px] text-emerald-700 font-semibold">
                  ✓ {calendarSavedNotice}
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </aside>
  );
};
