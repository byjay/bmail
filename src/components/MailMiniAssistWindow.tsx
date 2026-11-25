import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  Check,
  CheckSquare,
  Copy,
  FileEdit,
  Globe,
  Languages,
  Loader2,
  MailPlus,
  RefreshCw,
  ShieldCheck,
  Sparkles,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';

interface MiniAssistData {
  mode: 'private-ai' | 'builtin-ai' | 'deterministic-local';
  detectedLanguage: 'en' | 'ja' | 'zh' | 'ko' | 'mixed';
  languageLabel: string;
  salesIntent: string;
  urgency: 'high' | 'normal' | 'low';
  promptInjectionDetected: boolean;
  translatedKorean: string;
  whatItMeans: string[];
  whatYouMustDo: string[];
  keyHighlights: {
    deadlines: string[];
    amounts: string[];
    references: string[];
  };
  replyDraft: {
    subject: string;
    body: string;
    toneLabel: string;
  };
  requiresUserConfirmation: true;
}

interface MailMiniAssistWindowProps {
  isOpen: boolean;
  onClose: () => void;
  activeEmail: MailItem | null;
  sgwStrictPrivateOnly?: boolean;
  onSaveReplyDraft: (draft: { to: string; subject: string; body: string }) => void;
  onOpenComposeWithDraft?: (draft: { to: string; subject: string; body: string }) => void;
}

export const MailMiniAssistWindow: React.FC<MailMiniAssistWindowProps> = ({
  isOpen,
  onClose,
  activeEmail,
  sgwStrictPrivateOnly = false,
  onSaveReplyDraft,
  onOpenComposeWithDraft,
}) => {
  const [activeTab, setActiveTab] = useState<'translate' | 'reply' | 'sales'>('translate');
  const [replyTone, setReplyTone] = useState<'formal' | 'concise' | 'english-reply'>('formal');
  const [useAi, setUseAi] = useState<boolean>(false); // Safe default: 0-LLM local first unless user enables
  const [manualSubject, setManualSubject] = useState('');
  const [manualSender, setManualSender] = useState('');
  const [manualBody, setManualBody] = useState('');
  const [editableReplyBody, setEditableReplyBody] = useState('');
  const [editableReplySubject, setEditableReplySubject] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<MiniAssistData | null>(null);
  const [copied, setCopied] = useState(false);
  const [savedNotice, setSavedNotice] = useState(false);

  // Sync active email into inputs automatically when selected mail changes
  useEffect(() => {
    if (!activeEmail) return;
    setManualSubject(activeEmail.subject);
    setManualSender(activeEmail.fromName || activeEmail.fromEmail);
    setManualBody(activeEmail.body || activeEmail.snippet);
    setResult(null);
    setError(null);
  }, [activeEmail?.id]);

  const buildClientDeterministicMiniAssist = (
    subj: string,
    sender: string,
    bodyText: string,
    tone: 'formal' | 'concise' | 'english-reply'
  ): MiniAssistData => {
    const combined = `${subj}\n${bodyText}`;
    const lower = combined.toLowerCase();
    const koCount = (combined.match(/[가-힣]/g) || []).length;
    const enCount = (combined.match(/[a-zA-Z]/g) || []).length;
    const isEnglish = enCount > koCount * 2;
    const salesIntent = /(pos|rfq|견적|단가|발주|정산|quote|price)/.test(lower)
      ? '견적 · POS · 발주 검토'
      : /(dwg|drawing|도면|승인도|선급|3d|modeling)/.test(lower)
        ? '도면 · 기술 사양 검토'
        : '일반 업무 · 일정 협의';
    const urgency: 'high' | 'normal' | 'low' = /(긴급|urgent|asap|미결|금일)/.test(lower)
      ? 'high'
      : 'normal';
    const deadlines = Array.from(
      new Set([
        ...(combined.match(/\b20\d{2}[-./]\d{1,2}[-./]\d{1,2}\b/g) || []),
        ...(combined.match(/\d{1,2}\s*월\s*\d{1,2}\s*일/g) || []),
      ])
    ).slice(0, 4);
    const amounts = Array.from(
      new Set([
        ...(combined.match(/(?:USD|EUR|KRW|\$|₩)\s?[\d,]+(?:\.\d+)?/gi) || []),
        ...(combined.match(/[\d,]+(?:\.\d+)?\s*(?:원|만원|억원|달러)/g) || []),
      ])
    ).slice(0, 4);
    const references = Array.from(
      new Set(
        (combined.match(/\b(?:KYTS|H\d{3}|22K|DWG|POS)[-_ ]?[A-Z0-9.-]{1,12}\b/gi) || []).filter(
          (r) => /\d/.test(r)
        )
      )
    ).slice(0, 5);
    const reSubj = subj.startsWith('RE:') ? subj : `RE: ${subj || '(제목 없음)'}`;
    const draftBody =
      tone === 'english-reply'
        ? `Dear ${sender || 'Partner'},\n\nThank you for your email regarding "${subj}".\nWe are reviewing the attached specifications and will get back to you shortly.\n\nBest regards,`
        : tone === 'concise'
          ? `${sender || '담당자'}님, 안녕하세요.\n\n"${subj}" 건 확인했습니다. 내부 검토 후 신속히 회신드리겠습니다.\n\n감사합니다.`
          : `안녕하세요, ${sender || '담당자'}님.\n\n보내주신 "${subj}" 관련 메일 및 자료 잘 받았습니다.\n요청하신 사항을 면밀히 검토한 후 조속히 회신드리겠습니다.\n\n감사합니다.`;

    return {
      mode: 'deterministic-local',
      detectedLanguage: isEnglish ? 'en' : 'ko',
      languageLabel: isEnglish ? '영어 (EN)' : '한국어 (KO)',
      salesIntent,
      urgency,
      promptInjectionDetected: false,
      translatedKorean: `[요약·번역 해설]\n발신: ${sender || '미상'}\n제목: ${subj}\n\n${bodyText.slice(0, 900) || '(본문 없음)'}`,
      whatItMeans: [
        `[${salesIntent}] ${sender || '발신자'} 측에서 "${subj}" 건으로 연락했습니다.`,
        ...(deadlines.length > 0 ? [`기한·일정: ${deadlines.join(', ')}`] : []),
        ...(references.length > 0 ? [`관련 코드·호선: ${references.join(', ')}`] : []),
      ],
      whatYouMustDo: [
        '첨부 문서 및 사양 변경 내역 확인 후 회신 일정 조율',
      ],
      keyHighlights: {
        deadlines,
        amounts,
        references,
      },
      replyDraft: {
        subject: reSubj,
        body: draftBody,
        toneLabel:
          tone === 'english-reply' ? '영문 비즈니스 회신' : tone === 'concise' ? '간결 확인 회신' : '정중 공식 회신',
      },
      requiresUserConfirmation: true,
    };
  };

  const runAnalysis = async (toneOverride?: 'formal' | 'concise' | 'english-reply') => {
    const subj = manualSubject.trim();
    const body = manualBody.trim();
    const effTone = toneOverride ?? replyTone;
    if (!subj && !body) {
      setError('메일을 선택하거나 번역·분석할 원문 텍스트를 입력해 주세요.');
      return;
    }

    const fallbackData = buildClientDeterministicMiniAssist(
      subj,
      manualSender.trim(),
      body,
      effTone
    );
    setLoading(true);
    setError(null);
    try {
      const res = await fetch('/api/mail/mini-assist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          subject: subj,
          senderName: manualSender.trim(),
          body,
          replyTone: effTone,
          useAi,
          sgwStrictPrivateOnly,
        }),
      });
      if (!res.ok) {
        setResult(fallbackData);
        setEditableReplySubject(fallbackData.replyDraft.subject);
        setEditableReplyBody(fallbackData.replyDraft.body);
        return;
      }
      const data = (await res.json()) as MiniAssistData;
      setResult(data);
      setEditableReplySubject(data.replyDraft.subject);
      setEditableReplyBody(data.replyDraft.body);
    } catch {
      // Never display raw Failed to fetch: always serve deterministic fallback cleanly
      setResult(fallbackData);
      setEditableReplySubject(fallbackData.replyDraft.subject);
      setEditableReplyBody(fallbackData.replyDraft.body);
    } finally {
      setLoading(false);
    }
  };

  // Auto-run fast local 0-LLM breakdown when opened with an active email
  useEffect(() => {
    if (isOpen && activeEmail && !result && !loading) {
      void runAnalysis();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, activeEmail?.id]);

  if (!isOpen) return null;

  const handleCopyReply = async () => {
    try {
      await navigator.clipboard.writeText(`${editableReplySubject}\n\n${editableReplyBody}`);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch {
      // ignore clipboard errors
    }
  };

  const handleSaveDraftClick = () => {
    onSaveReplyDraft({
      to: activeEmail?.fromEmail || '',
      subject: editableReplySubject || `Re: ${manualSubject}`,
      body: editableReplyBody,
    });
    setSavedNotice(true);
    setTimeout(() => setSavedNotice(false), 2000);
  };

  return (
    <section
      aria-label="영업·외국어 메일 번역 요약 및 회신 초안 미니창"
      className="fixed bottom-3 right-3 z-40 w-[330px] max-w-[calc(100vw-1.5rem)] max-h-[58vh] bg-white border border-[#D1D1D1] rounded-lg shadow-[0_12px_32px_rgba(0,0,0,0.22)] flex flex-col overflow-hidden"
    >
      {/* Fluent 2 Mini-Window Header */}
      <div className="bg-[#0F6CBD] text-white px-3.5 py-2.5 flex items-center justify-between gap-2 shrink-0">
        <div className="flex items-center gap-2 min-w-0">
          <Languages className="w-4 h-4 shrink-0" />
          <div className="min-w-0">
            <h3 className="text-xs font-semibold truncate">
              Bmail 영업 · 외국어 번역 · 회신 미니창
            </h3>
            <p className="text-[10px] text-white/80 truncate">
              {activeEmail
                ? `선택 메일 연동: ${activeEmail.subject}`
                : '메일을 선택하거나 외국어/영업 원문을 직접 붙여넣으세요'}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onClick={() => void runAnalysis()}
            disabled={loading}
            title="다시 번역·분석"
            className="p-1 rounded hover:bg-white/15 transition-colors disabled:opacity-50"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />
          </button>
          <button
            type="button"
            onClick={onClose}
            aria-label="미니창 닫기"
            className="p-1 rounded hover:bg-white/15 transition-colors"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Mode & Tabs Strip */}
      <div className="px-3 py-2 bg-[#FAF9F8] border-b border-[#E0E0E0] flex items-center justify-between gap-2 shrink-0">
        <div role="tablist" className="flex items-center gap-1">
          {[
            { id: 'translate' as const, label: '번역 · 할 일' },
            { id: 'reply' as const, label: '회신 초안' },
            { id: 'sales' as const, label: '영업 · 핵심수치' },
          ].map((tab) => (
            <button
              key={tab.id}
              type="button"
              role="tab"
              aria-selected={activeTab === tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`px-2.5 py-1 text-[11px] font-semibold rounded transition-colors ${
                activeTab === tab.id
                  ? 'bg-[#0F6CBD] text-white'
                  : 'text-[#424242] hover:bg-[#EDEBE9]'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        <label className="inline-flex items-center gap-1.5 text-[11px] text-[#424242] cursor-pointer select-none">
          <input
            type="checkbox"
            checked={useAi}
            onChange={(e) => setUseAi(e.target.checked)}
            className="rounded border-[#C8C6C4] text-[#0F6CBD] focus:ring-[#0F6CBD]"
          />
          <span>{sgwStrictPrivateOnly ? '비공개 AI' : 'AI 정밀번역'}</span>
        </label>
      </div>

      {/* Scrollable Body */}
      <div className="flex-1 overflow-y-auto p-3.5 space-y-3 text-xs">
        {/* Input Box (collapsible / editable when no email is selected or user wants custom paste) */}
        {!activeEmail && (
          <div className="space-y-1.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded-md p-2.5">
            <input
              type="text"
              value={manualSubject}
              onChange={(e) => setManualSubject(e.target.value)}
              placeholder="외국어 / 영업 메일 제목 입력"
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
            />
            <textarea
              rows={3}
              value={manualBody}
              onChange={(e) => setManualBody(e.target.value)}
              placeholder="영어·일본어·중국어 또는 영업 문의 본문을 붙여넣으세요..."
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD] resize-none"
            />
            <button
              type="button"
              onClick={() => void runAnalysis()}
              disabled={loading}
              className="w-full py-1.5 bg-[#0F6CBD] hover:bg-[#115EA3] text-white font-semibold rounded text-xs flex items-center justify-center gap-1.5 disabled:opacity-60"
            >
              {loading ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Sparkles className="w-3.5 h-3.5" />}
              번역 · 요약 · 회신초안 즉시 분석
            </button>
          </div>
        )}

        {error && (
          <div className="p-2.5 rounded border border-[#D13438]/40 bg-[#FDF3F4] text-[#A4262C] text-[11px]">
            {error}
          </div>
        )}

        {loading && (
          <div className="py-8 flex flex-col items-center justify-center gap-2 text-[#616161]">
            <Loader2 className="w-5 h-5 animate-spin text-[#0F6CBD]" />
            <span>메일 원문 번역 및 수행 과제 분석 중...</span>
          </div>
        )}

        {!loading && result && (
          <>
            {/* Status Metadata Line (Zero-pill clean inline bar) */}
            <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-[#EDEBE9] text-[11px] text-[#616161]">
              <div className="flex items-center gap-2">
                <Globe className="w-3.5 h-3.5 text-[#0F6CBD]" />
                <span className="font-semibold text-[#242424]">{result.languageLabel}</span>
                <span>·</span>
                <span className="font-semibold text-[#0F6CBD]">{result.salesIntent}</span>
              </div>
              <span className="font-mono text-[10px] text-[#616161]">
                {result.mode === 'private-ai'
                  ? '사내 비공개 AI'
                  : result.mode === 'builtin-ai'
                    ? '서버 AI 번역'
                    : '로컬 0-LLM 즉시분석'}
              </span>
            </div>

            {result.promptInjectionDetected && (
              <div className="p-2 rounded border border-[#C57B00]/50 bg-[#FFF8EB] text-[#8A5600] text-[11px] flex items-start gap-1.5">
                <AlertTriangle className="w-3.5 h-3.5 shrink-0 mt-0.5" />
                <span>
                  본문 내 의심스러운 명령/전달 지시문이 감지되어 데이터로 격리 처리했습니다. (자동 발송 차단됨)
                </span>
              </div>
            )}

            {/* TAB 1: Translation + What it means + What I must do */}
            {activeTab === 'translate' && (
              <div className="space-y-3">
                <div className="space-y-1">
                  <div className="text-[11px] font-bold text-[#242424]">무슨 내용인가? (핵심 해설)</div>
                  <ul className="space-y-1 text-[11.5px] text-[#323130] list-disc pl-4 leading-relaxed">
                    {result.whatItMeans.map((line, i) => (
                      <li key={i}>{line}</li>
                    ))}
                  </ul>
                </div>

                <div className="space-y-1.5 bg-[#EBF3FC]/60 border border-[#CFE4FA] rounded-md p-2.5">
                  <div className="text-[11px] font-bold text-[#0F6CBD] flex items-center gap-1.5">
                    <CheckSquare className="w-3.5 h-3.5" />
                    내가 무엇을 해야 하는가? (후속 조치)
                  </div>
                  <ul className="space-y-1 text-[11.5px] text-[#242424] list-decimal pl-4 leading-relaxed">
                    {result.whatYouMustDo.map((todo, i) => (
                      <li key={i}>{todo}</li>
                    ))}
                  </ul>
                </div>

                <div className="space-y-1">
                  <div className="text-[11px] font-bold text-[#242424]">한국어 번역 / 원문 정리</div>
                  <div className="p-2.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded text-[11.5px] text-[#323130] whitespace-pre-wrap leading-relaxed max-h-40 overflow-y-auto">
                    {result.translatedKorean}
                  </div>
                </div>
              </div>
            )}

            {/* TAB 2: Editable Reply Draft (NEVER auto-sends) */}
            {activeTab === 'reply' && (
              <div className="space-y-2.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[11px] font-bold text-[#242424]">회신 어조 선택</span>
                  <div className="flex items-center gap-1">
                    {[
                      { id: 'formal' as const, label: '국문 격식' },
                      { id: 'concise' as const, label: '국문 간결' },
                      { id: 'english-reply' as const, label: '영문 회신' },
                    ].map((t) => (
                      <button
                        key={t.id}
                        type="button"
                        onClick={() => {
                          setReplyTone(t.id);
                          void runAnalysis(t.id);
                        }}
                        className={`px-2 py-0.5 text-[10.5px] rounded border transition-colors ${
                          replyTone === t.id
                            ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-semibold'
                            : 'bg-white border-[#D1D1D1] text-[#424242] hover:bg-[#F5F5F5]'
                        }`}
                      >
                        {t.label}
                      </button>
                    ))}
                  </div>
                </div>

                <input
                  type="text"
                  value={editableReplySubject}
                  onChange={(e) => setEditableReplySubject(e.target.value)}
                  className="w-full px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                  placeholder="회신 제목"
                />
                <textarea
                  rows={7}
                  value={editableReplyBody}
                  onChange={(e) => setEditableReplyBody(e.target.value)}
                  className="w-full px-2.5 py-2 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD] leading-relaxed"
                />

                <div className="flex items-center gap-1.5 text-[10.5px] text-[#616161]">
                  <ShieldCheck className="w-3.5 h-3.5 text-[#107C10] shrink-0" />
                  <span>자동 발송 금지: 반드시 내용 확인 후 작성창 또는 임시보관함에서 발송하세요.</span>
                </div>

                <div className="flex items-center justify-end gap-1.5 pt-1">
                  <button
                    type="button"
                    onClick={handleCopyReply}
                    className="px-2.5 py-1.5 text-[11px] font-medium border border-[#D1D1D1] rounded hover:bg-[#F5F5F5] text-[#242424] inline-flex items-center gap-1"
                  >
                    {copied ? <Check className="w-3.5 h-3.5 text-[#107C10]" /> : <Copy className="w-3.5 h-3.5" />}
                    {copied ? '복사됨' : '복사'}
                  </button>
                  <button
                    type="button"
                    onClick={handleSaveDraftClick}
                    className="px-2.5 py-1.5 text-[11px] font-medium border border-[#0F6CBD] text-[#0F6CBD] rounded hover:bg-[#EBF3FC] inline-flex items-center gap-1"
                  >
                    <FileEdit className="w-3.5 h-3.5" />
                    {savedNotice ? '임시보관 저장됨' : '임시보관 저장'}
                  </button>
                  {onOpenComposeWithDraft && (
                    <button
                      type="button"
                      onClick={() =>
                        onOpenComposeWithDraft({
                          to: activeEmail?.fromEmail || '',
                          subject: editableReplySubject,
                          body: editableReplyBody,
                        })
                      }
                      className="px-2.5 py-1.5 text-[11px] font-semibold bg-[#0F6CBD] hover:bg-[#115EA3] text-white rounded inline-flex items-center gap-1"
                    >
                      <MailPlus className="w-3.5 h-3.5" />
                      회신창 열기
                    </button>
                  )}
                </div>
              </div>
            )}

            {/* TAB 3: Sales & Highlights */}
            {activeTab === 'sales' && (
              <div className="space-y-2.5">
                <div className="grid grid-cols-2 gap-2">
                  <div className="p-2.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded">
                    <div className="text-[10px] text-[#616161]">영업 · 업무 의도</div>
                    <div className="text-xs font-bold text-[#242424] mt-0.5">{result.salesIntent}</div>
                  </div>
                  <div className="p-2.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded">
                    <div className="text-[10px] text-[#616161]">긴급도 판정</div>
                    <div
                      className={`text-xs font-bold mt-0.5 ${
                        result.urgency === 'high'
                          ? 'text-[#D13438]'
                          : result.urgency === 'low'
                            ? 'text-[#616161]'
                            : 'text-[#0F6CBD]'
                      }`}
                    >
                      {result.urgency === 'high' ? '높음 (즉시 대응)' : result.urgency === 'low' ? '낮음 (참고)' : '보통'}
                    </div>
                  </div>
                </div>

                <div className="p-2.5 bg-white border border-[#E0E0E0] rounded space-y-2 text-[11px]">
                  <div>
                    <span className="font-semibold text-[#616161]">추출된 마감·일정: </span>
                    <span className="text-[#242424]">
                      {result.keyHighlights.deadlines.length > 0
                        ? result.keyHighlights.deadlines.join(', ')
                        : '명시된 날짜 없음'}
                    </span>
                  </div>
                  <div>
                    <span className="font-semibold text-[#616161]">추출된 금액·단가: </span>
                    <span className="text-[#242424] font-mono">
                      {result.keyHighlights.amounts.length > 0
                        ? result.keyHighlights.amounts.join(', ')
                        : '명시된 금액 없음'}
                    </span>
                  </div>
                  <div>
                    <span className="font-semibold text-[#616161]">프로젝트·도면·문서 코드: </span>
                    <span className="text-[#0F6CBD] font-mono">
                      {result.keyHighlights.references.length > 0
                        ? result.keyHighlights.references.join(', ')
                        : '미감지'}
                    </span>
                  </div>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </section>
  );
};
