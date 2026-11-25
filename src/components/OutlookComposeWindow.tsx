import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlignCenter,
  AlignLeft,
  Bold,
  Clock,
  ExternalLink,
  FileEdit,
  Italic,
  Paperclip,
  PenSquare,
  Send,
  ShieldCheck,
  Trash2,
  Underline,
  X,
} from 'lucide-react';
import type { MailAttachmentMeta } from '../types/mail';
import { coachEmailDraft } from '../utils/aiMailIntelligence';
import {
  isInlineBusinessCardImage,
  type OutlookPreferences,
} from '../utils/outlookWorkspaceModel';

export interface OutlookSignatureEntry {
  name: string;
  htmlContent: string;
  textContent: string;
  isDefaultNew: boolean;
  isDefaultReply: boolean;
}

export interface ComposeDraftState {
  kind: 'new' | 'reply' | 'replyAll' | 'forward' | 'worklog';
  title: string;
  to: string;
  cc: string;
  subject: string;
  introText?: string;
  sourceEmailId?: string;
  quoteFromLabel?: string;
  quoteSentLabel?: string;
  quoteToLabel?: string;
  quoteCcLabel?: string;
  quoteSubjectLabel?: string;
  quotePlainBody?: string;
  quoteHtmlBody?: string;
  attachments?: MailAttachmentMeta[];
  importance?: 'low' | 'normal' | 'high';
  htmlSnapshot?: string;
}

export interface OutlookComposeWindowProps {
  mode: 'modal' | 'standalone';
  draft: ComposeDraftState;
  signatures: OutlookSignatureEntry[];
  outlookPrefs: OutlookPreferences;
  pop3ConnectedEmail?: string | null;
  onSendSmtp?: (payload: {
    to: string;
    cc: string;
    subject: string;
    body: string;
    htmlBody?: string;
  }) => Promise<{ ok: boolean; message: string }>;
  onSaveDraft?: (draft: {
    to: string;
    cc: string;
    subject: string;
    body: string;
    kind: ComposeDraftState['kind'];
  }) => void;
  onPopOutNewWindow?: (snapshot: ComposeDraftState) => void;
  onClose: () => void;
}

function escapeHtmlText(text: string): string {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function textToParagraphHtml(text: string): string {
  const escaped = escapeHtmlText(text);
  if (!escaped.trim()) return '<p><br></p>';
  return escaped
    .split(/\r?\n/)
    .map((line) => (line ? `<div>${line}</div>` : '<div><br></div>'))
    .join('');
}

function buildSignatureBlockHtml(sig: OutlookSignatureEntry | null, fallbackPlain: string): string {
  const html = sig?.htmlContent?.trim() || '';
  if (html) {
    return `<div data-bmail-sig="1" style="margin-top:14px;">${html}</div>`;
  }
  const plain = sig?.textContent?.trim() || fallbackPlain.trim();
  if (!plain) return '';
  return `<div data-bmail-sig="1" style="margin-top:14px;color:#323130;font-size:10pt;line-height:1.5;">${textToParagraphHtml(
    plain
  )}</div>`;
}

function buildQuoteBlockHtml(draft: ComposeDraftState, richHtml?: string): string {
  if (!draft.quoteFromLabel && !draft.quotePlainBody && !richHtml) {
    return '';
  }
  const innerBody =
    richHtml && richHtml.trim()
      ? richHtml
      : `<div style="white-space:pre-wrap;">${escapeHtmlText(draft.quotePlainBody || '')}</div>`;

  const headers: string[] = [];
  if (draft.quoteFromLabel) {
    headers.push(`<div><b>From:</b> ${escapeHtmlText(draft.quoteFromLabel)}</div>`);
  }
  if (draft.quoteSentLabel) {
    headers.push(`<div><b>Sent:</b> ${escapeHtmlText(draft.quoteSentLabel)}</div>`);
  }
  if (draft.quoteToLabel) {
    headers.push(`<div><b>To:</b> ${escapeHtmlText(draft.quoteToLabel)}</div>`);
  }
  if (draft.quoteCcLabel) {
    headers.push(`<div><b>Cc:</b> ${escapeHtmlText(draft.quoteCcLabel)}</div>`);
  }
  if (draft.quoteSubjectLabel) {
    headers.push(`<div><b>Subject:</b> ${escapeHtmlText(draft.quoteSubjectLabel)}</div>`);
  }

  return [
    '<div data-bmail-quote="1" style="margin-top:18px;padding-top:10px;border-top:1px solid #B5C4DF;font-family:\'Malgun Gothic\',\'Segoe UI\',sans-serif;font-size:10pt;color:#242424;">',
    ...headers,
    '<br/>',
    innerBody,
    '</div>',
  ].join('');
}

export const OutlookComposeWindow: React.FC<OutlookComposeWindowProps> = ({
  mode,
  draft,
  signatures,
  outlookPrefs,
  pop3ConnectedEmail,
  onSendSmtp,
  onSaveDraft,
  onPopOutNewWindow,
  onClose,
}) => {
  const [to, setTo] = useState(draft.to);
  const [cc, setCc] = useState(draft.cc);
  const [subject, setSubject] = useState(draft.subject);
  const [importance, setImportance] = useState<'low' | 'normal' | 'high'>(
    draft.importance || outlookPrefs.mail.defaultImportance || 'normal'
  );
  const [sendDelaySec, setSendDelaySec] = useState<number>(outlookPrefs.mail.undoSendSeconds ?? 5);
  const [fontSizePt, setFontSizePt] = useState<string>('10pt');
  const [selectedSigName, setSelectedSigName] = useState<string>(() => {
    const isReplyOrFwd =
      draft.kind === 'reply' || draft.kind === 'replyAll' || draft.kind === 'forward';
    const preferred =
      (isReplyOrFwd
        ? signatures.find((s) => s.isDefaultReply)
        : signatures.find((s) => s.isDefaultNew)) ||
      signatures.find((s) => s.name === outlookPrefs.mail.signatureName) ||
      signatures[0];
    return preferred?.name || '';
  });

  const realAttachments = useMemo(
    () => (draft.attachments || []).filter((att) => !isInlineBusinessCardImage(att)),
    [draft.attachments]
  );

  const editorRef = useRef<HTMLDivElement | null>(null);
  const [plainPreview, setPlainPreview] = useState('');
  const [notice, setNotice] = useState<string | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [isSending, setIsSending] = useState(false);

  const activeSig = useMemo(() => {
    if (selectedSigName === '__none__') return null;
    const byName = signatures.find((s) => s.name === selectedSigName);
    if (byName) return byName;
    if (outlookPrefs.mail.signatureHtml || outlookPrefs.mail.signature) {
      return {
        name: outlookPrefs.mail.signatureName || 'Outlook 기본 서명',
        htmlContent: outlookPrefs.mail.signatureHtml || '',
        textContent: outlookPrefs.mail.signature || '',
        isDefaultNew: true,
        isDefaultReply: true,
      };
    }
    return null;
  }, [selectedSigName, signatures, outlookPrefs.mail]);

  // Populate initial rich HTML editor content once per compose session
  useEffect(() => {
    if (!editorRef.current) return;
    if (draft.htmlSnapshot && draft.htmlSnapshot.trim()) {
      editorRef.current.innerHTML = draft.htmlSnapshot;
      setPlainPreview(editorRef.current.innerText || '');
      return;
    }

    const introHtml = draft.introText
      ? textToParagraphHtml(draft.introText)
      : '<div><br></div><div><br></div>';
    const sigHtml = buildSignatureBlockHtml(activeSig, outlookPrefs.mail.signature || '');
    const quoteHtml = buildQuoteBlockHtml(draft, draft.quoteHtmlBody);
    editorRef.current.innerHTML = `${introHtml}${sigHtml}${quoteHtml}`;
    setPlainPreview(editorRef.current.innerText || '');
  }, [draft.sourceEmailId, draft.kind]);

  // Upgrade quoted section to full original Outlook rich HTML + inline images from local cache/PST
  useEffect(() => {
    const srcId = draft.sourceEmailId;
    if (!srcId || draft.htmlSnapshot || draft.quoteHtmlBody) return;
    let cancelled = false;
    fetch(`/api/local/emails/${encodeURIComponent(srcId)}/render`)
      .then((r) => (r.ok ? r.json() : null))
      .then((data: { bodyHtml?: string } | null) => {
        if (cancelled || !data?.bodyHtml || !editorRef.current) return;
        const quoteEl = editorRef.current.querySelector('[data-bmail-quote="1"]');
        if (quoteEl) {
          const freshWrapper = document.createElement('div');
          freshWrapper.innerHTML = buildQuoteBlockHtml(draft, data.bodyHtml);
          const nextQuote = freshWrapper.firstElementChild;
          if (nextQuote) {
            quoteEl.replaceWith(nextQuote);
            setPlainPreview(editorRef.current.innerText || '');
          }
        }
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, [draft.sourceEmailId]);

  const handleApplySignatureChange = (sigName: string) => {
    setSelectedSigName(sigName);
    if (!editorRef.current) return;
    const picked =
      sigName === '__none__' ? null : signatures.find((s) => s.name === sigName) || activeSig;
    const sigHtml = sigName === '__none__' ? '' : buildSignatureBlockHtml(picked, '');
    const existingSig = editorRef.current.querySelector('[data-bmail-sig="1"]');
    if (existingSig) {
      if (!sigHtml) {
        existingSig.remove();
      } else {
        const tmp = document.createElement('div');
        tmp.innerHTML = sigHtml;
        if (tmp.firstElementChild) existingSig.replaceWith(tmp.firstElementChild);
      }
    } else if (sigHtml) {
      const quoteEl = editorRef.current.querySelector('[data-bmail-quote="1"]');
      const tmp = document.createElement('div');
      tmp.innerHTML = sigHtml;
      if (tmp.firstElementChild) {
        if (quoteEl) quoteEl.before(tmp.firstElementChild);
        else editorRef.current.appendChild(tmp.firstElementChild);
      }
    }
    setPlainPreview(editorRef.current.innerText || '');
  };

  const execFormat = (command: string, value?: string) => {
    editorRef.current?.focus();
    try {
      document.execCommand(command, false, value);
    } catch {
      // ignore unsupported execCommand in synthetic environments
    }
  };

  const composeCoaching = useMemo(() => {
    return coachEmailDraft({
      to: to.trim(),
      subject: subject.trim(),
      body: plainPreview || draft.introText || '',
      hasAttachment: realAttachments.length > 0,
    });
  }, [to, subject, plainPreview, draft.introText, realAttachments.length]);

  const collectForwardAttachmentIndices = (): number[] => {
    if (draft.kind !== 'forward') return [];
    return realAttachments
      .map((att, idx) => {
        const m = String(att.id || '').match(/-att-(\d+)$/);
        return m ? Number(m[1]) : idx;
      })
      .filter((n) => Number.isInteger(n) && n >= 0);
  };

  const handleSend = async (comAction: 'send' | 'display' = 'send') => {
    setErrorMsg(null);
    setNotice(null);
    const htmlBody = editorRef.current?.innerHTML || '';
    const textBody = editorRef.current?.innerText || plainPreview || '';
    const cleanSubject = subject.trim() || '(제목 없음)';

    if (comAction === 'send' && !to.trim()) {
      setErrorMsg('받는 사람 이메일 주소를 입력해 주세요.');
      return;
    }

    setIsSending(true);
    try {
      if (comAction === 'send' && pop3ConnectedEmail && onSendSmtp) {
        const smtpRes = await onSendSmtp({
          to: to.trim(),
          cc: cc.trim(),
          subject: cleanSubject,
          body: textBody,
          htmlBody,
        });
        if (smtpRes.ok) {
          setNotice(smtpRes.message);
          window.setTimeout(() => onClose(), 1000);
          return;
        }
        setErrorMsg(smtpRes.message);
        return;
      }

      // Send or open directly via PC Outlook COM (preserves HTML business card, inline images & forwarded files)
      const res = await fetch('/api/local/outlook/send-via-com', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          to: to.trim(),
          cc: cc.trim(),
          subject: cleanSubject,
          body: textBody,
          htmlBody,
          action: comAction,
          sourceEmailId: draft.sourceEmailId,
          forwardAttachmentIndices: collectForwardAttachmentIndices(),
        }),
      });
      const data = (await res.json()) as {
        ok?: boolean;
        error?: string;
        attachmentsCount?: number;
        inlineImagesCount?: number;
      };
      if (!res.ok || !data.ok) {
        setErrorMsg(
          data.error ||
            'PC Outlook COM 발송에 실패했습니다. 임시 보관함에 저장하거나 POP3/SMTP 계정을 연결해 주세요.'
        );
        return;
      }
      setNotice(
        comAction === 'display'
          ? `PC Outlook 새 메세지 창을 열었습니다 (명함 이미지 ${data.inlineImagesCount ?? 0}개 · 첨부 ${data.attachmentsCount ?? 0}개 포함).`
          : `Outlook으로 HTML 메일을 즉시 발송했습니다 (명함 이미지 ${data.inlineImagesCount ?? 0}개 · 첨부 ${data.attachmentsCount ?? 0}개).`
      );
      if (comAction === 'send') {
        window.setTimeout(() => onClose(), 1100);
      }
    } catch {
      setErrorMsg('메일 발송 중 통신 오류가 발생했습니다.');
    } finally {
      setIsSending(false);
    }
  };

  const handleSave = () => {
    const textBody = editorRef.current?.innerText || plainPreview || '';
    onSaveDraft?.({
      to: to.trim(),
      cc: cc.trim(),
      subject: subject.trim() || '(제목 없음)',
      body: textBody,
      kind: draft.kind,
    });
    setNotice('임시 보관함에 저장했습니다.');
  };

  const shellContent = (
    <div
      data-testid="outlook-compose-window"
      className={
        mode === 'standalone'
          ? 'w-screen h-screen bg-white flex flex-col overflow-hidden'
          : 'w-[min(1080px,96vw)] h-[min(860px,90vh)] bg-white rounded-lg border border-[#C8C6C4] shadow-2xl flex flex-col overflow-hidden'
      }
    >
      {/* Classic Outlook Window Title Bar */}
      <div className="px-4 py-2 bg-[#0F6CBD] text-white flex items-center justify-between gap-2 shrink-0 select-none">
        <div className="flex items-center gap-2 min-w-0">
          <PenSquare className="w-4 h-4 shrink-0" />
          <span className="text-xs sm:text-sm font-semibold truncate">
            {subject.trim() ? `${subject.trim()} — 메시지 (HTML)` : `${draft.title} — 메시지 (HTML)`}
          </span>
        </div>
        <div className="flex items-center gap-1.5 shrink-0">
          {mode === 'modal' && onPopOutNewWindow && (
            <button
              type="button"
              onClick={() =>
                onPopOutNewWindow({
                  ...draft,
                  to,
                  cc,
                  subject,
                  importance,
                  htmlSnapshot: editorRef.current?.innerHTML || '',
                })
              }
              title="독립 새 창으로 분리하여 작성"
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-white/15 hover:bg-white/25 text-[11px] font-semibold"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              새 창으로 열기
            </button>
          )}
          <button
            type="button"
            onClick={onClose}
            aria-label="작성 창 닫기"
            className="p-1 rounded hover:bg-white/20"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Classic Outlook Formatting & Action Ribbon */}
      <div className="px-3 py-1.5 bg-[#F3F2F1] border-b border-[#D1D1D1] flex items-center gap-1.5 flex-wrap text-xs shrink-0">
        <select
          aria-label="글꼴 크기"
          value={fontSizePt}
          onChange={(e) => {
            setFontSizePt(e.target.value);
            execFormat('fontSize', e.target.value === '14pt' ? '4' : e.target.value === '12pt' ? '3' : '2');
          }}
          className="px-2 py-1 rounded border border-[#C8C6C4] bg-white text-[11px] text-[#242424]"
        >
          <option value="10pt">맑은 고딕 · 10pt</option>
          <option value="11pt">맑은 고딕 · 11pt</option>
          <option value="12pt">맑은 고딕 · 12pt</option>
          <option value="14pt">맑은 고딕 · 14pt</option>
        </select>

        <div className="inline-flex items-center rounded border border-[#C8C6C4] bg-white overflow-hidden">
          <button
            type="button"
            onClick={() => execFormat('bold')}
            title="굵게"
            className="px-2 py-1 hover:bg-[#EDEBE9]"
          >
            <Bold className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => execFormat('italic')}
            title="기울임꼴"
            className="px-2 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]"
          >
            <Italic className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => execFormat('underline')}
            title="밑줄"
            className="px-2 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]"
          >
            <Underline className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => execFormat('justifyLeft')}
            title="왼쪽 맞춤"
            className="px-2 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]"
          >
            <AlignLeft className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => execFormat('justifyCenter')}
            title="가운데 맞춤"
            className="px-2 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]"
          >
            <AlignCenter className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="flex items-center gap-1 ml-1">
          <span className="text-[11px] font-semibold text-[#616161]">서명(명함):</span>
          <select
            aria-label="Outlook 명함 서명 선택"
            value={selectedSigName}
            onChange={(e) => handleApplySignatureChange(e.target.value)}
            className="px-2 py-1 rounded border border-[#C8C6C4] bg-white text-[11px] text-[#242424] max-w-[220px]"
          >
            {signatures.length === 0 ? (
              <option value="">기본 서명</option>
            ) : (
              signatures.map((s) => (
                <option key={s.name} value={s.name}>
                  {s.name}
                  {s.isDefaultNew ? ' (기본)' : ''}
                </option>
              ))
            )}
            <option value="__none__">(서명 없음)</option>
          </select>
        </div>

        <button
          type="button"
          onClick={() => setImportance((v) => (v === 'high' ? 'normal' : 'high'))}
          className={`px-2 py-1 rounded border text-[11px] font-semibold ${
            importance === 'high'
              ? 'bg-rose-50 border-rose-300 text-rose-700'
              : 'bg-white border-[#C8C6C4] text-[#424242]'
          }`}
        >
          {importance === 'high' ? '! 중요도 높음' : '! 중요도'}
        </button>

        <button
          type="button"
          onClick={() => void handleSend('display')}
          disabled={isSending}
          title="PC Outlook 실제 창으로 초안 바로 띄우기"
          className="px-2.5 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] font-semibold text-[#0F6CBD]"
        >
          Outlook으로 띄우기
        </button>

        {onSaveDraft && (
          <button
            type="button"
            onClick={handleSave}
            className="inline-flex items-center gap-1 px-2.5 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] font-medium text-[#242424]"
          >
            <FileEdit className="w-3 h-3" />
            임시저장
          </button>
        )}

        <label className="inline-flex items-center gap-1 text-[11px] text-[#616161] ml-auto">
          <Clock className="w-3.5 h-3.5 text-[#0F6CBD]" />
          <select
            aria-label="발송 지연 대기 시간"
            value={sendDelaySec}
            onChange={(e) => setSendDelaySec(Number(e.target.value))}
            className="px-1.5 py-0.5 bg-white border border-[#C8C6C4] rounded text-[11px]"
          >
            <option value={0}>즉시 발송</option>
            <option value={5}>5초 지연</option>
            <option value={10}>10초 지연</option>
          </select>
        </label>
      </div>

      {/* Outlook Header Block: Left tall Send Button + To / Cc / Subject / Attachments */}
      <div className="px-4 py-3 bg-white border-b border-[#E1DFDD] flex items-stretch gap-3 shrink-0">
        <button
          type="button"
          disabled={isSending}
          onClick={() => void handleSend('send')}
          className="w-[74px] shrink-0 rounded-[4px] border border-[#0F6CBD] bg-[#0F6CBD] hover:bg-[#115EA3] disabled:opacity-60 text-white flex flex-col items-center justify-center gap-1 text-xs font-bold shadow-xs transition-colors"
        >
          <Send className="w-4 h-4" />
          <span>{isSending ? '발송중' : '보내기(S)'}</span>
        </button>

        <div className="flex-1 min-w-0 space-y-1.5 text-xs">
          <div className="flex items-center gap-2 border-b border-[#EDEBE9] pb-1">
            <label
              htmlFor="compose-to"
              className="w-20 shrink-0 px-2 py-0.5 rounded bg-[#F3F2F1] border border-[#E1DFDD] text-center font-semibold text-[#424242]"
            >
              받는 사람(T)
            </label>
            <input
              id="compose-to"
              type="text"
              value={to}
              onChange={(e) => setTo(e.target.value)}
              placeholder="recipient@seastargo.com"
              className="flex-1 text-xs text-[#242424] bg-transparent focus:outline-none"
            />
          </div>

          <div className="flex items-center gap-2 border-b border-[#EDEBE9] pb-1">
            <label
              htmlFor="compose-cc"
              className="w-20 shrink-0 px-2 py-0.5 rounded bg-[#F3F2F1] border border-[#E1DFDD] text-center font-semibold text-[#424242]"
            >
              참조(C)
            </label>
            <input
              id="compose-cc"
              type="text"
              value={cc}
              onChange={(e) => setCc(e.target.value)}
              placeholder="cc@company.com"
              className="flex-1 text-xs text-[#242424] bg-transparent focus:outline-none"
            />
          </div>

          <div className="flex items-center gap-2">
            <label
              htmlFor="compose-subject"
              className="w-20 shrink-0 text-right pr-1 font-semibold text-[#616161]"
            >
              제목(U)
            </label>
            <input
              id="compose-subject"
              type="text"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              placeholder="제목 추가"
              className="flex-1 text-xs font-semibold text-[#242424] bg-transparent focus:outline-none"
            />
          </div>

          {draft.kind === 'forward' && realAttachments.length > 0 && (
            <div className="flex items-center gap-1.5 flex-wrap pt-1">
              <span className="text-[10.5px] font-semibold text-[#616161]">전달 첨부:</span>
              {realAttachments.map((att, idx) => (
                <span
                  key={att.id || idx}
                  className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#FAF9F8] border border-[#D1D1D1] text-[11px] text-[#242424]"
                >
                  <Paperclip className="w-3 h-3 text-[#0F6CBD]" />
                  <span>{att.filename}</span>
                </span>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Rich HTML Compose Canvas (Preserves Signature Business Card + Original Rich HTML) */}
      <div
        ref={editorRef}
        contentEditable
        suppressContentEditableWarning
        data-testid="compose-html-editor"
        onInput={() => setPlainPreview(editorRef.current?.innerText || '')}
        style={{ fontSize: fontSizePt }}
        className="flex-1 overflow-y-auto px-6 py-4 bg-white text-[#242424] focus:outline-none leading-relaxed select-text [&_img]:max-w-full [&_img]:h-auto [&_table]:border-collapse"
      />

      {/* Footer Status / Compact Pre-Send Coaching Bar */}
      <div className="px-4 py-2 bg-[#FAF9F8] border-t border-[#E1DFDD] space-y-1.5 shrink-0">
        {notice && (
          <div className="px-3 py-1.5 rounded bg-emerald-50 border border-emerald-200 text-xs text-emerald-800 font-medium">
            {notice}
          </div>
        )}
        {errorMsg && (
          <div className="px-3 py-1.5 rounded bg-rose-50 border border-rose-200 text-xs text-rose-800 font-medium">
            {errorMsg}
          </div>
        )}
        <div className="flex items-center justify-between gap-2 text-[11px] text-[#616161]">
          <div className="flex items-center gap-1.5 min-w-0 truncate">
            <ShieldCheck className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />
            <span className="truncate">
              {composeCoaching.warnings.length > 0
                ? `${composeCoaching.warnings[0].title}: ${composeCoaching.warnings[0].suggestion}`
                : '발송 전 점검 완료 · 명함 HTML 서명 및 원본 스레드 유지'}
            </span>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <span className="font-mono text-[10px]">
              {pop3ConnectedEmail ? `SMTP: ${pop3ConnectedEmail}` : 'PC Outlook COM 발송 대기'}
            </span>
            <button
              type="button"
              onClick={onClose}
              title="작성 창 닫기"
              className="p-1 rounded hover:bg-rose-50 text-[#616161] hover:text-rose-700"
            >
              <Trash2 className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      </div>
    </div>
  );

  if (mode === 'standalone') {
    return shellContent;
  }

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Outlook 메세지 작성 창"
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 p-3"
    >
      {shellContent}
    </div>
  );
};
