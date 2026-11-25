import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  BookUser,
  CheckCircle2,
  Clipboard,
  Clock,
  Copy,
  ExternalLink,
  FileEdit,
  Italic,
  Link2,
  List,
  ListOrdered,
  Paperclip,
  PenSquare,
  Scissors,
  Search,
  Send,
  ShieldCheck,
  Strikethrough,
  Table,
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

export interface OutlookContactEntry {
  name?: string;
  fullName?: string;
  displayName?: string;
  email: string;
  company?: string;
  department?: string;
}

function getContactDisplayName(c: OutlookContactEntry): string {
  return String(c.name || c.fullName || c.displayName || '').trim();
}

export interface ComposeLocalAttachment {
  id: string;
  filename: string;
  contentType: string;
  sizeBytes: number;
  contentBase64: string;
}

export interface ComposeDraftState {
  kind: 'new' | 'reply' | 'replyAll' | 'forward' | 'worklog';
  title: string;
  to: string;
  cc: string;
  bcc?: string;
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
  contacts?: OutlookContactEntry[];
  outlookPrefs: OutlookPreferences;
  pop3ConnectedEmail?: string | null;
  onSendSmtp?: (payload: {
    to: string;
    cc: string;
    bcc?: string;
    subject: string;
    body: string;
    htmlBody?: string;
    importance?: 'low' | 'normal' | 'high';
    requestReadReceipt?: boolean;
    attachments?: Array<{ filename: string; contentType?: string; contentBase64: string }>;
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

type ComposeRibbonTab = 'message' | 'insert' | 'options';
type AddressTargetField = 'to' | 'cc' | 'bcc';

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

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

async function readBrowserFileAsBase64(file: File): Promise<string> {
  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);
  const chunkSize = 0x8000;
  const parts: string[] = [];
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const sub = bytes.subarray(i, i + chunkSize);
    parts.push(String.fromCharCode(...sub));
  }
  return btoa(parts.join(''));
}

function appendAddressToken(existing: string, formattedEmail: string): string {
  const cleanNew = formattedEmail.trim();
  if (!cleanNew) return existing;
  const tokens = existing
    .split(/[;,]/)
    .map((t) => t.trim())
    .filter(Boolean);
  const lowerEmailMatch = cleanNew.match(/<([^>]+)>/);
  const targetAddr = (lowerEmailMatch ? lowerEmailMatch[1] : cleanNew).toLowerCase();
  const alreadyPresent = tokens.some((t) => t.toLowerCase().includes(targetAddr));
  if (alreadyPresent) return tokens.join('; ');
  return [...tokens, cleanNew].join('; ');
}

export const OutlookComposeWindow: React.FC<OutlookComposeWindowProps> = ({
  mode,
  draft,
  signatures,
  contacts = [],
  outlookPrefs,
  pop3ConnectedEmail,
  onSendSmtp,
  onSaveDraft,
  onPopOutNewWindow,
  onClose,
}) => {
  const [ribbonTab, setRibbonTab] = useState<ComposeRibbonTab>('message');
  const [to, setTo] = useState(draft.to);
  const [cc, setCc] = useState(draft.cc);
  const [bcc, setBcc] = useState(draft.bcc || '');
  const [showBcc, setShowBcc] = useState(Boolean(draft.bcc?.trim()));
  const [subject, setSubject] = useState(draft.subject);
  const [importance, setImportance] = useState<'low' | 'normal' | 'high'>(
    draft.importance || outlookPrefs.mail.defaultImportance || 'normal'
  );
  const [requestReadReceipt, setRequestReadReceipt] = useState<boolean>(
    Boolean(outlookPrefs.mail.requestReadReceipt)
  );
  const [sendDelaySec, setSendDelaySec] = useState<number>(outlookPrefs.mail.undoSendSeconds ?? 5);
  const [countdownLeft, setCountdownLeft] = useState<number | null>(null);
  const [fontFamily, setFontFamily] = useState<string>('맑은 고딕');
  const [fontSizePt, setFontSizePt] = useState<string>('10pt');
  const [textColor, setTextColor] = useState<string>('#242424');
  const [hiliteColor, setHiliteColor] = useState<string>('#FFF59D');

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

  const initialForwardAtts = useMemo(
    () =>
      draft.kind === 'forward'
        ? (draft.attachments || []).filter((att) => !isInlineBusinessCardImage(att))
        : [],
    [draft.kind, draft.attachments]
  );
  const [keptForwardAtts, setKeptForwardAtts] = useState<MailAttachmentMeta[]>(initialForwardAtts);
  useEffect(() => {
    setKeptForwardAtts(initialForwardAtts);
  }, [initialForwardAtts]);

  const [userAttachments, setUserAttachments] = useState<ComposeLocalAttachment[]>([]);
  const [isDraggingFile, setIsDraggingFile] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [addressBookOpen, setAddressBookOpen] = useState(false);
  const [addressTarget, setAddressTarget] = useState<AddressTargetField>('to');
  const [addressSearch, setAddressSearch] = useState('');
  const [focusedField, setFocusedField] = useState<AddressTargetField | null>(null);

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

  const handleAttachFiles = async (fileList: FileList | File[]) => {
    setErrorMsg(null);
    const list = Array.from(fileList);
    if (list.length === 0) return;

    const existingBytes = userAttachments.reduce((sum, a) => sum + a.sizeBytes, 0);
    let runningBytes = existingBytes;
    const nextItems: ComposeLocalAttachment[] = [];

    for (const file of list.slice(0, 10 - userAttachments.length)) {
      if (runningBytes + file.size > 20 * 1024 * 1024) {
        setErrorMsg('첨부 파일 합계 용량은 20MB 이하까지만 추가할 수 있습니다.');
        break;
      }
      runningBytes += file.size;
      const b64 = await readBrowserFileAsBase64(file);
      nextItems.push({
        id: `ua-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        filename: file.name,
        contentType: file.type || 'application/octet-stream',
        sizeBytes: file.size,
        contentBase64: b64,
      });
    }
    if (nextItems.length > 0) {
      setUserAttachments((prev) => [...prev, ...nextItems]);
    }
  };

  const totalAttachmentCount = keptForwardAtts.length + userAttachments.length;

  const composeCoaching = useMemo(() => {
    return coachEmailDraft({
      to: to.trim() || cc.trim() || bcc.trim(),
      subject: subject.trim(),
      body: plainPreview || draft.introText || '',
      hasAttachment: totalAttachmentCount > 0,
    });
  }, [to, cc, bcc, subject, plainPreview, draft.introText, totalAttachmentCount]);

  const collectForwardAttachmentIndices = (): number[] => {
    if (draft.kind !== 'forward') return [];
    return keptForwardAtts
      .map((att, idx) => {
        const m = String(att.id || '').match(/-att-(\d+)$/);
        return m ? Number(m[1]) : idx;
      })
      .filter((n) => Number.isInteger(n) && n >= 0);
  };

  const executeSendNow = async (comAction: 'send' | 'display' = 'send') => {
    setCountdownLeft(null);
    setErrorMsg(null);
    setNotice(null);
    const htmlBody = editorRef.current?.innerHTML || '';
    const textBody = editorRef.current?.innerText || plainPreview || '';
    const cleanSubject = subject.trim() || '(제목 없음)';

    if (comAction === 'send' && !to.trim() && !cc.trim() && !bcc.trim()) {
      setErrorMsg('받는 사람 이메일 주소를 입력해 주세요.');
      return;
    }

    setIsSending(true);
    try {
      const payloadAttachments = userAttachments.map((a) => ({
        filename: a.filename,
        contentType: a.contentType,
        contentBase64: a.contentBase64,
      }));

      if (comAction === 'send' && pop3ConnectedEmail && onSendSmtp && keptForwardAtts.length === 0) {
        const smtpRes = await onSendSmtp({
          to: to.trim(),
          cc: cc.trim(),
          bcc: bcc.trim() || undefined,
          subject: cleanSubject,
          body: textBody,
          htmlBody,
          importance,
          requestReadReceipt,
          attachments: payloadAttachments.length > 0 ? payloadAttachments : undefined,
        });
        if (smtpRes.ok) {
          setNotice(smtpRes.message);
          window.setTimeout(() => onClose(), 1000);
          return;
        }
        setErrorMsg(smtpRes.message);
        return;
      }

      // Send or open directly via PC Outlook COM (preserves HTML business card, inline images, forwarded PST files & local files)
      const res = await fetch('/api/local/outlook/send-via-com', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          to: to.trim(),
          cc: cc.trim(),
          bcc: bcc.trim(),
          subject: cleanSubject,
          body: textBody,
          htmlBody,
          action: comAction,
          importance,
          requestReadReceipt,
          sourceEmailId: draft.sourceEmailId,
          forwardAttachmentIndices: collectForwardAttachmentIndices(),
          userAttachments: payloadAttachments,
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

  // Real Undo-Send Countdown Loop when sendDelaySec > 0
  useEffect(() => {
    if (countdownLeft === null) return;
    if (countdownLeft <= 0) {
      void executeSendNow('send');
      return;
    }
    const timer = window.setTimeout(() => {
      setCountdownLeft((prev) => (prev === null ? null : prev - 1));
    }, 1000);
    return () => window.clearTimeout(timer);
  }, [countdownLeft]);

  const handleRequestSend = () => {
    setErrorMsg(null);
    if (!to.trim() && !cc.trim() && !bcc.trim()) {
      setErrorMsg('받는 사람 이메일 주소를 입력해 주세요.');
      return;
    }
    if (sendDelaySec > 0) {
      setCountdownLeft(sendDelaySec);
      return;
    }
    void executeSendNow('send');
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

  const insertContactIntoField = (contact: OutlookContactEntry, field: AddressTargetField) => {
    const contactName = getContactDisplayName(contact);
    const formatted = contactName
      ? `${contactName.replace(/[<>]/g, '')} <${contact.email}>`
      : contact.email;
    if (field === 'to') setTo((prev) => appendAddressToken(prev, formatted));
    else if (field === 'cc') setCc((prev) => appendAddressToken(prev, formatted));
    else {
      setShowBcc(true);
      setBcc((prev) => appendAddressToken(prev, formatted));
    }
  };

  const filteredAddressContacts = useMemo(() => {
    const q = addressSearch.trim().toLowerCase();
    const list = q
      ? contacts.filter(
          (c) =>
            getContactDisplayName(c).toLowerCase().includes(q) ||
            c.email.toLowerCase().includes(q) ||
            (c.company || '').toLowerCase().includes(q) ||
            (c.department || '').toLowerCase().includes(q)
        )
      : contacts;
    return list.slice(0, 60);
  }, [contacts, addressSearch]);

  const activeFieldValue = focusedField === 'to' ? to : focusedField === 'cc' ? cc : focusedField === 'bcc' ? bcc : '';
  const lastTokenQuery = activeFieldValue.split(/[;,]/).pop()?.trim().toLowerCase() || '';
  const autocompleteSuggestions = useMemo(() => {
    if (!focusedField || lastTokenQuery.length < 1 || contacts.length === 0) return [];
    return contacts
      .filter(
        (c) =>
          getContactDisplayName(c).toLowerCase().includes(lastTokenQuery) ||
          c.email.toLowerCase().includes(lastTokenQuery) ||
          (c.company || '').toLowerCase().includes(lastTokenQuery)
      )
      .slice(0, 6);
  }, [focusedField, lastTokenQuery, contacts]);

  const applyAutocompletePick = (contact: OutlookContactEntry, field: AddressTargetField) => {
    const contactName = getContactDisplayName(contact);
    const formatted = contactName
      ? `${contactName.replace(/[<>]/g, '')} <${contact.email}>`
      : contact.email;
    const replaceTail = (raw: string) => {
      const parts = raw.split(/[;,]/);
      parts.pop();
      const prefix = parts.map((p) => p.trim()).filter(Boolean).join('; ');
      return appendAddressToken(prefix, formatted);
    };
    if (field === 'to') setTo(replaceTail);
    else if (field === 'cc') setCc(replaceTail);
    else setBcc(replaceTail);
    setFocusedField(null);
  };

  const handleInsertLink = () => {
    const url = window.prompt('삽입할 웹/도면 URL을 입력하세요 (https://...)', 'https://');
    if (url && /^https?:\/\//i.test(url.trim())) {
      execFormat('createLink', url.trim());
    }
  };

  const handleInsertSimpleTable = () => {
    editorRef.current?.focus();
    const html =
      '<table style="border-collapse:collapse;width:100%;max-width:560px;margin:8px 0;font-size:10pt;">' +
      '<tr><th style="border:1px solid #C8C6C4;background:#F3F2F1;padding:4px 8px;text-align:left;">항목</th>' +
      '<th style="border:1px solid #C8C6C4;background:#F3F2F1;padding:4px 8px;text-align:left;">내용</th>' +
      '<th style="border:1px solid #C8C6C4;background:#F3F2F1;padding:4px 8px;text-align:left;">비고</th></tr>' +
      '<tr><td style="border:1px solid #C8C6C4;padding:4px 8px;">&nbsp;</td>' +
      '<td style="border:1px solid #C8C6C4;padding:4px 8px;">&nbsp;</td>' +
      '<td style="border:1px solid #C8C6C4;padding:4px 8px;">&nbsp;</td></tr></table><div><br></div>';
    execFormat('insertHTML', html);
  };

  const shellContent = (
    <div
      data-testid="outlook-compose-window"
      onKeyDown={(e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
          e.preventDefault();
          handleRequestSend();
        } else if (e.altKey && (e.key === 's' || e.key === 'S')) {
          e.preventDefault();
          handleRequestSend();
        }
      }}
      onDragOver={(e) => {
        if (e.dataTransfer?.types?.includes('Files')) {
          e.preventDefault();
          setIsDraggingFile(true);
        }
      }}
      onDragLeave={() => setIsDraggingFile(false)}
      onDrop={(e) => {
        e.preventDefault();
        setIsDraggingFile(false);
        if (e.dataTransfer?.files?.length) {
          void handleAttachFiles(e.dataTransfer.files);
        }
      }}
      className={
        mode === 'standalone'
          ? 'relative w-screen h-screen bg-white flex flex-col overflow-hidden'
          : 'relative w-[min(1120px,96vw)] h-[min(880px,92vh)] bg-white rounded-lg border border-[#C8C6C4] shadow-2xl flex flex-col overflow-hidden'
      }
    >
      <input
        ref={fileInputRef}
        type="file"
        multiple
        className="hidden"
        onChange={(e) => {
          if (e.target.files) void handleAttachFiles(e.target.files);
          e.target.value = '';
        }}
      />

      {isDraggingFile && (
        <div className="absolute inset-0 z-30 bg-[#0F6CBD]/10 border-2 border-dashed border-[#0F6CBD] flex items-center justify-center pointer-events-none">
          <div className="px-5 py-3 rounded-lg bg-white shadow-lg border border-[#0F6CBD] text-sm font-bold text-[#0F6CBD]">
            여기에 파일을 놓으면 바로 메일에 첨부됩니다
          </div>
        </div>
      )}

      {/* Classic Outlook Window Title Bar */}
      <div className="px-4 py-1.5 bg-[#0F6CBD] text-white flex items-center justify-between gap-2 shrink-0 select-none">
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
                  bcc,
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

      {/* Classic Outlook Ribbon Tabs Bar */}
      <div className="px-3 pt-1 bg-[#F3F2F1] border-b border-[#E1DFDD] flex items-center gap-1 text-[11px] shrink-0 select-none">
        {(
          [
            { id: 'message', label: '메시지' },
            { id: 'insert', label: '삽입' },
            { id: 'options', label: '옵션' },
          ] as const
        ).map((tab) => (
          <button
            key={tab.id}
            type="button"
            onClick={() => setRibbonTab(tab.id)}
            className={`px-3 py-1 rounded-t border-t border-x transition-colors ${
              ribbonTab === tab.id
                ? 'bg-white border-[#D1D1D1] font-bold text-[#0F6CBD] -mb-px pb-1.5'
                : 'border-transparent text-[#424242] hover:bg-white/60'
            }`}
          >
            {tab.label}
          </button>
        ))}
        <span className="ml-auto text-[10px] text-[#616161] pr-1 hidden sm:inline">
          단축키: Ctrl+Enter / Alt+S 보내기
        </span>
      </div>

      {/* Classic Outlook Ribbon Commands Band */}
      <div className="px-3 py-1.5 bg-white border-b border-[#D1D1D1] flex items-stretch gap-2 flex-wrap text-xs shrink-0">
        {ribbonTab === 'message' && (
          <>
            {/* Group: 클립보드 */}
            <div className="flex flex-col justify-between pr-2.5 border-r border-[#E1DFDD]">
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => execFormat('paste')}
                  title="붙여넣기"
                  className="px-2 py-1 rounded hover:bg-[#F3F2F1] flex items-center gap-1 text-[11px] text-[#242424]"
                >
                  <Clipboard className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  <span>붙여넣기</span>
                </button>
                <button
                  type="button"
                  onClick={() => execFormat('copy')}
                  title="복사"
                  className="p-1 rounded hover:bg-[#F3F2F1] text-[#424242]"
                >
                  <Copy className="w-3.5 h-3.5" />
                </button>
                <button
                  type="button"
                  onClick={() => execFormat('cut')}
                  title="잘라내기"
                  className="p-1 rounded hover:bg-[#F3F2F1] text-[#424242]"
                >
                  <Scissors className="w-3.5 h-3.5" />
                </button>
              </div>
              <span className="text-[9.5px] text-[#8A8886] text-center leading-none pt-0.5">
                클립보드
              </span>
            </div>

            {/* Group: 기본 텍스트 */}
            <div className="flex flex-col justify-between pr-2.5 border-r border-[#E1DFDD]">
              <div className="flex items-center gap-1 flex-wrap">
                <select
                  aria-label="글꼴"
                  value={fontFamily}
                  onChange={(e) => {
                    setFontFamily(e.target.value);
                    execFormat('fontName', e.target.value);
                  }}
                  className="px-1.5 py-0.5 rounded border border-[#C8C6C4] bg-white text-[11px] text-[#242424]"
                >
                  <option value="맑은 고딕">맑은 고딕 (본문)</option>
                  <option value="굴림">굴림</option>
                  <option value="바탕">바탕</option>
                  <option value="Segoe UI">Segoe UI</option>
                  <option value="Arial">Arial</option>
                </select>
                <select
                  aria-label="글꼴 크기"
                  value={fontSizePt}
                  onChange={(e) => {
                    const val = e.target.value;
                    setFontSizePt(val);
                    const mapSize: Record<string, string> = {
                      '9pt': '1',
                      '10pt': '2',
                      '11pt': '3',
                      '12pt': '3',
                      '14pt': '4',
                      '16pt': '5',
                    };
                    execFormat('fontSize', mapSize[val] || '2');
                  }}
                  className="px-1.5 py-0.5 rounded border border-[#C8C6C4] bg-white text-[11px] text-[#242424]"
                >
                  <option value="9pt">9</option>
                  <option value="10pt">10</option>
                  <option value="11pt">11</option>
                  <option value="12pt">12</option>
                  <option value="14pt">14</option>
                  <option value="16pt">16</option>
                </select>

                <div className="inline-flex items-center rounded border border-[#C8C6C4] bg-white overflow-hidden">
                  <button type="button" onClick={() => execFormat('bold')} title="굵게" className="px-1.5 py-1 hover:bg-[#EDEBE9]">
                    <Bold className="w-3.5 h-3.5" />
                  </button>
                  <button type="button" onClick={() => execFormat('italic')} title="기울임꼴" className="px-1.5 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]">
                    <Italic className="w-3.5 h-3.5" />
                  </button>
                  <button type="button" onClick={() => execFormat('underline')} title="밑줄" className="px-1.5 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]">
                    <Underline className="w-3.5 h-3.5" />
                  </button>
                  <button type="button" onClick={() => execFormat('strikeThrough')} title="취소선" className="px-1.5 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]">
                    <Strikethrough className="w-3.5 h-3.5" />
                  </button>
                </div>

                <label title="글꼴 색" className="inline-flex items-center gap-0.5 px-1 py-0.5 rounded border border-[#C8C6C4] bg-white cursor-pointer">
                  <span className="text-[10px] font-bold" style={{ color: textColor }}>
                    A
                  </span>
                  <input
                    type="color"
                    aria-label="글꼴 색"
                    value={textColor}
                    onChange={(e) => {
                      setTextColor(e.target.value);
                      execFormat('foreColor', e.target.value);
                    }}
                    className="w-3.5 h-3.5 border-0 p-0 bg-transparent cursor-pointer"
                  />
                </label>

                <label title="강조 표시 색" className="inline-flex items-center gap-0.5 px-1 py-0.5 rounded border border-[#C8C6C4] bg-white cursor-pointer">
                  <span className="text-[10px] font-bold px-0.5 rounded" style={{ backgroundColor: hiliteColor }}>
                    ab
                  </span>
                  <input
                    type="color"
                    aria-label="텍스트 강조 색"
                    value={hiliteColor}
                    onChange={(e) => {
                      setHiliteColor(e.target.value);
                      execFormat('hiliteColor', e.target.value);
                    }}
                    className="w-3.5 h-3.5 border-0 p-0 bg-transparent cursor-pointer"
                  />
                </label>

                <div className="inline-flex items-center rounded border border-[#C8C6C4] bg-white overflow-hidden">
                  <button type="button" onClick={() => execFormat('insertUnorderedList')} title="글머리 기호" className="px-1.5 py-1 hover:bg-[#EDEBE9]">
                    <List className="w-3.5 h-3.5" />
                  </button>
                  <button type="button" onClick={() => execFormat('insertOrderedList')} title="번호 매기기" className="px-1.5 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]">
                    <ListOrdered className="w-3.5 h-3.5" />
                  </button>
                  <button type="button" onClick={() => execFormat('justifyLeft')} title="왼쪽 맞춤" className="px-1.5 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]">
                    <AlignLeft className="w-3.5 h-3.5" />
                  </button>
                  <button type="button" onClick={() => execFormat('justifyCenter')} title="가운데 맞춤" className="px-1.5 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]">
                    <AlignCenter className="w-3.5 h-3.5" />
                  </button>
                  <button type="button" onClick={() => execFormat('justifyRight')} title="오른쪽 맞춤" className="px-1.5 py-1 border-l border-[#E1DFDD] hover:bg-[#EDEBE9]">
                    <AlignRight className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
              <span className="text-[9.5px] text-[#8A8886] text-center leading-none pt-0.5">
                기본 텍스트
              </span>
            </div>

            {/* Group: 이름 (주소록) */}
            <div className="flex flex-col justify-between pr-2.5 border-r border-[#E1DFDD]">
              <button
                type="button"
                onClick={() => {
                  setAddressTarget('to');
                  setAddressBookOpen(true);
                }}
                aria-label="Outlook 주소록 열기"
                className="px-2.5 py-1 rounded border border-[#C8C6C4] bg-[#FAF9F8] hover:bg-[#EBF3FC] hover:border-[#0F6CBD] text-[11px] font-semibold text-[#242424] flex items-center gap-1"
              >
                <BookUser className="w-3.5 h-3.5 text-[#0F6CBD]" />
                <span>주소록 ({contacts.length})</span>
              </button>
              <span className="text-[9.5px] text-[#8A8886] text-center leading-none pt-0.5">
                이름
              </span>
            </div>

            {/* Group: 포함 (첨부 + 링크 + 명함 서명) */}
            <div className="flex flex-col justify-between pr-2.5 border-r border-[#E1DFDD]">
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  aria-label="파일 첨부"
                  className="px-2 py-1 rounded border border-[#C8C6C4] bg-[#FAF9F8] hover:bg-[#EBF3FC] hover:border-[#0F6CBD] text-[11px] font-semibold text-[#242424] flex items-center gap-1"
                >
                  <Paperclip className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  <span>파일 첨부</span>
                </button>
                <button
                  type="button"
                  onClick={handleInsertLink}
                  title="하이퍼링크 삽입"
                  className="px-2 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] text-[#242424] flex items-center gap-1"
                >
                  <Link2 className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  <span>링크</span>
                </button>
                <select
                  aria-label="Outlook 명함 서명 선택"
                  value={selectedSigName}
                  onChange={(e) => handleApplySignatureChange(e.target.value)}
                  className="px-2 py-1 rounded border border-[#C8C6C4] bg-white text-[11px] text-[#242424] max-w-[175px]"
                >
                  {signatures.length === 0 ? (
                    <option value="">기본 서명</option>
                  ) : (
                    signatures.map((s) => (
                      <option key={s.name} value={s.name}>
                        서명: {s.name}
                        {s.isDefaultNew ? ' (기본)' : ''}
                      </option>
                    ))
                  )}
                  <option value="__none__">(서명 없음)</option>
                </select>
              </div>
              <span className="text-[9.5px] text-[#8A8886] text-center leading-none pt-0.5">
                포함
              </span>
            </div>

            {/* Group: 태그 & Outlook COM */}
            <div className="flex flex-col justify-between">
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => setImportance((v) => (v === 'high' ? 'normal' : 'high'))}
                  className={`px-2 py-1 rounded border text-[11px] font-semibold ${
                    importance === 'high'
                      ? 'bg-rose-50 border-rose-300 text-rose-700'
                      : 'bg-white border-[#C8C6C4] text-[#424242] hover:bg-[#F3F2F1]'
                  }`}
                >
                  ! 중요도 높음
                </button>
                <button
                  type="button"
                  onClick={() => setImportance((v) => (v === 'low' ? 'normal' : 'low'))}
                  className={`px-2 py-1 rounded border text-[11px] font-semibold ${
                    importance === 'low'
                      ? 'bg-sky-50 border-sky-300 text-sky-700'
                      : 'bg-white border-[#C8C6C4] text-[#424242] hover:bg-[#F3F2F1]'
                  }`}
                >
                  ↓ 중요도 낮음
                </button>
                <button
                  type="button"
                  onClick={() => void executeSendNow('display')}
                  disabled={isSending}
                  title="PC Outlook 실제 창으로 초안 바로 띄우기"
                  className="px-2 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] font-semibold text-[#0F6CBD]"
                >
                  Outlook으로 띄우기
                </button>
                {onSaveDraft && (
                  <button
                    type="button"
                    onClick={handleSave}
                    className="inline-flex items-center gap-1 px-2 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] font-medium text-[#242424]"
                  >
                    <FileEdit className="w-3 h-3" />
                    임시저장
                  </button>
                )}
              </div>
              <span className="text-[9.5px] text-[#8A8886] text-center leading-none pt-0.5">
                태그 · 발송 연동
              </span>
            </div>
          </>
        )}

        {ribbonTab === 'insert' && (
          <div className="flex items-center gap-2 py-0.5">
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="px-2.5 py-1 rounded border border-[#C8C6C4] bg-[#FAF9F8] hover:bg-[#EBF3FC] text-[11px] font-semibold text-[#242424] flex items-center gap-1"
            >
              <Paperclip className="w-3.5 h-3.5 text-[#0F6CBD]" />
              파일 첨부
            </button>
            <button
              type="button"
              onClick={handleInsertSimpleTable}
              className="px-2.5 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] font-medium text-[#242424] flex items-center gap-1"
            >
              <Table className="w-3.5 h-3.5 text-[#0F6CBD]" />
              표 (3열) 삽입
            </button>
            <button
              type="button"
              onClick={handleInsertLink}
              className="px-2.5 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] font-medium text-[#242424] flex items-center gap-1"
            >
              <Link2 className="w-3.5 h-3.5 text-[#0F6CBD]" />
              웹·도면 링크
            </button>
            <button
              type="button"
              onClick={() => execFormat('insertHorizontalRule')}
              className="px-2.5 py-1 rounded border border-[#C8C6C4] bg-white hover:bg-[#EDEBE9] text-[11px] font-medium text-[#242424]"
            >
              가로 구분선 삽입
            </button>
          </div>
        )}

        {ribbonTab === 'options' && (
          <div className="flex items-center gap-3 py-0.5 flex-wrap">
            <button
              type="button"
              onClick={() => setShowBcc((v) => !v)}
              className={`px-2.5 py-1 rounded border text-[11px] font-semibold ${
                showBcc
                  ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD]'
                  : 'bg-white border-[#C8C6C4] text-[#424242]'
              }`}
            >
              숨은 참조(B) {showBcc ? '표시 중' : '표시'}
            </button>

            <label className="inline-flex items-center gap-1.5 text-[11px] font-medium text-[#242424] cursor-pointer">
              <input
                type="checkbox"
                checked={requestReadReceipt}
                onChange={(e) => setRequestReadReceipt(e.target.checked)}
                className="rounded border-[#C8C6C4] text-[#0F6CBD]"
              />
              <span>읽음 확인 요청 (Disposition-Notification)</span>
            </label>

            <label className="inline-flex items-center gap-1 text-[11px] text-[#424242]">
              <Clock className="w-3.5 h-3.5 text-[#0F6CBD]" />
              <span>발송 지연 취소:</span>
              <select
                aria-label="발송 지연 대기 시간"
                value={sendDelaySec}
                onChange={(e) => setSendDelaySec(Number(e.target.value))}
                className="px-2 py-0.5 bg-white border border-[#C8C6C4] rounded text-[11px]"
              >
                <option value={0}>즉시 발송 (0초)</option>
                <option value={5}>5초 지연</option>
                <option value={10}>10초 지연</option>
                <option value={30}>30초 지연</option>
              </select>
            </label>
          </div>
        )}
      </div>

      {/* Real Undo-Send Countdown Banner */}
      {countdownLeft !== null && (
        <div
          role="status"
          className="px-4 py-2 bg-amber-50 border-b border-amber-200 flex items-center justify-between gap-2 text-xs shrink-0"
        >
          <span className="font-semibold text-amber-900 flex items-center gap-1.5">
            <Clock className="w-4 h-4 text-amber-600 animate-pulse" />
            {countdownLeft}초 후 메일이 발송됩니다 ({subject.trim() || '(제목 없음)'})
          </span>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => void executeSendNow('send')}
              className="px-2.5 py-1 rounded bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-[11px] font-semibold"
            >
              지금 즉시 발송
            </button>
            <button
              type="button"
              onClick={() => {
                setCountdownLeft(null);
                setNotice('발송 대기를 취소했습니다. 내용을 이어서 편집할 수 있습니다.');
              }}
              className="px-2.5 py-1 rounded bg-white border border-amber-400 hover:bg-amber-100 text-amber-900 text-[11px] font-bold"
            >
              발송 취소
            </button>
          </div>
        </div>
      )}

      {/* Outlook Header Block: Left tall Send Button + To / Cc / Bcc / Subject / Attachments */}
      <div className="px-4 py-2.5 bg-white border-b border-[#E1DFDD] flex items-stretch gap-3 shrink-0">
        <button
          type="button"
          disabled={isSending || countdownLeft !== null}
          onClick={handleRequestSend}
          className="w-[74px] shrink-0 rounded-[4px] border border-[#0F6CBD] bg-[#0F6CBD] hover:bg-[#115EA3] disabled:opacity-60 text-white flex flex-col items-center justify-center gap-1 text-xs font-bold shadow-xs transition-colors"
        >
          <Send className="w-4 h-4" />
          <span>{isSending ? '발송중' : countdownLeft !== null ? `${countdownLeft}s` : '보내기(S)'}</span>
        </button>

        <div className="flex-1 min-w-0 space-y-1.5 text-xs relative">
          {/* To Field */}
          <div className="flex items-center gap-2 border-b border-[#EDEBE9] pb-1 relative">
            <button
              type="button"
              onClick={() => {
                setAddressTarget('to');
                setAddressBookOpen(true);
              }}
              title="클릭하여 주소록에서 받는 사람 선택"
              className="w-20 shrink-0 px-2 py-0.5 rounded bg-[#F3F2F1] hover:bg-[#E1DFDD] border border-[#E1DFDD] text-center font-semibold text-[#424242]"
            >
              받는 사람(T)
            </button>
            <input
              id="compose-to"
              type="text"
              value={to}
              onFocus={() => setFocusedField('to')}
              onBlur={() => window.setTimeout(() => setFocusedField((f) => (f === 'to' ? null : f)), 180)}
              onChange={(e) => setTo(e.target.value)}
              placeholder="recipient@seastargo.com (; 또는 , 로 여러 명 구분)"
              className="flex-1 text-xs text-[#242424] bg-transparent focus:outline-none"
            />
          </div>

          {/* Cc Field + Toggle Bcc */}
          <div className="flex items-center gap-2 border-b border-[#EDEBE9] pb-1 relative">
            <button
              type="button"
              onClick={() => {
                setAddressTarget('cc');
                setAddressBookOpen(true);
              }}
              title="클릭하여 주소록에서 참조 선택"
              className="w-20 shrink-0 px-2 py-0.5 rounded bg-[#F3F2F1] hover:bg-[#E1DFDD] border border-[#E1DFDD] text-center font-semibold text-[#424242]"
            >
              참조(C)
            </button>
            <input
              id="compose-cc"
              type="text"
              value={cc}
              onFocus={() => setFocusedField('cc')}
              onBlur={() => window.setTimeout(() => setFocusedField((f) => (f === 'cc' ? null : f)), 180)}
              onChange={(e) => setCc(e.target.value)}
              placeholder="cc@company.com"
              className="flex-1 text-xs text-[#242424] bg-transparent focus:outline-none"
            />
            {!showBcc && (
              <button
                type="button"
                onClick={() => setShowBcc(true)}
                className="px-1.5 py-0.5 rounded text-[10.5px] font-medium text-[#0F6CBD] hover:bg-[#EBF3FC] shrink-0"
              >
                + 숨은 참조(B)
              </button>
            )}
          </div>

          {/* Bcc Field (Shown when toggled or non-empty) */}
          {showBcc && (
            <div className="flex items-center gap-2 border-b border-[#EDEBE9] pb-1 relative">
              <button
                type="button"
                onClick={() => {
                  setAddressTarget('bcc');
                  setAddressBookOpen(true);
                }}
                title="클릭하여 주소록에서 숨은 참조 선택"
                className="w-20 shrink-0 px-2 py-0.5 rounded bg-[#F3F2F1] hover:bg-[#E1DFDD] border border-[#E1DFDD] text-center font-semibold text-[#424242]"
              >
                숨은 참조(B)
              </button>
              <input
                id="compose-bcc"
                type="text"
                value={bcc}
                onFocus={() => setFocusedField('bcc')}
                onBlur={() => window.setTimeout(() => setFocusedField((f) => (f === 'bcc' ? null : f)), 180)}
                onChange={(e) => setBcc(e.target.value)}
                placeholder="bcc@company.com (수신자에게 비공개)"
                className="flex-1 text-xs text-[#242424] bg-transparent focus:outline-none"
              />
            </div>
          )}

          {/* Contact Autocomplete Dropdown */}
          {focusedField && autocompleteSuggestions.length > 0 && (
            <div
              role="listbox"
              aria-label="Outlook 주소록 자동완성"
              className="absolute left-22 right-0 top-14 z-30 bg-white border border-[#C8C6C4] rounded-md shadow-lg max-h-40 overflow-y-auto"
            >
              {autocompleteSuggestions.map((c) => {
                const cname = getContactDisplayName(c);
                return (
                  <button
                    key={`${c.email}-${cname}`}
                    type="button"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      applyAutocompletePick(c, focusedField);
                    }}
                    className="w-full px-3 py-1.5 text-left text-[11px] hover:bg-[#EBF3FC] flex items-center justify-between gap-2 border-b last:border-b-0 border-[#F3F2F1]"
                  >
                    <span className="font-semibold text-[#242424] truncate">{cname || c.email}</span>
                    <span className="font-mono text-[#0F6CBD] truncate">{c.email}</span>
                  </button>
                );
              })}
            </div>
          )}

          {/* Subject Field */}
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

          {/* Combined Attachments Strip (Forwarded PST + Newly Uploaded Local Files) */}
          {totalAttachmentCount > 0 && (
            <div className="flex items-center gap-1.5 flex-wrap pt-1">
              <span className="text-[10.5px] font-semibold text-[#616161]">
                첨부 ({totalAttachmentCount}):
              </span>
              {keptForwardAtts.map((att, idx) => (
                <span
                  key={att.id || `fwd-${idx}`}
                  className="inline-flex items-center gap-1 pl-2 pr-1 py-0.5 rounded bg-[#FAF9F8] border border-[#D1D1D1] text-[11px] text-[#242424]"
                >
                  <Paperclip className="w-3 h-3 text-[#0F6CBD]" />
                  <span className="max-w-[200px] truncate">{att.filename}</span>
                  <button
                    type="button"
                    aria-label={`${att.filename} 첨부 제거`}
                    onClick={() =>
                      setKeptForwardAtts((prev) => prev.filter((_, i) => i !== idx))
                    }
                    className="p-0.5 rounded hover:bg-rose-100 text-[#616161] hover:text-rose-700"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))}
              {userAttachments.map((ua) => (
                <span
                  key={ua.id}
                  className="inline-flex items-center gap-1 pl-2 pr-1 py-0.5 rounded bg-[#EBF3FC] border border-[#B4D6FA] text-[11px] text-[#0F6CBD] font-medium"
                >
                  <Paperclip className="w-3 h-3 shrink-0" />
                  <span className="max-w-[180px] truncate">{ua.filename}</span>
                  <span className="text-[10px] text-[#424242]">({formatFileSize(ua.sizeBytes)})</span>
                  <button
                    type="button"
                    aria-label={`${ua.filename} 첨부 제거`}
                    onClick={() =>
                      setUserAttachments((prev) => prev.filter((item) => item.id !== ua.id))
                    }
                    className="p-0.5 rounded hover:bg-rose-100 text-[#0F6CBD] hover:text-rose-700"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Inline Address Book Picker Drawer Modal */}
      {addressBookOpen && (
        <div
          role="dialog"
          aria-label="Outlook 주소록 선택기"
          className="absolute inset-x-6 top-28 bottom-16 z-40 bg-white border border-[#0F6CBD] rounded-lg shadow-2xl flex flex-col overflow-hidden"
        >
          <div className="px-4 py-2 bg-[#0F6CBD] text-white flex items-center justify-between gap-2 text-xs">
            <span className="font-bold flex items-center gap-1.5">
              <BookUser className="w-4 h-4" />
              Outlook 주소록 ({contacts.length}명) · 대상:{' '}
              {addressTarget === 'to' ? '받는 사람(T)' : addressTarget === 'cc' ? '참조(C)' : '숨은 참조(B)'}
            </span>
            <button
              type="button"
              onClick={() => setAddressBookOpen(false)}
              aria-label="주소록 닫기"
              className="p-1 rounded hover:bg-white/20"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
          <div className="p-2.5 bg-[#FAF9F8] border-b border-[#E1DFDD] flex items-center gap-2 flex-wrap">
            <div className="flex-1 min-w-[200px] flex items-center gap-1.5 px-2.5 py-1 bg-white border border-[#C8C6C4] rounded">
              <Search className="w-3.5 h-3.5 text-[#616161]" />
              <input
                type="text"
                value={addressSearch}
                onChange={(e) => setAddressSearch(e.target.value)}
                placeholder="이름, 이메일, 회사 또는 부서 검색..."
                className="flex-1 text-xs focus:outline-none"
              />
            </div>
            <div className="flex items-center gap-1 text-[11px]">
              {(['to', 'cc', 'bcc'] as const).map((field) => (
                <button
                  key={field}
                  type="button"
                  onClick={() => setAddressTarget(field)}
                  className={`px-2 py-1 rounded border font-semibold ${
                    addressTarget === field
                      ? 'bg-[#0F6CBD] text-white border-[#0F6CBD]'
                      : 'bg-white text-[#424242] border-[#C8C6C4]'
                  }`}
                >
                  {field === 'to' ? '받는 사람' : field === 'cc' ? '참조' : '숨은 참조'}
                </button>
              ))}
            </div>
          </div>
          <div className="flex-1 overflow-y-auto divide-y divide-[#F0F0F0] text-xs">
            {filteredAddressContacts.length === 0 ? (
              <div className="p-6 text-center text-[#707070]">
                {contacts.length === 0
                  ? '가져온 Outlook 연락처가 없습니다. 일반 설정에서 [아웃룩 현재 설정 탐색·적용]을 실행하세요.'
                  : '검색어와 일치하는 연락처가 없습니다.'}
              </div>
            ) : (
              filteredAddressContacts.map((c) => {
                const cname = getContactDisplayName(c);
                return (
                  <div
                    key={`${c.email}-${cname}`}
                    className="px-4 py-2 flex items-center justify-between gap-2 hover:bg-[#FAF9F8]"
                  >
                    <div className="min-w-0">
                      <div className="font-semibold text-[#242424] truncate">
                        {cname || c.email}
                        {c.company ? (
                          <span className="ml-1.5 text-[11px] font-normal text-[#616161]">
                            ({c.company}
                            {c.department ? ` · ${c.department}` : ''})
                          </span>
                        ) : null}
                      </div>
                      <div className="font-mono text-[11px] text-[#0F6CBD] truncate">{c.email}</div>
                    </div>
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        onClick={() => insertContactIntoField(c, addressTarget)}
                        className="px-2.5 py-1 rounded bg-[#EBF3FC] hover:bg-[#0F6CBD] hover:text-white text-[#0F6CBD] text-[11px] font-semibold transition-colors"
                      >
                        추가
                      </button>
                    </div>
                  </div>
                );
              })
            )}
          </div>
          <div className="px-4 py-2 bg-[#FAF9F8] border-t border-[#E1DFDD] flex items-center justify-between text-[11px] text-[#616161]">
            <span>받는 사람: {to || '(없음)'}</span>
            <button
              type="button"
              onClick={() => setAddressBookOpen(false)}
              className="px-3 py-1 rounded bg-[#0F6CBD] text-white font-semibold flex items-center gap-1"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              완료
            </button>
          </div>
        </div>
      )}

      {/* Rich HTML Compose Canvas (Preserves Signature Business Card + Original Rich HTML) */}
      <div
        ref={editorRef}
        contentEditable
        suppressContentEditableWarning
        data-testid="compose-html-editor"
        onInput={() => setPlainPreview(editorRef.current?.innerText || '')}
        style={{ fontSize: fontSizePt, fontFamily }}
        className="flex-1 overflow-y-auto px-6 py-4 bg-white text-[#242424] focus:outline-none leading-relaxed select-text [&_img]:max-w-full [&_img]:h-auto [&_table]:border-collapse"
      />

      {/* Footer Status / Compact Pre-Send Coaching Bar */}
      <div className="px-4 py-1.5 bg-[#FAF9F8] border-t border-[#E1DFDD] space-y-1 shrink-0">
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
                : '발송 전 점검 완료 · 명함 HTML 서명·원본 스레드·첨부 유지'}
            </span>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            {requestReadReceipt && (
              <span className="px-1.5 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] text-[10px] font-semibold">
                읽음확인 요청
              </span>
            )}
            <span className="font-mono text-[10px]">
              {pop3ConnectedEmail && keptForwardAtts.length === 0
                ? `SMTP: ${pop3ConnectedEmail}`
                : 'PC Outlook COM + 로컬 보관함 연동'}
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
