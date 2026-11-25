import React, { useMemo, useState } from 'react';
import {
  ArrowLeft,
  Database,
  FileText,
  Flag,
  Forward,
  Inbox,
  Loader2,
  Mail,
  Monitor,
  Paperclip,
  Plug,
  RefreshCw,
  Reply,
  Search,
  Send,
  SquarePen,
  Trash2,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';
import { formatSyncClockTime } from '../utils/headerSyncState';
import type { OutlookDraftItem } from './OutlookWorkbench';
import type { Pop3ConnectedAccount } from './Pop3QuickConnectModal';

type MobileTab = 'inbox' | 'sent' | 'drafts' | 'compose';

interface MobileMailShellProps {
  emails: MailItem[];
  drafts: OutlookDraftItem[];
  pop3Account: Pop3ConnectedAccount | null;
  isSyncingPop3: boolean;
  microsoftConnected?: boolean;
  microsoftEmail?: string | null;
  isSyncingGraph?: boolean;
  lastSyncedAt?: string | null;
  notice: string | null;
  onDismissNotice: () => void;
  onOpenPop3Modal: () => void;
  onOpenPstModal: () => void;
  onSyncPop3: () => void;
  onSwitchToDesktop: () => void;
  onToggleRead: (id: string) => void;
  onToggleFlag: (id: string) => void;
  onDeleteEmail: (id: string) => void;
  onSaveDraft: (draft: Omit<OutlookDraftItem, 'id' | 'updatedAt'>) => void;
  onDeleteDraft: (id: string) => void;
  onSendViaSmtp: (payload: {
    to: string;
    cc: string;
    subject: string;
    body: string;
  }) => Promise<{ ok: boolean; message: string }>;
}

/**
 * Dedicated ultra-clean Mobile Mail Mode:
 * Strictly icons + core mail reading, writing, and company POP3 connect only.
 */
export const MobileMailShell: React.FC<MobileMailShellProps> = ({
  emails,
  drafts,
  pop3Account,
  isSyncingPop3,
  microsoftConnected = false,
  microsoftEmail = null,
  isSyncingGraph = false,
  lastSyncedAt = null,
  notice,
  onDismissNotice,
  onOpenPop3Modal,
  onOpenPstModal,
  onSyncPop3,
  onSwitchToDesktop,
  onToggleRead,
  onToggleFlag,
  onDeleteEmail,
  onSaveDraft,
  onDeleteDraft,
  onSendViaSmtp,
}) => {
  const [activeTab, setActiveTab] = useState<MobileTab>('inbox');
  const [selectedMailId, setSelectedMailId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');

  // Compose state
  const [composeTo, setComposeTo] = useState('');
  const [composeCc, setComposeCc] = useState('');
  const [composeSubject, setComposeSubject] = useState('');
  const [composeBody, setComposeBody] = useState('');
  const [composeKind, setComposeKind] = useState<OutlookDraftItem['kind']>('new');
  const [isSending, setIsSending] = useState(false);
  const [composeFeedback, setComposeFeedback] = useState<string | null>(null);

  const inboxEmails = useMemo(
    () =>
      emails
        .filter((m) => m.folder === 'inbox')
        .sort((a, b) => b.timestamp.localeCompare(a.timestamp)),
    [emails]
  );

  const sentEmails = useMemo(
    () =>
      emails
        .filter((m) => m.folder === 'sent')
        .sort((a, b) => b.timestamp.localeCompare(a.timestamp)),
    [emails]
  );

  const unreadCount = useMemo(
    () => inboxEmails.filter((m) => !m.isRead).length,
    [inboxEmails]
  );

  const visibleList = useMemo(() => {
    const base = activeTab === 'sent' ? sentEmails : inboxEmails;
    const q = searchQuery.trim().toLowerCase();
    if (!q) return base;
    return base.filter(
      (m) =>
        m.subject.toLowerCase().includes(q) ||
        m.fromName.toLowerCase().includes(q) ||
        m.fromEmail.toLowerCase().includes(q) ||
        m.snippet.toLowerCase().includes(q)
    );
  }, [activeTab, inboxEmails, sentEmails, searchQuery]);

  const activeMail = useMemo(
    () => emails.find((m) => m.id === selectedMailId) || null,
    [emails, selectedMailId]
  );

  const openMailReader = (mail: MailItem) => {
    setSelectedMailId(mail.id);
    if (!mail.isRead) {
      onToggleRead(mail.id);
    }
  };

  const startCompose = (preset?: {
    to?: string;
    cc?: string;
    subject?: string;
    body?: string;
    kind?: OutlookDraftItem['kind'];
  }) => {
    setSelectedMailId(null);
    setComposeTo(preset?.to || '');
    setComposeCc(preset?.cc || '');
    setComposeSubject(preset?.subject || '');
    setComposeBody(preset?.body || '');
    setComposeKind(preset?.kind || 'new');
    setComposeFeedback(null);
    setActiveTab('compose');
  };

  const handleReply = (mail: MailItem) => {
    startCompose({
      to: mail.fromEmail || mail.fromName,
      subject: mail.subject.startsWith('Re:') ? mail.subject : `Re: ${mail.subject}`,
      body: `\n\n--- 원본 메일 (${mail.date} ${mail.fromName}) ---\n${mail.body.slice(0, 800)}`,
      kind: 'reply',
    });
  };

  const handleForward = (mail: MailItem) => {
    startCompose({
      to: '',
      subject: mail.subject.startsWith('Fwd:') ? mail.subject : `Fwd: ${mail.subject}`,
      body: `\n\n--- 전달된 메일 (${mail.date} ${mail.fromName} <${mail.fromEmail}>) ---\n${mail.body.slice(0, 1200)}`,
      kind: 'forward',
    });
  };

  const handleSaveDraftClick = () => {
    onSaveDraft({
      to: composeTo.trim(),
      cc: composeCc.trim(),
      subject: composeSubject.trim() || '(제목 없음)',
      body: composeBody,
      kind: composeKind,
    });
    setComposeFeedback('임시 보관함에 저장했습니다.');
  };

  const handleSendClick = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!composeTo.trim()) {
      setComposeFeedback('받는 사람 이메일 주소를 입력해 주세요.');
      return;
    }
    if (!pop3Account) {
      onSaveDraft({
        to: composeTo.trim(),
        cc: composeCc.trim(),
        subject: composeSubject.trim() || '(제목 없음)',
        body: composeBody,
        kind: composeKind,
      });
      setComposeFeedback('연결된 POP3/SMTP 계정이 없어 임시 보관함에 저장했습니다. 상단 연결 아이콘으로 계정을 연결하면 바로 발송할 수 있습니다.');
      return;
    }
    setIsSending(true);
    setComposeFeedback(null);
    const result = await onSendViaSmtp({
      to: composeTo.trim(),
      cc: composeCc.trim(),
      subject: composeSubject.trim() || '(제목 없음)',
      body: composeBody,
    });
    setIsSending(false);
    setComposeFeedback(result.message);
    if (result.ok) {
      setComposeTo('');
      setComposeCc('');
      setComposeSubject('');
      setComposeBody('');
      setActiveTab('sent');
    }
  };

  const headerTitle = activeMail
    ? '메일 읽기'
    : activeTab === 'inbox'
      ? '받은 편지함'
      : activeTab === 'sent'
        ? '보낸 편지함'
        : activeTab === 'drafts'
          ? '임시 보관함'
          : '메일 쓰기';

  return (
    <div className="h-screen flex flex-col bg-[#FAF9F8] text-[#242424] overflow-hidden max-w-full">
      {/* Compact Mobile Header (#0F6CBD) */}
      <header className="h-12 px-3 bg-[#0F6CBD] text-white flex items-center justify-between gap-2 shrink-0">
        <div className="flex items-center gap-2 min-w-0">
          {activeMail || activeTab === 'compose' ? (
            <button
              type="button"
              onClick={() => {
                setSelectedMailId(null);
                if (activeTab === 'compose') setActiveTab('inbox');
              }}
              aria-label="목록으로 돌아가기"
              className="w-8 h-8 flex items-center justify-center rounded-full hover:bg-white/15"
            >
              <ArrowLeft className="w-5 h-5" />
            </button>
          ) : (
            <div className="w-8 h-8 flex items-center justify-center rounded-full bg-white/15 shrink-0">
              <Mail className="w-4 h-4 text-white" />
            </div>
          )}
          <div className="min-w-0">
            <h1 className="text-sm font-bold truncate leading-tight">{headerTitle}</h1>
            {(pop3Account || microsoftConnected || isSyncingPop3 || isSyncingGraph) && (
              <div
                className="flex items-center gap-1 text-[10px] text-white/90 truncate"
                aria-live="polite"
              >
                {isSyncingPop3 || isSyncingGraph ? (
                  <>
                    <Loader2 className="w-2.5 h-2.5 animate-spin shrink-0 text-sky-100" />
                    <span className="truncate font-medium">
                      {isSyncingGraph && isSyncingPop3
                        ? 'M365 · 메일 동기화 중...'
                        : isSyncingGraph
                          ? 'Microsoft 365 동기화 중...'
                          : `${(pop3Account?.protocol || 'pop3').toUpperCase()} 동기화 중...`}
                    </span>
                  </>
                ) : (
                  <>
                    <span className="relative flex h-1.5 w-1.5 shrink-0">
                      <span className="absolute inline-flex h-full w-full rounded-full bg-emerald-300 opacity-65 animate-ping" />
                      <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-emerald-300" />
                    </span>
                    <span className="truncate font-mono">
                      {pop3Account?.email || microsoftEmail || '연결됨'}
                    </span>
                    {formatSyncClockTime(lastSyncedAt) && (
                      <span className="text-white/75 shrink-0">
                        · {formatSyncClockTime(lastSyncedAt)?.slice(0, 5)}
                      </span>
                    )}
                  </>
                )}
              </div>
            )}
          </div>
        </div>

        {/* Essential Icon-only Actions */}
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onClick={onOpenPop3Modal}
            aria-label={isSyncingPop3 ? '회사 메일 동기화 중' : '회사 POP3 메일 연결'}
            aria-busy={isSyncingPop3}
            title="메일플러그 · 회사 POP3/IMAP 연결"
            className="relative w-9 h-9 flex items-center justify-center rounded-full hover:bg-white/15"
          >
            {isSyncingPop3 ? (
              <Loader2 className="w-[18px] h-[18px] animate-spin" />
            ) : (
              <Plug className="w-[18px] h-[18px]" />
            )}
            {pop3Account && (
              <span
                aria-hidden="true"
                className="absolute top-1.5 right-1.5 flex h-2 w-2 items-center justify-center"
              >
                {isSyncingPop3 && (
                  <span className="absolute inline-flex h-full w-full rounded-full bg-emerald-300 opacity-80 animate-ping" />
                )}
                <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-300 ring-1 ring-[#0F6CBD]" />
              </span>
            )}
          </button>

          {pop3Account && (
            <button
              type="button"
              onClick={onSyncPop3}
              disabled={isSyncingPop3}
              aria-label="메일 새로고침"
              title="POP3/IMAP 메일 새로고침"
              className="w-9 h-9 flex items-center justify-center rounded-full hover:bg-white/15 disabled:opacity-50"
            >
              <RefreshCw className={`w-[18px] h-[18px] ${isSyncingPop3 ? 'animate-spin' : ''}`} />
            </button>
          )}

          <button
            type="button"
            onClick={onSwitchToDesktop}
            aria-label="데스크톱 모드로 전환"
            title="데스크톱 모드"
            className="w-9 h-9 flex items-center justify-center rounded-full hover:bg-white/15"
          >
            <Monitor className="w-[18px] h-[18px]" />
          </button>
        </div>
      </header>

      {/* Optional compact notice banner */}
      {notice && (
        <div className="px-3 py-2 bg-[#EBF3FC] border-b border-[#B4D6FA] text-xs text-[#0F6CBD] flex items-center justify-between gap-2 shrink-0">
          <span className="truncate">{notice}</span>
          <button
            type="button"
            onClick={onDismissNotice}
            aria-label="알림 닫기"
            className="p-0.5 shrink-0"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Main Content Area: Reader OR Compose OR Drafts OR Mail List */}
      <main className="flex-1 min-h-0 overflow-y-auto">
        {activeMail ? (
          /* 1. MOBILE MAIL READER */
          <article className="bg-white min-h-full flex flex-col">
            {/* Reader Icon Toolbar */}
            <div className="px-3 py-2 border-b border-[#EDEBE9] bg-[#FAF9F8] flex items-center justify-around gap-1 shrink-0">
              <button
                type="button"
                onClick={() => handleReply(activeMail)}
                className="flex flex-col items-center gap-0.5 px-3 py-1 rounded text-[#242424] hover:bg-[#EBF3FC] hover:text-[#0F6CBD]"
              >
                <Reply className="w-4 h-4" />
                <span className="text-[10px]">회신</span>
              </button>
              <button
                type="button"
                onClick={() => handleForward(activeMail)}
                className="flex flex-col items-center gap-0.5 px-3 py-1 rounded text-[#242424] hover:bg-[#EBF3FC] hover:text-[#0F6CBD]"
              >
                <Forward className="w-4 h-4" />
                <span className="text-[10px]">전달</span>
              </button>
              <button
                type="button"
                onClick={() => onToggleFlag(activeMail.id)}
                className={`flex flex-col items-center gap-0.5 px-3 py-1 rounded ${
                  activeMail.isFlagged ? 'text-red-600' : 'text-[#242424] hover:bg-[#F3F2F1]'
                }`}
              >
                <Flag className="w-4 h-4" />
                <span className="text-[10px]">중요</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  onDeleteEmail(activeMail.id);
                  setSelectedMailId(null);
                }}
                className="flex flex-col items-center gap-0.5 px-3 py-1 rounded text-[#242424] hover:bg-red-50 hover:text-red-600"
              >
                <Trash2 className="w-4 h-4" />
                <span className="text-[10px]">삭제</span>
              </button>
            </div>

            {/* Mail Header & Body */}
            <div className="p-4 space-y-3 flex-1">
              <h2 className="text-base font-bold text-[#242424] leading-snug">
                {activeMail.subject}
              </h2>
              <div className="flex items-start justify-between gap-2 pb-3 border-b border-[#EDEBE9]">
                <div className="min-w-0">
                  <div className="text-xs font-semibold text-[#242424] truncate">
                    {activeMail.fromName}
                  </div>
                  <div className="text-[11px] text-[#616161] truncate font-mono">
                    {activeMail.fromEmail}
                  </div>
                  {activeMail.to.length > 0 && (
                    <div className="text-[11px] text-[#616161] truncate mt-0.5">
                      받는 사람: {activeMail.to.join(', ')}
                    </div>
                  )}
                </div>
                <div className="text-[11px] text-[#616161] tabular-nums shrink-0 text-right">
                  <div>{activeMail.date}</div>
                  <div>{activeMail.timeLabel}</div>
                </div>
              </div>

              {activeMail.attachments && activeMail.attachments.length > 0 && (
                <div className="flex flex-wrap gap-1.5 pb-2">
                  {activeMail.attachments.map((att) => (
                    <span
                      key={att.id}
                      className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] bg-[#FAF9F8] border border-[#E0E0E0] rounded text-[#424242]"
                    >
                      <Paperclip className="w-3 h-3 text-[#0F6CBD]" />
                      <span className="truncate max-w-[200px]">{att.filename}</span>
                    </span>
                  ))}
                </div>
              )}

              <div className="text-sm text-[#242424] leading-relaxed whitespace-pre-wrap break-words">
                {activeMail.body}
              </div>
            </div>
          </article>
        ) : activeTab === 'compose' ? (
          /* 2. MOBILE MAIL COMPOSE */
          <form onSubmit={handleSendClick} className="bg-white min-h-full p-4 space-y-3 flex flex-col">
            {composeFeedback && (
              <div className="p-2.5 rounded-lg bg-[#EBF3FC] border border-[#B4D6FA] text-xs text-[#0F6CBD]">
                {composeFeedback}
              </div>
            )}

            <div className="space-y-1">
              <label htmlFor="mob-to" className="text-[11px] font-semibold text-[#616161]">
                받는 사람
              </label>
              <input
                id="mob-to"
                type="text"
                value={composeTo}
                onChange={(e) => setComposeTo(e.target.value)}
                placeholder="email@company.co.kr"
                className="w-full px-3 py-2 text-sm bg-[#FAF9F8] border border-[#D1D1D1] rounded-lg focus:outline-none focus:border-[#0F6CBD]"
              />
            </div>

            <div className="space-y-1">
              <label htmlFor="mob-subject" className="text-[11px] font-semibold text-[#616161]">
                제목
              </label>
              <input
                id="mob-subject"
                type="text"
                value={composeSubject}
                onChange={(e) => setComposeSubject(e.target.value)}
                placeholder="메일 제목 입력"
                className="w-full px-3 py-2 text-sm bg-[#FAF9F8] border border-[#D1D1D1] rounded-lg focus:outline-none focus:border-[#0F6CBD]"
              />
            </div>

            <div className="space-y-1 flex-1 flex flex-col">
              <label htmlFor="mob-body" className="text-[11px] font-semibold text-[#616161]">
                내용
              </label>
              <textarea
                id="mob-body"
                rows={10}
                value={composeBody}
                onChange={(e) => setComposeBody(e.target.value)}
                placeholder="메일 본문을 작성하세요"
                className="w-full flex-1 px-3 py-2.5 text-sm bg-[#FAF9F8] border border-[#D1D1D1] rounded-lg focus:outline-none focus:border-[#0F6CBD] resize-none"
              />
            </div>

            <div className="flex items-center justify-between gap-2 pt-2 border-t border-[#EDEBE9]">
              <button
                type="button"
                onClick={handleSaveDraftClick}
                className="px-3.5 py-2.5 text-xs font-medium text-[#424242] bg-[#FAF9F8] border border-[#D1D1D1] rounded-lg inline-flex items-center gap-1.5"
              >
                <FileText className="w-4 h-4" />
                임시 저장
              </button>
              <button
                type="submit"
                disabled={isSending}
                className="flex-1 py-2.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-lg inline-flex items-center justify-center gap-1.5"
              >
                <Send className="w-4 h-4" />
                {isSending ? '발송 중...' : pop3Account ? '메일 보내기' : '임시함에 저장'}
              </button>
            </div>
          </form>
        ) : activeTab === 'drafts' ? (
          /* 3. MOBILE DRAFTS LIST */
          <div className="divide-y divide-[#EDEBE9] bg-white min-h-full">
            {drafts.length === 0 ? (
              <div className="p-8 text-center space-y-3">
                <FileText className="w-8 h-8 text-[#9E9E9E] mx-auto" />
                <p className="text-xs text-[#616161]">저장된 임시 메일이 없습니다.</p>
                <button
                  type="button"
                  onClick={() => startCompose()}
                  className="px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] rounded-lg inline-flex items-center gap-1.5"
                >
                  <SquarePen className="w-4 h-4" />
                  새 메일 쓰기
                </button>
              </div>
            ) : (
              drafts.map((d) => (
                <div key={d.id} className="px-4 py-3 flex items-center justify-between gap-3">
                  <button
                    type="button"
                    onClick={() =>
                      startCompose({
                        to: d.to,
                        cc: d.cc,
                        subject: d.subject,
                        body: d.body,
                        kind: d.kind,
                      })
                    }
                    className="flex-1 min-w-0 text-left"
                  >
                    <div className="text-sm font-semibold text-[#242424] truncate">
                      {d.subject || '(제목 없음)'}
                    </div>
                    <div className="text-xs text-[#616161] truncate">
                      받는 사람: {d.to || '(미지정)'}
                    </div>
                  </button>
                  <button
                    type="button"
                    onClick={() => onDeleteDraft(d.id)}
                    aria-label="임시 메일 삭제"
                    className="p-2 text-[#616161] hover:text-red-600"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              ))
            )}
          </div>
        ) : (
          /* 4. MOBILE INBOX / SENT LIST */
          <div className="flex flex-col min-h-full bg-white">
            {/* Compact 1-line Search */}
            <div className="p-2.5 bg-[#FAF9F8] border-b border-[#EDEBE9]">
              <div className="relative">
                <Search className="w-4 h-4 text-[#616161] absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="search"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="제목 · 보낸 사람 검색"
                  aria-label="모바일 메일 검색"
                  className="w-full pl-9 pr-3 py-2 text-xs bg-white border border-[#D1D1D1] rounded-lg focus:outline-none focus:border-[#0F6CBD]"
                />
              </div>
            </div>

            {visibleList.length === 0 ? (
              <div className="p-8 text-center space-y-4 my-auto">
                <div className="w-12 h-12 rounded-2xl bg-[#EBF3FC] text-[#0F6CBD] flex items-center justify-center mx-auto">
                  <Inbox className="w-6 h-6" />
                </div>
                <div className="space-y-1">
                  <p className="text-sm font-semibold text-[#242424]">
                    {activeTab === 'sent' ? '보낸 메일이 없습니다' : '연결된 메일이 없습니다'}
                  </p>
                  <p className="text-xs text-[#616161] leading-relaxed">
                    메일플러그 등 회사 이메일과 비밀번호만으로 POP3 설정을 자동 감지해 바로 읽어옵니다.
                  </p>
                </div>
                <div className="flex flex-col gap-2 max-w-xs mx-auto">
                  <button
                    type="button"
                    onClick={onOpenPop3Modal}
                    className="w-full py-2.5 px-4 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg inline-flex items-center justify-center gap-2"
                  >
                    <Plug className="w-4 h-4" />
                    메일플러그 · 회사 POP3 자동 연결
                  </button>
                  <button
                    type="button"
                    onClick={onOpenPstModal}
                    className="w-full py-2.5 px-4 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-lg inline-flex items-center justify-center gap-2"
                  >
                    <Database className="w-4 h-4" />
                    PST · EML 파일 열기
                  </button>
                </div>
              </div>
            ) : (
              <div className="divide-y divide-[#EDEBE9]">
                {visibleList.map((mail) => (
                  <button
                    key={mail.id}
                    type="button"
                    onClick={() => openMailReader(mail)}
                    className="w-full px-4 py-3 text-left flex items-start gap-2.5 active:bg-[#EBF3FC] transition-colors"
                  >
                    <span
                      aria-hidden="true"
                      className={`mt-1.5 w-2 h-2 rounded-full shrink-0 ${
                        mail.isRead ? 'bg-transparent' : 'bg-[#0F6CBD]'
                      }`}
                    />
                    <div className="flex-1 min-w-0 space-y-0.5">
                      <div className="flex items-center justify-between gap-2">
                        <span
                          className={`text-xs truncate ${
                            mail.isRead ? 'font-medium text-[#424242]' : 'font-bold text-[#242424]'
                          }`}
                        >
                          {activeTab === 'sent' ? mail.to[0] || mail.fromName : mail.fromName}
                        </span>
                        <span className="text-[11px] text-[#616161] tabular-nums shrink-0 flex items-center gap-1">
                          {mail.hasAttachment && <Paperclip className="w-3 h-3 text-[#616161]" />}
                          {mail.date.slice(5)}
                        </span>
                      </div>
                      <div
                        className={`text-sm truncate ${
                          mail.isRead ? 'font-normal text-[#242424]' : 'font-bold text-[#242424]'
                        }`}
                      >
                        {mail.subject}
                      </div>
                      <div className="text-xs text-[#616161] truncate">{mail.snippet}</div>
                    </div>
                  </button>
                ))}
              </div>
            )}
          </div>
        )}
      </main>

      {/* Bottom Essential Icon Navigation Bar */}
      <nav
        aria-label="모바일 메일 기본 메뉴"
        className="h-14 bg-white border-t border-[#E0E0E0] grid grid-cols-5 shrink-0 select-none"
      >
        <button
          type="button"
          onClick={() => {
            setSelectedMailId(null);
            setActiveTab('inbox');
          }}
          className={`flex flex-col items-center justify-center gap-0.5 relative ${
            activeTab === 'inbox' && !activeMail ? 'text-[#0F6CBD] font-semibold' : 'text-[#616161]'
          }`}
        >
          <Inbox className="w-5 h-5" />
          <span className="text-[10px]">받은메일</span>
          {unreadCount > 0 && (
            <span className="absolute top-1.5 right-3 min-w-[16px] h-4 px-1 rounded-full bg-[#0F6CBD] text-white text-[9px] font-bold flex items-center justify-center">
              {unreadCount > 99 ? '99+' : unreadCount}
            </span>
          )}
        </button>

        <button
          type="button"
          onClick={() => {
            setSelectedMailId(null);
            setActiveTab('sent');
          }}
          className={`flex flex-col items-center justify-center gap-0.5 ${
            activeTab === 'sent' && !activeMail ? 'text-[#0F6CBD] font-semibold' : 'text-[#616161]'
          }`}
        >
          <Send className="w-5 h-5" />
          <span className="text-[10px]">보낸메일</span>
        </button>

        <button
          type="button"
          onClick={() => startCompose()}
          className={`flex flex-col items-center justify-center gap-0.5 ${
            activeTab === 'compose' ? 'text-[#0F6CBD] font-semibold' : 'text-[#616161]'
          }`}
        >
          <SquarePen className="w-5 h-5" />
          <span className="text-[10px]">메일쓰기</span>
        </button>

        <button
          type="button"
          onClick={() => {
            setSelectedMailId(null);
            setActiveTab('drafts');
          }}
          className={`flex flex-col items-center justify-center gap-0.5 relative ${
            activeTab === 'drafts' && !activeMail ? 'text-[#0F6CBD] font-semibold' : 'text-[#616161]'
          }`}
        >
          <FileText className="w-5 h-5" />
          <span className="text-[10px]">임시함</span>
          {drafts.length > 0 && (
            <span className="absolute top-1.5 right-3 min-w-[16px] h-4 px-1 rounded-full bg-[#616161] text-white text-[9px] font-bold flex items-center justify-center">
              {drafts.length}
            </span>
          )}
        </button>

        <button
          type="button"
          onClick={onOpenPop3Modal}
          className={`flex flex-col items-center justify-center gap-0.5 ${
            pop3Account ? 'text-emerald-700 font-semibold' : 'text-[#616161]'
          }`}
        >
          <Plug className="w-5 h-5" />
          <span className="text-[10px]">POP3연결</span>
        </button>
      </nav>
    </div>
  );
};
