import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Check,
  Download,
  ExternalLink,
  Lock,
  LogOut,
  Mail,
  MessageSquare,
  Paperclip,
  RefreshCw,
  Send,
  Users,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';

interface SgwChatUser {
  sub: string;
  name: string;
  email: string;
  role: string;
  department: string;
  position: string;
}

interface SgwEmployeeItem {
  id: string;
  name: string;
  email: string;
  role: string;
  department: string;
  position: string;
  status: string;
}

interface SgwChatRoom {
  id: string;
  participants: string[];
  lastMessage: string;
  lastMessageTime: number;
  lastMessageSenderId?: string;
  unreadCount?: number;
  participantProfiles?: SgwEmployeeItem[];
}

interface SgwChatMessage {
  id: string;
  roomId: string;
  senderId: string;
  content: string;
  timestamp: number;
  readBy: string[];
  hasFile: boolean;
  fileName?: string;
  fileData?: string;
  isDeleted?: boolean;
  reactions?: Record<string, string[]>;
}

interface SgwDirectMessage {
  id: string;
  senderId: string;
  senderName: string;
  receiverId: string;
  receiverName: string;
  content: string;
  timestamp: number;
  isRead: boolean;
  hasFile: boolean;
  fileName?: string;
  fileData?: string;
  isDeleted?: boolean;
}

export interface SgwChatPanelProps {
  mode?: 'drawer' | 'full';
  defaultEmail?: string;
  pendingShareMail?: MailItem | null;
  onClearPendingShare?: () => void;
  initialPeerEmail?: { email: string; name?: string } | null;
  onClearInitialPeer?: () => void;
  onComposeFromChat?: (draft: { to: string; subject: string; body: string }) => void;
  onOpenMailById?: (emailId: string) => void;
  onUnreadCountChange?: (total: number) => void;
  onClose?: () => void;
}

const BMAIL_DEEP_LINK_RE = /bmail:\/\/open\?id=([^\s]+)/i;
const QUICK_REACTIONS = ['👍', '✅', '👀'];

function formatChatTime(ts: number): string {
  if (!ts) return '';
  try {
    const d = new Date(ts);
    return d.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', hour12: false });
  } catch {
    return '';
  }
}

export const SgwChatPanel: React.FC<SgwChatPanelProps> = ({
  mode = 'drawer',
  defaultEmail = '',
  pendingShareMail = null,
  onClearPendingShare,
  initialPeerEmail = null,
  onClearInitialPeer,
  onComposeFromChat,
  onOpenMailById,
  onUnreadCountChange,
  onClose,
}) => {
  const [connected, setConnected] = useState(false);
  const [sgwBaseUrl, setSgwBaseUrl] = useState('https://sgw.seastar.work');
  const [user, setUser] = useState<SgwChatUser | null>(null);
  const [loginEmail, setLoginEmail] = useState(defaultEmail);
  const [loginPassword, setLoginPassword] = useState('');
  const [showAdvancedUrl, setShowAdvancedUrl] = useState(false);
  const [showDirectCredentialForm, setShowDirectCredentialForm] = useState(false);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [isOpeningLoginWin, setIsOpeningLoginWin] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const [activeTab, setActiveTab] = useState<'rooms' | 'dm' | 'people'>('rooms');
  const [rooms, setRooms] = useState<SgwChatRoom[]>([]);
  const [employees, setEmployees] = useState<SgwEmployeeItem[]>([]);
  const [selectedRoomId, setSelectedRoomId] = useState<string>('');
  const [roomMessages, setRoomMessages] = useState<SgwChatMessage[]>([]);
  const [dms, setDms] = useState<SgwDirectMessage[]>([]);
  const [selectedDmPeerId, setSelectedDmPeerId] = useState<string>('');
  const [peopleSearch, setPeopleSearch] = useState('');

  const [draftText, setDraftText] = useState('');
  const [attachFile, setAttachFile] = useState<{ name: string; b64: string } | null>(null);
  const [isSending, setIsSending] = useState(false);
  const [shareNote, setShareNote] = useState('');
  const [shareAttachmentIdx, setShareAttachmentIdx] = useState<number | null>(null);
  const [isSharingMail, setIsSharingMail] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!loginEmail && defaultEmail) {
      setLoginEmail(defaultEmail);
    }
  }, [defaultEmail, loginEmail]);

  const loadStatus = useCallback(async () => {
    try {
      const res = await fetch('/api/sgw-chat/status');
      if (!res.ok) return;
      const data = (await res.json()) as {
        connected: boolean;
        sgwBaseUrl: string;
        user: SgwChatUser | null;
      };
      setConnected(Boolean(data.connected));
      if (data.sgwBaseUrl) setSgwBaseUrl(data.sgwBaseUrl);
      setUser(data.user || null);
      if (!data.connected) {
        onUnreadCountChange?.(0);
      }
    } catch {
      // ignore startup error
    }
  }, [onUnreadCountChange]);

  const refreshDirectoryAndRooms = useCallback(async () => {
    if (!connected) return;
    try {
      const [roomsRes, empRes, dmRes] = await Promise.all([
        fetch('/api/sgw-chat/rooms'),
        fetch('/api/sgw-chat/employees'),
        fetch('/api/sgw-chat/dm'),
      ]);
      if (roomsRes.status === 401 || empRes.status === 401 || dmRes.status === 401) {
        setConnected(false);
        setUser(null);
        setErrorMsg('SGW 로그인 세션이 만료되었습니다. 다시 로그인해 주세요.');
        onUnreadCountChange?.(0);
        return;
      }
      let unreadSum = 0;
      if (roomsRes.ok) {
        const d = (await roomsRes.json()) as { rooms?: SgwChatRoom[]; unreadTotal?: number };
        const nextRooms = Array.isArray(d.rooms) ? d.rooms : [];
        setRooms(nextRooms);
        unreadSum += Number(d.unreadTotal || 0);
        if (!selectedRoomId && nextRooms.length > 0) {
          setSelectedRoomId(nextRooms[0].id);
        }
      }
      if (empRes.ok) {
        const d = (await empRes.json()) as { employees?: SgwEmployeeItem[] };
        setEmployees(Array.isArray(d.employees) ? d.employees : []);
      }
      if (dmRes.ok) {
        const d = (await dmRes.json()) as { messages?: SgwDirectMessage[] };
        const nextDms = Array.isArray(d.messages) ? d.messages : [];
        setDms(nextDms);
        if (user) {
          unreadSum += nextDms.filter((m) => m.receiverId === user.sub && !m.isRead && !m.isDeleted).length;
        }
      }
      onUnreadCountChange?.(unreadSum);
    } catch {
      // ignore transient network glitches during background polling
    }
  }, [connected, onUnreadCountChange, selectedRoomId, user]);

  const loadSelectedRoomMessages = useCallback(async () => {
    if (!connected || !selectedRoomId) return;
    try {
      const res = await fetch(
        `/api/sgw-chat/rooms/${encodeURIComponent(selectedRoomId)}/messages`
      );
      if (res.status === 401) {
        setConnected(false);
        setUser(null);
        return;
      }
      if (!res.ok) return;
      const d = (await res.json()) as { messages?: SgwChatMessage[] };
      const msgs = Array.isArray(d.messages) ? d.messages : [];
      setRoomMessages(msgs);
      const lastMsg = msgs[msgs.length - 1];
      if (lastMsg && user && !lastMsg.readBy.includes(user.sub)) {
        void fetch(`/api/sgw-chat/rooms/${encodeURIComponent(selectedRoomId)}/read`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ lastReadMessageId: lastMsg.id }),
        });
      }
    } catch {
      // ignore
    }
  }, [connected, selectedRoomId, user]);

  useEffect(() => {
    void loadStatus();
  }, [loadStatus]);

  useEffect(() => {
    if (!connected) return;
    void refreshDirectoryAndRooms();
    const timer = window.setInterval(() => {
      if (document.visibilityState === 'visible') {
        void refreshDirectoryAndRooms();
      }
    }, 8000);
    return () => window.clearInterval(timer);
  }, [connected, refreshDirectoryAndRooms]);

  useEffect(() => {
    if (!connected || activeTab !== 'rooms' || !selectedRoomId) return;
    void loadSelectedRoomMessages();
    const timer = window.setInterval(() => {
      if (document.visibilityState === 'visible') {
        void loadSelectedRoomMessages();
      }
    }, 6000);
    return () => window.clearInterval(timer);
  }, [connected, activeTab, selectedRoomId, loadSelectedRoomMessages]);

  useEffect(() => {
    if (!connected || !initialPeerEmail?.email || employees.length === 0) return;
    const match = employees.find(
      (e) => e.email.toLowerCase() === initialPeerEmail.email.toLowerCase()
    );
    if (match) {
      void handleStartChatWithEmployee(match);
    } else {
      setActiveTab('people');
      setPeopleSearch(initialPeerEmail.name || initialPeerEmail.email.split('@')[0] || '');
    }
    onClearInitialPeer?.();
  }, [connected, initialPeerEmail, employees]);

  const normalizeSgwError = (err: unknown, fallbackMsg: string): string => {
    const raw = err instanceof Error ? err.message : String(err || '');
    if (/failed to fetch|networkerror|load failed/i.test(raw)) {
      return '로컬 Bmail 서버 연결을 확인 중입니다. 위의 [SGW 로그인 창 열기] 버튼을 눌러 공식 로그인 창으로 연동해 주세요.';
    }
    return raw || fallbackMsg;
  };

  const handleOpenSgwLoginWindow = async () => {
    setErrorMsg(null);
    setIsOpeningLoginWin(true);
    try {
      const desktopBridge = (
        window as unknown as {
          bmailDesktop?: {
            openSgwLoginWindow?: (url: string) => Promise<{
              connected?: boolean;
              canceled?: boolean;
              sgwBaseUrl?: string;
              user?: SgwChatUser;
              error?: string;
            }>;
          };
        }
      ).bmailDesktop;

      if (desktopBridge?.openSgwLoginWindow) {
        const res = await desktopBridge.openSgwLoginWindow(sgwBaseUrl.trim());
        if (res?.connected && res.user) {
          setConnected(true);
          setUser(res.user);
          if (res.sgwBaseUrl) setSgwBaseUrl(res.sgwBaseUrl);
          await loadStatus();
          return;
        }
        if (res?.canceled) {
          return;
        }
        if (res?.error) {
          setErrorMsg(res.error);
          return;
        }
      } else {
        // Browser popup fallback + status check on focus return
        const popup = window.open(
          sgwBaseUrl.trim(),
          'bmail_sgw_login_window',
          'width=520,height=740,resizable=yes,scrollbars=yes'
        );
        if (!popup) {
          setErrorMsg('브라우저 팝업이 차단되었습니다. 팝업 허용 후 다시 클릭하거나 아래 계정 직접 연결을 이용해 주세요.');
          setShowDirectCredentialForm(true);
          return;
        }
        const checkTimer = window.setInterval(() => {
          if (popup.closed) {
            window.clearInterval(checkTimer);
            void loadStatus();
          }
        }, 1000);
      }
    } catch (err) {
      setErrorMsg(normalizeSgwError(err, 'SGW 로그인 창을 열지 못했습니다.'));
    } finally {
      setIsOpeningLoginWin(false);
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg(null);
    setIsLoggingIn(true);
    try {
      const res = await fetch('/api/sgw-chat/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email: loginEmail.trim(),
          password: loginPassword,
          sgwBaseUrl: sgwBaseUrl.trim(),
        }),
      });
      const data = (await res.json()) as {
        connected?: boolean;
        sgwBaseUrl?: string;
        user?: SgwChatUser;
        error?: string;
      };
      if (!res.ok || !data.connected) {
        setErrorMsg(data.error || 'SGW 로그인에 실패했습니다.');
        return;
      }
      setLoginPassword('');
      setConnected(true);
      setUser(data.user || null);
      if (data.sgwBaseUrl) setSgwBaseUrl(data.sgwBaseUrl);
    } catch (err) {
      setErrorMsg(normalizeSgwError(err, 'SGW 로그인 중 오류가 발생했습니다.'));
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    await fetch('/api/sgw-chat/logout', { method: 'POST' }).catch(() => {});
    const clearFn = (
      window as unknown as {
        bmailDesktop?: { clearSgwAuthSession?: (url: string) => Promise<unknown> };
      }
    ).bmailDesktop?.clearSgwAuthSession;
    if (clearFn) {
      await clearFn(sgwBaseUrl).catch(() => {});
    }
    setConnected(false);
    setUser(null);
    setRooms([]);
    setRoomMessages([]);
    setDms([]);
    onUnreadCountChange?.(0);
  };

  const empById = useMemo(
    () => new Map(employees.map((e) => [String(e.id), e])),
    [employees]
  );

  const getRoomDisplayName = (room: SgwChatRoom): string => {
    const others = (room.participants || [])
      .filter((pid) => !user || pid !== user.sub)
      .map((pid) => empById.get(pid)?.name || room.participantProfiles?.find((p) => p.id === pid)?.name || pid);
    if (others.length === 0) return '내 메모방';
    return others.join(', ');
  };

  const handleStartChatWithEmployee = async (emp: SgwEmployeeItem) => {
    if (!user) return;
    setErrorMsg(null);
    try {
      const pair = [user.sub, emp.id].sort();
      const deterministicId = `dm_${pair[0]}_${pair[1]}`;
      const existing = rooms.find(
        (r) =>
          r.id === deterministicId ||
          (r.participants.length === 2 &&
            r.participants.includes(user.sub) &&
            r.participants.includes(emp.id))
      );
      if (existing) {
        setSelectedRoomId(existing.id);
        setActiveTab('rooms');
        return;
      }
      const res = await fetch('/api/sgw-chat/rooms', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: deterministicId, participants: pair }),
      });
      const data = (await res.json()) as { ok?: boolean; id?: string; error?: string };
      if (!res.ok || !data.id) {
        setErrorMsg(data.error || '대화방 생성에 실패했습니다.');
        return;
      }
      setSelectedRoomId(data.id);
      setActiveTab('rooms');
      await refreshDirectoryAndRooms();
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : '대화방 열기 실패');
    }
  };

  const handleAttachFileChange = (files: FileList | null) => {
    const file = files?.[0];
    if (!file) return;
    if (file.size > 5 * 1024 * 1024) {
      setErrorMsg('SGW 채팅 첨부파일은 5MB 이하만 업로드할 수 있습니다.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const result = typeof reader.result === 'string' ? reader.result : '';
      const commaIdx = result.indexOf(',');
      const b64 = commaIdx !== -1 ? result.slice(commaIdx + 1) : result;
      setAttachFile({ name: file.name, b64 });
      setErrorMsg(null);
    };
    reader.readAsDataURL(file);
  };

  const handleSendMessage = async () => {
    if ((!draftText.trim() && !attachFile) || isSending) return;
    setErrorMsg(null);
    setIsSending(true);
    try {
      if (activeTab === 'rooms') {
        if (!selectedRoomId) {
          setErrorMsg('먼저 대화방을 선택해 주세요.');
          return;
        }
        const res = await fetch(
          `/api/sgw-chat/rooms/${encodeURIComponent(selectedRoomId)}/messages`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              content: draftText.trim(),
              ...(attachFile ? { fileName: attachFile.name, fileData: attachFile.b64 } : {}),
            }),
          }
        );
        const data = (await res.json()) as { ok?: boolean; error?: string };
        if (!res.ok) {
          setErrorMsg(data.error || '메시지 전송 실패');
          return;
        }
        setDraftText('');
        setAttachFile(null);
        await Promise.all([loadSelectedRoomMessages(), refreshDirectoryAndRooms()]);
      } else if (activeTab === 'dm') {
        if (!selectedDmPeerId) {
          setErrorMsg('쪽지를 받을 임직원을 선택해 주세요.');
          return;
        }
        const peer = empById.get(selectedDmPeerId);
        const res = await fetch('/api/sgw-chat/dm', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            receiverId: selectedDmPeerId,
            receiverName: peer?.name || '',
            content: draftText.trim(),
            ...(attachFile ? { fileName: attachFile.name, fileData: attachFile.b64 } : {}),
          }),
        });
        const data = (await res.json()) as { ok?: boolean; error?: string };
        if (!res.ok) {
          setErrorMsg(data.error || '쪽지 전송 실패');
          return;
        }
        setDraftText('');
        setAttachFile(null);
        await refreshDirectoryAndRooms();
      }
    } finally {
      setIsSending(false);
    }
  };

  const handleShareMailNow = async () => {
    if (!pendingShareMail || isSharingMail) return;
    setErrorMsg(null);
    const targetType = activeTab === 'dm' ? 'dm' : 'room';
    if (targetType === 'room' && !selectedRoomId) {
      setErrorMsg('메일을 공유할 대화방을 왼쪽 또는 목록에서 먼저 선택해 주세요.');
      return;
    }
    if (targetType === 'dm' && !selectedDmPeerId) {
      setErrorMsg('메일을 공유할 직원을 먼저 선택해 주세요.');
      return;
    }
    setIsSharingMail(true);
    try {
      const peer = selectedDmPeerId ? empById.get(selectedDmPeerId) : undefined;
      const res = await fetch('/api/sgw-chat/share-mail', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          emailId: pendingShareMail.id,
          mail: pendingShareMail,
          targetType,
          roomId: selectedRoomId,
          receiverId: selectedDmPeerId,
          receiverName: peer?.name || '',
          note: shareNote.trim(),
          ...(shareAttachmentIdx !== null ? { attachmentIdx: shareAttachmentIdx } : {}),
        }),
      });
      const data = (await res.json()) as { ok?: boolean; error?: string };
      if (!res.ok) {
        setErrorMsg(data.error || '메일 공유 전송에 실패했습니다.');
        return;
      }
      setShareNote('');
      setShareAttachmentIdx(null);
      onClearPendingShare?.();
      if (targetType === 'room') {
        await Promise.all([loadSelectedRoomMessages(), refreshDirectoryAndRooms()]);
      } else {
        await refreshDirectoryAndRooms();
      }
    } finally {
      setIsSharingMail(false);
    }
  };

  const handleDownloadChatAttachment = async (
    target: { kind: 'room'; roomId: string; msgId: string; fileName?: string } | { kind: 'dm'; msgId: string; fileName?: string }
  ) => {
    try {
      const url =
        target.kind === 'room'
          ? `/api/sgw-chat/rooms/${encodeURIComponent(target.roomId)}/messages?msgId=${encodeURIComponent(target.msgId)}`
          : `/api/sgw-chat/dm?id=${encodeURIComponent(target.msgId)}`;
      const res = await fetch(url);
      if (!res.ok) return;
      const data = (await res.json()) as { messages?: Array<{ fileData?: string; fileName?: string }> };
      const rawData = data.messages?.[0]?.fileData;
      if (!rawData) return;
      const dataUrl = rawData.startsWith('data:')
        ? rawData
        : `data:application/octet-stream;base64,${rawData}`;
      const a = document.createElement('a');
      a.href = dataUrl;
      a.download = data.messages?.[0]?.fileName || target.fileName || 'sgw-attachment';
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch {
      // ignore
    }
  };

  const handleReaction = async (msgId: string, emoji: string, currentlyReacted: boolean) => {
    if (!selectedRoomId) return;
    await fetch(`/api/sgw-chat/rooms/${encodeURIComponent(selectedRoomId)}/react`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id: msgId, emoji, on: !currentlyReacted }),
    });
    await loadSelectedRoomMessages();
  };

  const renderMessageContent = (content: string) => {
    const match = content.match(BMAIL_DEEP_LINK_RE);
    const cleaned = content.replace(BMAIL_DEEP_LINK_RE, '').trim();
    const decodedEmailId = match ? decodeURIComponent(match[1]) : '';
    return (
      <div className="space-y-1.5">
        {cleaned && <p className="whitespace-pre-wrap break-words text-xs leading-relaxed">{cleaned}</p>}
        {decodedEmailId && onOpenMailById && (
          <button
            type="button"
            onClick={() => onOpenMailById(decodedEmailId)}
            className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-semibold bg-white/95 text-[#0F6CBD] border border-[#CFE4FA] hover:bg-[#EBF3FC] rounded"
          >
            <Mail className="w-3 h-3" />
            Bmail 원문 메일 바로 열기
          </button>
        )}
      </div>
    );
  };

  const dmConversation = useMemo(() => {
    if (!user || !selectedDmPeerId) return [];
    return dms.filter(
      (m) =>
        (m.senderId === user.sub && m.receiverId === selectedDmPeerId) ||
        (m.senderId === selectedDmPeerId && m.receiverId === user.sub)
    );
  }, [dms, user, selectedDmPeerId]);

  const filteredEmployees = useMemo(() => {
    const q = peopleSearch.trim().toLowerCase();
    return employees.filter((e) => {
      if (user && e.id === user.sub) return false;
      if (!q) return true;
      return `${e.name} ${e.department} ${e.position} ${e.email}`.toLowerCase().includes(q);
    });
  }, [employees, peopleSearch, user]);

  useEffect(() => {
    if (activeTab === 'dm' && selectedDmPeerId && user) {
      const hasUnread = dms.some(
        (m) => m.senderId === selectedDmPeerId && m.receiverId === user.sub && !m.isRead
      );
      if (hasUnread) {
        void fetch('/api/sgw-chat/dm/read', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ otherId: selectedDmPeerId }),
        }).then(() => refreshDirectoryAndRooms());
      }
    }
  }, [activeTab, selectedDmPeerId, dms, user, refreshDirectoryAndRooms]);

  const containerCls =
    mode === 'drawer'
      ? 'w-[380px] max-w-[92vw] shrink-0 h-full bg-white border-l border-[#E0E0E0] flex flex-col'
      : 'flex-1 h-full bg-[#F5F5F5] flex flex-col min-h-0';

  if (!connected) {
    return (
      <aside className={containerCls} aria-label="SGW 실시간 사내 채팅">
        <div className="px-4 py-3 bg-white border-b border-[#E0E0E0] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <MessageSquare className="w-4 h-4 text-[#0F6CBD]" />
            <span className="text-xs font-bold text-[#242424]">SGW 실시간 사내 채팅</span>
          </div>
          {onClose && (
            <button
              type="button"
              onClick={onClose}
              aria-label="채팅 창 닫기"
              className="p-1 text-[#616161] hover:text-[#242424] rounded hover:bg-[#F5F5F5]"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>
        <div className="flex-1 overflow-y-auto p-5 flex items-center justify-center">
          <div className="w-full max-w-sm bg-white border border-[#E0E0E0] rounded-xl p-5 space-y-4 shadow-sm">
            <div className="space-y-1">
              <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] text-[11px] font-semibold">
                <Lock className="w-3 h-3" />
                사내 그룹웨어 공식 로그인 창 직결
              </div>
              <h2 className="text-sm font-bold text-[#242424]">SGW 실시간 메신저 로그인 연동</h2>
              <p className="text-xs text-[#616161] leading-relaxed break-keep">
                SGW 그룹웨어({sgwBaseUrl}) 공식 로그인 창을 띄워 로그인하면 인증 세션을 자동으로 연결합니다. 대화방·쪽지를 Bmail 안에서 실시간으로 주고받고 메일을 원클릭 공유합니다.
              </p>
            </div>

            {errorMsg && (
              <div className="p-2.5 text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-md leading-relaxed">
                {errorMsg}
              </div>
            )}

            <button
              type="button"
              disabled={isOpeningLoginWin || isLoggingIn}
              onClick={() => void handleOpenSgwLoginWindow()}
              className="w-full py-2.5 px-3 text-xs font-bold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-md shadow-sm flex items-center justify-center gap-2 transition-colors"
            >
              <ExternalLink className="w-4 h-4 shrink-0" />
              <span>
                {isOpeningLoginWin
                  ? 'SGW 공식 로그인 창에서 인증 대기 중...'
                  : 'SGW 로그인 창 띄워 완전 연동하기'}
              </span>
            </button>

            <div className="space-y-2 pt-1 border-t border-[#EDEBE9]">
              <div className="flex items-center justify-between">
                <button
                  type="button"
                  onClick={() => setShowDirectCredentialForm((v) => !v)}
                  className="text-[11px] font-medium text-[#616161] hover:text-[#0F6CBD] underline"
                >
                  {showDirectCredentialForm
                    ? '계정 직접 입력 닫기'
                    : '브라우저용 사원 계정 직접 입력으로 연결'}
                </button>
                <button
                  type="button"
                  onClick={() => setShowAdvancedUrl((v) => !v)}
                  className="text-[11px] text-[#616161] hover:text-[#0F6CBD] underline"
                >
                  {showAdvancedUrl ? '서버 주소 숨기기' : '서버 주소 변경'}
                </button>
              </div>

              {showAdvancedUrl && (
                <input
                  type="url"
                  value={sgwBaseUrl}
                  onChange={(e) => setSgwBaseUrl(e.target.value)}
                  className="w-full px-2.5 py-1.5 text-xs font-mono border border-[#D1D1D1] rounded-md"
                />
              )}

              {showDirectCredentialForm && (
                <form onSubmit={handleLogin} className="space-y-2.5 pt-1">
                  <div>
                    <label className="block text-[11px] font-semibold text-[#424242] mb-1">
                      사내 이메일 / 성명 / 아이디
                    </label>
                    <input
                      type="text"
                      required
                      value={loginEmail}
                      onChange={(e) => setLoginEmail(e.target.value)}
                      placeholder="사내 이메일 입력"
                      className="w-full px-3 py-2 text-xs border border-[#D1D1D1] rounded-md focus:outline-none focus:border-[#0F6CBD]"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-semibold text-[#424242] mb-1">
                      SGW 비밀번호
                    </label>
                    <input
                      type="password"
                      required
                      value={loginPassword}
                      onChange={(e) => setLoginPassword(e.target.value)}
                      placeholder="SGW 로그인 비밀번호"
                      className="w-full px-3 py-2 text-xs border border-[#D1D1D1] rounded-md focus:outline-none focus:border-[#0F6CBD]"
                    />
                  </div>
                  <button
                    type="submit"
                    disabled={isLoggingIn}
                    className="w-full py-2 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] border border-[#0F6CBD]/30 rounded-md"
                  >
                    {isLoggingIn ? '연결 중...' : '입력한 계정으로 직접 연결'}
                  </button>
                </form>
              )}
            </div>

            <div className="pt-2 border-t border-[#EDEBE9] flex items-center justify-between text-[11px] text-[#616161]">
              <span>세션 키: 로컬 AES-256-GCM 암호화</span>
              <a
                href={sgwBaseUrl}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-1 text-[#0F6CBD] hover:underline"
              >
                SGW 웹 열기 <ExternalLink className="w-3 h-3" />
              </a>
            </div>
          </div>
        </div>
      </aside>
    );
  }

  return (
    <aside className={containerCls} aria-label="SGW 실시간 사내 채팅">
      {/* Header */}
      <div className="px-4 py-2.5 bg-white border-b border-[#E0E0E0] flex items-center justify-between gap-2 shrink-0">
        <div className="min-w-0">
          <div className="flex items-center gap-1.5">
            <MessageSquare className="w-4 h-4 text-[#0F6CBD] shrink-0" />
            <span className="text-xs font-bold text-[#242424] truncate">SGW 실시간 채팅</span>
            <span className="px-1.5 py-0.5 text-[10px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 rounded">
              연결됨
            </span>
          </div>
          <p className="text-[11px] text-[#616161] truncate">
            {user?.name} {user?.position} · {user?.department || '사내 그룹웨어'}
          </p>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onClick={() => void refreshDirectoryAndRooms()}
            title="새로고침"
            className="p-1.5 text-[#616161] hover:text-[#0F6CBD] rounded hover:bg-[#F5F5F5]"
          >
            <RefreshCw className="w-3.5 h-3.5" />
          </button>
          <a
            href={sgwBaseUrl}
            target="_blank"
            rel="noreferrer"
            title="SGW 웹앱 열기"
            className="p-1.5 text-[#616161] hover:text-[#0F6CBD] rounded hover:bg-[#F5F5F5]"
          >
            <ExternalLink className="w-3.5 h-3.5" />
          </a>
          <button
            type="button"
            onClick={() => void handleLogout()}
            title="SGW 채팅 로그아웃"
            className="p-1.5 text-[#616161] hover:text-rose-600 rounded hover:bg-[#F5F5F5]"
          >
            <LogOut className="w-3.5 h-3.5" />
          </button>
          {onClose && (
            <button
              type="button"
              onClick={onClose}
              aria-label="채팅 창 닫기"
              className="p-1.5 text-[#616161] hover:text-[#242424] rounded hover:bg-[#F5F5F5]"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* Mode Tabs */}
      <div className="px-3 pt-2 bg-[#FAFAFA] border-b border-[#E0E0E0] flex items-center gap-1 shrink-0">
        {(
          [
            { id: 'rooms', label: `대화방 (${rooms.length})` },
            { id: 'dm', label: `쪽지 (${dms.length})` },
            { id: 'people', label: `임직원 (${filteredEmployees.length})` },
          ] as const
        ).map((tab) => (
          <button
            key={tab.id}
            type="button"
            onClick={() => setActiveTab(tab.id)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-t-md border-x border-t transition-colors ${
              activeTab === tab.id
                ? 'bg-white border-[#E0E0E0] text-[#0F6CBD]'
                : 'bg-transparent border-transparent text-[#616161] hover:text-[#242424]'
            }`}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {errorMsg && (
        <div className="mx-3 mt-2 px-3 py-1.5 bg-rose-50 border border-rose-200 rounded text-[11px] text-rose-700 flex items-center justify-between gap-2 shrink-0">
          <span>{errorMsg}</span>
          <button type="button" onClick={() => setErrorMsg(null)} className="underline">
            닫기
          </button>
        </div>
      )}

      {/* Pending Mail Share Bar */}
      {pendingShareMail && (
        <div className="mx-3 mt-2 p-2.5 bg-[#EBF3FC] border border-[#CFE4FA] rounded-lg space-y-2 shrink-0">
          <div className="flex items-start justify-between gap-2">
            <div className="min-w-0">
              <div className="text-[11px] font-bold text-[#0F6CBD]">선택한 메일 채팅 공유</div>
              <div className="text-xs font-medium text-[#242424] truncate">
                {pendingShareMail.subject || '(제목 없음)'}
              </div>
            </div>
            <button
              type="button"
              onClick={onClearPendingShare}
              className="text-[11px] text-[#616161] hover:text-[#242424]"
            >
              취소
            </button>
          </div>
          {pendingShareMail.id.startsWith('pst-') &&
            pendingShareMail.attachments &&
            pendingShareMail.attachments.length > 0 && (
              <select
                value={shareAttachmentIdx ?? ''}
                onChange={(e) =>
                  setShareAttachmentIdx(e.target.value === '' ? null : Number(e.target.value))
                }
                aria-label="함께 첨부할 PST 파일"
                className="w-full px-2 py-1 text-[11px] bg-white border border-[#CFE4FA] rounded"
              >
                <option value="">메일 요약·딥링크만 공유 (첨부 제외)</option>
                {pendingShareMail.attachments.map((att, idx) => (
                  <option key={`${att.filename}-${idx}`} value={idx}>
                    첨부 포함: {att.filename}
                    {att.sizeBytes ? ` (${Math.round(att.sizeBytes / 1024)} KB)` : ''}
                  </option>
                ))}
              </select>
            )}
          <div className="flex gap-1.5">
            <input
              type="text"
              value={shareNote}
              onChange={(e) => setShareNote(e.target.value)}
              placeholder="전달 메모 (선택)"
              className="flex-1 px-2 py-1 text-xs bg-white border border-[#CFE4FA] rounded"
            />
            <button
              type="button"
              disabled={isSharingMail}
              onClick={() => void handleShareMailNow()}
              className="px-3 py-1 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded whitespace-nowrap"
            >
              {isSharingMail ? '전송 중...' : '채팅으로 공유'}
            </button>
          </div>
        </div>
      )}

      {/* Body Area */}
      {activeTab === 'people' ? (
        <div className="flex-1 flex flex-col min-h-0 p-3 space-y-2">
          <input
            type="search"
            value={peopleSearch}
            onChange={(e) => setPeopleSearch(e.target.value)}
            placeholder="이름, 부서, 이메일 검색..."
            className="w-full px-3 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-md"
          />
          <div className="flex-1 overflow-y-auto divide-y divide-[#EDEBE9] bg-white border border-[#E0E0E0] rounded-lg">
            {filteredEmployees.length === 0 ? (
              <div className="p-6 text-center text-xs text-[#616161]">표시할 사내 임직원이 없습니다.</div>
            ) : (
              filteredEmployees.map((emp) => (
                <div key={emp.id} className="p-2.5 flex items-center justify-between gap-2 hover:bg-[#FAF9F8]">
                  <div className="min-w-0">
                    <div className="text-xs font-semibold text-[#242424] truncate">
                      {emp.name} {emp.position}
                    </div>
                    <div className="text-[11px] text-[#616161] truncate">
                      {emp.department || '부서 미지정'} · {emp.email}
                    </div>
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <button
                      type="button"
                      onClick={() => void handleStartChatWithEmployee(emp)}
                      className="px-2.5 py-1 text-[11px] font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded"
                    >
                      대화
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setSelectedDmPeerId(emp.id);
                        setActiveTab('dm');
                      }}
                      className="px-2 py-1 text-[11px] text-[#424242] bg-[#F5F5F5] hover:bg-[#E0E0E0] rounded"
                    >
                      쪽지
                    </button>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      ) : (
        <div
          className={`flex-1 flex ${mode === 'full' ? 'flex-row' : 'flex-col'} min-h-0 overflow-hidden`}
        >
          {/* Room / Peer Selector */}
          <div
            className={
              mode === 'full'
                ? 'w-72 shrink-0 border-r border-[#E0E0E0] bg-white overflow-y-auto divide-y divide-[#EDEBE9]'
                : 'max-h-36 shrink-0 border-b border-[#E0E0E0] bg-white overflow-y-auto divide-y divide-[#EDEBE9]'
            }
          >
            {activeTab === 'rooms' ? (
              rooms.length === 0 ? (
                <div className="p-4 text-center text-xs text-[#616161] flex flex-col items-center gap-1.5">
                  <Users className="w-4 h-4 text-[#0F6CBD]" />
                  <span>참여 중인 대화방이 없습니다. 임직원 탭에서 대화를 시작하세요.</span>
                </div>
              ) : (
                rooms.map((room) => {
                  const active = room.id === selectedRoomId;
                  return (
                    <button
                      key={room.id}
                      type="button"
                      onClick={() => setSelectedRoomId(room.id)}
                      className={`w-full text-left px-3 py-2 flex items-center justify-between gap-2 transition-colors ${
                        active ? 'bg-[#EBF3FC]' : 'hover:bg-[#FAF9F8]'
                      }`}
                    >
                      <div className="min-w-0 flex-1">
                        <div className="text-xs font-semibold text-[#242424] truncate">
                          {getRoomDisplayName(room)}
                        </div>
                        <div className="text-[11px] text-[#616161] truncate">
                          {room.lastMessage || '(메시지 없음)'}
                        </div>
                      </div>
                      {(room.unreadCount || 0) > 0 && (
                        <span className="px-1.5 py-0.5 text-[10px] font-bold bg-rose-600 text-white rounded-full shrink-0">
                          {room.unreadCount}
                        </span>
                      )}
                    </button>
                  );
                })
              )
            ) : (
              employees
                .filter((e) => !user || e.id !== user.sub)
                .map((emp) => {
                  const active = emp.id === selectedDmPeerId;
                  const unread = user
                    ? dms.filter((m) => m.senderId === emp.id && m.receiverId === user.sub && !m.isRead).length
                    : 0;
                  return (
                    <button
                      key={emp.id}
                      type="button"
                      onClick={() => setSelectedDmPeerId(emp.id)}
                      className={`w-full text-left px-3 py-2 flex items-center justify-between gap-2 ${
                        active ? 'bg-[#EBF3FC]' : 'hover:bg-[#FAF9F8]'
                      }`}
                    >
                      <div className="min-w-0">
                        <div className="text-xs font-semibold text-[#242424] truncate">
                          {emp.name} {emp.position}
                        </div>
                        <div className="text-[11px] text-[#616161] truncate">{emp.department}</div>
                      </div>
                      {unread > 0 && (
                        <span className="px-1.5 py-0.5 text-[10px] font-bold bg-rose-600 text-white rounded-full">
                          {unread}
                        </span>
                      )}
                    </button>
                  );
                })
            )}
          </div>

          {/* Messages Stream + Composer */}
          <div className="flex-1 flex flex-col min-h-0 bg-[#FAF9F8]">
            <div className="flex-1 overflow-y-auto p-3 space-y-2.5">
              {activeTab === 'rooms' ? (
                roomMessages.length === 0 ? (
                  <div className="py-10 text-center text-xs text-[#616161]">
                    대화 내용을 선택하거나 첫 메시지를 보내 보세요.
                  </div>
                ) : (
                  roomMessages.map((msg) => {
                    const isMine = Boolean(user && msg.senderId === user.sub);
                    const sender = empById.get(msg.senderId);
                    const senderLabel = sender ? `${sender.name} ${sender.position}`.trim() : msg.senderId;
                    return (
                      <div
                        key={msg.id}
                        className={`flex flex-col ${isMine ? 'items-end' : 'items-start'} gap-1`}
                      >
                        <div className="text-[10px] text-[#616161]">
                          {isMine ? '나' : senderLabel} · {formatChatTime(msg.timestamp)}
                        </div>
                        <div
                          className={`max-w-[88%] rounded-xl px-3 py-2 border ${
                            isMine
                              ? 'bg-[#EBF3FC] border-[#CFE4FA] text-[#242424]'
                              : 'bg-white border-[#E0E0E0] text-[#242424]'
                          }`}
                        >
                          {msg.isDeleted ? (
                            <span className="text-xs italic text-[#8A8886]">삭제된 메시지입니다.</span>
                          ) : (
                            renderMessageContent(msg.content)
                          )}
                          {msg.hasFile && !msg.isDeleted && (
                            <button
                              type="button"
                              onClick={() =>
                                void handleDownloadChatAttachment({
                                  kind: 'room',
                                  roomId: selectedRoomId,
                                  msgId: msg.id,
                                  fileName: msg.fileName,
                                })
                              }
                              className="mt-1.5 inline-flex items-center gap-1 px-2 py-1 text-[11px] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded text-[#0F6CBD]"
                            >
                              <Download className="w-3 h-3" />
                              {msg.fileName || '첨부파일 다운로드'}
                            </button>
                          )}
                        </div>
                        <div className="flex items-center gap-1 flex-wrap">
                          {QUICK_REACTIONS.map((em) => {
                            const usersReacted = msg.reactions?.[em] || [];
                            const mineReacted = Boolean(user && usersReacted.includes(user.sub));
                            return (
                              <button
                                key={em}
                                type="button"
                                onClick={() => void handleReaction(msg.id, em, mineReacted)}
                                className={`px-1.5 py-0.5 text-[10px] rounded border ${
                                  mineReacted
                                    ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD]'
                                    : 'bg-white border-[#E0E0E0] text-[#616161] hover:border-[#0F6CBD]'
                                }`}
                              >
                                {em} {usersReacted.length > 0 ? usersReacted.length : ''}
                              </button>
                            );
                          })}
                          {onComposeFromChat && !msg.isDeleted && (
                            <button
                              type="button"
                              onClick={() =>
                                onComposeFromChat({
                                  to: sender?.email || '',
                                  subject: `[SGW 채팅 회신] ${senderLabel}님 전달 건`,
                                  body: `${msg.content}\n\n---\nSGW 대화방에서 메일로 회신 작성`,
                                })
                              }
                              className="inline-flex items-center gap-1 px-1.5 py-0.5 text-[10px] text-[#0F6CBD] hover:underline"
                            >
                              <Mail className="w-2.5 h-2.5" />
                              메일로 회신
                            </button>
                          )}
                        </div>
                      </div>
                    );
                  })
                )
              ) : dmConversation.length === 0 ? (
                <div className="py-10 text-center text-xs text-[#616161]">
                  상대 직원을 선택하고 1:1 쪽지를 주고받을 수 있습니다.
                </div>
              ) : (
                dmConversation.map((m) => {
                  const isMine = Boolean(user && m.senderId === user.sub);
                  const peer = empById.get(isMine ? m.receiverId : m.senderId);
                  return (
                    <div
                      key={m.id}
                      className={`flex flex-col ${isMine ? 'items-end' : 'items-start'} gap-1`}
                    >
                      <div className="text-[10px] text-[#616161] flex items-center gap-1">
                        <span>
                          {isMine ? '나' : m.senderName || peer?.name} · {formatChatTime(m.timestamp)}
                        </span>
                        {isMine && m.isRead && <Check className="w-2.5 h-2.5 text-[#0F6CBD]" />}
                      </div>
                      <div
                        className={`max-w-[88%] rounded-xl px-3 py-2 border ${
                          isMine
                            ? 'bg-[#EBF3FC] border-[#CFE4FA]'
                            : 'bg-white border-[#E0E0E0]'
                        }`}
                      >
                        {renderMessageContent(m.content)}
                        {m.hasFile && (
                          <button
                            type="button"
                            onClick={() =>
                              void handleDownloadChatAttachment({
                                kind: 'dm',
                                msgId: m.id,
                                fileName: m.fileName,
                              })
                            }
                            className="mt-1.5 inline-flex items-center gap-1 px-2 py-1 text-[11px] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded text-[#0F6CBD]"
                          >
                            <Download className="w-3 h-3" />
                            {m.fileName || '첨부파일'}
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })
              )}
              <div ref={messagesEndRef} />
            </div>

            {/* Composer */}
            <div className="p-2.5 bg-white border-t border-[#E0E0E0] space-y-1.5 shrink-0">
              {attachFile && (
                <div className="flex items-center justify-between px-2 py-1 bg-[#EBF3FC] rounded text-[11px] text-[#0F6CBD]">
                  <span className="truncate">{attachFile.name}</span>
                  <button type="button" onClick={() => setAttachFile(null)}>
                    <X className="w-3 h-3" />
                  </button>
                </div>
              )}
              <div className="flex items-end gap-1.5">
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  onChange={(e) => handleAttachFileChange(e.target.files)}
                />
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  title="파일 첨부 (최대 5MB)"
                  className="p-2 text-[#616161] hover:text-[#0F6CBD] hover:bg-[#F5F5F5] rounded-md shrink-0"
                >
                  <Paperclip className="w-4 h-4" />
                </button>
                <textarea
                  rows={2}
                  value={draftText}
                  onChange={(e) => setDraftText(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && !e.shiftKey) {
                      e.preventDefault();
                      void handleSendMessage();
                    }
                  }}
                  placeholder={
                    activeTab === 'rooms'
                      ? 'SGW 대화방 메시지 입력 (Enter 전송)'
                      : '쪽지 내용 입력 (Enter 전송)'
                  }
                  className="flex-1 px-2.5 py-1.5 text-xs bg-[#FAF9F8] border border-[#D1D1D1] rounded-md resize-none focus:outline-none focus:border-[#0F6CBD]"
                />
                <button
                  type="button"
                  disabled={isSending || (!draftText.trim() && !attachFile)}
                  onClick={() => void handleSendMessage()}
                  aria-label="메시지 전송"
                  className="p-2 text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-md shrink-0"
                >
                  <Send className="w-4 h-4" />
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </aside>
  );
};
