import React, { useEffect, useState } from 'react';
import {
  AlertCircle,
  Check,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  Copy,
  Database,
  ExternalLink,
  Eye,
  EyeOff,
  Globe,
  KeyRound,
  Loader2,
  Lock,
  Mail,
  Plug,
  RefreshCw,
  Search,
  Server,
  ShieldCheck,
  Wand2,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';

export interface Pop3ConnectedAccount {
  email: string;
  providerId: string;
  providerName: string;
  protocol?: 'imap' | 'pop3';
  pop3Host: string;
  pop3Port: number;
  pop3Tls: boolean;
  imapHost?: string;
  imapPort?: number;
  imapTls?: boolean;
  smtpHost: string;
  smtpPort: number;
  smtpTls: boolean;
}

interface DiscoveredConfigView {
  email: string;
  domain: string;
  detectedProviderId: string;
  detectedProviderName: string;
  confidence: 'high' | 'medium' | 'fallback';
  detectionMethod: string;
  mxRecords: string[];
  supportedProtocols: Array<'imap' | 'pop3'>;
  recommendedProtocol: 'imap' | 'pop3';
  pop3Host: string;
  pop3Port: number;
  pop3Tls: boolean;
  pop3Reachable: boolean;
  pop3Banner?: string;
  imapHost: string;
  imapPort: number;
  imapTls: boolean;
  imapReachable: boolean;
  imapBanner?: string;
  smtpHost: string;
  smtpPort: number;
  smtpTls: boolean;
  reachable: boolean;
  usernameFormat: 'full-email' | 'local-or-full';
  appPasswordPolicy: 'required' | 'recommended' | 'optional';
  overseasIpWarning: boolean;
  appPasswordGuide: string;
  protocolSetupGuide: string;
  guidanceNote: string;
}

interface ServerEgressView {
  ipv4: string | null;
  countryCode: string | null;
  countryName: string | null;
  city: string | null;
  isKorea: boolean | null;
}

interface AutoHealAttemptView {
  step: number;
  protocol: 'imap' | 'pop3';
  host: string;
  port: number;
  strategyLabel: string;
  outcome: 'success' | 'failed';
  serverCode?: string;
}

interface ResolutionGuideStepView {
  id: string;
  priority: 'critical' | 'recommended' | 'alternative';
  title: string;
  menuPath: string;
  detail: string;
  copyValue?: string;
  copyLabel?: string;
  helpUrl?: string;
  helpLabel?: string;
}

interface MailConnectionDiagnosisView {
  code: string;
  headline: string;
  checklist: string[];
  suggestProtocolSwitch?: 'imap' | 'pop3';
  rawServerError: string;
  serverEgress?: ServerEgressView;
  autoHealAttempts?: AutoHealAttemptView[];
  resolutionSteps?: ResolutionGuideStepView[];
}

const SERVICE_OPTIONS = [
  { id: 'auto', label: '자동 감지 (회사 도메인 MX · IMAP/POP3 자동 조회)' },
  { id: 'mailplug', label: '메일플러그 (Mailplug) — POP3 권장 (2025.04~ IMAP 제한) / 앱비번' },
  { id: 'hiworks', label: '하이웍스 (Hiworks) — POP3 전용 (995 SSL)' },
  { id: 'gabia', label: '가비아 기업메일 / 하이메일 (Gabia) — POP3' },
  { id: 'naverworks', label: '네이버 웍스 (NAVER WORKS) — IMAP·POP3' },
  { id: 'daouoffice', label: '다우오피스 (DaouOffice) — IMAP·POP3' },
  { id: 'dooray', label: '두레이 (NHN Dooray) — IMAP 전용 (993 SSL)' },
  { id: 'cafe24', label: '카페24 웹메일 (Cafe24) — POP3 전용 (webmail.도메인)' },
  { id: 'hanbiro', label: '한비로 그룹웨어 (Hanbiro) — IMAP·POP3' },
  { id: 'kakaowork', label: '카카오워크 메일 (Kakao Work) — IMAP·POP3' },
  { id: 'whois', label: '후이즈 메일 (Whois) — IMAP·POP3' },
  { id: 'google', label: 'Google Workspace / Gmail — IMAP 권장 / 16자리 앱비번' },
  { id: 'microsoft', label: 'Microsoft 365 / Outlook — IMAP·POP3 (OAuth 권장)' },
  { id: 'naver', label: '네이버 메일 (Naver) — IMAP·POP3 / 앱비번' },
  { id: 'daum', label: '다음 / 카카오 메일 (Daum) — IMAP·POP3 / 앱비번' },
  { id: 'nate', label: '네이트 메일 (Nate) — IMAP·POP3' },
  { id: 'icloud', label: 'Apple iCloud 메일 — IMAP 전용 (993 SSL)' },
  { id: 'zoho', label: 'Zoho Mail — IMAP·POP3 / 앱비번' },
  { id: 'custom', label: '자사 메일 서버 (직접 입력 / 자동 탐색)' },
];

const FLUENT_INPUT_CLS =
  'w-full px-3 py-2 text-xs bg-white border border-[var(--fluent-stroke-1)] rounded-[4px] text-[var(--fluent-fg-1)] placeholder:text-[var(--fluent-fg-3)] focus:outline-none focus:border-[var(--fluent-brand)] focus:ring-2 focus:ring-[#0F6CBD]/25 transition-all duration-150';

interface InitialMailAccountHint {
  email?: string;
  providerPreset?: string;
  protocol?: 'pop3' | 'imap';
  pop3Host?: string;
  pop3Port?: number;
  smtpHost?: string;
  smtpPort?: number;
}

interface Pop3QuickConnectModalProps {
  isOpen: boolean;
  connectedAccount: Pop3ConnectedAccount | null;
  initialAccount?: InitialMailAccountHint | null;
  onClose: () => void;
  onConnected: (payload: {
    sessionToken: string;
    account: Pop3ConnectedAccount;
    emails: MailItem[];
    totalMailboxCount: number;
  }) => void;
  onDisconnect: () => void;
  onOpenPstFallback?: () => void;
}

export const Pop3QuickConnectModal: React.FC<Pop3QuickConnectModalProps> = ({
  isOpen,
  connectedAccount,
  initialAccount,
  onClose,
  onConnected,
  onDisconnect,
  onOpenPstFallback,
}) => {
  const [providerPreset, setProviderPreset] = useState(initialAccount?.providerPreset || 'auto');
  const [protocol, setProtocol] = useState<'imap' | 'pop3'>(initialAccount?.protocol || 'pop3');
  const [email, setEmail] = useState(connectedAccount?.email || initialAccount?.email || '');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);

  const [isDiscovering, setIsDiscovering] = useState(false);
  const [discovered, setDiscovered] = useState<DiscoveredConfigView | null>(null);
  const [serverEgress, setServerEgress] = useState<ServerEgressView | null>(null);
  const [showAdvanced, setShowAdvanced] = useState(false);

  const [pop3Host, setPop3Host] = useState(initialAccount?.pop3Host || '');
  const [pop3Port, setPop3Port] = useState(initialAccount?.pop3Port || 995);
  const [pop3Tls, setPop3Tls] = useState(true);
  const [imapHost, setImapHost] = useState('');
  const [imapPort, setImapPort] = useState(993);
  const [imapTls, setImapTls] = useState(true);
  const [smtpHost, setSmtpHost] = useState(initialAccount?.smtpHost || '');
  const [smtpPort, setSmtpPort] = useState(initialAccount?.smtpPort || 465);
  const [smtpTls, setSmtpTls] = useState(true);
  const [maxMessages, setMaxMessages] = useState(30);

  const [isConnecting, setIsConnecting] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [diagnosis, setDiagnosis] = useState<MailConnectionDiagnosisView | null>(null);
  const [checkedSteps, setCheckedSteps] = useState<Record<string, boolean>>({});
  const [copiedValue, setCopiedValue] = useState<string | null>(null);

  useEffect(() => {
    if (!isOpen || connectedAccount || !initialAccount) return;
    if (!email && initialAccount.email) setEmail(initialAccount.email);
    if (providerPreset === 'auto' && initialAccount.providerPreset) {
      setProviderPreset(initialAccount.providerPreset);
    }
    if (initialAccount.protocol) setProtocol(initialAccount.protocol);
    if (!pop3Host && initialAccount.pop3Host) setPop3Host(initialAccount.pop3Host);
    if (initialAccount.pop3Port) setPop3Port(initialAccount.pop3Port);
    if (!smtpHost && initialAccount.smtpHost) setSmtpHost(initialAccount.smtpHost);
    if (initialAccount.smtpPort) setSmtpPort(initialAccount.smtpPort);
  }, [isOpen, connectedAccount, initialAccount, email, providerPreset, pop3Host, smtpHost]);

  if (!isOpen) return null;

  const handleCopyText = async (value: string) => {
    try {
      await navigator.clipboard.writeText(value);
      setCopiedValue(value);
      window.setTimeout(() => {
        setCopiedValue((prev) => (prev === value ? null : prev));
      }, 2200);
    } catch {
      // ignore clipboard denial
    }
  };

  const toggleStepChecked = (stepId: string) => {
    setCheckedSteps((prev) => ({ ...prev, [stepId]: !prev[stepId] }));
  };

  const handleAutoDiscover = async (overridePreset?: string) => {
    const cleanEmail = email.trim();
    if (!cleanEmail || !cleanEmail.includes('@')) {
      setErrorMsg('올바른 회사 이메일 주소를 먼저 입력해 주세요.');
      return;
    }
    setIsDiscovering(true);
    setErrorMsg(null);
    setDiagnosis(null);
    try {
      const res = await fetch('/api/pop3/autodiscover', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email: cleanEmail,
          providerPreset: overridePreset ?? providerPreset,
        }),
      });
      const data = (await res.json()) as {
        config?: DiscoveredConfigView;
        serverEgress?: ServerEgressView;
        error?: string;
      };
      if (!res.ok || !data.config) {
        setErrorMsg(data.error || '메일 서버 설정 자동 조회에 실패했습니다.');
        return;
      }
      setDiscovered(data.config);
      if (data.serverEgress) {
        setServerEgress(data.serverEgress);
      }
      setProtocol(data.config.recommendedProtocol);
      setPop3Host(data.config.pop3Host);
      setPop3Port(data.config.pop3Port);
      setPop3Tls(data.config.pop3Tls);
      setImapHost(data.config.imapHost);
      setImapPort(data.config.imapPort);
      setImapTls(data.config.imapTls);
      setSmtpHost(data.config.smtpHost);
      setSmtpPort(data.config.smtpPort);
      setSmtpTls(data.config.smtpTls);
    } catch {
      setErrorMsg('서버와 통신하지 못해 메일 설정 조회에 실패했습니다.');
    } finally {
      setIsDiscovering(false);
    }
  };

  const runConnect = async (overrideProtocol?: 'imap' | 'pop3') => {
    if (!email.trim() || !password) {
      setErrorMsg('이메일 주소와 비밀번호(또는 앱 비밀번호)를 모두 입력해 주세요.');
      return;
    }
    const targetProtocol = overrideProtocol ?? protocol;
    if (overrideProtocol) setProtocol(overrideProtocol);

    setIsConnecting(true);
    setErrorMsg(null);
    setDiagnosis(null);
    try {
      const res = await fetch('/api/pop3/connect', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email: email.trim(),
          password,
          providerPreset,
          protocol: targetProtocol,
          pop3Host: showAdvanced && pop3Host.trim() ? pop3Host.trim() : undefined,
          pop3Port: showAdvanced ? pop3Port : undefined,
          pop3Tls: showAdvanced ? pop3Tls : undefined,
          imapHost: showAdvanced && imapHost.trim() ? imapHost.trim() : undefined,
          imapPort: showAdvanced ? imapPort : undefined,
          imapTls: showAdvanced ? imapTls : undefined,
          smtpHost: showAdvanced && smtpHost.trim() ? smtpHost.trim() : undefined,
          smtpPort: showAdvanced ? smtpPort : undefined,
          smtpTls: showAdvanced ? smtpTls : undefined,
          maxMessages,
        }),
      });
      const data = (await res.json()) as {
        sessionToken?: string;
        account?: Pop3ConnectedAccount;
        discovered?: DiscoveredConfigView;
        emails?: MailItem[];
        totalMailboxCount?: number;
        error?: string;
        diagnosis?: MailConnectionDiagnosisView;
      };
      if (!res.ok || !data.sessionToken || !data.account) {
        setErrorMsg(data.error || '메일 서버 연결 및 메일 로드에 실패했습니다.');
        setDiagnosis(data.diagnosis ?? null);
        if (data.diagnosis?.serverEgress) {
          setServerEgress(data.diagnosis.serverEgress);
        }
        return;
      }
      setPassword('');
      onConnected({
        sessionToken: data.sessionToken,
        account: data.account,
        emails: data.emails || [],
        totalMailboxCount: data.totalMailboxCount ?? (data.emails?.length || 0),
      });
      onClose();
    } catch {
      setErrorMsg('메일 서버 연결 중 네트워크 오류가 발생했습니다.');
    } finally {
      setIsConnecting(false);
    }
  };

  const handleConnect = async (e: React.FormEvent) => {
    e.preventDefault();
    await runConnect();
  };

  const supportsImap = !discovered || discovered.supportedProtocols.includes('imap');
  const supportsPop3 = !discovered || discovered.supportedProtocols.includes('pop3');

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="회사 메일 IMAP · POP3 자동 분석 및 연결"
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 p-3 sm:p-4"
    >
      <div className="w-full max-w-xl bg-white rounded-[8px] border border-[var(--fluent-stroke-2)] shadow-[var(--fluent-shadow-16)] overflow-hidden flex flex-col max-h-[92vh]">
        {/* Fluent 2 Header */}
        <div className="px-5 py-3.5 bg-[var(--fluent-brand)] text-white flex items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2 min-w-0">
            <Plug className="w-4 h-4 shrink-0" strokeWidth={1.75} />
            <h2 className="text-sm font-semibold truncate">
              메일 서비스 자동 분석 · IMAP / POP3 연결
            </h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="닫기"
            className="p-1 rounded-[4px] hover:bg-white/15 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleConnect} className="p-5 space-y-4 overflow-y-auto">
          {connectedAccount && (
            <div className="p-3 rounded-[6px] bg-[var(--fluent-success-bg)] border border-[var(--fluent-success-border)] flex items-center justify-between gap-2 text-xs">
              <div className="min-w-0">
                <div className="font-semibold text-[#107C41] truncate">
                  연결됨 ({(connectedAccount.protocol || 'pop3').toUpperCase()}): {connectedAccount.email}
                </div>
                <div className="text-[11px] text-[#107C41] font-mono truncate">
                  {connectedAccount.pop3Host}:{connectedAccount.pop3Port} · {connectedAccount.providerName}
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  onDisconnect();
                  onClose();
                }}
                className="px-2.5 py-1 text-[11px] font-semibold text-[#C50F1F] bg-white border border-[#EEACB2] hover:bg-[#FDF3F4] rounded-[4px] shrink-0"
              >
                연결 해제
              </button>
            </div>
          )}

          {/* 1. Service Preset */}
          <div className="space-y-1">
            <label htmlFor="pop3-service-preset" className="block text-xs font-semibold text-[var(--fluent-fg-1)]">
              1. 메일 서비스 형태 (국내·글로벌 18개 서비스 특징 반영)
            </label>
            <select
              id="pop3-service-preset"
              value={providerPreset}
              onChange={(e) => {
                const next = e.target.value;
                setProviderPreset(next);
                if (email.includes('@')) {
                  void handleAutoDiscover(next);
                }
              }}
              className={FLUENT_INPUT_CLS}
            >
              {SERVICE_OPTIONS.map((opt) => (
                <option key={opt.id} value={opt.id}>
                  {opt.label}
                </option>
              ))}
            </select>
          </div>

          {/* 2. Email + Auto Discover */}
          <div className="space-y-1">
            <label htmlFor="pop3-email" className="block text-xs font-semibold text-[var(--fluent-fg-1)]">
              2. 회사 / 개인 이메일 주소
            </label>
            <div className="flex gap-2">
              <div className="relative flex-1">
                <Mail className="w-3.5 h-3.5 text-[var(--fluent-fg-3)] absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  id="pop3-email"
                  type="email"
                  required
                  autoComplete="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  onBlur={() => {
                    if (email.includes('@') && !discovered) {
                      void handleAutoDiscover();
                    }
                  }}
                  placeholder="예: hong@mycompany.co.kr"
                  className={`${FLUENT_INPUT_CLS} pl-8`}
                />
              </div>
              <button
                type="button"
                disabled={isDiscovering || !email.includes('@')}
                onClick={() => void handleAutoDiscover()}
                className="px-3 py-2 text-xs font-semibold text-[var(--fluent-brand)] bg-[var(--fluent-brand-tint)] hover:bg-[var(--fluent-brand-tint-hover)] disabled:opacity-50 rounded-[4px] whitespace-nowrap inline-flex items-center gap-1.5 transition-colors"
              >
                {isDiscovering ? (
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <Search className="w-3.5 h-3.5" />
                )}
                서버 자동 분석
              </button>
            </div>
          </div>

          {/* Auto-discovered provider profile card */}
          {discovered && (
            <div
              className={`p-3.5 rounded-[6px] border text-xs space-y-2.5 ${
                discovered.reachable
                  ? 'bg-[#F3F9F1] border-[#9FD89F] text-[var(--fluent-fg-1)]'
                  : 'bg-[var(--fluent-bg-subtle)] border-[var(--fluent-stroke-2)] text-[var(--fluent-fg-1)]'
              }`}
            >
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="font-semibold flex items-center gap-1.5">
                  <Server className="w-3.5 h-3.5 text-[var(--fluent-brand)]" />
                  {discovered.detectedProviderName}
                </span>
                <div className="flex flex-wrap items-center gap-1.5">
                  {discovered.imapReachable && (
                    <span className="text-[10px] font-semibold font-mono px-1.5 py-0.5 rounded-[4px] bg-emerald-50 border border-emerald-300 text-emerald-800">
                      IMAP OK
                    </span>
                  )}
                  {discovered.pop3Reachable && (
                    <span className="text-[10px] font-semibold font-mono px-1.5 py-0.5 rounded-[4px] bg-emerald-50 border border-emerald-300 text-emerald-800">
                      POP3 OK
                    </span>
                  )}
                  {!discovered.reachable && (
                    <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded-[4px] bg-amber-50 border border-amber-300 text-amber-800">
                      포트 응답 미확인
                    </span>
                  )}
                </div>
              </div>

              {/* Provider Key Characteristics Tags */}
              <div className="flex flex-wrap items-center gap-1.5 text-[10px]">
                <span
                  className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded-[4px] font-semibold border ${
                    discovered.appPasswordPolicy === 'required'
                      ? 'bg-[#FDF3F4] text-[#C50F1F] border-[#EEACB2]'
                      : discovered.appPasswordPolicy === 'recommended'
                        ? 'bg-[#FFF9E6] text-[#B45309] border-[#F6D88C]'
                        : 'bg-white text-[var(--fluent-fg-2)] border-[var(--fluent-stroke-1)]'
                  }`}
                >
                  <KeyRound className="w-3 h-3" />
                  {discovered.appPasswordPolicy === 'required'
                    ? '앱 비밀번호 필수'
                    : discovered.appPasswordPolicy === 'recommended'
                      ? '전용 비밀번호 권장'
                      : '일반 비밀번호 지원'}
                </span>
                <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-[4px] font-medium bg-white text-[var(--fluent-fg-2)] border border-[var(--fluent-stroke-1)]">
                  지원 프로토콜: {discovered.supportedProtocols.map((p) => p.toUpperCase()).join(' · ')}
                </span>
                {discovered.usernameFormat === 'local-or-full' && (
                  <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-[4px] font-medium bg-white text-[var(--fluent-brand)] border border-[var(--fluent-brand-border)]">
                    아이디/이메일 자동 재시도
                  </span>
                )}
                {discovered.overseasIpWarning && (
                  <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-[4px] font-semibold bg-[#FFF9E6] text-[#B45309] border border-[#F6D88C]">
                    <Globe className="w-3 h-3" />
                    해외 IP 차단 확인 필요
                  </span>
                )}
              </div>

              {/* Protocol Switcher (IMAP vs POP3) */}
              <div className="flex items-center gap-2 pt-0.5">
                <span className="text-[11px] font-medium text-[var(--fluent-fg-2)]">수신 방식:</span>
                <div className="inline-flex rounded-[4px] border border-[var(--fluent-stroke-1)] bg-white overflow-hidden">
                  <button
                    type="button"
                    disabled={!supportsImap}
                    onClick={() => setProtocol('imap')}
                    className={`px-2.5 py-1 text-[11px] font-semibold transition-colors disabled:opacity-40 disabled:cursor-not-allowed ${
                      protocol === 'imap'
                        ? 'bg-[var(--fluent-brand)] text-white'
                        : 'text-[var(--fluent-fg-2)] hover:bg-[var(--fluent-bg-canvas)]'
                    }`}
                  >
                    IMAP ({discovered.imapPort || 993})
                  </button>
                  <button
                    type="button"
                    disabled={!supportsPop3}
                    onClick={() => setProtocol('pop3')}
                    className={`px-2.5 py-1 text-[11px] font-semibold border-l border-[var(--fluent-stroke-1)] transition-colors disabled:opacity-40 disabled:cursor-not-allowed ${
                      protocol === 'pop3'
                        ? 'bg-[var(--fluent-brand)] text-white'
                        : 'text-[var(--fluent-fg-2)] hover:bg-[var(--fluent-bg-canvas)]'
                    }`}
                  >
                    POP3 ({discovered.pop3Port || 995})
                  </button>
                </div>
              </div>

              {/* Endpoints summary */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-1 font-mono text-[11px] text-[var(--fluent-fg-2)]">
                {protocol === 'imap' && discovered.imapHost ? (
                  <div className="truncate">
                    IMAP: {discovered.imapHost}:{discovered.imapPort} ({discovered.imapTls ? 'SSL' : 'PLAIN'})
                  </div>
                ) : (
                  <div className="truncate">
                    POP3: {discovered.pop3Host}:{discovered.pop3Port} ({discovered.pop3Tls ? 'SSL' : 'PLAIN'})
                  </div>
                )}
                <div className="truncate">
                  SMTP: {discovered.smtpHost}:{discovered.smtpPort} ({discovered.smtpTls ? 'SSL' : 'STARTTLS'})
                </div>
              </div>

              {discovered.mxRecords.length > 0 && (
                <div className="text-[11px] text-[var(--fluent-fg-3)] truncate">
                  감지 MX: <span className="font-mono">{discovered.mxRecords.slice(0, 2).join(', ')}</span>
                </div>
              )}

              {/* Provider-Specific Essential Setup Notes */}
              <div className="space-y-1 pt-1.5 border-t border-black/[0.06] text-[11px] text-[var(--fluent-fg-2)] leading-relaxed">
                <p>
                  <strong className="text-[var(--fluent-fg-1)]">비밀번호 정책:</strong> {discovered.appPasswordGuide}
                </p>
                <p>
                  <strong className="text-[var(--fluent-fg-1)]">서버 설정:</strong> {discovered.protocolSetupGuide}
                </p>
              </div>

              {/* Proactive Live Overseas Cloud Egress Warning + 1-Click IP Copy */}
              {discovered.overseasIpWarning && serverEgress && serverEgress.isKorea === false && (
                <div className="mt-1.5 p-2.5 rounded-[5px] bg-[#FFF9E6] border border-[#F6D88C] text-[11px] text-[#7A4400] space-y-1.5">
                  <div className="flex items-center justify-between gap-2 flex-wrap">
                    <span className="inline-flex items-center gap-1.5 font-semibold">
                      <Globe className="w-3.5 h-3.5 text-[#B45309] shrink-0" />
                      현재 클라우드 접속 위치: {serverEgress.countryName || serverEgress.countryCode} ({serverEgress.countryCode})
                      {serverEgress.ipv4 && (
                        <code className="px-1.5 py-0.5 rounded bg-white/90 border border-[#F0C76E] font-mono text-[10px] text-[#242424]">
                          {serverEgress.ipv4}
                        </code>
                      )}
                    </span>
                    {serverEgress.ipv4 && (
                      <button
                        type="button"
                        onClick={() => void handleCopyText(serverEgress.ipv4!)}
                        className="inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] bg-white border border-[#E5B856] hover:bg-[#FFF3CC] text-[10px] font-semibold text-[#7A4400] transition-colors"
                      >
                        {copiedValue === serverEgress.ipv4 ? (
                          <>
                            <Check className="w-3 h-3 text-[#107C10]" />
                            IP 복사됨
                          </>
                        ) : (
                          <>
                            <Copy className="w-3 h-3" />
                            예외 등록용 IP 복사
                          </>
                        )}
                      </button>
                    )}
                  </div>
                  <p className="text-[10.5px] leading-relaxed text-[#6E4008]">
                    앱 비밀번호가 정확해도 웹메일의 <strong>[해외 로그인 차단]</strong>이 켜져 있으면{' '}
                    <code className="font-mono">#5.7.5</code>로 거부됩니다. 연결 전{' '}
                    <strong>{serverEgress.countryName || '해외'}</strong> 또는 위 IP를 허용하거나 사내 Windows <code>.exe</code>에서 실행하세요.
                  </p>
                </div>
              )}
            </div>
          )}

          {/* 3. Password / App Password */}
          <div className="space-y-1">
            <label htmlFor="pop3-password" className="block text-xs font-semibold text-[var(--fluent-fg-1)]">
              3. {discovered?.appPasswordPolicy === 'required' ? '앱 비밀번호 (외부 연동 전용)' : '이메일 비밀번호 (또는 앱 비밀번호)'}
            </label>
            <div className="relative">
              <Lock className="w-3.5 h-3.5 text-[var(--fluent-fg-3)] absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                id="pop3-password"
                type={showPassword ? 'text' : 'password'}
                required
                autoComplete="current-password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder={
                  discovered?.appPasswordPolicy === 'required'
                    ? '웹메일 보안 설정에서 발급한 앱 비밀번호 입력'
                    : '메일 계정 비밀번호 입력'
                }
                className={`${FLUENT_INPUT_CLS} pl-8 pr-9`}
              />
              <button
                type="button"
                onClick={() => setShowPassword((v) => !v)}
                aria-label={showPassword ? '비밀번호 숨기기' : '비밀번호 보기'}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[var(--fluent-fg-3)] hover:text-[var(--fluent-fg-1)]"
              >
                {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          {/* Optional Advanced Host/Port Override */}
          <div className="border-t border-[var(--fluent-stroke-subtle)] pt-2">
            <button
              type="button"
              onClick={() => setShowAdvanced((v) => !v)}
              className="inline-flex items-center gap-1 text-[11px] font-medium text-[var(--fluent-fg-3)] hover:text-[var(--fluent-brand)]"
            >
              {showAdvanced ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
              고급 IMAP / POP3 / SMTP 서버 주소 · 가져올 개수 조정 (선택)
            </button>

            {showAdvanced && (
              <div className="mt-2.5 p-3 rounded-[6px] bg-[var(--fluent-bg-subtle)] border border-[var(--fluent-stroke-subtle)] grid grid-cols-2 gap-2.5 text-xs">
                {protocol === 'imap' ? (
                  <>
                    <div>
                      <label className="block text-[11px] text-[var(--fluent-fg-3)] mb-1">IMAP 서버</label>
                      <input
                        type="text"
                        value={imapHost}
                        onChange={(e) => setImapHost(e.target.value)}
                        placeholder="imap.mailplug.co.kr"
                        className={`${FLUENT_INPUT_CLS} font-mono py-1.5`}
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] text-[var(--fluent-fg-3)] mb-1">IMAP 포트 / SSL</label>
                      <div className="flex items-center gap-2">
                        <input
                          type="number"
                          value={imapPort}
                          onChange={(e) => setImapPort(Number(e.target.value) || 993)}
                          className="w-20 px-2 py-1.5 text-xs font-mono bg-white border border-[var(--fluent-stroke-1)] rounded-[4px] focus:outline-none focus:border-[var(--fluent-brand)] focus:ring-2 focus:ring-[#0F6CBD]/25"
                        />
                        <label className="inline-flex items-center gap-1 text-[11px]">
                          <input
                            type="checkbox"
                            checked={imapTls}
                            onChange={(e) => setImapTls(e.target.checked)}
                            className="accent-[#0F6CBD]"
                          />
                          SSL
                        </label>
                      </div>
                    </div>
                  </>
                ) : (
                  <>
                    <div>
                      <label className="block text-[11px] text-[var(--fluent-fg-3)] mb-1">POP3 서버</label>
                      <input
                        type="text"
                        value={pop3Host}
                        onChange={(e) => setPop3Host(e.target.value)}
                        placeholder="pop3.mailplug.co.kr"
                        className={`${FLUENT_INPUT_CLS} font-mono py-1.5`}
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] text-[var(--fluent-fg-3)] mb-1">POP3 포트 / SSL</label>
                      <div className="flex items-center gap-2">
                        <input
                          type="number"
                          value={pop3Port}
                          onChange={(e) => setPop3Port(Number(e.target.value) || 995)}
                          className="w-20 px-2 py-1.5 text-xs font-mono bg-white border border-[var(--fluent-stroke-1)] rounded-[4px] focus:outline-none focus:border-[var(--fluent-brand)] focus:ring-2 focus:ring-[#0F6CBD]/25"
                        />
                        <label className="inline-flex items-center gap-1 text-[11px]">
                          <input
                            type="checkbox"
                            checked={pop3Tls}
                            onChange={(e) => setPop3Tls(e.target.checked)}
                            className="accent-[#0F6CBD]"
                          />
                          SSL
                        </label>
                      </div>
                    </div>
                  </>
                )}

                <div>
                  <label className="block text-[11px] text-[var(--fluent-fg-3)] mb-1">SMTP 서버 (발송용)</label>
                  <input
                    type="text"
                    value={smtpHost}
                    onChange={(e) => setSmtpHost(e.target.value)}
                    placeholder="smtp.mailplug.co.kr"
                    className={`${FLUENT_INPUT_CLS} font-mono py-1.5`}
                  />
                </div>
                <div>
                  <label className="block text-[11px] text-[var(--fluent-fg-3)] mb-1">최근 가져올 메일 수</label>
                  <input
                    type="number"
                    min={5}
                    max={50}
                    value={maxMessages}
                    onChange={(e) => setMaxMessages(Math.max(5, Math.min(50, Number(e.target.value) || 30)))}
                    className={`${FLUENT_INPUT_CLS} font-mono py-1.5`}
                  />
                </div>
              </div>
            )}
          </div>

          {/* Interactive Auto-Heal Diagnostics & Step-by-Step Resolution Wizard */}
          {(errorMsg || diagnosis) && (
            <div
              role="alert"
              className="p-3.5 rounded-[6px] bg-[#FDF3F4] border border-[#EEACB2] text-xs text-[#6B0F1A] space-y-3"
            >
              <div className="flex items-start gap-2">
                <AlertCircle className="w-4 h-4 text-[#C50F1F] shrink-0 mt-0.5" />
                <div className="min-w-0 space-y-0.5">
                  <div className="font-semibold leading-snug">
                    {diagnosis ? diagnosis.headline : errorMsg}
                  </div>
                  {diagnosis?.rawServerError && (
                    <div className="text-[10px] font-mono text-[#707070] break-all">
                      서버 원문 응답: {diagnosis.rawServerError}
                    </div>
                  )}
                </div>
              </div>

              {/* Transparent Auto-Heal Attempt Log */}
              {diagnosis?.autoHealAttempts && diagnosis.autoHealAttempts.length > 0 && (
                <div className="p-2.5 rounded-[5px] bg-white/90 border border-[#F1CDD1] space-y-1.5">
                  <div className="flex items-center gap-1.5 text-[11px] font-semibold text-[#242424]">
                    <Wand2 className="w-3.5 h-3.5 text-[#0F6CBD]" />
                    자동 해결 엔진이 이미 시도한 항목 ({diagnosis.autoHealAttempts.length}단계 자동 점검 완료)
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {diagnosis.autoHealAttempts.map((att) => (
                      <span
                        key={`${att.step}-${att.protocol}-${att.host}`}
                        className="inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] bg-[#FAF9F8] border border-[#E1DFDD] text-[10px] text-[#424242]"
                      >
                        <span className="font-semibold text-[#C50F1F]">✕</span>
                        <span>{att.strategyLabel}</span>
                        <span className="font-mono text-[#707070]">
                          ({att.protocol.toUpperCase()}:{att.port}
                          {att.serverCode ? ` · ${att.serverCode}` : ''})
                        </span>
                      </span>
                    ))}
                  </div>
                </div>
              )}

              {/* Live Detected Server IP / Country Box for Instant Whitelisting */}
              {(diagnosis?.serverEgress || serverEgress)?.ipv4 &&
                (diagnosis?.serverEgress || serverEgress)?.isKorea === false && (
                  <div className="p-2.5 rounded-[5px] bg-[#FFF9E6] border border-[#F0C76E] flex items-center justify-between gap-2 flex-wrap text-[11px] text-[#7A4400]">
                    <div className="flex items-center gap-1.5 min-w-0">
                      <Globe className="w-3.5 h-3.5 text-[#B45309] shrink-0" />
                      <span className="font-semibold">
                        차단된 접속 IP:{' '}
                        <code className="px-1.5 py-0.5 rounded bg-white border border-[#E5B856] font-mono text-[#242424]">
                          {(diagnosis?.serverEgress || serverEgress)?.ipv4}
                        </code>
                      </span>
                      <span className="text-[10.5px]">
                        ({(diagnosis?.serverEgress || serverEgress)?.countryName || '해외'}{' '}
                        {(diagnosis?.serverEgress || serverEgress)?.countryCode})
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() =>
                        void handleCopyText((diagnosis?.serverEgress || serverEgress)!.ipv4!)
                      }
                      className="inline-flex items-center gap-1 px-2.5 py-1 rounded-[4px] bg-white border border-[#E5B856] hover:bg-[#FFF3CC] text-[11px] font-semibold text-[#7A4400] transition-colors shrink-0"
                    >
                      {copiedValue === (diagnosis?.serverEgress || serverEgress)?.ipv4 ? (
                        <>
                          <Check className="w-3.5 h-3.5 text-[#107C10]" />
                          IP 복사 완료
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          접속 IP 복사
                        </>
                      )}
                    </button>
                  </div>
                )}

              {/* Detailed Interactive Resolution Steps */}
              {diagnosis?.resolutionSteps && diagnosis.resolutionSteps.length > 0 ? (
                <div className="space-y-2">
                  <div className="text-[11px] font-semibold text-[#424242]">
                    아래 항목을 순서대로 확인·체크한 뒤 바로 다시 연결하세요:
                  </div>
                  <div className="space-y-1.5">
                    {diagnosis.resolutionSteps.map((step) => {
                      const isDone = Boolean(checkedSteps[step.id]);
                      return (
                        <div
                          key={step.id}
                          className={`p-2.5 rounded-[5px] border transition-colors ${
                            isDone
                              ? 'bg-[#F1FAF1] border-[#9FD89F]'
                              : 'bg-white border-[#EBD2D5]'
                          }`}
                        >
                          <label className="flex items-start gap-2 cursor-pointer">
                            <input
                              type="checkbox"
                              checked={isDone}
                              onChange={() => toggleStepChecked(step.id)}
                              className="mt-0.5 accent-[#0F6CBD] shrink-0"
                            />
                            <div className="min-w-0 flex-1 space-y-1">
                              <div className="flex items-center gap-1.5 flex-wrap">
                                <span
                                  className={`px-1.5 py-0.5 rounded-[3px] text-[9.5px] font-semibold ${
                                    step.priority === 'critical'
                                      ? 'bg-[#FDF3F4] text-[#C50F1F] border border-[#EEACB2]'
                                      : step.priority === 'recommended'
                                        ? 'bg-[#EBF3FC] text-[#0F6CBD] border border-[#B4D6FA]'
                                        : 'bg-[#F5F5F5] text-[#424242] border border-[#D1D1D1]'
                                  }`}
                                >
                                  {step.priority === 'critical'
                                    ? '필수 확인'
                                    : step.priority === 'recommended'
                                      ? '권장'
                                      : '즉시 대안'}
                                </span>
                                <span
                                  className={`text-[11.5px] font-semibold leading-snug ${
                                    isDone ? 'line-through text-[#616161]' : 'text-[#242424]'
                                  }`}
                                >
                                  {step.title}
                                </span>
                              </div>
                              <div className="text-[10.5px] font-mono px-2 py-0.5 rounded bg-[#F5F5F5] text-[#0F6CBD] break-words">
                                경로: {step.menuPath}
                              </div>
                              <p className="text-[11px] text-[#424242] leading-relaxed">
                                {step.detail}
                              </p>
                              {(step.copyValue || step.helpUrl) && (
                                <div className="flex items-center gap-2 flex-wrap pt-0.5">
                                  {step.copyValue && (
                                    <button
                                      type="button"
                                      onClick={() => void handleCopyText(step.copyValue!)}
                                      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] bg-[#EBF3FC] hover:bg-[#D6E8FA] text-[#0F6CBD] border border-[#B4D6FA] text-[10.5px] font-semibold transition-colors"
                                    >
                                      {copiedValue === step.copyValue ? (
                                        <>
                                          <Check className="w-3 h-3 text-[#107C10]" />
                                          복사됨
                                        </>
                                      ) : (
                                        <>
                                          <Copy className="w-3 h-3" />
                                          {step.copyLabel || '값 복사'}
                                        </>
                                      )}
                                    </button>
                                  )}
                                  {step.helpUrl && (
                                    <a
                                      href={step.helpUrl}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] bg-white hover:bg-[#F5F5F5] text-[#424242] border border-[#D1D1D1] text-[10.5px] font-medium transition-colors"
                                    >
                                      <ExternalLink className="w-3 h-3" />
                                      {step.helpLabel || '공식 가이드 열기'}
                                    </a>
                                  )}
                                </div>
                              )}
                            </div>
                          </label>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ) : (
                diagnosis &&
                diagnosis.checklist.length > 0 && (
                  <ul className="space-y-1.5 pl-6 text-[11px] text-[#424242] leading-relaxed list-none">
                    {diagnosis.checklist.map((step, i) => (
                      <li key={i}>{step}</li>
                    ))}
                  </ul>
                )
              )}

              {/* One-Click Auto-Heal Retry & Immediate PST/EML Escape Hatch */}
              <div className="flex items-center gap-2 flex-wrap pt-1">
                <button
                  type="button"
                  disabled={isConnecting}
                  onClick={() =>
                    void runConnect(
                      discovered?.recommendedProtocol || diagnosis?.suggestProtocolSwitch || protocol
                    )
                  }
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[11px] font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:opacity-60 rounded-[4px] transition-colors"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isConnecting ? 'animate-spin' : ''}`} />
                  지금 자동 해결로 다시 연결 (POP3·IMAP·공백보정 전체 시도)
                </button>
                {onOpenPstFallback && (
                  <button
                    type="button"
                    onClick={() => {
                      onClose();
                      onOpenPstFallback();
                    }}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-[11px] font-semibold text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[4px] transition-colors"
                  >
                    <Database className="w-3.5 h-3.5 text-[#0F6CBD]" />
                    Outlook .pst / .eml 파일로 바로 시작
                  </button>
                )}
              </div>
              <p className="text-[11px] text-[#424242] bg-white/80 border border-[#EEACB2]/60 rounded px-2.5 py-1.5 leading-relaxed">
                <strong>Windows 로컬 Outlook 안내:</strong> 회사 관리자 정책으로 외부 POP3가 <code className="font-mono">#5.7.5</code> 차단되어 있어도, 로컬 Outlook PC 자동 연동으로 수신함·서명(명함)·전체 설정 및 HTML 발송(COM)이 별도 앱 비밀번호 없이 바로 동작합니다.
              </p>
            </div>
          )}

          {/* Security assurance */}
          <div className="flex items-start gap-2 p-2.5 rounded-[6px] bg-[var(--fluent-bg-subtle)] border border-[var(--fluent-stroke-subtle)] text-[11px] text-[var(--fluent-fg-3)]">
            <ShieldCheck className="w-4 h-4 text-[var(--fluent-brand)] shrink-0 mt-0.5" />
            <span>
              비밀번호는 <code className="font-mono">config.json</code>이나 브라우저 저장소에 저장되지 않으며, 서버 메모리에만 임시 유지 후 연결 해제 시 즉시 삭제됩니다.
            </span>
          </div>

          {/* Actions */}
          <div className="flex items-center justify-end gap-2 pt-2 border-t border-[var(--fluent-stroke-subtle)]">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-2 text-xs font-medium text-[var(--fluent-fg-2)] bg-white border border-[var(--fluent-stroke-1)] hover:bg-[var(--fluent-bg-canvas)] rounded-[4px] transition-colors"
            >
              취소
            </button>
            <button
              type="submit"
              disabled={isConnecting || !email.trim() || !password}
              className="px-4 py-2 text-xs font-semibold text-white bg-[var(--fluent-brand)] hover:bg-[var(--fluent-brand-hover)] disabled:bg-slate-300 rounded-[4px] inline-flex items-center gap-1.5 transition-colors"
            >
              {isConnecting ? (
                <>
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  {protocol.toUpperCase()} 연결 및 메일 로드 중...
                </>
              ) : (
                <>
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  {protocol.toUpperCase()} 자동 설정으로 메일 불러오기
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
