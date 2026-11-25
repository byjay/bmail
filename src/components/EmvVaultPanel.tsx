import React, { useCallback, useEffect, useState } from 'react';
import {
  CheckCircle2,
  Eye,
  EyeOff,
  KeyRound,
  Lock,
  LogOut,
  Plus,
  Search,
  Shield,
  ShieldAlert,
  Trash2,
  Unlock,
  UserCheck,
  Users,
} from 'lucide-react';

type EmvScope = 'global' | 'employee';

interface EmvKeyDefinitionView {
  key: string;
  label: string;
  category: 'ai-primary' | 'ai-rag' | 'ai-fallback' | 'oauth';
  sensitive: boolean;
  description: string;
  placeholder: string;
}

interface PublicEmvEntryView {
  key: string;
  scope: EmvScope;
  employeeId?: string;
  maskedPreview: string;
  fingerprint: string;
  updatedAt: string;
  updatedBy: string;
}

interface EmvEmployeeBucket {
  employeeId: string;
  entries: PublicEmvEntryView[];
}

interface EmvAuthStatusView {
  envAdminPasswordConfigured: boolean;
  adminClaimed: boolean;
  requiresInitialClaim: boolean;
  authenticated: boolean;
  persistenceMode: 'encrypted-disk' | 'encrypted-memory';
  sessionExpiresAt: number | null;
}

interface EmvResolutionSourceItem {
  key: string;
  configured: boolean;
  source: 'employee-emv' | 'global-emv' | 'server-env' | 'unconfigured';
  maskedPreview: string;
}

interface EmvVaultPanelProps {
  activeEmployeeEmail?: string | null;
  pop3SessionToken?: string | null;
  onVaultChanged?: () => void;
}

const FLUENT_INPUT_CLASS =
  'w-full px-3 py-2 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] placeholder:text-[#707070] transition-colors duration-150 focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD] disabled:bg-[#F5F5F5] disabled:text-[#A19F9D]';

const FLUENT_SELECT_CLASS =
  'w-full px-3 py-2 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] transition-colors duration-150 focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD] disabled:bg-[#F5F5F5] disabled:text-[#A19F9D]';

export const EmvVaultPanel: React.FC<EmvVaultPanelProps> = ({
  activeEmployeeEmail,
  pop3SessionToken,
  onVaultChanged,
}) => {
  // Admin Bearer token is strictly held in React memory (never persisted to localStorage/sessionStorage)
  const [adminToken, setAdminToken] = useState<string | null>(null);
  const [authStatus, setAuthStatus] = useState<EmvAuthStatusView | null>(null);
  const [catalog, setCatalog] = useState<EmvKeyDefinitionView[]>([]);
  const [globalEntries, setGlobalEntries] = useState<PublicEmvEntryView[]>([]);
  const [employees, setEmployees] = useState<EmvEmployeeBucket[]>([]);
  const [verifiedEmployeeId, setVerifiedEmployeeId] = useState<string | null>(null);

  // Admin auth forms
  const [passwordInput, setPasswordInput] = useState('');
  const [confirmPasswordInput, setConfirmPasswordInput] = useState('');
  const [authLoading, setAuthLoading] = useState(false);

  // EMV Entry Editor
  const [activeScope, setActiveScope] = useState<EmvScope>('global');
  const [employeeIdInput, setEmployeeIdInput] = useState(activeEmployeeEmail || '');
  const [selectedKey, setSelectedKey] = useState('GEMINI_API_KEY');
  const [keyValueInput, setKeyValueInput] = useState('');
  const [showKeyValue, setShowKeyValue] = useState(false);
  const [isMutating, setIsMutating] = useState(false);

  // Server-side Resolution Preview
  const [previewEmployeeInput, setPreviewEmployeeInput] = useState(activeEmployeeEmail || '');
  const [previewItems, setPreviewItems] = useState<EmvResolutionSourceItem[] | null>(null);
  const [previewTargetId, setPreviewTargetId] = useState<string | null>(null);
  const [isPreviewLoading, setIsPreviewLoading] = useState(false);

  const [notice, setNotice] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  useEffect(() => {
    if (activeEmployeeEmail && !employeeIdInput) {
      setEmployeeIdInput(activeEmployeeEmail);
    }
    if (activeEmployeeEmail && !previewEmployeeInput) {
      setPreviewEmployeeInput(activeEmployeeEmail);
    }
  }, [activeEmployeeEmail, employeeIdInput, previewEmployeeInput]);

  const fetchStatus = useCallback(
    async (tokenOverride?: string | null) => {
      const effectiveToken = tokenOverride !== undefined ? tokenOverride : adminToken;
      const headers: Record<string, string> = {};
      if (effectiveToken) {
        headers.Authorization = `Bearer ${effectiveToken}`;
      } else if (pop3SessionToken) {
        headers.Authorization = `Bearer ${pop3SessionToken}`;
      }
      try {
        const res = await fetch('/api/emv/status', { headers });
        if (!res.ok) return;
        const data = (await res.json()) as {
          auth: EmvAuthStatusView;
          catalog: EmvKeyDefinitionView[];
          verifiedEmployeeId: string | null;
          globalEntries: PublicEmvEntryView[];
          employees: EmvEmployeeBucket[];
        };
        setAuthStatus(data.auth);
        if (Array.isArray(data.catalog)) setCatalog(data.catalog);
        setVerifiedEmployeeId(data.verifiedEmployeeId);
        if (data.auth.authenticated) {
          setGlobalEntries(data.globalEntries || []);
          setEmployees(data.employees || []);
        } else {
          setGlobalEntries([]);
          setEmployees([]);
          if (effectiveToken) setAdminToken(null);
        }
      } catch {
        // ignore transient network errors
      }
    },
    [adminToken, pop3SessionToken]
  );

  useEffect(() => {
    void fetchStatus();
  }, [fetchStatus]);

  const handleAdminClaim = async (e: React.FormEvent) => {
    e.preventDefault();
    setNotice(null);
    if (passwordInput.trim().length < 8) {
      setNotice({ type: 'error', text: '종합 관리자 비밀번호는 8자 이상으로 설정해 주세요.' });
      return;
    }
    if (passwordInput !== confirmPasswordInput) {
      setNotice({ type: 'error', text: '비밀번호 확인이 일치하지 않습니다.' });
      return;
    }
    setAuthLoading(true);
    try {
      const res = await fetch('/api/emv/admin/claim', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password: passwordInput }),
      });
      const data = (await res.json()) as {
        token?: string;
        auth?: EmvAuthStatusView;
        globalEntries?: PublicEmvEntryView[];
        employees?: EmvEmployeeBucket[];
        error?: string;
      };
      if (!res.ok || !data.token) {
        setNotice({ type: 'error', text: data.error || '초기 관리자 등록에 실패했습니다.' });
        return;
      }
      setAdminToken(data.token);
      if (data.auth) setAuthStatus(data.auth);
      setGlobalEntries(data.globalEntries || []);
      setEmployees(data.employees || []);
      setPasswordInput('');
      setConfirmPasswordInput('');
      setNotice({
        type: 'success',
        text: '종합 관리자 비밀번호가 안전하게 등록되었고 EMV 볼트가 잠금 해제되었습니다.',
      });
    } finally {
      setAuthLoading(false);
    }
  };

  const handleAdminLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setNotice(null);
    if (!passwordInput.trim()) {
      setNotice({ type: 'error', text: '종합 관리자 비밀번호를 입력해 주세요.' });
      return;
    }
    setAuthLoading(true);
    try {
      const res = await fetch('/api/emv/admin/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password: passwordInput }),
      });
      const data = (await res.json()) as {
        token?: string;
        auth?: EmvAuthStatusView;
        globalEntries?: PublicEmvEntryView[];
        employees?: EmvEmployeeBucket[];
        error?: string;
      };
      if (!res.ok || !data.token) {
        setNotice({ type: 'error', text: data.error || '관리자 인증에 실패했습니다.' });
        return;
      }
      setAdminToken(data.token);
      if (data.auth) setAuthStatus(data.auth);
      setGlobalEntries(data.globalEntries || []);
      setEmployees(data.employees || []);
      setPasswordInput('');
      setNotice({ type: 'success', text: '종합 관리자 인증 완료 · EMV 키 관리가 활성화되었습니다.' });
    } finally {
      setAuthLoading(false);
    }
  };

  const handleAdminLogout = async () => {
    if (adminToken) {
      await fetch('/api/emv/admin/logout', {
        method: 'POST',
        headers: { Authorization: `Bearer ${adminToken}` },
      }).catch(() => {});
    }
    setAdminToken(null);
    setGlobalEntries([]);
    setEmployees([]);
    setPreviewItems(null);
    setNotice({ type: 'success', text: 'EMV 관리 세션이 잠금 처리되었습니다.' });
    void fetchStatus(null);
  };

  const handleSaveEntry = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!adminToken) return;
    setNotice(null);
    if (activeScope === 'employee' && !employeeIdInput.trim()) {
      setNotice({ type: 'error', text: '개별 사원 이메일 또는 사원 ID를 입력해 주세요.' });
      return;
    }
    if (!keyValueInput.trim()) {
      setNotice({ type: 'error', text: '저장할 API 키 또는 환경 변수 값을 입력해 주세요.' });
      return;
    }

    setIsMutating(true);
    try {
      const res = await fetch('/api/emv/entries', {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${adminToken}`,
        },
        body: JSON.stringify({
          scope: activeScope,
          employeeId: activeScope === 'employee' ? employeeIdInput.trim() : undefined,
          key: selectedKey,
          value: keyValueInput.trim(),
        }),
      });
      const data = (await res.json()) as {
        entry?: PublicEmvEntryView;
        globalEntries?: PublicEmvEntryView[];
        employees?: EmvEmployeeBucket[];
        error?: string;
      };
      if (!res.ok) {
        setNotice({ type: 'error', text: data.error || 'EMV 키 저장에 실패했습니다.' });
        return;
      }
      setGlobalEntries(data.globalEntries || []);
      setEmployees(data.employees || []);
      setKeyValueInput('');
      setShowKeyValue(false);
      onVaultChanged?.();
      setNotice({
        type: 'success',
        text:
          activeScope === 'global'
            ? `전사 공통 EMV [${selectedKey}] 암호화 저장 완료 (미리보기: ${data.entry?.maskedPreview || '••••'})`
            : `사원 [${data.entry?.employeeId || employeeIdInput.trim()}] 개별 EMV [${selectedKey}] 암호화 저장 완료`,
      });
    } finally {
      setIsMutating(false);
    }
  };

  const handleDeleteEntry = async (scope: EmvScope, key: string, employeeId?: string) => {
    if (!adminToken) return;
    setNotice(null);
    setIsMutating(true);
    try {
      const res = await fetch('/api/emv/entries', {
        method: 'DELETE',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${adminToken}`,
        },
        body: JSON.stringify({ scope, key, employeeId }),
      });
      const data = (await res.json()) as {
        globalEntries?: PublicEmvEntryView[];
        employees?: EmvEmployeeBucket[];
        error?: string;
      };
      if (!res.ok) {
        setNotice({ type: 'error', text: data.error || 'EMV 키 삭제에 실패했습니다.' });
        return;
      }
      setGlobalEntries(data.globalEntries || []);
      setEmployees(data.employees || []);
      onVaultChanged?.();
      setNotice({
        type: 'success',
        text:
          scope === 'global'
            ? `전사 공통 EMV [${key}] 항목을 삭제했습니다.`
            : `사원 [${employeeId}] 개별 EMV [${key}] 항목을 삭제했습니다.`,
      });
    } finally {
      setIsMutating(false);
    }
  };

  const handleRunResolutionPreview = async (targetEmp?: string) => {
    if (!adminToken) return;
    const emp = (targetEmp ?? previewEmployeeInput).trim();
    setIsPreviewLoading(true);
    try {
      const res = await fetch('/api/emv/resolve-preview', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${adminToken}`,
        },
        body: JSON.stringify({ employeeId: emp || undefined }),
      });
      if (res.ok) {
        const data = (await res.json()) as {
          employeeId: string | null;
          items: EmvResolutionSourceItem[];
        };
        setPreviewTargetId(data.employeeId);
        setPreviewItems(data.items || []);
      }
    } finally {
      setIsPreviewLoading(false);
    }
  };

  const selectedKeyDef = catalog.find((c) => c.key === selectedKey);
  const visibleEntries =
    activeScope === 'global'
      ? globalEntries
      : employees.find((b) => b.employeeId === employeeIdInput.trim().toLowerCase())?.entries || [];

  const sourceBadge = (source: EmvResolutionSourceItem['source']) => {
    switch (source) {
      case 'employee-emv':
        return { label: '개별 사원 EMV 우선', cls: 'bg-[#EBF3FC] border-[#B4D6FA] text-[#0F6CBD]' };
      case 'global-emv':
        return { label: '전사 공통 EMV', cls: 'bg-emerald-50 border-emerald-200 text-emerald-700' };
      case 'server-env':
        return { label: '서버 .env 기본값', cls: 'bg-slate-100 border-slate-200 text-slate-700' };
      default:
        return { label: '미설정', cls: 'bg-white border-slate-200 text-slate-400' };
    }
  };

  return (
    <section
      aria-label="EMV 환경 변수 및 사원별 API 키 관리"
      className="bg-white border border-[#E0E0E0] rounded-[8px] p-6 space-y-5 shadow-[0_1px_2px_rgba(0,0,0,0.04)]"
    >
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-4 border-b border-[#EDEBE9]">
        <div className="space-y-1">
          <div className="flex items-center gap-2 flex-wrap">
            <KeyRound className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
            <h2 className="text-sm font-bold text-[#242424]">
              EMV (Environment Variable) 중앙 · 개별 사원별 API 키 관리 볼트
            </h2>
            {authStatus && (
              <span
                className={`text-[10px] font-semibold px-2 py-0.5 rounded-[4px] border ${
                  authStatus.persistenceMode === 'encrypted-disk'
                    ? 'bg-emerald-50 border-emerald-200 text-emerald-700'
                    : 'bg-amber-50 border-amber-200 text-amber-800'
                }`}
              >
                {authStatus.persistenceMode === 'encrypted-disk'
                  ? 'AES-256-GCM 암호화 영구 저장'
                  : 'AES-256-GCM 메모리 보안 모드'}
              </span>
            )}
          </div>
          <p className="text-xs text-[#616161] leading-relaxed">
            종합 관리자가 <strong>전사 공통 API 키</strong>와 <strong>개별 사원별 API 키</strong>를 서버 암호화 볼트에 분리 저장합니다. 실제 AI 호출 시 서버가 인증된 사원 세션(Microsoft OAuth · POP3 로그인)을 식별해 <code className="font-mono text-[#242424]">개별 사원 EMV → 전사 공통 EMV → 서버 .env</code> 순으로 자동 적용합니다.
          </p>
        </div>

        {authStatus?.authenticated && (
          <button
            type="button"
            onClick={() => void handleAdminLogout()}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#FAF9F8] rounded-[4px] shrink-0 transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
          >
            <LogOut className="w-3.5 h-3.5" strokeWidth={1.75} />
            관리자 잠금
          </button>
        )}
      </div>

      {/* Persistence Mode Hint */}
      {authStatus && authStatus.persistenceMode === 'encrypted-memory' && (
        <div className="flex items-start gap-2 px-3.5 py-2.5 rounded-lg bg-amber-50/70 border border-amber-200 text-[11px] text-amber-900">
          <ShieldAlert className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
          <span>
            현재 평문 저장 방지를 위해 <strong>서버 메모리 암호화 모드</strong>로 동작 중입니다. 서버 재시작 후에도 암호화 파일(<code className="font-mono">data/emv-vault.enc.json</code>)로 영구 보관하려면 서버 <code className="font-mono">.env</code>에 <code className="font-mono font-semibold">EMV_MASTER_KEY</code>(8자 이상) 또는 <code className="font-mono font-semibold">CONFIG_ADMIN_PASSWORD</code>를 설정하세요.
          </span>
        </div>
      )}

      {notice && (
        <div
          role="status"
          className={`flex items-center gap-2 px-3.5 py-2.5 rounded-lg border text-xs ${
            notice.type === 'success'
              ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
              : 'bg-red-50 border-red-200 text-red-800'
          }`}
        >
          <CheckCircle2 className="w-4 h-4 shrink-0" />
          <span>{notice.text}</span>
        </div>
      )}

      {/* Auth Gate: First-Run Claim OR Admin Sign-In */}
      {!authStatus?.authenticated ? (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 items-start">
          {authStatus?.requiresInitialClaim ? (
            <form
              onSubmit={(e) => void handleAdminClaim(e)}
              className="p-4 rounded-[8px] border border-[#CFE4FA] bg-[#F5F9FE] space-y-3"
            >
              <div className="flex items-center gap-2 text-xs font-bold text-[#0F6CBD]">
                <Lock className="w-4 h-4" strokeWidth={1.75} />
                최초 1회 · 종합 관리자 EMV 비밀번호 설정
              </div>
              <p className="text-[11px] text-[#616161] leading-relaxed">
                서버 <code className="font-mono text-[#242424]">.env</code>에 <code className="font-mono text-[#242424]">CONFIG_ADMIN_PASSWORD</code>가 아직 설정되지 않았습니다. 최초 1회 종합 관리자 비밀번호를 등록하면 서버에서 단방향 해시로 잠기며 즉시 EMV 키 관리가 열립니다.
              </p>
              <div className="space-y-2.5">
                <div className="space-y-1">
                  <label htmlFor="emv-claim-pw" className="block text-xs font-semibold text-[#242424]">
                    종합 관리자 비밀번호 (8자 이상)
                  </label>
                  <input
                    id="emv-claim-pw"
                    type="password"
                    value={passwordInput}
                    onChange={(e) => setPasswordInput(e.target.value)}
                    placeholder="예: 영문·숫자·기호 조합 8자 이상"
                    autoComplete="new-password"
                    className={FLUENT_INPUT_CLASS}
                  />
                </div>
                <div className="space-y-1">
                  <label htmlFor="emv-claim-confirm" className="block text-xs font-semibold text-[#242424]">
                    비밀번호 확인
                  </label>
                  <input
                    id="emv-claim-confirm"
                    type="password"
                    value={confirmPasswordInput}
                    onChange={(e) => setConfirmPasswordInput(e.target.value)}
                    placeholder="비밀번호 다시 입력"
                    autoComplete="new-password"
                    className={FLUENT_INPUT_CLASS}
                  />
                </div>
              </div>
              <button
                type="submit"
                disabled={authLoading}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-[#C8C6C4] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
              >
                <Unlock className="w-3.5 h-3.5" strokeWidth={1.75} />
                {authLoading ? '등록 중...' : '종합 관리자 등록 및 볼트 열기'}
              </button>
            </form>
          ) : (
            <form
              onSubmit={(e) => void handleAdminLogin(e)}
              className="p-4 rounded-[8px] border border-[#E0E0E0] bg-[#FAF9F8] space-y-3"
            >
              <div className="flex items-center gap-2 text-xs font-bold text-[#242424]">
                <Lock className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
                종합 관리자 인증 (EMV 볼트 잠금 해제)
              </div>
              <p className="text-[11px] text-[#616161] leading-relaxed">
                {authStatus?.envAdminPasswordConfigured
                  ? '서버 .env의 CONFIG_ADMIN_PASSWORD를 입력하면 공통 및 사원별 API 키를 안전하게 등록·교체할 수 있습니다.'
                  : '등록된 종합 관리자 비밀번호를 입력해 EMV 키 관리 콘솔을 잠금 해제하세요.'}
              </p>
              <div className="space-y-1">
                <label htmlFor="emv-login-pw" className="block text-xs font-semibold text-[#242424]">
                  종합 관리자 비밀번호
                </label>
                <input
                  id="emv-login-pw"
                  type="password"
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  placeholder="종합 관리자 비밀번호 입력"
                  autoComplete="current-password"
                  className={FLUENT_INPUT_CLASS}
                />
              </div>
              <button
                type="submit"
                disabled={authLoading}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-[#C8C6C4] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
              >
                <Unlock className="w-3.5 h-3.5" strokeWidth={1.75} />
                {authLoading ? '인증 중...' : 'EMV 관리 콘솔 열기'}
              </button>
            </form>
          )}

          {/* Server-Side Identity Safety Explanation */}
          <div className="p-4 rounded-[8px] border border-[#E0E0E0] bg-white space-y-2.5 text-xs">
            <div className="flex items-center gap-1.5 font-bold text-[#242424]">
              <Shield className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
              서버 측 사원별 키 식별 및 격리 원칙
            </div>
            <ul className="space-y-1.5 text-[11px] text-[#616161] list-disc pl-4 leading-relaxed">
              <li>
                저장된 키 원문은 브라우저로 다시 반환되지 않으며, 화면에는 마스킹 미리보기(<code className="font-mono text-[#242424]">AIza••••9aBc</code>)와 지문만 표시됩니다.
              </li>
              <li>
                사원별 키는 서버에서 <strong>인증된 세션</strong>(Microsoft OAuth 또는 POP3 로그인)의 이메일을 기준으로 HMAC 파티션 키로 격리 조회되며, 임의 사원 ID 위조 요청은 차단됩니다.
              </li>
              <li>
                현재 서버 식별 세션:{' '}
                <strong className="font-mono text-[#242424]">
                  {verifiedEmployeeId || '미인증 (전사 공통 EMV/.env 적용)'}
                </strong>
              </li>
            </ul>
          </div>
        </div>
      ) : (
        /* Authenticated Super-Admin EMV Management Console */
        <div className="space-y-5">
          {/* Scope Selector Tabs */}
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="inline-flex p-1 rounded-lg bg-slate-100 border border-slate-200">
              <button
                type="button"
                onClick={() => setActiveScope('global')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md transition-colors ${
                  activeScope === 'global'
                    ? 'bg-white text-[#0F6CBD] shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Users className="w-3.5 h-3.5" />
                전사 공통 EMV ({globalEntries.length})
              </button>
              <button
                type="button"
                onClick={() => setActiveScope('employee')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md transition-colors ${
                  activeScope === 'employee'
                    ? 'bg-white text-[#0F6CBD] shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <UserCheck className="w-3.5 h-3.5" />
                개별 사원별 EMV ({employees.length}명)
              </button>
            </div>

            <div className="text-[11px] text-slate-500">
              적용 우선순위: <strong className="text-[#0F6CBD]">1. 개별 사원 EMV</strong> &gt;{' '}
              <strong className="text-emerald-700">2. 전사 공통 EMV</strong> &gt; 3. 서버 .env
            </div>
          </div>

          {/* Per-Employee Quick Directory */}
          {activeScope === 'employee' && (
            <div className="p-4 rounded-[8px] bg-[#FAF9F8] border border-[#E0E0E0] space-y-2.5">
              <div className="flex flex-col sm:flex-row sm:items-center gap-2.5">
                <label htmlFor="emv-emp-id" className="text-xs font-semibold text-[#242424] shrink-0">
                  대상 사원 이메일 / ID
                </label>
                <input
                  id="emv-emp-id"
                  type="text"
                  value={employeeIdInput}
                  onChange={(e) => setEmployeeIdInput(e.target.value)}
                  placeholder="예: hong@company.co.kr 또는 emp_01"
                  className={`flex-1 ${FLUENT_INPUT_CLASS} font-mono`}
                />
                {activeEmployeeEmail && (
                  <button
                    type="button"
                    onClick={() => setEmployeeIdInput(activeEmployeeEmail)}
                    className="px-3 py-2 text-[11px] font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[4px] whitespace-nowrap transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                  >
                    현재 접속 사원 선택
                  </button>
                )}
              </div>

              {employees.length > 0 && (
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="text-[11px] text-[#616161] mr-1">등록된 사원:</span>
                  {employees.map((bucket) => {
                    const selected = bucket.employeeId === employeeIdInput.trim().toLowerCase();
                    return (
                      <button
                        key={bucket.employeeId}
                        type="button"
                        onClick={() => setEmployeeIdInput(bucket.employeeId)}
                        className={`px-2.5 py-1 text-[11px] font-mono rounded-[4px] border transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] ${
                          selected
                            ? 'bg-[#0F6CBD] text-white border-[#0F6CBD]'
                            : 'bg-white text-[#242424] border-[#D1D1D1] hover:bg-[#F5F5F5]'
                        }`}
                      >
                        {bucket.employeeId} ({bucket.entries.length})
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          )}

          {/* Add / Update EMV Key Form */}
          <form
            onSubmit={(e) => void handleSaveEntry(e)}
            className="p-4 rounded-[8px] border border-[#E0E0E0] bg-[#FAF9F8] space-y-3.5"
          >
            <div className="text-xs font-bold text-[#242424]">
              {activeScope === 'global'
                ? '전사 공통 EMV 키 등록 · 갱신 (모든 사원 설치본에 자동 반영)'
                : `개별 사원 전용 EMV 키 등록 · 갱신 (${employeeIdInput.trim() || '사원 ID 입력 필요'})`}
            </div>

            <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end">
              <div className="md:col-span-4 space-y-1">
                <label htmlFor="emv-key-select" className="block text-xs font-semibold text-[#242424]">
                  환경 변수 (EMV) 키
                </label>
                <select
                  id="emv-key-select"
                  value={selectedKey}
                  onChange={(e) => setSelectedKey(e.target.value)}
                  className={`${FLUENT_SELECT_CLASS} font-mono`}
                >
                  {catalog.map((item) => (
                    <option key={item.key} value={item.key}>
                      {item.key} — {item.label}
                    </option>
                  ))}
                </select>
              </div>

              <div className="md:col-span-6 space-y-1">
                <label htmlFor="emv-key-value" className="block text-xs font-semibold text-[#242424]">
                  {selectedKeyDef?.label || selectedKey} 값
                </label>
                <div className="relative">
                  <input
                    id="emv-key-value"
                    type={showKeyValue || !selectedKeyDef?.sensitive ? 'text' : 'password'}
                    value={keyValueInput}
                    onChange={(e) => setKeyValueInput(e.target.value)}
                    placeholder={selectedKeyDef?.placeholder || '값 입력'}
                    autoComplete="off"
                    className={`${FLUENT_INPUT_CLASS} pr-9 font-mono`}
                  />
                  {selectedKeyDef?.sensitive && (
                    <button
                      type="button"
                      onClick={() => setShowKeyValue((v) => !v)}
                      aria-label={showKeyValue ? '값 숨기기' : '값 표시'}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[#616161] hover:text-[#242424] rounded-[4px] p-0.5 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                    >
                      {showKeyValue ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                    </button>
                  )}
                </div>
              </div>

              <div className="md:col-span-2">
                <button
                  type="submit"
                  disabled={isMutating}
                  className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-[#C8C6C4] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
                >
                  <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
                  {isMutating ? '저장 중' : '암호화 저장'}
                </button>
              </div>
            </div>

            {selectedKeyDef && (
              <p className="text-[11px] text-[#616161]">{selectedKeyDef.description}</p>
            )}
          </form>

          {/* Stored Entries Table (Masked Previews Only) */}
          <div className="border border-slate-200 rounded-xl overflow-hidden">
            <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-xs">
              <span className="font-bold text-slate-900">
                {activeScope === 'global'
                  ? `저장된 전사 공통 EMV 목록 (${visibleEntries.length}건)`
                  : `사원 [${employeeIdInput.trim() || '-'}] 개별 EMV 목록 (${visibleEntries.length}건)`}
              </span>
              <span className="text-[11px] text-slate-500">원문 미노출 · 마스킹 표시</span>
            </div>

            {visibleEntries.length === 0 ? (
              <div className="px-4 py-6 text-center text-xs text-slate-500">
                {activeScope === 'global'
                  ? '아직 EMV 볼트에 저장된 전사 공통 키가 없습니다. (현재 서버 .env 기본값 사용 중)'
                  : '이 사원에 등록된 개별 오버라이드 키가 없습니다. (전사 공통 EMV → 서버 .env 자동 상속)'}
              </div>
            ) : (
              <div className="divide-y divide-slate-100">
                {visibleEntries.map((entry) => (
                  <div
                    key={`${entry.scope}-${entry.employeeId || 'global'}-${entry.key}`}
                    className="px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 text-xs bg-white"
                  >
                    <div className="min-w-0 space-y-0.5">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-mono font-bold text-slate-900">{entry.key}</span>
                        <span className="font-mono text-[11px] px-2 py-0.5 rounded bg-slate-100 text-slate-700">
                          {entry.maskedPreview}
                        </span>
                        <span className="font-mono text-[10px] text-slate-400">
                          SHA256:{entry.fingerprint}
                        </span>
                      </div>
                      <div className="text-[11px] text-slate-500">
                        수정: {new Date(entry.updatedAt).toLocaleString('ko-KR')} · 등록자: {entry.updatedBy}
                      </div>
                    </div>

                    <button
                      type="button"
                      disabled={isMutating}
                      onClick={() => void handleDeleteEntry(entry.scope, entry.key, entry.employeeId)}
                      className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium text-red-700 bg-red-50 hover:bg-red-100 border border-red-200 rounded-lg"
                    >
                      <Trash2 className="w-3 h-3" />
                      삭제
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Server-Side Per-Employee Key Resolution Inspector */}
          <div className="p-4 rounded-[8px] border border-[#E0E0E0] bg-white space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="text-xs font-bold text-[#242424] flex items-center gap-1.5">
                <Search className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
                사원별 최종 적용 키 우선순위 시뮬레이터 (서버 식별 기준)
              </div>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  value={previewEmployeeInput}
                  onChange={(e) => setPreviewEmployeeInput(e.target.value)}
                  placeholder="사원 이메일 입력 (비우면 공통 기준)"
                  aria-label="적용 결과 확인할 사원 이메일"
                  className={`${FLUENT_INPUT_CLASS} font-mono sm:w-60`}
                />
                <button
                  type="button"
                  disabled={isPreviewLoading}
                  onClick={() => void handleRunResolutionPreview()}
                  className="px-3.5 py-2 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[4px] whitespace-nowrap transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                >
                  {isPreviewLoading ? '조회 중...' : '적용 결과 확인'}
                </button>
              </div>
            </div>

            {previewItems && (
              <div className="space-y-2">
                <div className="text-[11px] text-slate-600">
                  조회 대상:{' '}
                  <strong className="font-mono text-slate-900">
                    {previewTargetId || '전사 공통 (미지정 사원)'}
                  </strong>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
                  {previewItems.map((item) => {
                    const badge = sourceBadge(item.source);
                    return (
                      <div
                        key={item.key}
                        className="p-2.5 rounded-lg border border-slate-200 bg-[#FAF9F8] space-y-1 text-[11px]"
                      >
                        <div className="flex items-center justify-between gap-1.5">
                          <span className="font-mono font-bold text-slate-900 truncate">{item.key}</span>
                          <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded border shrink-0 ${badge.cls}`}>
                            {badge.label}
                          </span>
                        </div>
                        <div className="font-mono text-[10px] text-slate-500 truncate">
                          {item.configured ? item.maskedPreview : '미설정'}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </section>
  );
};
