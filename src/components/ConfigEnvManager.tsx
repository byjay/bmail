import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Activity,
  AlertCircle,
  AlertTriangle,
  Check,
  CheckCircle2,
  CloudUpload,
  Download,
  FileCode2,
  FolderSearch,
  HardDrive,
  KeyRound,
  Plug,
  RotateCcw,
  Save,
  ShieldCheck,
  Sparkles,
  Upload,
  UserCheck,
  Users,
} from 'lucide-react';
import type {
  AiProviderId,
  AppConfig,
  BlogTone,
  DateMode,
  EnvSecretStatus,
  ExternalEngineProbeResult,
  MailItem,
  MailScope,
} from '../types/mail';
import { DEFAULT_APP_CONFIG } from '../utils/appConfig';
import {
  buildUnifiedAddressBook,
  exportContactsToVcf,
  type AccessPolicyScope,
  type OutlookPcProfileSnapshot,
} from '../services/outlookLocalBridge';
import { EmvVaultPanel } from './EmvVaultPanel';
import { AdminLlmUsageCompactCard } from './AdminLlmUsageCompactCard';

interface HybridCreditStatusView {
  realtimeGeminiCredit: {
    label: string;
    role: string;
    model: string;
    configured: boolean;
    activePriority: number;
  };
  agentBuilderRagCredit: {
    label: string;
    role: string;
    configured: boolean;
    mode: 'cloudflare-worker' | 'discovery-engine-direct' | 'unconfigured';
    gcsArchiveBucketConfigured: boolean;
    activePriority: number;
    expiryGuard: {
      expiryDate: string;
      hardStopAfterExpiry: boolean;
      daysRemaining: number;
      isExpired: boolean;
      isNearExpiry: boolean;
      allowAgentBuilderCall: boolean;
      statusMessage: string;
    };
  };
  setupChecklist: Array<{
    step: number;
    title: string;
    creditTarget: string;
    envKeys: string[];
    ready: boolean;
  }>;
}

interface ConfigEnvManagerProps {
  config: AppConfig;
  envSecrets: EnvSecretStatus[];
  oauthCallbackUrl: string;
  configPasswordRequired: boolean;
  emails?: MailItem[];
  activeEmployeeEmail?: string | null;
  pop3SessionToken?: string | null;
  outlookPcProfile?: OutlookPcProfileSnapshot | null;
  onImportPcProfile?: (profile: OutlookPcProfileSnapshot) => void;
  onOpenPop3Modal?: () => void;
  onVaultUpdated?: () => void;
  onSaveConfig: (nextConfig: AppConfig, adminPassword?: string) => Promise<{ ok: boolean; error?: string }>;
}

interface ApiConnectionCheckItem {
  id: string;
  label: string;
  configured: boolean;
  scope: 'org-wide' | 'per-user';
  detail: string;
}

const FLUENT_INPUT_CLASS =
  'w-full px-3 py-2 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] placeholder:text-[#707070] transition-colors duration-150 focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD] disabled:bg-[#F5F5F5] disabled:text-[#A19F9D]';

const FLUENT_SELECT_CLASS =
  'w-full px-3 py-2 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] transition-colors duration-150 focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD] disabled:bg-[#F5F5F5] disabled:text-[#A19F9D]';

const FLUENT_CHECKBOX_CLASS =
  'w-4 h-4 accent-[#0F6CBD] rounded-[3px] focus:outline-none focus:ring-2 focus:ring-[#0F6CBD] focus:ring-offset-1';

export const ConfigEnvManager: React.FC<ConfigEnvManagerProps> = ({
  config,
  envSecrets,
  oauthCallbackUrl,
  configPasswordRequired,
  emails = [],
  activeEmployeeEmail,
  pop3SessionToken,
  outlookPcProfile,
  onImportPcProfile,
  onOpenPop3Modal,
  onVaultUpdated,
  onSaveConfig,
}) => {
  const [draft, setDraft] = useState<AppConfig>(config);
  const [adminPassword, setAdminPassword] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [isProbing, setIsProbing] = useState(false);
  const [hybridStatus, setHybridStatus] = useState<HybridCreditStatusView | null>(null);
  const [maskJsonlSensitive, setMaskJsonlSensitive] = useState(true);
  const [isExportingJsonl, setIsExportingJsonl] = useState(false);
  const [jsonlNotice, setJsonlNotice] = useState<string | null>(null);

  // Org-Wide vs Per-User API Connection Policy & Live Check
  const [policyScope, setPolicyScope] = useState<AccessPolicyScope>('org-wide');
  const [isTestingApis, setIsTestingApis] = useState(false);
  const [apiChecks, setApiChecks] = useState<ApiConnectionCheckItem[] | null>(null);
  const [pcProfileNotice, setPcProfileNotice] = useState<string | null>(null);
  const pcProfileInputRef = useRef<HTMLInputElement | null>(null);

  const unifiedContacts = useMemo(
    () => buildUnifiedAddressBook(emails, outlookPcProfile?.contacts || []),
    [emails, outlookPcProfile]
  );

  const [probeResults, setProbeResults] = useState<{
    pdf: ExternalEngineProbeResult;
    excel: ExternalEngineProbeResult;
  } | null>(null);
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [backupStatus, setBackupStatus] = useState<{
    dbExists: boolean;
    dbSizeBytes: number;
    maxRetention: number;
    snapshots: Array<{ filename: string; createdAt: string; sizeBytes: number; reason: string }>;
    lastBackupAt: string | null;
  } | null>(null);
  const [isRunningBackup, setIsRunningBackup] = useState(false);

  const loadBackupStatus = async () => {
    try {
      const res = await fetch('/api/admin/backups');
      if (res.ok) {
        const data = await res.json();
        setBackupStatus(data);
      }
    } catch {
      // ignore offline
    }
  };

  useEffect(() => {
    void loadBackupStatus();
  }, []);

  const handleRunBackupNow = async () => {
    setIsRunningBackup(true);
    try {
      const res = await fetch('/api/admin/backups/run', { method: 'POST' });
      const data = (await res.json()) as { ok?: boolean; status?: typeof backupStatus };
      if (data.status) {
        setBackupStatus(data.status);
      } else {
        await loadBackupStatus();
      }
    } finally {
      setIsRunningBackup(false);
    }
  };

  useEffect(() => {
    setDraft(config);
  }, [config]);

  const handleTestApiConnections = async () => {
    setIsTestingApis(true);
    try {
      const headers: Record<string, string> = {};
      if (pop3SessionToken) {
        headers.Authorization = `Bearer ${pop3SessionToken}`;
      }
      const res = await fetch('/api/config/test-connections', { method: 'POST', headers });
      if (res.ok) {
        const data = (await res.json()) as { connections?: ApiConnectionCheckItem[] };
        if (Array.isArray(data.connections)) {
          setApiChecks(data.connections);
        }
      }
    } finally {
      setIsTestingApis(false);
    }
  };

  const handleUploadPcProfile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setPcProfileNotice(null);
    try {
      const text = await file.text();
      const parsedJson = JSON.parse(text);
      const res = await fetch('/api/outlook/pc-profile/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(parsedJson),
      });
      const data = (await res.json()) as { profile?: OutlookPcProfileSnapshot; error?: string };
      if (!res.ok || !data.profile) {
        setPcProfileNotice(data.error || 'PC Outlook 프로필 JSON 가져오기에 실패했습니다.');
        return;
      }
      onImportPcProfile?.(data.profile);
      setPcProfileNotice(
        `PC Outlook 프로필 반영 완료: 계정 ${data.profile.accounts.length}개 · PST ${data.profile.pstPaths.length}개 · 주소록 ${data.profile.contacts.length}명`
      );
    } catch {
      setPcProfileNotice('올바른 bmail-outlook-pc-profile.json 파일을 선택해 주세요.');
    } finally {
      if (pcProfileInputRef.current) pcProfileInputRef.current.value = '';
    }
  };

  const handleDownloadContactsVcf = () => {
    if (unifiedContacts.length === 0) {
      setPcProfileNotice('내보낼 주소록 연락처가 없습니다. 먼저 PC Outlook 프로필 또는 PST/POP3 메일을 불러오세요.');
      return;
    }
    const vcf = exportContactsToVcf(unifiedContacts);
    const blob = new Blob([vcf], { type: 'text/vcard;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `bmail-outlook-contacts-${new Date().toISOString().slice(0, 10)}.vcf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setPcProfileNotice(`통합 주소록 ${unifiedContacts.length}명을 Outlook 가져오기용 vCard(.vcf)로 저장했습니다.`);
  };

  useEffect(() => {
    let mounted = true;
    void fetch('/api/hybrid-credits/status')
      .then((r) => (r.ok ? r.json() : null))
      .then((data) => {
        if (mounted && data) {
          setHybridStatus(data as HybridCreditStatusView);
        }
      })
      .catch(() => {
        // ignore transient error
      });
    return () => {
      mounted = false;
    };
  }, []);

  const handleExportAgentBuilderJsonl = async () => {
    if (emails.length === 0) {
      setJsonlNotice('내보낼 실제 메일이 없습니다. 상단 PST · EML 열기 또는 Microsoft 365 동기화를 먼저 실행하세요.');
      return;
    }
    setIsExportingJsonl(true);
    setJsonlNotice(null);
    try {
      const res = await fetch('/api/hybrid-credits/export-jsonl', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ emails, maskSensitive: maskJsonlSensitive }),
      });
      if (!res.ok) {
        const err = (await res.json().catch(() => ({}))) as { error?: string };
        setJsonlNotice(err.error || 'JSONL 내보내기에 실패했습니다.');
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `bmail-agent-builder-archive-${new Date().toISOString().slice(0, 10)}.jsonl`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setJsonlNotice(
        `실제 메일 ${emails.length}건을 Cloud Storage / Vertex AI Agent Builder용 NDJSON(.jsonl)로 내보냈습니다.`
      );
    } catch {
      setJsonlNotice('서버와 통신하지 못해 JSONL 내보내기에 실패했습니다.');
    } finally {
      setIsExportingJsonl(false);
    }
  };

  const handleProbeEngines = async () => {
    setIsProbing(true);
    try {
      const res = await fetch('/api/external-engines/probe', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          pdfEngineDir: draft.externalEngines.pdfEngineDir,
          pdfEngineCommand: draft.externalEngines.pdfEngineCommand,
          excelEngineDir: draft.externalEngines.excelEngineDir,
          excelEngineCommand: draft.externalEngines.excelEngineCommand,
        }),
      });
      if (res.ok) {
        const data = (await res.json()) as {
          pdf: ExternalEngineProbeResult;
          excel: ExternalEngineProbeResult;
        };
        setProbeResults(data);
      }
    } finally {
      setIsProbing(false);
    }
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSaving(true);
    setStatusMessage(null);
    const res = await onSaveConfig(draft, adminPassword || undefined);
    setIsSaving(false);
    if (res.ok) {
      setStatusMessage({ type: 'success', text: 'config.json 파일에 변경 사항을 저장했습니다.' });
    } else {
      setStatusMessage({ type: 'error', text: res.error || 'config.json 저장에 실패했습니다.' });
    }
  };

  const handleResetDefaults = () => {
    setDraft(DEFAULT_APP_CONFIG);
    setStatusMessage(null);
  };

  return (
    <div className="flex-1 overflow-y-auto bg-[#F5F5F5]">
      <div className="max-w-6xl mx-auto px-6 py-7 space-y-6">
        <header className="flex flex-col md:flex-row md:items-end justify-between gap-4 pb-5 border-b border-[#E0E0E0]">
          <div className="space-y-1.5">
            <h1 className="text-xl font-bold text-[#242424] tracking-tight">
              설정 분리 관리 (config.json · .env 시크릿)
            </h1>
            <p className="text-xs text-[#616161] max-w-2xl leading-relaxed">
              일반 앱·플러그인·AI 에이전트 동작 규칙은 <code className="font-mono text-[#242424]">config.json</code>에 저장하고, API 키·OAuth 시크릿·관리 비밀번호는 <code className="font-mono text-[#242424]">.env</code> 파일(Secrets)로 완전히 분리 관리합니다.
            </p>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <a
              href="/api/config/download"
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px] transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
            >
              <Download className="w-3.5 h-3.5" strokeWidth={1.75} />
              config.json 받기
            </a>
            <a
              href="/api/env-template/download"
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#FAF9F8] rounded-[4px] transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
            >
              <KeyRound className="w-3.5 h-3.5" strokeWidth={1.75} />
              .env 템플릿 받기
            </a>
          </div>
        </header>

        {/* Compact LLM Token & Cost Usage Monitor for Admin */}
        <AdminLlmUsageCompactCard />

        {/* Compact Offline Local SQLite Daily Compressed Backup Snapshot Strip */}
        <section
          aria-label="로컬 SQLite 자동 백업 스냅샷"
          className="bg-white border border-[#E0E0E0] rounded-[8px] px-4 py-2.5 flex flex-wrap items-center justify-between gap-2 text-xs shadow-[0_1px_2px_rgba(0,0,0,0.03)]"
        >
          <div className="flex items-center gap-2 min-w-0">
            <HardDrive className="w-4 h-4 text-[#0F6CBD] shrink-0" strokeWidth={1.75} />
            <span className="font-bold text-[#242424]">로컬 SQLite 자동 백업 스냅샷 (일 1회 gzip · 최대 7세대)</span>
            {backupStatus && (
              <span className="text-[11px] font-mono text-[#616161]">
                보관 {backupStatus.snapshots.length}건 ·{' '}
                {backupStatus.lastBackupAt
                  ? `최근 ${backupStatus.lastBackupAt.slice(0, 16).replace('T', ' ')} (${Math.max(1, Math.round((backupStatus.snapshots[0]?.sizeBytes || 0) / 1024))} KB)`
                  : '백업 대기 중'}
              </span>
            )}
          </div>
          <button
            type="button"
            disabled={isRunningBackup}
            onClick={() => void handleRunBackupNow()}
            className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] disabled:opacity-50 rounded-[4px] border border-[#B4D6FA]"
          >
            <HardDrive className="w-3 h-3" />
            {isRunningBackup ? '압축 백업 중...' : '지금 백업 생성'}
          </button>
        </section>

        {/* Section 00: Hybrid GCP Credit Priority Manager (Gemini 13.43만원 + Agent Builder 143만원) */}
        <section className="bg-white border border-[#E0E0E0] rounded-[8px] p-6 space-y-5 shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-4 border-b border-[#EDEBE9]">
            <div className="space-y-1">
              <div className="flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
                <h2 className="text-sm font-bold text-[#242424]">
                  00. GCP 하이브리드 크레딧 우선 소진 관리 (Gemini 13.43만 원 + Agent Builder 143만 원)
                </h2>
              </div>
              <p className="text-xs text-[#616161]">
                모든 프로젝트·메일 분석 실행 시 두 GCP 크레딧을 역할별로 최우선 소진하고, 장애·한도 초과·만료 시 2순위 프로바이더 및 로컬 엔진으로 무중단 폴백합니다.
              </p>
            </div>
            <div className="flex items-center gap-2.5 shrink-0">
              <label className="inline-flex items-center gap-1.5 text-[11px] text-[#424242] cursor-pointer">
                <input
                  type="checkbox"
                  checked={maskJsonlSensitive}
                  onChange={(e) => setMaskJsonlSensitive(e.target.checked)}
                  className={FLUENT_CHECKBOX_CLASS}
                />
                민감정보 마스킹
              </label>
              <button
                type="button"
                disabled={isExportingJsonl}
                onClick={() => void handleExportAgentBuilderJsonl()}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] disabled:opacity-50 rounded-[4px] transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
              >
                <CloudUpload className="w-3.5 h-3.5" strokeWidth={1.75} />
                {isExportingJsonl
                  ? 'JSONL 생성 중...'
                  : `GCS 아카이브 내보내기 (${emails.length}건)`}
              </button>
            </div>
          </div>

          {jsonlNotice && (
            <div className="px-3.5 py-2.5 text-xs rounded-lg border border-[#B4D6FA] bg-[#EBF3FC] text-[#0F6CBD]">
              {jsonlNotice}
            </div>
          )}

          {/* Role Split Matrix */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="p-4 rounded-lg border border-slate-200 bg-[#FAF9F8] space-y-2">
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs font-bold text-slate-900">
                  1순위 · 실시간 메일 요약 · 업무일지 · 블로그
                </span>
                <span
                  className={`text-[11px] font-semibold px-2 py-0.5 rounded border ${
                    hybridStatus?.realtimeGeminiCredit.configured
                      ? 'bg-emerald-50 border-emerald-200 text-emerald-700'
                      : 'bg-white border-slate-200 text-slate-500'
                  }`}
                >
                  {hybridStatus?.realtimeGeminiCredit.configured ? '크레딧 활성' : '키 미설정'}
                </span>
              </div>
              <div className="text-[11px] text-slate-600 leading-relaxed">
                <strong>소진 대상:</strong> Gemini API 13.43만 원 크레딧 (모델:{' '}
                <code className="font-mono text-slate-900">
                  {hybridStatus?.realtimeGeminiCredit.model || 'gemini-4-hydrogen'}
                </code>
                )
              </div>
              <div className="text-[11px] text-slate-500">
                표준 fetch · 20초 타임아웃 · HTTP 429/5xx 시 Groq/Claude/OpenAI 자동 폴백
              </div>
            </div>

            <div className="p-4 rounded-lg border border-slate-200 bg-[#FAF9F8] space-y-2">
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs font-bold text-slate-900">
                  1순위 · 과거 누적 메일 검색 · 맥락 분석 (RAG)
                </span>
                <span
                  className={`text-[11px] font-semibold px-2 py-0.5 rounded border ${
                    hybridStatus?.agentBuilderRagCredit.configured &&
                    hybridStatus?.agentBuilderRagCredit.expiryGuard.allowAgentBuilderCall
                      ? 'bg-[#EBF3FC] border-[#B4D6FA] text-[#0F6CBD]'
                      : 'bg-white border-slate-200 text-slate-500'
                  }`}
                >
                  {hybridStatus?.agentBuilderRagCredit.configured
                    ? hybridStatus.agentBuilderRagCredit.mode === 'cloudflare-worker'
                      ? 'Worker 프록시 활성'
                      : 'Discovery API 활성'
                    : '미연결 (Gemini 폴백)'}
                </span>
              </div>
              <div className="text-[11px] text-slate-600 leading-relaxed">
                <strong>소진 대상:</strong> Vertex AI Agent Builder 143만 원 크레딧 (Cloud Storage + Data Store)
              </div>
              <div className="text-[11px] text-slate-500">
                Ask AI 질의 시 Agent Builder 최우선 호출 → 실패·만료 시 Gemini Primary → 로컬 0-LLM 전환
              </div>
            </div>
          </div>

          {/* Expiry Guard Banner */}
          <div
            className={`flex items-start gap-2.5 p-3.5 rounded-lg border text-xs ${
              hybridStatus?.agentBuilderRagCredit.expiryGuard.isExpired ||
              hybridStatus?.agentBuilderRagCredit.expiryGuard.isNearExpiry
                ? 'bg-amber-50 border-amber-200 text-amber-900'
                : 'bg-slate-50 border-slate-200 text-slate-700'
            }`}
          >
            <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
            <div className="space-y-0.5 leading-relaxed">
              <div className="font-semibold text-slate-900">
                143만 원 체험판 만료일 자동 과금 차단 가드 (
                <span className="font-mono tabular-nums">
                  {hybridStatus?.agentBuilderRagCredit.expiryGuard.expiryDate || '2027-02-02'}
                </span>
                {hybridStatus
                  ? ` · D-${Math.max(0, hybridStatus.agentBuilderRagCredit.expiryGuard.daysRemaining)}`
                  : ''}
                )
              </div>
              <div className="text-[11px]">
                {hybridStatus?.agentBuilderRagCredit.expiryGuard.statusMessage ||
                  '만료일(2027-02-02) 이후 자동 과금을 막기 위해 AGENT_BUILDER_HARD_STOP_AFTER_EXPIRY=true가 기본 적용됩니다. GCP 결제 > 예산 및 알림도 함께 설정하세요.'}
              </div>
            </div>
          </div>

          {/* 4-Step GCP / AI Studio Setup Checklist */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3">
            {(
              hybridStatus?.setupChecklist || [
                {
                  step: 1,
                  title: 'AI Studio 유료 등급(Pay-as-you-go) 활성화',
                  creditTarget: '13.43만 원 크레딧',
                  envKeys: ['GEMINI_API_KEY'],
                  ready: false,
                },
                {
                  step: 2,
                  title: 'GCP 필수 API 2종 활성화 (Generative Language · Discovery Engine)',
                  creditTarget: '공통 GCP 프로젝트',
                  envKeys: ['GCP_PROJECT_ID'],
                  ready: false,
                },
                {
                  step: 3,
                  title: 'Agent Builder 서비스 계정 / Cloudflare Worker 연결',
                  creditTarget: '143만 원 크레딧',
                  envKeys: ['AGENT_BUILDER_WORKER_URL'],
                  ready: false,
                },
                {
                  step: 4,
                  title: 'Cloud Storage 버킷 및 데이터 스토어 연결',
                  creditTarget: '메일 아카이브 RAG',
                  envKeys: ['GCS_MAIL_ARCHIVE_BUCKET', 'AGENT_BUILDER_DATA_STORE_ID'],
                  ready: false,
                },
              ]
            ).map((step) => (
              <div
                key={step.step}
                className={`p-3 rounded-lg border text-xs space-y-1.5 ${
                  step.ready
                    ? 'bg-emerald-50/60 border-emerald-200'
                    : 'bg-white border-slate-200'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="font-mono text-[11px] font-bold text-[#0F6CBD]">
                    STEP 0{step.step}
                  </span>
                  <span
                    className={`text-[10px] font-semibold ${
                      step.ready ? 'text-emerald-700' : 'text-slate-400'
                    }`}
                  >
                    {step.ready ? '준비됨' : '설정 필요'}
                  </span>
                </div>
                <div className="font-semibold text-slate-900 leading-snug">{step.title}</div>
                <div className="text-[11px] text-slate-500">{step.creditTarget}</div>
                <div className="font-mono text-[10px] text-slate-500 truncate">
                  {step.envKeys.join(' · ')}
                </div>
              </div>
            ))}
          </div>
        </section>

        {/* Section 00-B: Enterprise Multi-User API Policy + PC Outlook Bridge + .exe / Cloudflare Build */}
        <section className="bg-white border border-[#E0E0E0] rounded-[8px] p-6 space-y-5 shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-4 border-b border-[#EDEBE9]">
            <div className="space-y-1">
              <div className="flex items-center gap-2">
                <Users className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
                <h2 className="text-sm font-bold text-[#242424]">
                  전사 공통(최고 관리자 .env) · 개인별 API 연결 및 PC Outlook · .exe / Cloudflare 배포
                </h2>
              </div>
              <p className="text-xs text-[#616161]">
                최고 관리자가 서버 <code className="font-mono text-[#242424]">.env</code> / Cloudflare Secrets에 API 키를 1회 등록하면 10~20명+ 직원이 각자 PC에 설치 후 개인 ID·POP3·PST로 로그인해 개인정보 하드코딩 없이 동일한 AI 기능을 사용합니다.
              </p>
            </div>
            <div className="flex items-center gap-2 shrink-0 flex-wrap">
              <button
                type="button"
                onClick={() => void handleTestApiConnections()}
                disabled={isTestingApis}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-[#C8C6C4] rounded-[4px] transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
              >
                <Activity className="w-3.5 h-3.5" strokeWidth={1.75} />
                {isTestingApis ? '연결 점검 중...' : '전체 API 연결 점검'}
              </button>
              {onOpenPop3Modal && (
                <button
                  type="button"
                  onClick={onOpenPop3Modal}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[4px] transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                >
                  <Plug className="w-3.5 h-3.5" strokeWidth={1.75} />
                  개인 POP3 · 메일플러그 연결
                </button>
              )}
            </div>
          </div>

          {/* Access Scope Switcher (Org-Wide vs Per-User) */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <button
              type="button"
              onClick={() => setPolicyScope('org-wide')}
              className={`p-3.5 rounded-lg border text-left transition-colors ${
                policyScope === 'org-wide'
                  ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#242424]'
                  : 'bg-[#FAF9F8] border-slate-200 text-slate-700 hover:bg-slate-50'
              }`}
            >
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs font-bold flex items-center gap-1.5">
                  <Users className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  전체 이용자 공통 (최고 관리자 중앙 .env / Cloudflare Secrets)
                </span>
                {policyScope === 'org-wide' && (
                  <span className="text-[10px] font-semibold text-[#0F6CBD]">적용 중</span>
                )}
              </div>
              <p className="text-[11px] text-slate-600 mt-1 leading-relaxed">
                관리자가 1회 세팅한 API를 전 직원(10~20명+) 설치본이 자동 공유합니다. 직원 PC에는 API 키가 저장되지 않습니다.
              </p>
            </button>

            <button
              type="button"
              onClick={() => setPolicyScope('per-user-override')}
              className={`p-3.5 rounded-lg border text-left transition-colors ${
                policyScope === 'per-user-override'
                  ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#242424]'
                  : 'bg-[#FAF9F8] border-slate-200 text-slate-700 hover:bg-slate-50'
              }`}
            >
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs font-bold flex items-center gap-1.5">
                  <UserCheck className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  개인별 메일 계정 · 로컬 데이터 격리 ({activeEmployeeEmail || '직원 로그인 대기'})
                </span>
                {policyScope === 'per-user-override' && (
                  <span className="text-[10px] font-semibold text-[#0F6CBD]">적용 중</span>
                )}
              </div>
              <p className="text-[11px] text-slate-600 mt-1 leading-relaxed">
                직원별 이메일/ID 로그인 시 프로젝트·초안·기록함·POP3 세션이 계정 단위로 완전 분리됩니다.
              </p>
            </button>
          </div>

          {apiChecks && (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2.5">
              {apiChecks.map((item) => (
                <div
                  key={item.id}
                  className={`p-3 rounded-lg border text-xs space-y-1 ${
                    item.configured
                      ? 'bg-emerald-50/60 border-emerald-200'
                      : 'bg-slate-50 border-slate-200'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-bold text-slate-900 truncate">{item.label}</span>
                    <span
                      className={`text-[10px] font-semibold shrink-0 ${
                        item.configured ? 'text-emerald-700' : 'text-slate-500'
                      }`}
                    >
                      {item.configured ? '연결 가능' : '미설정'}
                    </span>
                  </div>
                  <div className="text-[11px] text-slate-600 leading-snug">{item.detail}</div>
                  <div className="text-[10px] font-mono text-slate-400">
                    범위: {item.scope === 'org-wide' ? '전체 이용자 공통' : '개인별 계정'}
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* Windows PC Outlook Profile / PST / Contacts Bridge + .exe & Cloudflare Build Guide */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 pt-2 border-t border-slate-100">
            <div className="p-4 rounded-lg border border-slate-200 bg-[#FAF9F8] space-y-3">
              <div className="flex items-center justify-between gap-2">
                <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                  <HardDrive className="w-4 h-4 text-[#0F6CBD]" />
                  설치된 PC Outlook 설정 · PST · 주소록 읽기/쓰기
                </div>
                <span className="text-[11px] font-mono text-slate-500">
                  주소록 {unifiedContacts.length}명
                </span>
              </div>
              <p className="text-[11px] text-slate-600 leading-relaxed">
                직원 PC에 설치된 Outlook의 계정·PST 경로·주소록·서명·폴더를 안전하게 읽어오고(비밀번호 제외), 통합 주소록을 Outlook용 vCard(.vcf)로 다시 내보낼 수 있습니다.
              </p>
              <div className="flex flex-wrap items-center gap-2">
                <a
                  href="/api/outlook/pc-inspector.ps1"
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg"
                >
                  <Download className="w-3.5 h-3.5" />
                  PC Outlook 분석 스크립트 (.ps1)
                </a>
                <button
                  type="button"
                  onClick={() => pcProfileInputRef.current?.click()}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg"
                >
                  <Upload className="w-3.5 h-3.5" />
                  프로필 JSON 불러오기
                </button>
                <button
                  type="button"
                  onClick={handleDownloadContactsVcf}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-lg"
                >
                  <Download className="w-3.5 h-3.5" />
                  주소록 쓰기 (.vcf)
                </button>
                <input
                  ref={pcProfileInputRef}
                  type="file"
                  accept=".json,application/json"
                  onChange={(e) => void handleUploadPcProfile(e)}
                  className="hidden"
                />
              </div>

              {pcProfileNotice && (
                <div className="px-3 py-2 rounded bg-[#EBF3FC] border border-[#B4D6FA] text-[11px] text-[#0F6CBD]">
                  {pcProfileNotice}
                </div>
              )}

              {outlookPcProfile && (
                <div className="p-2.5 rounded bg-white border border-slate-200 text-[11px] space-y-1">
                  <div className="font-semibold text-slate-900">
                    감지 PC: {outlookPcProfile.machineName || 'Windows PC'} (Outlook {outlookPcProfile.outlookVersion || '-'})
                  </div>
                  <div className="text-slate-600">
                    계정 {outlookPcProfile.accounts.length}개 · PST/OST {outlookPcProfile.pstPaths.length}개 · 주소록 {outlookPcProfile.contacts.length}명 · 서명 {outlookPcProfile.signatureNames.length}개
                  </div>
                </div>
              )}
            </div>

            <div className="p-4 rounded-lg border border-slate-200 bg-[#FAF9F8] space-y-2.5">
              <div className="text-xs font-bold text-slate-900">
                최적 사내 배포 구조: Windows .exe 설치본 + Cloudflare 중앙 게이트웨이
              </div>
              <ol className="list-decimal pl-4 text-[11px] text-slate-600 space-y-1.5 leading-relaxed">
                <li>
                  <strong>최고 관리자 1회 설정 (Cloudflare Secrets):</strong>{' '}
                  <code className="font-mono text-slate-900">npx wrangler secret put GEMINI_API_KEY</code> 후{' '}
                  <code className="font-mono text-slate-900">npm run build &amp;&amp; npx wrangler deploy</code>로 사내 공용 게이트웨이 배포 (<code className="font-mono">wrangler.toml</code> 포함)
                </li>
                <li>
                  <strong>Windows .exe 인스톨러 빌드:</strong>{' '}
                  <code className="font-mono text-slate-900">npx electron-builder --config desktop/electron-builder.json</code>으로 직원 PC별 설치 파일(<code className="font-mono">Bmail-Outlook-AI-Setup.exe</code>) 생성
                </li>
                <li>
                  <strong>직원 10~20명 개별 사용:</strong> 각 직원이 설치 후 본인 회사 이메일+비밀번호(POP3 자동 감지) 또는 PC PST 파일로 즉시 로그인·분석 (개인정보/키 하드코딩 0건)
                </li>
              </ol>
            </div>
          </div>
        </section>

        {/* Section 00-C: Super-Admin Encrypted EMV Vault (Org-Wide & Per-Employee API Key Management) */}
        <EmvVaultPanel
          activeEmployeeEmail={activeEmployeeEmail}
          pop3SessionToken={pop3SessionToken}
          onVaultChanged={() => {
            onVaultUpdated?.();
            void handleTestApiConnections();
          }}
        />

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-start">
          {/* Left 2 Columns: config.json Editable Form */}
          <form onSubmit={handleSave} className="lg:col-span-2 space-y-6">
            {/* Section 1: Autonomous AI Agent Config */}
            <section className="bg-white border border-[#E0E0E0] rounded-[8px] p-6 space-y-5 shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
              <div className="flex items-center justify-between gap-2 pb-3 border-b border-[#EDEBE9]">
                <h2 className="text-sm font-bold text-[#242424]">
                  01. AI 에이전트 자동화 설정 (config.json → agent)
                </h2>
                <span className="text-[11px] font-mono text-[#616161]">비비밀 설정</span>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="space-y-1.5">
                  <label htmlFor="cfg-default-scope" className="block text-xs font-semibold text-[#242424]">
                    기본 메일 범위
                  </label>
                  <select
                    id="cfg-default-scope"
                    value={draft.agent.defaultScope}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        agent: { ...prev.agent, defaultScope: e.target.value as MailScope },
                      }))
                    }
                    className={FLUENT_SELECT_CLASS}
                  >
                    <option value="all">수신+발신 전체</option>
                    <option value="received">받은 메일만</option>
                    <option value="sent">보낸 메일만</option>
                    <option value="selected">선택한 메일</option>
                  </select>
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-default-date-mode" className="block text-xs font-semibold text-[#242424]">
                    기본 일자 모드
                  </label>
                  <select
                    id="cfg-default-date-mode"
                    value={draft.agent.defaultDateMode}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        agent: { ...prev.agent, defaultDateMode: e.target.value as DateMode },
                      }))
                    }
                    className={FLUENT_SELECT_CLASS}
                  >
                    <option value="single">특정 일자</option>
                    <option value="range">기간 지정</option>
                  </select>
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-default-tone" className="block text-xs font-semibold text-[#242424]">
                    기본 블로그 문체
                  </label>
                  <select
                    id="cfg-default-tone"
                    value={draft.agent.defaultBlogTone}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        agent: { ...prev.agent, defaultBlogTone: e.target.value as BlogTone },
                      }))
                    }
                    className={FLUENT_SELECT_CLASS}
                  >
                    <option value="insight">실무 인사이트</option>
                    <option value="tech-dev">개발 기술 블로그</option>
                    <option value="essay">회고 에세이</option>
                    <option value="newsletter">주간 뉴스레터</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                <label className="flex items-start gap-2.5 text-xs text-[#424242] cursor-pointer">
                  <input
                    type="checkbox"
                    checked={draft.agent.autoRunOnImport}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        agent: { ...prev.agent, autoRunOnImport: e.target.checked },
                      }))
                    }
                    className={`mt-0.5 ${FLUENT_CHECKBOX_CLASS}`}
                  />
                  <span>
                    <strong className="block text-[#242424]">PST/EML 연동 즉시 AI 에이전트 자동 실행</strong>
                    <span className="text-[11px] text-[#616161]">파일을 불러오면 클릭 없이 자동으로 일지·블로그 생성</span>
                  </span>
                </label>

                <label className="flex items-start gap-2.5 text-xs text-[#424242] cursor-pointer">
                  <input
                    type="checkbox"
                    checked={draft.agent.autoMaskSensitiveInfo}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        agent: { ...prev.agent, autoMaskSensitiveInfo: e.target.checked },
                      }))
                    }
                    className={`mt-0.5 ${FLUENT_CHECKBOX_CLASS}`}
                  />
                  <span>
                    <strong className="block text-[#242424]">민감 정보 자동 마스킹 기본 적용</strong>
                    <span className="text-[11px] text-[#616161]">이메일·전화번호·금액을 LLM 전송 전 비식별화</span>
                  </span>
                </label>

                <label className="flex items-start gap-2.5 text-xs text-[#424242] cursor-pointer">
                  <input
                    type="checkbox"
                    checked={draft.agent.autoSaveArchive}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        agent: { ...prev.agent, autoSaveArchive: e.target.checked },
                      }))
                    }
                    className={`mt-0.5 ${FLUENT_CHECKBOX_CLASS}`}
                  />
                  <span>
                    <strong className="block text-[#242424]">생성 결과 자동 기록 저장</strong>
                    <span className="text-[11px] text-[#616161]">완료된 업무 일지와 블로그를 기록함에 자동 보관</span>
                  </span>
                </label>

                <label className="flex items-start gap-2.5 text-xs text-[#424242] cursor-pointer">
                  <input
                    type="checkbox"
                    checked={draft.ai.fallbackToBuiltIn}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        ai: { ...prev.ai, fallbackToBuiltIn: e.target.checked },
                      }))
                    }
                    className={`mt-0.5 ${FLUENT_CHECKBOX_CLASS}`}
                  />
                  <span>
                    <strong className="block text-[#242424]">선택 AI 키 미등록 시 내장 엔진 자동 폴백</strong>
                    <span className="text-[11px] text-[#616161]">에이전트가 중단 없이 가용 엔진으로 자동 전환</span>
                  </span>
                </label>
              </div>

              <div className="space-y-1.5">
                <label htmlFor="cfg-directive" className="block text-xs font-semibold text-[#242424]">
                  AI 에이전트 상시 지침 (선택)
                </label>
                <input
                  id="cfg-directive"
                  type="text"
                  value={draft.agent.customSystemDirective}
                  onChange={(e) =>
                    setDraft((prev) => ({
                      ...prev,
                      agent: { ...prev.agent, customSystemDirective: e.target.value },
                    }))
                  }
                  placeholder="예: 프로젝트명과 일정 지연 리스크를 최우선으로 요약해 줘"
                  className={FLUENT_INPUT_CLASS}
                />
              </div>
            </section>

            {/* Section 2: AI Models & Plugin Settings */}
            <section className="bg-white border border-[#E0E0E0] rounded-[8px] p-6 space-y-5 shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
              <div className="pb-3 border-b border-[#EDEBE9]">
                <h2 className="text-sm font-bold text-[#242424]">
                  02. AI 모델 및 Outlook 플러그인 설정 (config.json → ai / plugin)
                </h2>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label htmlFor="cfg-default-provider" className="block text-xs font-semibold text-[#242424]">
                    우선 AI 엔진
                  </label>
                  <select
                    id="cfg-default-provider"
                    value={draft.ai.defaultProvider}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        ai: { ...prev.ai, defaultProvider: e.target.value as AiProviderId },
                      }))
                    }
                    className={FLUENT_SELECT_CLASS}
                  >
                    <option value="groq">Groq (무료 API)</option>
                    <option value="claude">Claude (Anthropic)</option>
                    <option value="openai">ChatGPT (OpenAI)</option>
                    <option value="builtin">기본 내장 엔진</option>
                  </select>
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-plugin-name" className="block text-xs font-semibold text-[#242424]">
                    Outlook 리본 표시 이름
                  </label>
                  <input
                    id="cfg-plugin-name"
                    type="text"
                    value={draft.plugin.displayName}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        plugin: { ...prev.plugin, displayName: e.target.value },
                      }))
                    }
                    className={FLUENT_INPUT_CLASS}
                  />
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-model-groq" className="block text-xs font-semibold text-[#242424]">
                    Groq 모델 ID
                  </label>
                  <input
                    id="cfg-model-groq"
                    type="text"
                    value={draft.ai.models.groq}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        ai: { ...prev.ai, models: { ...prev.ai.models, groq: e.target.value } },
                      }))
                    }
                    className={`${FLUENT_INPUT_CLASS} font-mono`}
                  />
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-model-claude" className="block text-xs font-semibold text-[#242424]">
                    Claude 모델 ID
                  </label>
                  <input
                    id="cfg-model-claude"
                    type="text"
                    value={draft.ai.models.claude}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        ai: { ...prev.ai, models: { ...prev.ai.models, claude: e.target.value } },
                      }))
                    }
                    className={`${FLUENT_INPUT_CLASS} font-mono`}
                  />
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-model-openai" className="block text-xs font-semibold text-[#242424]">
                    ChatGPT 모델 ID
                  </label>
                  <input
                    id="cfg-model-openai"
                    type="text"
                    value={draft.ai.models.openai}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        ai: { ...prev.ai, models: { ...prev.ai.models, openai: e.target.value } },
                      }))
                    }
                    className={`${FLUENT_INPUT_CLASS} font-mono`}
                  />
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-max-emails" className="block text-xs font-semibold text-[#242424]">
                    회당 최대 분석 메일 수
                  </label>
                  <input
                    id="cfg-max-emails"
                    type="number"
                    min={1}
                    max={100}
                    value={draft.ai.maxEmailsPerRun}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        ai: { ...prev.ai, maxEmailsPerRun: Number(e.target.value) || 40 },
                      }))
                    }
                    className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                  />
                </div>
              </div>

              {configPasswordRequired && (
                <div className="pt-3 border-t border-[#EDEBE9] space-y-1.5">
                  <label htmlFor="cfg-admin-pw" className="block text-xs font-semibold text-[#242424]">
                    관리자 비밀번호 (.env의 CONFIG_ADMIN_PASSWORD)
                  </label>
                  <input
                    id="cfg-admin-pw"
                    type="password"
                    value={adminPassword}
                    onChange={(e) => setAdminPassword(e.target.value)}
                    placeholder="config.json 저장 시 필요"
                    className={`${FLUENT_INPUT_CLASS} max-w-sm`}
                  />
                </div>
              )}

              {statusMessage && (
                <div
                  role="alert"
                  className={`flex items-center gap-2 p-3 rounded-[4px] text-xs border ${
                    statusMessage.type === 'success'
                      ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                      : 'bg-red-50 border-red-200 text-red-800'
                  }`}
                >
                  {statusMessage.type === 'success' ? (
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" strokeWidth={1.75} />
                  ) : (
                    <AlertCircle className="w-4 h-4 text-red-600 shrink-0" strokeWidth={1.75} />
                  )}
                  <span>{statusMessage.text}</span>
                </div>
              )}

              <div className="flex items-center justify-between pt-3 border-t border-[#EDEBE9]">
                <button
                  type="button"
                  onClick={handleResetDefaults}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#FAF9F8] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                >
                  <RotateCcw className="w-3.5 h-3.5" strokeWidth={1.75} />
                  기본값 복원
                </button>
                <button
                  type="submit"
                  disabled={isSaving}
                  className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-[#C8C6C4] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
                >
                  <Save className="w-3.5 h-3.5" strokeWidth={1.75} />
                  {isSaving ? '저장 중...' : 'config.json 저장'}
                </button>
              </div>
            </section>

            {/* Section 3: External PDF / Excel Project Folder Adapter */}
            <section className="bg-white border border-[#E0E0E0] rounded-[8px] p-6 space-y-5 shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
              <div className="flex items-center justify-between gap-2 pb-3 border-b border-[#EDEBE9]">
                <h2 className="text-sm font-bold text-[#242424]">
                  03. 외부 PDF · Excel 분석 프로젝트 폴더 연동 (config.json → externalEngines)
                </h2>
                <button
                  type="button"
                  disabled={isProbing}
                  onClick={() => void handleProbeEngines()}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                >
                  <FolderSearch className="w-3.5 h-3.5" strokeWidth={1.75} />
                  {isProbing ? '폴더 감지 중...' : '폴더 감지 테스트'}
                </button>
              </div>

              <p className="text-xs text-[#616161] leading-relaxed">
                로컬 PC나 다른 에이전트 프로젝트에 이미 구축해 둔 고성능 PDF·Excel 분석 폴더 경로만 입력하면 해당 엔진을 우선 호출합니다. 미입력 시 내장 MIT 파서(pdf-parse · exceljs)로 자동 실행됩니다.
              </p>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label htmlFor="cfg-pdf-dir" className="block text-xs font-semibold text-[#242424]">
                    외부 PDF 분석 프로젝트 폴더 경로
                  </label>
                  <input
                    id="cfg-pdf-dir"
                    type="text"
                    value={draft.externalEngines.pdfEngineDir}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        externalEngines: { ...prev.externalEngines, pdfEngineDir: e.target.value },
                      }))
                    }
                    placeholder="예: D:\projects\my-pdf-analyzer 또는 /opt/pdf-engine"
                    className={`${FLUENT_INPUT_CLASS} font-mono`}
                  />
                </div>
                <div className="space-y-1.5">
                  <label htmlFor="cfg-pdf-cmd" className="block text-xs font-semibold text-[#242424]">
                    외부 PDF 실행 명령어 (선택, {'{file}'} 치환)
                  </label>
                  <input
                    id="cfg-pdf-cmd"
                    type="text"
                    value={draft.externalEngines.pdfEngineCommand}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        externalEngines: { ...prev.externalEngines, pdfEngineCommand: e.target.value },
                      }))
                    }
                    placeholder="예: python maillog_adapter.py {file} --json"
                    className={`${FLUENT_INPUT_CLASS} font-mono`}
                  />
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="cfg-excel-dir" className="block text-xs font-semibold text-[#242424]">
                    외부 Excel 분석 프로젝트 폴더 경로
                  </label>
                  <input
                    id="cfg-excel-dir"
                    type="text"
                    value={draft.externalEngines.excelEngineDir}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        externalEngines: { ...prev.externalEngines, excelEngineDir: e.target.value },
                      }))
                    }
                    placeholder="예: D:\projects\my-excel-engine 또는 /opt/excel-engine"
                    className={`${FLUENT_INPUT_CLASS} font-mono`}
                  />
                </div>
                <div className="space-y-1.5">
                  <label htmlFor="cfg-excel-cmd" className="block text-xs font-semibold text-[#242424]">
                    외부 Excel 실행 명령어 (선택, {'{file}'} 치환)
                  </label>
                  <input
                    id="cfg-excel-cmd"
                    type="text"
                    value={draft.externalEngines.excelEngineCommand}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        externalEngines: { ...prev.externalEngines, excelEngineCommand: e.target.value },
                      }))
                    }
                    placeholder="예: python excel_adapter.py {file} --json"
                    className={`${FLUENT_INPUT_CLASS} font-mono`}
                  />
                </div>
              </div>

              <label className="flex items-center gap-2 text-xs text-[#424242] cursor-pointer">
                <input
                  type="checkbox"
                  checked={draft.externalEngines.fallbackToBuiltIn}
                  onChange={(e) =>
                    setDraft((prev) => ({
                      ...prev,
                      externalEngines: { ...prev.externalEngines, fallbackToBuiltIn: e.target.checked },
                    }))
                  }
                  className={FLUENT_CHECKBOX_CLASS}
                />
                <span>외부 폴더 미감지 또는 오류 시 내장 PDF(pdf-parse)·Excel(exceljs) 엔진으로 자동 폴백</span>
              </label>

              {probeResults && (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                  {([probeResults.pdf, probeResults.excel] as const).map((pr) => (
                    <div
                      key={pr.kind}
                      className={`p-3 rounded-[6px] border ${
                        pr.exists && pr.detectedRuntime !== 'none'
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                          : 'bg-[#FAF9F8] border-[#E0E0E0] text-[#424242]'
                      }`}
                    >
                      <div className="font-bold font-mono uppercase">{pr.kind} 엔진 감지 결과</div>
                      <div className="text-[11px] mt-0.5">{pr.message}</div>
                    </div>
                  ))}
                </div>
              )}
            </section>

            {/* Live config.json JSON Preview */}
            <section className="bg-white border border-[#E0E0E0] rounded-[8px] overflow-hidden shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
              <div className="px-5 py-3 border-b border-[#EDEBE9] flex items-center justify-between bg-[#FAF9F8]">
                <span className="inline-flex items-center gap-1.5 text-xs font-bold text-[#242424]">
                  <FileCode2 className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
                  실시간 config.json 미리보기 (비밀 키 자동 차단)
                </span>
                <span className="text-[11px] font-mono text-[#707070]">/config.json</span>
              </div>
              <pre className="p-4 text-[11px] leading-relaxed font-mono text-[#242424] bg-[#FAF9F8]/60 overflow-auto max-h-[260px]">
                <code>{JSON.stringify(draft, null, 2)}</code>
              </pre>
            </section>
          </form>

          {/* Right Column: .env Secrets Read-Only Status (Strict Separation) */}
          <aside className="bg-white border border-[#E0E0E0] rounded-[8px] p-5 space-y-4 shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
              <h2 className="text-sm font-bold text-[#242424]">
                .env 시크릿 · 비밀번호 분리 상태
              </h2>
            </div>
            <p className="text-xs text-[#616161] leading-relaxed">
              서버의 <code className="font-mono text-[#242424]">.env</code> 파일(Secrets)과 상단 <strong>EMV 암호화 볼트</strong>의 종합 적용 상태입니다. 키 원문은 클라이언트로 내려오지 않고 설정 여부만 표시됩니다.
            </p>

            <div className="divide-y divide-[#EDEBE9] border border-[#E0E0E0] rounded-[6px] overflow-hidden text-xs">
              {envSecrets.map((secret) => (
                <div key={secret.key} className="px-3 py-2.5 bg-white space-y-0.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-mono font-semibold text-[#242424] text-[11px]">
                      {secret.key}
                    </span>
                    <span
                      className={`inline-flex items-center gap-1 font-medium whitespace-nowrap ${
                        secret.configured ? 'text-[#107C41]' : 'text-[#8A8886]'
                      }`}
                    >
                      {secret.configured && <Check className="w-3 h-3" strokeWidth={2} />}
                      {secret.configured ? '설정됨' : '미설정'}
                    </span>
                  </div>
                  <p className="text-[11px] text-[#616161] leading-snug">{secret.description}</p>
                </div>
              ))}
            </div>

            <div className="p-3 rounded-[6px] bg-[#FAF9F8] border border-[#EDEBE9] space-y-1.5 text-[11px] text-[#616161]">
              <div className="font-semibold text-[#242424]">로컬 PC / 사내 서버 배포 시</div>
              <ol className="list-decimal pl-4 space-y-1 leading-relaxed">
                <li>우측 상단 <strong>.env 템플릿 받기</strong>로 <code className="font-mono text-[#242424]">.env.example</code> 저장</li>
                <li>파일명을 <code className="font-mono text-[#242424]">.env</code>로 변경 후 키/비밀번호 입력</li>
                <li>일반 동작 규칙은 <code className="font-mono text-[#242424]">config.json</code>만 수정 (팀 공유 가능)</li>
              </ol>
              <div className="pt-1 text-[10px] font-mono text-[#707070] break-all">
                OAuth Callback: {oauthCallbackUrl}
              </div>
            </div>
          </aside>
        </div>
      </div>
    </div>
  );
};
