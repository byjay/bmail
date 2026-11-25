import React, { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  ArrowRightLeft,
  BookMarked,
  Building2,
  Check,
  CheckCircle2,
  Database,
  Download,
  ExternalLink,
  FileCode2,
  FolderSync,
  Inbox,
  KeyRound,
  Lock,
  Mail,
  MessageSquare,
  Network,
  Plus,
  RefreshCw,
  RotateCcw,
  Search,
  Send,
  ShieldAlert,
  ShieldCheck,
  Sliders,
  Trash2,
} from 'lucide-react';
import type {
  DomainGlossaryEntry,
  MigrationDryRunReport,
  MigrationExecutionReport,
  SedmsReviewQueueItem,
  SgwEmployee,
  SgwMailRule,
  SgwMailViewItem,
  SgwMailboxFolder,
  SgwUserAiPreferences,
  SgwUserIdentity,
} from '../types/sgwSedms';

interface SgwStatusResponse {
  appName: string;
  sgwBaseUrlConfigured: boolean;
  ssoSecretConfigured: boolean;
  privateAiConfigured: boolean;
  privateAiAllowed: boolean;
  privateAiStatusMessage: string;
  sedmsHmacConfigured: boolean;
  authenticated: boolean;
  user: SgwUserIdentity | null;
  limits: {
    maxSubjectLength: number;
    maxBodyLength: number;
    maxAttachmentBytes: number;
    allowedAttachmentExtensions: string[];
  };
}

interface HandoffFileMeta {
  name: string;
  relPath: string;
  exists: boolean;
  sizeBytes: number;
}

interface SgwSedmsHubViewProps {
  embeddedMode?: boolean;
  initialTab?: 'sgw-mail' | 'ai-glossary' | 'sedms-hub' | 'migration' | 'handoff';
  onOpenSgwChat?: () => void;
  localTotalEmails?: number;
}

export const SgwSedmsHubView: React.FC<SgwSedmsHubViewProps> = ({
  embeddedMode = false,
  initialTab = 'sgw-mail',
  onOpenSgwChat,
  localTotalEmails = 0,
}) => {
  const [activeTab, setActiveTab] = useState<
    'sgw-mail' | 'ai-glossary' | 'sedms-hub' | 'migration' | 'handoff'
  >(initialTab);
  const [status, setStatus] = useState<SgwStatusResponse | null>(null);
  const [ssoTokenInput, setSsoTokenInput] = useState('');
  const [banner, setBanner] = useState<{ type: 'ok' | 'err'; text: string } | null>(null);

  // SGW Mail State
  const [sgwBox, setSgwBox] = useState<'all' | SgwMailboxFolder>('inbox');
  const [sgwSearch, setSgwSearch] = useState('');
  const [sgwUnreadOnly, setSgwUnreadOnly] = useState(false);
  const [sgwMails, setSgwMails] = useState<SgwMailViewItem[]>([]);
  const [selectedMail, setSelectedMail] = useState<SgwMailViewItem | null>(null);
  const [employees, setEmployees] = useState<SgwEmployee[]>([]);
  const [composeOpen, setComposeOpen] = useState(false);
  const [composeTo, setComposeTo] = useState('');
  const [composeCc, setComposeCc] = useState('');
  const [composeBcc, setComposeBcc] = useState('');
  const [composeSubject, setComposeSubject] = useState('');
  const [composeBody, setComposeBody] = useState('');
  const [aiAssistOutput, setAiAssistOutput] = useState<string | null>(null);

  // Preferences, Glossary, Rules
  const [prefs, setPrefs] = useState<SgwUserAiPreferences | null>(null);
  const [glossary, setGlossary] = useState<DomainGlossaryEntry[]>([]);
  const [newTerm, setNewTerm] = useState('');
  const [newDef, setNewDef] = useState('');
  const [newCat, setNewCat] = useState('조선·전장');
  const [rules, setRules] = useState<SgwMailRule[]>([]);
  const [newRuleName, setNewRuleName] = useState('');
  const [newRuleKeyword, setNewRuleKeyword] = useState('');

  // SEDMS Review Queue & Deterministic Match Preview
  const [sedmsTenant, setSedmsTenant] = useState('TENANT-DEFAULT');
  const [reviewItems, setReviewItems] = useState<SedmsReviewQueueItem[]>([]);
  const [previewSubject, setPreviewSubject] = useState('');
  const [previewBody, setPreviewBody] = useState('');
  const [previewFrom, setPreviewFrom] = useState('');
  const [previewResult, setPreviewResult] = useState<string | null>(null);

  // Migration Tool State
  const [migrationJson, setMigrationJson] = useState('[]');
  const [dryRunReport, setDryRunReport] = useState<MigrationDryRunReport | null>(null);
  const [execReport, setExecReport] = useState<MigrationExecutionReport | null>(null);

  // Handoff Files
  const [handoffFiles, setHandoffFiles] = useState<HandoffFileMeta[]>([]);
  const [previewHandoffName, setPreviewHandoffName] = useState<string | null>(null);
  const [previewHandoffContent, setPreviewHandoffContent] = useState<string>('');

  // Live SGW (sgw.seastar.work) & SEDMS (sedms.app) Bridge Control State
  const [sgwChatLive, setSgwChatLive] = useState<{
    connected: boolean;
    sgwBaseUrl: string;
    user?: { name: string; email: string } | null;
  } | null>(null);
  const [sedmsLive, setSedmsLive] = useState<{
    sedmsBaseUrl: string;
    hasToken: boolean;
    tenantId: string;
    health?: { ok: boolean; status: number; latencyMs: number };
  } | null>(null);
  const [showSedmsConfig, setShowSedmsConfig] = useState(false);
  const [sedmsTokenInput, setSedmsTokenInput] = useState('');
  const [sedmsTenantInput, setSedmsTenantInput] = useState('TENANT-DEFAULT');
  const [isSyncingLedger, setIsSyncingLedger] = useState(false);
  const [isScanningSedms, setIsScanningSedms] = useState(false);
  const [sedmsScanReport, setSedmsScanReport] = useState<{
    totalScanned: number;
    drawingMatchedCount: number;
    autoLinkedCount: number;
    reviewQueueCount: number;
    topProjects: Array<{ projectCode: string; mailCount: number; drawingCount: number }>;
    items: Array<{
      emailId: string;
      date: string;
      matchedProject: string | null;
      decision: string;
      drawingRefs: string[];
    }>;
  } | null>(null);

  const loadStatus = useCallback(async () => {
    try {
      const res = await fetch('/api/sgw/status');
      if (res.ok) {
        const data = (await res.json()) as SgwStatusResponse;
        setStatus(data);
      }
    } catch {
      // ignore
    }
  }, []);

  const loadSgwMailData = useCallback(async () => {
    if (!status?.authenticated) return;
    try {
      const params = new URLSearchParams({ box: sgwBox });
      if (sgwUnreadOnly) params.set('unread', '1');
      if (sgwSearch.trim()) params.set('q', sgwSearch.trim());

      const [mailRes, prefRes, gloRes, ruleRes] = await Promise.all([
        fetch(`/api/sgw/mail?${params.toString()}`),
        fetch('/api/sgw/preferences'),
        fetch('/api/sgw/glossary'),
        fetch('/api/sgw/rules'),
      ]);
      if (mailRes.ok) {
        const d = (await mailRes.json()) as { items: SgwMailViewItem[] };
        setSgwMails(d.items || []);
      }
      if (prefRes.ok) {
        const d = (await prefRes.json()) as { preferences: SgwUserAiPreferences };
        setPrefs(d.preferences);
      }
      if (gloRes.ok) {
        const d = (await gloRes.json()) as { entries: DomainGlossaryEntry[] };
        setGlossary(d.entries || []);
      }
      if (ruleRes.ok) {
        const d = (await ruleRes.json()) as { rules: SgwMailRule[] };
        setRules(d.rules || []);
      }

      const empRes = await fetch('/api/sgw/employees');
      if (empRes.ok) {
        const d = (await empRes.json()) as { employees: SgwEmployee[] };
        setEmployees(d.employees || []);
      }
    } catch {
      // ignore
    }
  }, [status?.authenticated, sgwBox, sgwUnreadOnly, sgwSearch]);

  const loadSedmsAndHandoff = useCallback(async () => {
    try {
      const [catRes, handRes] = await Promise.all([
        fetch(`/api/sedms/catalog?tenantId=${encodeURIComponent(sedmsTenant)}`),
        fetch('/api/sedms/handoff/list'),
      ]);
      if (catRes.ok) {
        const d = (await catRes.json()) as { reviewQueue: SedmsReviewQueueItem[] };
        setReviewItems(d.reviewQueue || []);
      }
      if (handRes.ok) {
        const d = (await handRes.json()) as { files: HandoffFileMeta[] };
        setHandoffFiles(d.files || []);
      }
    } catch {
      // ignore
    }
  }, [sedmsTenant]);

  const loadLiveEcosystemStatus = useCallback(async () => {
    try {
      const [sgwChatRes, sedmsRes] = await Promise.all([
        fetch('/api/sgw-chat/status'),
        fetch('/api/sedms-live/status'),
      ]);
      if (sgwChatRes.ok) {
        const d = (await sgwChatRes.json()) as {
          connected: boolean;
          sgwBaseUrl: string;
          user?: { name: string; email: string } | null;
        };
        setSgwChatLive(d);
      }
      if (sedmsRes.ok) {
        const d = (await sedmsRes.json()) as {
          sedmsBaseUrl: string;
          hasToken: boolean;
          tenantId: string;
          health?: { ok: boolean; status: number; latencyMs: number };
        };
        setSedmsLive(d);
        if (d.tenantId) setSedmsTenantInput(d.tenantId);
      }
    } catch {
      // ignore offline probe error
    }
  }, []);

  useEffect(() => {
    void loadStatus();
    void loadSedmsAndHandoff();
    void loadLiveEcosystemStatus();
  }, [loadStatus, loadSedmsAndHandoff, loadLiveEcosystemStatus]);

  useEffect(() => {
    void loadSgwMailData();
  }, [loadSgwMailData]);

  // Auto-exchange `?sso_token=...` when SGW launches Bmail in embedded mode
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const token = params.get('sso_token');
    if (token) {
      void (async () => {
        const res = await fetch('/api/sgw/auth/exchange', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ token }),
        });
        if (res.ok) {
          params.delete('sso_token');
          const clean = params.toString() ? `${window.location.pathname}?${params.toString()}` : window.location.pathname;
          window.history.replaceState({}, '', clean);
          await loadStatus();
        }
      })();
    }
  }, [loadStatus]);

  const handleSsoExchange = async () => {
    setBanner(null);
    const res = await fetch('/api/sgw/auth/exchange', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: ssoTokenInput.trim() }),
    });
    const data = (await res.json()) as { error?: string; identity?: SgwUserIdentity };
    if (!res.ok) {
      setBanner({ type: 'err', text: data.error || 'SSO 토큰 교환 실패' });
      return;
    }
    setSsoTokenInput('');
    setBanner({ type: 'ok', text: `SGW SSO 인증 완료: ${data.identity?.name} (${data.identity?.sub})` });
    await loadStatus();
  };

  const handleSendSgwMail = async (saveAsDraft: boolean) => {
    setBanner(null);
    const parseRecs = (raw: string) =>
      raw
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean)
        .map((sub) => {
          const found = employees.find((e) => e.sub === sub || e.email === sub || e.name === sub);
          return { sub: found?.sub || sub, name: found?.name || sub };
        });

    const res = await fetch('/api/sgw/mail', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        to: parseRecs(composeTo),
        cc: parseRecs(composeCc),
        bcc: parseRecs(composeBcc),
        subject: composeSubject,
        body: composeBody,
        saveAsDraft,
      }),
    });
    const data = (await res.json()) as { error?: string };
    if (!res.ok) {
      setBanner({ type: 'err', text: data.error || '메일 저장/발송 실패' });
      return;
    }
    setComposeOpen(false);
    setComposeTo('');
    setComposeCc('');
    setComposeBcc('');
    setComposeSubject('');
    setComposeBody('');
    setBanner({ type: 'ok', text: saveAsDraft ? '임시보관함에 저장되었습니다.' : 'SGW 사내메일을 발송했습니다.' });
    await loadSgwMailData();
  };

  const handleToggleRead = async (mail: SgwMailViewItem) => {
    const res = await fetch(`/api/sgw/mail/${encodeURIComponent(mail.id)}/read`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ isRead: !mail.isRead }),
    });
    const data = (await res.json()) as { error?: string; item?: SgwMailViewItem };
    if (!res.ok) {
      setBanner({ type: 'err', text: data.error || '읽음 상태 변경 실패' });
      return;
    }
    if (data.item) setSelectedMail(data.item);
    await loadSgwMailData();
  };

  const handleDeleteMail = async (mail: SgwMailViewItem, permanent: boolean) => {
    const url = `/api/sgw/mail/${encodeURIComponent(mail.id)}${permanent ? '?permanent=true' : ''}`;
    const res = await fetch(url, { method: 'DELETE' });
    const data = (await res.json()) as { error?: string };
    if (!res.ok) {
      setBanner({ type: 'err', text: data.error || '삭제 실패' });
      return;
    }
    setSelectedMail(null);
    setBanner({ type: 'ok', text: permanent ? '영구 삭제되었습니다.' : '휴지통으로 이동했습니다.' });
    await loadSgwMailData();
  };

  const handleRunSgwPrivateAi = async (action: 'summarize' | 'reply_draft' | 'classify_and_rules') => {
    if (!selectedMail) return;
    setAiAssistOutput(null);
    setBanner(null);
    const res = await fetch('/api/sgw/ai/assist', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messageId: selectedMail.id, action }),
    });
    const data = (await res.json()) as { error?: string; result?: unknown };
    if (!res.ok) {
      setBanner({ type: 'err', text: data.error || 'AI 호출 차단됨' });
      return;
    }
    setAiAssistOutput(JSON.stringify(data.result, null, 2));
  };

  const handleSavePreferences = async (next: Partial<SgwUserAiPreferences>) => {
    const res = await fetch('/api/sgw/preferences', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(next),
    });
    if (res.ok) {
      const d = (await res.json()) as { preferences: SgwUserAiPreferences };
      setPrefs(d.preferences);
      setBanner({ type: 'ok', text: 'AI 맞춤설정이 저장되었습니다.' });
    }
  };

  const handleAddGlossary = async () => {
    if (!newTerm.trim() || !newDef.trim()) return;
    const res = await fetch('/api/sgw/glossary', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ term: newTerm, definition: newDef, category: newCat, scope: 'user' }),
    });
    if (res.ok) {
      setNewTerm('');
      setNewDef('');
      await loadSgwMailData();
    }
  };

  const handleAddRule = async () => {
    if (!newRuleName.trim() || !newRuleKeyword.trim()) return;
    const res = await fetch('/api/sgw/rules', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: newRuleName,
        matchField: 'subject',
        matchKeyword: newRuleKeyword,
        targetCategory: 'approval',
        targetPriority: 'high',
        enabled: true,
      }),
    });
    if (res.ok) {
      setNewRuleName('');
      setNewRuleKeyword('');
      await loadSgwMailData();
    }
  };

  const handleRunMigrationDryRun = async () => {
    setBanner(null);
    setExecReport(null);
    try {
      const rows = JSON.parse(migrationJson);
      const res = await fetch('/api/sgw/migration/dry-run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows }),
      });
      const d = (await res.json()) as { report?: MigrationDryRunReport; error?: string };
      if (!res.ok) throw new Error(d.error || '드라이런 실패');
      setDryRunReport(d.report || null);
    } catch (e) {
      setBanner({ type: 'err', text: e instanceof Error ? e.message : 'JSON 형식을 확인하세요.' });
    }
  };

  const handleExecuteMigration = async () => {
    setBanner(null);
    try {
      const rows = JSON.parse(migrationJson);
      const res = await fetch('/api/sgw/migration/execute', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows, approved: true }),
      });
      const d = (await res.json()) as { report?: MigrationExecutionReport; error?: string };
      if (!res.ok) throw new Error(d.error || '이전 실행 실패');
      setExecReport(d.report || null);
      setBanner({ type: 'ok', text: `이전 완료: ${d.report?.migratedCount ?? 0}건 반영` });
      await loadSgwMailData();
    } catch (e) {
      setBanner({ type: 'err', text: e instanceof Error ? e.message : '이전 실행 오류' });
    }
  };

  const handleRollbackMigration = async () => {
    const res = await fetch('/api/sgw/migration/rollback', { method: 'POST' });
    const d = (await res.json()) as { removedMessages?: number };
    setExecReport(null);
    setBanner({ type: 'ok', text: `롤백 완료: 이전된 메일 ${d.removedMessages ?? 0}건 제거` });
    await loadSgwMailData();
  };

  const handlePreviewSedmsMatch = async () => {
    const res = await fetch('/api/sedms/mail/preview-match', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: {
          tenantId: sedmsTenant,
          accountId: 'acc-preview',
          messageId: `<preview-${Date.now()}@example.test>`,
          direction: 'in',
          from: previewFrom || 'partner@example.test',
          to: ['team@example.test'],
          subject: previewSubject,
          bodyText: previewBody,
          sentAt: new Date().toISOString(),
        },
      }),
    });
    const d = await res.json();
    setPreviewResult(JSON.stringify(d, null, 2));
  };

  const handleSyncRecentMailsToSgwLedger = async () => {
    setBanner(null);
    setIsSyncingLedger(true);
    try {
      const res = await fetch('/api/sgw-bridge/ledger/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ limit: 30 }),
      });
      const data = (await res.json()) as { inserted?: number; error?: string };
      if (!res.ok) {
        setBanner({ type: 'err', text: data.error || 'SGW 원장 동기화에 실패했습니다.' });
        return;
      }
      setBanner({
        type: 'ok',
        text: `SGW(bmail-ledger) 업무원장 동기화 완료: 최근 메일 ${data.inserted ?? 0}건 반영`,
      });
    } catch (err) {
      setBanner({ type: 'err', text: err instanceof Error ? err.message : 'SGW 원장 동기화 오류' });
    } finally {
      setIsSyncingLedger(false);
    }
  };

  const handlePushWeeklyReport = async () => {
    setBanner(null);
    const today = new Date().toISOString().slice(0, 10);
    const res = await fetch('/api/sgw-bridge/ledger/push-weekly', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ logDate: today }),
    });
    const data = (await res.json()) as { ok?: boolean; week?: string; error?: string };
    if (!res.ok) {
      setBanner({ type: 'err', text: data.error || '주간업무보고 반영에 실패했습니다.' });
      return;
    }
    setBanner({ type: 'ok', text: `SGW 주간업무보고(${data.week || today})에 일일 업무로그 반영 완료` });
  };

  const handleScanLocalMailsForSedms = async () => {
    setBanner(null);
    setIsScanningSedms(true);
    try {
      const res = await fetch('/api/sedms-live/scan-local-mails', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ limit: 150, tenantId: sedmsTenantInput }),
      });
      const data = await res.json();
      if (!res.ok) {
        setBanner({ type: 'err', text: data.error || '로컬 메일 SEDMS 스캔 실패' });
        return;
      }
      setSedmsScanReport(data);
      setBanner({
        type: 'ok',
        text: `SEDMS 도면·호선 스캔 완료: ${data.totalScanned}건 중 도면 포함 ${data.drawingMatchedCount}건, 자동연결 ${data.autoLinkedCount}건`,
      });
    } finally {
      setIsScanningSedms(false);
    }
  };

  const handleSaveSedmsConfig = async () => {
    setBanner(null);
    const res = await fetch('/api/sedms-live/config', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        bearerToken: sedmsTokenInput.trim() || undefined,
        tenantId: sedmsTenantInput.trim() || 'TENANT-DEFAULT',
      }),
    });
    if (res.ok) {
      setSedmsTokenInput('');
      setShowSedmsConfig(false);
      await loadLiveEcosystemStatus();
      setBanner({ type: 'ok', text: 'SEDMS 연결 설정이 로컬 AES-256-GCM 금고에 암호화 저장되었습니다.' });
    }
  };

  const handleSyncContactsToSedms = async (commit: boolean) => {
    setBanner(null);
    const endpoint = commit ? '/api/sedms-live/contacts/commit' : '/api/sedms-live/contacts/preview';
    const res = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(commit ? { confirmed: true } : {}),
    });
    const data = (await res.json()) as {
      validCount?: number;
      insertedCount?: number;
      skippedCount?: number;
      error?: string;
    };
    if (!res.ok) {
      setBanner({ type: 'err', text: data.error || 'SEDMS 연락처 동기화 실패' });
      return;
    }
    setBanner({
      type: 'ok',
      text: commit
        ? `SEDMS 수신처 일괄 등록 완료: 신규 ${data.insertedCount ?? 0}명 (건너뜀 ${data.skippedCount ?? 0}명)`
        : `SEDMS 수신처 미리보기: 유효 ${data.validCount ?? 0}명 등록 가능`,
    });
  };

  const handlePreviewHandoffFile = async (name: string) => {
    const res = await fetch(`/api/sedms/handoff/file/${encodeURIComponent(name)}`);
    if (res.ok) {
      const d = (await res.json()) as { content: string };
      setPreviewHandoffName(name);
      setPreviewHandoffContent(d.content);
    }
  };

  return (
    <div className="h-full flex flex-col bg-[#F5F5F5] text-[#242424] overflow-hidden">
      {/* Header */}
      <div className="bg-white border-b border-[#E0E0E0] px-5 py-3 flex flex-wrap items-center justify-between gap-3 shrink-0">
        <div className="flex items-center gap-2.5">
          <Building2 className="w-5 h-5 text-[#0F6CBD]" />
          <div>
            <h1 className="text-[15px] font-bold text-[#242424]">
              {embeddedMode
                ? 'Bmail · SGW 사내메일'
                : 'Bmail · SGW 그룹웨어 & SEDMS 메일 허브 연동 센터'}
            </h1>
            <p className="text-xs text-[#616161]">
              SGW HMAC SSO(sub 신원 통일) · AI 기본 꺼짐/외부 무료 LLM 차단 · SEDMS 5순위 결정적 매칭 · 이관 핸드오프
            </p>
          </div>
        </div>

        <div role="tablist" className="flex flex-wrap items-center gap-1 bg-[#F3F2F1] p-1 rounded-md">
          {[
            { id: 'sgw-mail' as const, label: 'SGW 사내메일', icon: Mail },
            { id: 'ai-glossary' as const, label: 'AI 맞춤 · 용어사전', icon: Sliders },
            { id: 'sedms-hub' as const, label: 'SEDMS 허브 · 대기열', icon: Network },
            { id: 'migration' as const, label: 'internal_mail 이전', icon: ArrowRightLeft },
            { id: 'handoff' as const, label: 'SEDMS·SGW 핸드오프', icon: FileCode2 },
          ].map((tab) => {
            const Icon = tab.icon;
            return (
              <button
                key={tab.id}
                type="button"
                role="tab"
                aria-selected={activeTab === tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`px-2.5 py-1.5 rounded text-xs font-semibold inline-flex items-center gap-1.5 transition-colors ${
                  activeTab === tab.id
                    ? 'bg-white text-[#0F6CBD] shadow-xs'
                    : 'text-[#424242] hover:text-[#242424]'
                }`}
              >
                <Icon className="w-3.5 h-3.5" />
                {tab.label}
              </button>
            );
          })}
        </div>
      </div>

      {banner && (
        <div
          className={`mx-5 mt-3 px-3.5 py-2 rounded-md border text-xs flex items-center justify-between ${
            banner.type === 'ok'
              ? 'bg-[#F3FAF3] border-[#107C10]/30 text-[#107C10]'
              : 'bg-[#FDF3F4] border-[#D13438]/30 text-[#A4262C]'
          }`}
        >
          <span>{banner.text}</span>
          <button type="button" onClick={() => setBanner(null)} className="text-xs underline">
            닫기
          </button>
        </div>
      )}

      {/* Main Content */}
      <div className="flex-1 overflow-y-auto p-5 space-y-4">
        {/* Live SGW (sgw.seastar.work) & SEDMS (sedms.app) Ecosystem Control Card */}
        {!embeddedMode && (
          <section className="bg-white border border-[#CFE4FA] rounded-xl p-4 space-y-3 shadow-xs">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div>
                <h2 className="text-sm font-bold text-[#242424]">
                  SGW(sgw.seastar.work) · SEDMS(sedms.app) 실시간 통합 연동 관제
                </h2>
                <p className="text-[11px] text-[#616161]">
                  구 bonglook을 대체하여 로컬 보관 메일·도면 첨부·사내 채팅·주간업무보고·SEDMS 수신처를 직접 연결합니다.
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                <a
                  href={sgwChatLive?.sgwBaseUrl || 'https://sgw.seastar.work'}
                  target="_blank"
                  rel="noreferrer"
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] rounded"
                >
                  SGW 열기 <ExternalLink className="w-3 h-3" />
                </a>
                <a
                  href={sedmsLive?.sedmsBaseUrl || 'https://sedms.app'}
                  target="_blank"
                  rel="noreferrer"
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] rounded"
                >
                  SEDMS 열기 <ExternalLink className="w-3 h-3" />
                </a>
                <button
                  type="button"
                  onClick={() => void loadLiveEcosystemStatus()}
                  className="inline-flex items-center gap-1 px-2 py-1 text-[11px] text-[#424242] bg-[#F5F5F5] hover:bg-[#E0E0E0] rounded"
                >
                  <RefreshCw className="w-3 h-3" />
                  상태 갱신
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-2.5 text-xs">
              <div className="p-2.5 rounded-lg border border-[#EDEBE9] bg-[#FAF9F8]">
                <div className="text-[11px] text-[#616161]">1. SGW 채팅 · 업무원장</div>
                <div className="font-bold mt-0.5 text-[#242424]">
                  {sgwChatLive?.connected
                    ? `연결됨 (${sgwChatLive.user?.name || sgwChatLive.user?.email})`
                    : 'SGW 채팅 로그인 대기'}
                </div>
              </div>
              <div className="p-2.5 rounded-lg border border-[#EDEBE9] bg-[#FAF9F8]">
                <div className="text-[11px] text-[#616161]">2. SEDMS 도면·수신처 서버</div>
                <div className="font-bold mt-0.5 text-[#242424]">
                  {sedmsLive?.health?.ok
                    ? `정상 응답 HTTP ${sedmsLive.health.status} (${sedmsLive.health.latencyMs}ms)`
                    : '오프라인 / 연결 확인 중'}
                </div>
              </div>
              <div className="p-2.5 rounded-lg border border-[#EDEBE9] bg-[#FAF9F8]">
                <div className="text-[11px] text-[#616161]">3. 로컬 Bmail 보관함</div>
                <div className="font-bold mt-0.5 text-[#242424]">
                  로컬 SQLite {localTotalEmails.toLocaleString()}건 색인됨
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {onOpenSgwChat && (
                <button
                  type="button"
                  onClick={onOpenSgwChat}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-md"
                >
                  <MessageSquare className="w-3.5 h-3.5" />
                  SGW 실시간 채팅 열기
                </button>
              )}
              <button
                type="button"
                disabled={isSyncingLedger}
                onClick={() => void handleSyncRecentMailsToSgwLedger()}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] disabled:opacity-50 rounded-md"
              >
                <FolderSync className="w-3.5 h-3.5" />
                {isSyncingLedger ? '동기화 중...' : '최근 메일 SGW 업무원장 동기화'}
              </button>
              <button
                type="button"
                onClick={() => void handlePushWeeklyReport()}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-[#242424] bg-[#F5F5F5] hover:bg-[#E0E0E0] rounded-md"
              >
                오늘 로그 SGW 주간보고 반영
              </button>
              <button
                type="button"
                disabled={isScanningSedms}
                onClick={() => void handleScanLocalMailsForSedms()}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 hover:bg-emerald-100 rounded-md"
              >
                <Search className="w-3.5 h-3.5" />
                {isScanningSedms ? '스캔 중...' : '로컬 메일 SEDMS 도면·호선 스캔'}
              </button>
              <button
                type="button"
                onClick={() => setShowSedmsConfig((v) => !v)}
                className="px-2.5 py-1.5 text-xs text-[#616161] hover:text-[#242424] border border-[#D1D1D1] rounded-md"
              >
                SEDMS 토큰·수신처 동기화
              </button>
            </div>

            {showSedmsConfig && (
              <div className="p-3 bg-[#FAF9F8] border border-[#E0E0E0] rounded-lg space-y-2 text-xs">
                <div className="flex flex-wrap gap-2">
                  <input
                    type="text"
                    value={sedmsTenantInput}
                    onChange={(e) => setSedmsTenantInput(e.target.value)}
                    placeholder="회사 Tenant ID (기본 TENANT-DEFAULT)"
                    className="px-2.5 py-1.5 border border-[#D1D1D1] rounded bg-white text-xs"
                  />
                  <input
                    type="password"
                    value={sedmsTokenInput}
                    onChange={(e) => setSedmsTokenInput(e.target.value)}
                    placeholder="SEDMS 관리자 Bearer 토큰 (로컬 AES-256 암호화 저장)"
                    className="flex-1 min-w-[220px] px-2.5 py-1.5 border border-[#D1D1D1] rounded bg-white text-xs"
                  />
                  <button
                    type="button"
                    onClick={() => void handleSaveSedmsConfig()}
                    className="px-3 py-1.5 font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded"
                  >
                    토큰 암호화 저장
                  </button>
                </div>
                <div className="flex flex-wrap items-center gap-2 text-[11px] text-[#616161]">
                  <span>토큰 상태: {sedmsLive?.hasToken ? '등록됨' : '미등록'}</span>
                  <button
                    type="button"
                    onClick={() => void handleSyncContactsToSedms(false)}
                    className="px-2.5 py-1 bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded text-[#242424]"
                  >
                    Outlook 연락처 → SEDMS 미리보기
                  </button>
                  <button
                    type="button"
                    onClick={() => void handleSyncContactsToSedms(true)}
                    className="px-2.5 py-1 bg-[#EBF3FC] border border-[#CFE4FA] hover:bg-[#CFE4FA] rounded text-[#0F6CBD] font-semibold"
                  >
                    SEDMS 수신처 일괄 반영
                  </button>
                </div>
              </div>
            )}

            {sedmsScanReport && (
              <div className="p-3 bg-[#FAF9F8] border border-[#E0E0E0] rounded-lg space-y-2 text-xs">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <span className="font-bold text-[#242424]">
                    SEDMS 로컬 메일 스캔 결과 ({sedmsScanReport.totalScanned}건 분석 · 도면 포함 {sedmsScanReport.drawingMatchedCount}건 · 자동매칭 {sedmsScanReport.autoLinkedCount}건)
                  </span>
                </div>
                {sedmsScanReport.topProjects.length > 0 && (
                  <div className="flex flex-wrap gap-1.5">
                    {sedmsScanReport.topProjects.map((p) => (
                      <span
                        key={p.projectCode}
                        className="px-2 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] text-[11px] font-semibold"
                      >
                        {p.projectCode}: 메일 {p.mailCount}건 · 도면 {p.drawingCount}종
                      </span>
                    ))}
                  </div>
                )}
              </div>
            )}
          </section>
        )}

        {/* Readiness Status Bar */}
        {status && (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
            <div className="bg-white border border-[#E0E0E0] rounded-lg p-3">
              <div className="text-[#616161] flex items-center gap-1.5">
                <KeyRound className="w-3.5 h-3.5 text-[#0F6CBD]" /> SGW 관문 SSO
              </div>
              <div className="font-bold mt-1">
                {status.authenticated && status.user
                  ? `${status.user.name} (${status.user.sub})`
                  : status.ssoSecretConfigured
                    ? 'SSO 토큰 대기 중'
                    : 'SGW_SSO_SECRET 미설정'}
              </div>
            </div>
            <div className="bg-white border border-[#E0E0E0] rounded-lg p-3">
              <div className="text-[#616161] flex items-center gap-1.5">
                <ShieldCheck className="w-3.5 h-3.5 text-[#107C10]" /> 비공개 AI 가드 (S1)
              </div>
              <div className="font-bold mt-1">
                {status.privateAiAllowed ? '사내 비공개 엔드포인트 준비됨' : '외부 무료 LLM 차단 · 비공개 미설정'}
              </div>
            </div>
            <div className="bg-white border border-[#E0E0E0] rounded-lg p-3">
              <div className="text-[#616161] flex items-center gap-1.5">
                <Database className="w-3.5 h-3.5 text-[#0F6CBD]" /> SEDMS HMAC 수집기
              </div>
              <div className="font-bold mt-1">
                {status.sedmsHmacConfigured ? '테넌트 HMAC 키 설정됨' : 'SEDMS_INGEST_HMAC_SECRETS_JSON 미설정'}
              </div>
            </div>
            <div className="bg-white border border-[#E0E0E0] rounded-lg p-3">
              <div className="text-[#616161] flex items-center gap-1.5">
                <Lock className="w-3.5 h-3.5 text-[#0F6CBD]" /> 첨부 정책
              </div>
              <div className="font-bold mt-1 font-mono">
                최대 {Math.round(status.limits.maxAttachmentBytes / (1024 * 1024))}MB ·{' '}
                {status.limits.allowedAttachmentExtensions.length}종 허용
              </div>
            </div>
          </div>
        )}

        {/* TAB 1: SGW Groupware Mail */}
        {activeTab === 'sgw-mail' && (
          <div className="space-y-4">
            {!status?.authenticated ? (
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 items-stretch">
                <div className="bg-white border border-[#E0E0E0] rounded-lg p-5 space-y-3 flex flex-col justify-between">
                  <div className="space-y-3">
                    <h2 className="text-sm font-bold flex items-center gap-2">
                      <KeyRound className="w-4 h-4 text-[#0F6CBD]" /> SGW 관문 SSO 로그인 (자체 비밀번호 미사용)
                    </h2>
                    <p className="text-xs text-[#616161] leading-relaxed">
                      SGW 그룹웨어 메인메뉴에서 진입 시 HMAC-SHA256 서명 토큰이 자동 교환됩니다. 수동 검증이 필요하면
                      SGW에서 발급한 단기 SSO 토큰을 아래에 입력하세요. (시크릿은 서버 <code>.env</code>의{' '}
                      <code>SGW_SSO_SECRET</code>으로만 검증)
                    </p>
                  </div>
                  <div className="flex gap-2 pt-2">
                    <input
                      type="password"
                      value={ssoTokenInput}
                      onChange={(e) => setSsoTokenInput(e.target.value)}
                      placeholder="SGW HMAC SSO 토큰 붙여넣기..."
                      className="flex-1 px-3 py-2 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                    />
                    <button
                      type="button"
                      onClick={() => void handleSsoExchange()}
                      className="px-4 py-2 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-semibold rounded whitespace-nowrap"
                    >
                      SSO 세션 교환
                    </button>
                  </div>
                </div>

                <div className="bg-white border border-[#E0E0E0] rounded-lg p-5 space-y-3 flex flex-col justify-between">
                  <div className="space-y-2.5">
                    <h2 className="text-sm font-bold flex items-center gap-2">
                      <Network className="w-4 h-4 text-[#0F6CBD]" /> 그룹웨어 메인메뉴 임베드 · SEDMS 이관 준비
                    </h2>
                    <p className="text-xs text-[#616161] leading-relaxed">
                      사내 그룹웨어(SGW) 메인메뉴에서 Bmail을 단독 임베드 모드로 바로 열거나, SEDMS 5순위 자동 매칭 테스트 및
                      PostgreSQL 마이그레이션·TLA+ 핸드오프 패키지를 즉시 확인할 수 있습니다.
                    </p>
                    <div className="text-[11px] font-mono text-[#424242] bg-[#FAF9F8] border border-[#EDEBE9] rounded px-3 py-2 space-y-1">
                      <div>임베드 진입 경로: /sgw/mail 또는 /?embed=sgw</div>
                      <div>CSP 허용 헤더: .env의 SGW_FRAME_ANCESTORS 자동 적용</div>
                    </div>
                  </div>
                  <div className="flex flex-wrap items-center gap-2 pt-1">
                    <button
                      type="button"
                      onClick={() => setActiveTab('sedms-hub')}
                      className="px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded whitespace-nowrap"
                    >
                      SEDMS 5순위 매칭 검증
                    </button>
                    <button
                      type="button"
                      onClick={() => setActiveTab('migration')}
                      className="px-3 py-1.5 text-xs font-medium text-[#242424] bg-[#F3F2F1] hover:bg-[#E8E6E4] rounded whitespace-nowrap"
                    >
                      internal_mail 드라이런
                    </button>
                    <button
                      type="button"
                      onClick={() => setActiveTab('handoff')}
                      className="px-3 py-1.5 text-xs font-medium text-[#242424] bg-[#F3F2F1] hover:bg-[#E8E6E4] rounded whitespace-nowrap"
                    >
                      SEDMS·SGW 핸드오프 패키지
                    </button>
                  </div>
                </div>
              </div>
            ) : (
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
                {/* Folder Sidebar */}
                <div className="lg:col-span-2 bg-white border border-[#E0E0E0] rounded-lg p-3 space-y-2">
                  <button
                    type="button"
                    onClick={() => setComposeOpen(true)}
                    className="w-full py-2 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-semibold rounded flex items-center justify-center gap-1.5"
                  >
                    <Send className="w-3.5 h-3.5" /> 새 사내메일
                  </button>
                  <div className="space-y-0.5 pt-1">
                    {[
                      { id: 'inbox' as const, label: '받은편지함' },
                      { id: 'sent' as const, label: '보낸편지함' },
                      { id: 'drafts' as const, label: '임시보관함' },
                      { id: 'trash' as const, label: '휴지통' },
                      { id: 'all' as const, label: '전체 메일' },
                    ].map((f) => (
                      <button
                        key={f.id}
                        type="button"
                        onClick={() => setSgwBox(f.id)}
                        className={`w-full text-left px-2.5 py-1.5 rounded text-xs font-medium ${
                          sgwBox === f.id
                            ? 'bg-[#EBF3FC] text-[#0F6CBD] font-semibold'
                            : 'text-[#323130] hover:bg-[#F5F5F5]'
                        }`}
                      >
                        {f.label}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Mail List */}
                <div className="lg:col-span-4 bg-white border border-[#E0E0E0] rounded-lg flex flex-col overflow-hidden">
                  <div className="p-2.5 border-b border-[#E0E0E0] flex items-center gap-2">
                    <input
                      type="search"
                      value={sgwSearch}
                      onChange={(e) => setSgwSearch(e.target.value)}
                      placeholder="보낸이·제목·본문 검색"
                      className="flex-1 px-2.5 py-1 text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
                    />
                    <label className="inline-flex items-center gap-1 text-[11px] text-[#424242] shrink-0">
                      <input
                        type="checkbox"
                        checked={sgwUnreadOnly}
                        onChange={(e) => setSgwUnreadOnly(e.target.checked)}
                      />
                      안읽음
                    </label>
                  </div>
                  <div className="divide-y divide-[#EDEBE9] max-h-[520px] overflow-y-auto">
                    {sgwMails.length === 0 ? (
                      <div className="p-8 text-center text-xs text-[#8A8886]">
                        해당 폴더에 메일이 없습니다.
                      </div>
                    ) : (
                      sgwMails.map((m) => (
                        <button
                          key={m.id}
                          type="button"
                          onClick={() => setSelectedMail(m)}
                          className={`w-full text-left p-3 hover:bg-[#FAF9F8] transition-colors ${
                            selectedMail?.id === m.id ? 'bg-[#EBF3FC]/70' : ''
                          }`}
                        >
                          <div className="flex items-center justify-between gap-2 text-xs">
                            <span className={m.isRead ? 'text-[#323130]' : 'font-bold text-[#0F6CBD]'}>
                              {m.senderName}
                            </span>
                            <span className="text-[10px] text-[#616161] font-mono">
                              {(m.sentAt || m.createdAt).slice(0, 10)}
                            </span>
                          </div>
                          <div className={`text-xs truncate mt-0.5 ${m.isRead ? '' : 'font-bold'}`}>
                            {m.subject}
                          </div>
                          <div className="text-[11px] text-[#616161] truncate mt-0.5">{m.body}</div>
                        </button>
                      ))
                    )}
                  </div>
                </div>

                {/* Reading / Compose Pane */}
                <div className="lg:col-span-6 bg-white border border-[#E0E0E0] rounded-lg p-4 space-y-3">
                  {composeOpen ? (
                    <div className="space-y-2.5">
                      <div className="flex items-center justify-between">
                        <h3 className="text-sm font-bold">SGW 사내메일 작성 (발신자 서버 강제: {status.user?.sub})</h3>
                        <button type="button" onClick={() => setComposeOpen(false)} className="text-xs text-[#616161]">
                          취소
                        </button>
                      </div>
                      <input
                        type="text"
                        value={composeTo}
                        onChange={(e) => setComposeTo(e.target.value)}
                        placeholder="받는사람 sub (쉼표 구분, SGW /api/employees 연동)"
                        className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded"
                      />
                      <div className="grid grid-cols-2 gap-2">
                        <input
                          type="text"
                          value={composeCc}
                          onChange={(e) => setComposeCc(e.target.value)}
                          placeholder="참조 (CC, 선택)"
                          className="px-3 py-1.5 text-xs border border-[#D1D1D1] rounded"
                        />
                        <input
                          type="text"
                          value={composeBcc}
                          onChange={(e) => setComposeBcc(e.target.value)}
                          placeholder="숨은참조 (BCC, 선택)"
                          className="px-3 py-1.5 text-xs border border-[#D1D1D1] rounded"
                        />
                      </div>
                      <input
                        type="text"
                        value={composeSubject}
                        onChange={(e) => setComposeSubject(e.target.value)}
                        placeholder="제목"
                        className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded"
                      />
                      <textarea
                        rows={8}
                        value={composeBody}
                        onChange={(e) => setComposeBody(e.target.value)}
                        placeholder="메일 본문을 입력하세요..."
                        className="w-full px-3 py-2 text-xs border border-[#D1D1D1] rounded"
                      />
                      <div className="flex justify-end gap-2">
                        <button
                          type="button"
                          onClick={() => void handleSendSgwMail(true)}
                          className="px-3 py-1.5 border border-[#D1D1D1] rounded text-xs font-medium"
                        >
                          임시저장
                        </button>
                        <button
                          type="button"
                          onClick={() => void handleSendSgwMail(false)}
                          className="px-4 py-1.5 bg-[#0F6CBD] text-white rounded text-xs font-semibold"
                        >
                          보내기
                        </button>
                      </div>
                    </div>
                  ) : selectedMail ? (
                    <div className="space-y-3">
                      <div className="flex items-start justify-between gap-2 border-b border-[#EDEBE9] pb-2.5">
                        <div>
                          <h3 className="text-sm font-bold">{selectedMail.subject}</h3>
                          <p className="text-[11px] text-[#616161] mt-0.5">
                            보낸사람: {selectedMail.senderName} ({selectedMail.senderSub}) · 스레드 {selectedMail.threadId}
                          </p>
                        </div>
                        <div className="flex items-center gap-1.5 shrink-0">
                          <button
                            type="button"
                            onClick={() => void handleToggleRead(selectedMail)}
                            className="px-2 py-1 text-[11px] border border-[#D1D1D1] rounded hover:bg-[#F5F5F5]"
                          >
                            {selectedMail.isRead ? '안읽음 표시' : '읽음 표시'}
                          </button>
                          <button
                            type="button"
                            onClick={() => void handleDeleteMail(selectedMail, selectedMail.folder === 'trash')}
                            className="px-2 py-1 text-[11px] border border-[#D13438]/40 text-[#A4262C] rounded hover:bg-[#FDF3F4]"
                          >
                            {selectedMail.folder === 'trash' ? '영구 삭제' : '휴지통'}
                          </button>
                        </div>
                      </div>
                      <div className="text-xs text-[#242424] whitespace-pre-wrap leading-relaxed min-h-[140px]">
                        {selectedMail.body}
                      </div>
                      <div className="pt-2 border-t border-[#EDEBE9] flex flex-wrap items-center gap-2">
                        <span className="text-[11px] font-semibold text-[#616161]">
                          비공개 AI 보조 (기본 꺼짐 · 확인 후 반영):
                        </span>
                        <button
                          type="button"
                          onClick={() => void handleRunSgwPrivateAi('summarize')}
                          className="px-2.5 py-1 text-[11px] border border-[#D1D1D1] rounded hover:bg-[#F5F5F5]"
                        >
                          요약·할 일
                        </button>
                        <button
                          type="button"
                          onClick={() => void handleRunSgwPrivateAi('reply_draft')}
                          className="px-2.5 py-1 text-[11px] border border-[#D1D1D1] rounded hover:bg-[#F5F5F5]"
                        >
                          답장 초안 (자동발송 금지)
                        </button>
                        <button
                          type="button"
                          onClick={() => void handleRunSgwPrivateAi('classify_and_rules')}
                          className="px-2.5 py-1 text-[11px] border border-[#D1D1D1] rounded hover:bg-[#F5F5F5]"
                        >
                          분류·규칙 제안
                        </button>
                      </div>
                      {aiAssistOutput && (
                        <pre className="p-2.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded text-[11px] overflow-x-auto">
                          {aiAssistOutput}
                        </pre>
                      )}
                    </div>
                  ) : (
                    <div className="h-64 flex flex-col items-center justify-center text-xs text-[#8A8886] gap-2">
                      <Inbox className="w-7 h-7 opacity-40" />
                      <span>왼쪽 목록에서 SGW 사내메일을 선택하세요.</span>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 2: AI Preferences (Default OFF) + Marine/Electrical Glossary + Rules */}
        {activeTab === 'ai-glossary' && (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <section className="bg-white border border-[#E0E0E0] rounded-lg p-4 space-y-3">
              <h2 className="text-sm font-bold flex items-center gap-2">
                <Sliders className="w-4 h-4 text-[#0F6CBD]" /> 사용자 AI 맞춤설정 (기본값 꺼짐 + 사용자 확인 후 반영)
              </h2>
              {!prefs ? (
                <p className="text-xs text-[#616161]">SGW SSO 로그인 후 사용자별 AI 설정을 편집할 수 있습니다.</p>
              ) : (
                <div className="space-y-3 text-xs">
                  <label className="flex items-center justify-between p-2.5 rounded border border-[#E0E0E0] bg-[#FAF9F8]">
                    <div>
                      <div className="font-semibold">사내 비공개 AI 보조 켜기</div>
                      <div className="text-[11px] text-[#616161]">꺼짐 상태에서는 AI 엔드포인트 호출 0회 보장</div>
                    </div>
                    <input
                      type="checkbox"
                      checked={prefs.aiEnabled}
                      onChange={(e) => void handleSavePreferences({ aiEnabled: e.target.checked })}
                    />
                  </label>
                  <div className="grid grid-cols-3 gap-2">
                    <label className="space-y-1">
                      <span className="text-[11px] text-[#616161]">어조</span>
                      <select
                        value={prefs.tone}
                        onChange={(e) =>
                          void handleSavePreferences({ tone: e.target.value as 'formal' | 'concise' })
                        }
                        className="w-full px-2 py-1.5 border border-[#D1D1D1] rounded"
                      >
                        <option value="formal">격식체</option>
                        <option value="concise">간결체</option>
                      </select>
                    </label>
                    <label className="space-y-1">
                      <span className="text-[11px] text-[#616161]">길이</span>
                      <select
                        value={prefs.length}
                        onChange={(e) =>
                          void handleSavePreferences({
                            length: e.target.value as 'short' | 'medium' | 'detailed',
                          })
                        }
                        className="w-full px-2 py-1.5 border border-[#D1D1D1] rounded"
                      >
                        <option value="short">짧게</option>
                        <option value="medium">보통</option>
                        <option value="detailed">상세</option>
                      </select>
                    </label>
                    <label className="space-y-1">
                      <span className="text-[11px] text-[#616161]">언어</span>
                      <select
                        value={prefs.language}
                        onChange={(e) =>
                          void handleSavePreferences({ language: e.target.value as 'ko' | 'en' })
                        }
                        className="w-full px-2 py-1.5 border border-[#D1D1D1] rounded"
                      >
                        <option value="ko">한국어 (기본)</option>
                        <option value="en">English</option>
                      </select>
                    </label>
                  </div>
                  <label className="block space-y-1">
                    <span className="text-[11px] text-[#616161]">답장 기본 서명</span>
                    <textarea
                      rows={3}
                      value={prefs.signature}
                      onChange={(e) => setPrefs({ ...prefs, signature: e.target.value })}
                      onBlur={() => void handleSavePreferences({ signature: prefs.signature })}
                      placeholder="예: 감사합니다. 전장설계팀 드림"
                      className="w-full px-2.5 py-1.5 border border-[#D1D1D1] rounded"
                    />
                  </label>
                  <label className="flex items-center gap-2 text-[11px] text-[#424242]">
                    <input
                      type="checkbox"
                      checked={prefs.allowTraining}
                      onChange={(e) => void handleSavePreferences({ allowTraining: e.target.checked })}
                    />
                    사내 비공개 모델 품질 개선 학습 사용 허용 (기본 꺼짐)
                  </label>
                </div>
              )}

              {/* User Rules */}
              <div className="pt-3 border-t border-[#EDEBE9] space-y-2">
                <h3 className="text-xs font-bold">자동 분류 규칙 (사용자 확인 등록분만 적용)</h3>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={newRuleName}
                    onChange={(e) => setNewRuleName(e.target.value)}
                    placeholder="규칙 이름"
                    className="flex-1 px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
                  />
                  <input
                    type="text"
                    value={newRuleKeyword}
                    onChange={(e) => setNewRuleKeyword(e.target.value)}
                    placeholder="제목 키워드"
                    className="flex-1 px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
                  />
                  <button
                    type="button"
                    onClick={() => void handleAddRule()}
                    className="px-3 py-1.5 bg-[#0F6CBD] text-white text-xs font-semibold rounded"
                  >
                    추가
                  </button>
                </div>
                <div className="space-y-1">
                  {rules.map((r) => (
                    <div
                      key={r.id}
                      className="flex items-center justify-between px-2.5 py-1.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded text-xs"
                    >
                      <span>
                        <strong>{r.name}</strong> · 키워드: <code>{r.matchKeyword}</code> → {r.targetCategory}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            </section>

            {/* Marine & Electrical Domain Glossary */}
            <section className="bg-white border border-[#E0E0E0] rounded-lg p-4 space-y-3">
              <h2 className="text-sm font-bold flex items-center gap-2">
                <BookMarked className="w-4 h-4 text-[#0F6CBD]" /> 조선 · 선박 전장 업무 용어 사전 (사용자/관리자 편집)
              </h2>
              <p className="text-xs text-[#616161]">
                특정 호선·회사명을 코드에 하드코딩하지 않으며, 요청자가 제공하는 용어 목록을 등록해 번역·요약·초안에
                반영합니다.
              </p>
              <div className="grid grid-cols-3 gap-2">
                <input
                  type="text"
                  value={newTerm}
                  onChange={(e) => setNewTerm(e.target.value)}
                  placeholder="용어 (예: MSBD)"
                  className="px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
                />
                <input
                  type="text"
                  value={newCat}
                  onChange={(e) => setNewCat(e.target.value)}
                  placeholder="분류"
                  className="px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
                />
                <button
                  type="button"
                  onClick={() => void handleAddGlossary()}
                  className="px-3 py-1.5 bg-[#0F6CBD] text-white text-xs font-semibold rounded inline-flex items-center justify-center gap-1"
                >
                  <Plus className="w-3.5 h-3.5" /> 용어 등록
                </button>
              </div>
              <input
                type="text"
                value={newDef}
                onChange={(e) => setNewDef(e.target.value)}
                placeholder="용어 설명 (예: Main Switchboard 주배전반)"
                className="w-full px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
              />
              <div className="divide-y divide-[#EDEBE9] border border-[#E0E0E0] rounded max-h-72 overflow-y-auto">
                {glossary.length === 0 ? (
                  <div className="p-6 text-center text-xs text-[#8A8886]">
                    등록된 용어가 없습니다. 요청자 제공 조선·전장 용어를 추가하세요.
                  </div>
                ) : (
                  glossary.map((g) => (
                    <div key={g.id} className="p-2.5 flex items-center justify-between gap-2 text-xs">
                      <div>
                        <span className="font-bold text-[#0F6CBD]">{g.term}</span>{' '}
                        <span className="text-[10px] text-[#616161]">[{g.category}]</span>
                        <div className="text-[#323130] mt-0.5">{g.definition}</div>
                      </div>
                      <button
                        type="button"
                        onClick={async () => {
                          await fetch(`/api/sgw/glossary/${encodeURIComponent(g.id)}`, { method: 'DELETE' });
                          await loadSgwMailData();
                        }}
                        className="p-1 text-[#616161] hover:text-[#D13438]"
                        title="삭제"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))
                )}
              </div>
            </section>
          </div>
        )}

        {/* TAB 3: SEDMS Mail Hub (20261120-MAILHUB) Deterministic Matcher & Review Queue */}
        {activeTab === 'sedms-hub' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
            <section className="lg:col-span-6 bg-white border border-[#E0E0E0] rounded-lg p-4 space-y-3">
              <h2 className="text-sm font-bold flex items-center gap-2">
                <Network className="w-4 h-4 text-[#0F6CBD]" /> SEDMS 5순위 결정적 매칭 엔진 검증기 (0-LLM · S1~S8)
              </h2>
              <p className="text-xs text-[#616161]">
                규칙 1(In-Reply-To) → 규칙 2(<code>drawingNoKey</code>) → 규칙 3(DB 프로젝트 토큰) → 규칙 4(등록 업체 후보)
                → 규칙 5/충돌(확인 대기열) 판정을 외부 AI 없이 100% 결정적으로 실행합니다.
              </p>
              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  value={sedmsTenant}
                  onChange={(e) => setSedmsTenant(e.target.value)}
                  placeholder="tenant_id"
                  className="px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded font-mono"
                />
                <input
                  type="text"
                  value={previewFrom}
                  onChange={(e) => setPreviewFrom(e.target.value)}
                  placeholder="발신자 이메일 (예: sender@example.test)"
                  className="px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
                />
              </div>
              <input
                type="text"
                value={previewSubject}
                onChange={(e) => setPreviewSubject(e.target.value)}
                placeholder="테스트 메일 제목 (도면번호 또는 프로젝트 토큰 포함)"
                className="w-full px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
              />
              <textarea
                rows={4}
                value={previewBody}
                onChange={(e) => setPreviewBody(e.target.value)}
                placeholder="테스트 메일 본문..."
                className="w-full px-2.5 py-1.5 text-xs border border-[#D1D1D1] rounded"
              />
              <button
                type="button"
                onClick={() => void handlePreviewSedmsMatch()}
                className="px-4 py-1.5 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-semibold rounded"
              >
                결정적 5순위 매칭 판정 실행
              </button>
              {previewResult && (
                <pre className="p-2.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded text-[11px] max-h-60 overflow-auto">
                  {previewResult}
                </pre>
              )}
            </section>

            <section className="lg:col-span-6 bg-white border border-[#E0E0E0] rounded-lg p-4 space-y-3">
              <div className="flex items-center justify-between">
                <h2 className="text-sm font-bold flex items-center gap-2">
                  <ShieldAlert className="w-4 h-4 text-[#C57B00]" /> 사람 확인 대기열 (mail_review_queue · 모달 금지)
                </h2>
                <button
                  type="button"
                  onClick={() => void loadSedmsAndHandoff()}
                  className="p-1 text-[#616161] hover:text-[#0F6CBD]"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                </button>
              </div>
              <p className="text-xs text-[#616161]">
                S6 불변식: 자동 매칭은 개정·봉인·배포 상태를 변경하지 않으며, 충돌·미분류 건은 이 그리드에서 담당자가
                확정/거부합니다.
              </p>
              {reviewItems.length === 0 ? (
                <div className="p-8 text-center text-xs text-[#8A8886] border border-dashed border-[#E0E0E0] rounded">
                  현재 대기 중인 미분류·충돌 메일이 없습니다.
                </div>
              ) : (
                <div className="overflow-x-auto border border-[#E0E0E0] rounded">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-[#FAF9F8] border-b border-[#E0E0E0] text-[#616161]">
                      <tr>
                        <th className="p-2">제목</th>
                        <th className="p-2">후보 · 규칙</th>
                        <th className="p-2">상태</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#EDEBE9]">
                      {reviewItems.map((item) => (
                        <tr key={item.id}>
                          <td className="p-2 font-medium">{item.subject}</td>
                          <td className="p-2 text-[11px] text-[#616161]">
                            {item.candidates.length > 0
                              ? item.candidates.map((c) => `${c.projectId}(R${c.rule}/${c.score})`).join(', ')
                              : '미분류'}
                          </td>
                          <td className="p-2 font-mono">{item.state}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </div>
        )}

        {/* TAB 4: SGW `internal_mail` Migration Tool (Dry-run -> Approve -> Execute -> Rollback) */}
        {activeTab === 'migration' && (
          <div className="bg-white border border-[#E0E0E0] rounded-lg p-5 space-y-4">
            <div>
              <h2 className="text-sm font-bold flex items-center gap-2">
                <FolderSync className="w-4 h-4 text-[#0F6CBD]" /> SGW D1 <code>internal_mail</code> 이전 도구 (드라이런 → 승인 실행 → ID 대조 → 롤백)
              </h2>
              <p className="text-xs text-[#616161] mt-1">
                SGW D1 <code>internal_mail</code> JSON 덤프를 붙여넣고 먼저 드라이런으로 건수·중복·인코딩 오류를 점검한
                뒤 승인 실행하세요. 대조 리포트는 실명 없이 <code>sub</code>/<code>ID</code>만 출력합니다.
              </p>
            </div>
            <textarea
              rows={6}
              value={migrationJson}
              onChange={(e) => setMigrationJson(e.target.value)}
              className="w-full p-3 font-mono text-xs border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
              placeholder='[{"id":"1","senderId":"u01","senderName":"User01","recipientId":"u02","recipientName":"User02","subject":"테스트","body":"본문","isRead":0,"createdAt":"2026-11-20T00:00:00Z"}]'
            />
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={() => void handleRunMigrationDryRun()}
                className="px-3.5 py-2 border border-[#0F6CBD] text-[#0F6CBD] hover:bg-[#EBF3FC] text-xs font-semibold rounded"
              >
                1. 드라이런 검증
              </button>
              <button
                type="button"
                disabled={!dryRunReport || dryRunReport.validRows === 0}
                onClick={() => void handleExecuteMigration()}
                className="px-3.5 py-2 bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-semibold rounded disabled:opacity-40 inline-flex items-center gap-1.5"
              >
                <CheckCircle2 className="w-3.5 h-3.5" /> 2. 승인 후 이전 실행
              </button>
              <button
                type="button"
                onClick={() => void handleRollbackMigration()}
                className="px-3.5 py-2 border border-[#D13438]/50 text-[#A4262C] hover:bg-[#FDF3F4] text-xs font-semibold rounded inline-flex items-center gap-1.5"
              >
                <RotateCcw className="w-3.5 h-3.5" /> 이전 데이터 롤백
              </button>
            </div>

            {dryRunReport && (
              <div className="p-3 bg-[#FAF9F8] border border-[#E0E0E0] rounded space-y-1.5 text-xs">
                <div className="font-bold">드라이런 결과 요약</div>
                <div className="font-mono text-[11px]">
                  원본 {dryRunReport.totalSourceRows}건 · 유효 {dryRunReport.validRows}건 · 중복 건너뜀{' '}
                  {dryRunReport.skippedDuplicateIds.length}건 · 누락 오류 {dryRunReport.missingFieldErrors.length}건 ·
                  인코딩 경고 {dryRunReport.encodingWarnings.length}건
                </div>
              </div>
            )}

            {execReport && (
              <div className="p-3 bg-[#F3FAF3] border border-[#107C10]/30 rounded space-y-1.5 text-xs text-[#107C10]">
                <div className="font-bold flex items-center gap-1.5">
                  <Check className="w-4 h-4" /> 이전 완료 및 ID 기준 대조 결과
                </div>
                <div className="font-mono text-[11px]">
                  이전 {execReport.migratedCount}건 / 건너뜀 {execReport.skippedCount}건 / 원본 {execReport.sourceTotal}건
                  (건수 일치: {String(execReport.verifiedCountMatch)})
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 5: SEDMS & SGW Handoff Package */}
        {activeTab === 'handoff' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
            <section className="lg:col-span-5 bg-white border border-[#E0E0E0] rounded-lg p-4 space-y-3">
              <h2 className="text-sm font-bold flex items-center gap-2">
                <FileCode2 className="w-4 h-4 text-[#0F6CBD]" /> SEDMS(`E:/code-project/SEDMS_G`) · SGW 이관 핸드오프 파일
              </h2>
              <p className="text-xs text-[#616161]">
                이 워크스페이스 밖(SEDMS D1 마이그레이션·NAS/R2 워커·SGW 관문 라우트)에서 실행할 정본 산출물 일체입니다.
              </p>
              <div className="divide-y divide-[#EDEBE9] border border-[#E0E0E0] rounded">
                {handoffFiles.map((f) => (
                  <div key={f.name} className="p-2.5 flex items-center justify-between gap-2 text-xs">
                    <button
                      type="button"
                      onClick={() => void handlePreviewHandoffFile(f.name)}
                      className="text-left min-w-0 flex-1 hover:text-[#0F6CBD]"
                    >
                      <div className="font-semibold truncate">{f.name}</div>
                      <div className="text-[10px] text-[#616161] font-mono truncate">{f.relPath}</div>
                    </button>
                    <a
                      href={`/api/sedms/handoff/file/${encodeURIComponent(f.name)}?download=1`}
                      className="px-2 py-1 border border-[#D1D1D1] rounded text-[11px] hover:bg-[#F5F5F5] inline-flex items-center gap-1 shrink-0"
                    >
                      <Download className="w-3 h-3" /> 받기
                    </a>
                  </div>
                ))}
              </div>
            </section>

            <section className="lg:col-span-7 bg-white border border-[#E0E0E0] rounded-lg p-4 space-y-2">
              <div className="flex items-center justify-between">
                <h3 className="text-xs font-bold text-[#242424]">
                  {previewHandoffName ? `미리보기: ${previewHandoffName}` : '왼쪽에서 핸드오프 파일을 선택해 미리보기'}
                </h3>
                {previewHandoffName && (
                  <span className="text-[10px] text-[#616161] font-mono">UTF-8 정본</span>
                )}
              </div>
              <pre className="p-3 bg-[#FAF9F8] border border-[#E0E0E0] rounded text-[11px] leading-relaxed h-[460px] overflow-auto whitespace-pre-wrap">
                {previewHandoffContent ||
                  'SEDMS 마이그레이션 SQL(0106/0107/0108), TLA+ 형식 명세(MailHubIngest.tla), SEDMS/SGW 이관 핸드오프 문서를 즉시 확인하고 내려받을 수 있습니다.'}
              </pre>
            </section>
          </div>
        )}
      </div>
    </div>
  );
};
