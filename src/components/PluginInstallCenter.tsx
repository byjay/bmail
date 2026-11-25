import React, { useEffect, useMemo, useState } from 'react';
import {
  Building2,
  Check,
  ClipboardCopy,
  Database,
  Download,
  ExternalLink,
  FolderCog,
  Laptop,
  RefreshCw,
  ShieldCheck,
  Sparkles,
} from 'lucide-react';
import type { ProviderStatus } from '../types/mail';
import {
  buildOutlookManifestXml,
  buildUnifiedManifestJson,
  normalizeBaseUrl,
} from '../utils/outlookManifest';

interface PluginInstallCenterProps {
  defaultAppUrl: string;
  oauthCallbackUrl: string;
  providers: ProviderStatus[];
  microsoftConfigured: boolean;
  onOpenPstModal: () => void;
}

type DeployTrack = 'pst-mailcraft' | 'local' | 'enterprise-m365' | 'enterprise-shared-folder';
type ManifestFormat = 'xml' | 'json';

export const PluginInstallCenter: React.FC<PluginInstallCenterProps> = ({
  defaultAppUrl,
  oauthCallbackUrl,
  providers,
  microsoftConfigured,
  onOpenPstModal,
}) => {
  const [hostUrl, setHostUrl] = useState(defaultAppUrl);
  const [addinName, setAddinName] = useState('MailLog Studio AI');
  const [sharedFolderPath, setSharedFolderPath] = useState('\\\\fileserver\\outlook-addins\\maillog');
  const [deployTrack, setDeployTrack] = useState<DeployTrack>('pst-mailcraft');
  const [manifestFormat, setManifestFormat] = useState<ManifestFormat>('xml');
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [updateStatus, setUpdateStatus] = useState<{
    checking: boolean;
    installing: boolean;
    currentVersion: string;
    latestVersion: string;
    updateAvailable: boolean;
    releaseNotes?: string;
    error?: string;
    message?: string;
  }>({
    checking: false,
    installing: false,
    currentVersion: '1.2.0',
    latestVersion: '1.2.0',
    updateAvailable: false,
  });

  const isDesktopApp =
    typeof window !== 'undefined' &&
    Boolean((window as { bmailDesktop?: { isDesktopExe?: boolean } }).bmailDesktop?.isDesktopExe);

  const handleCheckDesktopUpdate = async () => {
    setUpdateStatus((prev) => ({ ...prev, checking: true, error: undefined, message: undefined }));
    try {
      const desktopBridge = (
        window as {
          bmailDesktop?: {
            checkUpdate?: () => Promise<{
              currentVersion: string;
              latestVersion: string;
              updateAvailable: boolean;
              releaseNotes?: string;
              error?: string;
            }>;
          };
        }
      ).bmailDesktop;
      const res = desktopBridge?.checkUpdate
        ? await desktopBridge.checkUpdate()
        : await fetch('/api/desktop/update-check').then((r) => r.json());
      setUpdateStatus((prev) => ({
        ...prev,
        checking: false,
        currentVersion: res.currentVersion || prev.currentVersion,
        latestVersion: res.latestVersion || prev.latestVersion,
        updateAvailable: Boolean(res.updateAvailable),
        releaseNotes: res.releaseNotes || '',
        error: res.error,
        message: res.error
          ? undefined
          : res.updateAvailable
            ? `새 버전 v${res.latestVersion} 업데이트를 바로 적용할 수 있습니다.`
            : `현재 최신 버전(v${res.currentVersion || prev.currentVersion})을 사용 중입니다.`,
      }));
    } catch (err) {
      setUpdateStatus((prev) => ({
        ...prev,
        checking: false,
        error: err instanceof Error ? err.message : '업데이트 확인 실패',
      }));
    }
  };

  const handleRunDesktopInstallUpdate = async () => {
    const installer = (
      window as { bmailDesktop?: { installUpdate?: () => Promise<{ started?: boolean; error?: string }> } }
    ).bmailDesktop?.installUpdate;
    if (!installer) return;
    setUpdateStatus((prev) => ({ ...prev, installing: true, error: undefined }));
    const res = await installer();
    if (!res.started) {
      setUpdateStatus((prev) => ({
        ...prev,
        installing: false,
        error: res.error || '자동 업데이트를 시작할 수 없습니다.',
      }));
    }
  };

  useEffect(() => {
    const getInfo = (
      window as {
        bmailDesktop?: { getAppInfo?: () => Promise<{ version?: string }> };
      }
    ).bmailDesktop?.getAppInfo;
    if (getInfo) {
      void getInfo()
        .then((info) => {
          if (info?.version) {
            setUpdateStatus((prev) => ({ ...prev, currentVersion: info.version! }));
          }
        })
        .catch(() => {});
    }
  }, []);

  const normalizedUrl = useMemo(() => normalizeBaseUrl(hostUrl || defaultAppUrl), [hostUrl, defaultAppUrl]);

  const manifestXml = useMemo(
    () => buildOutlookManifestXml({ baseUrl: normalizedUrl, displayName: addinName }),
    [normalizedUrl, addinName]
  );

  const manifestJson = useMemo(
    () => buildUnifiedManifestJson({ baseUrl: normalizedUrl, displayName: addinName }),
    [normalizedUrl, addinName]
  );

  const liveManifestDownloadHref = `/manifest.xml?download=1&baseUrl=${encodeURIComponent(
    normalizedUrl
  )}&name=${encodeURIComponent(addinName)}`;

  const liveUnifiedDownloadHref = `/manifest.json?download=1&baseUrl=${encodeURIComponent(
    normalizedUrl
  )}&name=${encodeURIComponent(addinName)}`;

  const liveRegDownloadHref = `/api/outlook/trusted-catalog.reg?path=${encodeURIComponent(
    sharedFolderPath
  )}`;

  const manifestDirectUrl = `${normalizedUrl}/manifest.xml`;

  const copyText = async (id: string, text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedId(id);
      window.setTimeout(() => setCopiedId((prev) => (prev === id ? null : prev)), 2000);
    } catch {
      // Clipboard blocked in non-secure context
    }
  };

  return (
    <div className="flex-1 overflow-y-auto bg-slate-50">
      <div className="max-w-6xl mx-auto px-6 py-8 space-y-8">
        {/* Intro Header */}
        <header className="flex flex-col md:flex-row md:items-end justify-between gap-4 pb-6 border-b border-slate-200">
          <div className="space-y-1.5">
            <h1 className="text-2xl font-bold text-slate-900 tracking-tight">
              Outlook 플러그인 설치 및 사내 배포 센터
            </h1>
            <p className="text-sm text-slate-600 max-w-2xl leading-relaxed">
              로컬 인증서 오류나 보안 정책 제약 없이, HTTPS로 즉시 호스팅되는 표준 Office.js Add-in 매니페스트를 통해 내 PC Outlook과 회사 전 직원 Outlook 리본에 바로 아이콘을 추가할 수 있습니다.
            </p>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <a
              href={liveManifestDownloadHref}
              className="inline-flex items-center gap-2 px-4 py-2.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg transition-colors whitespace-nowrap"
            >
              <Download className="w-4 h-4" />
              manifest.xml 다운로드
            </a>
            <a
              href="/taskpane"
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-1.5 px-3.5 py-2.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors whitespace-nowrap"
            >
              Taskpane 단독 열기
              <ExternalLink className="w-3.5 h-3.5" />
            </a>
          </div>
        </header>

        {/* Section 00: Company-Wide 1-Click Installer & Auto-Updater */}
        <section className="bg-gradient-to-br from-[#F0F7FF] via-white to-[#F8FBFF] border border-[#CFE4FA] rounded-xl p-6 space-y-4 shadow-sm">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="space-y-1">
              <div className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-[#EBF3FC] text-[#0F6CBD] text-[11px] font-semibold">
                <Sparkles className="w-3.5 h-3.5" />
                전 직원 원클릭 PC 설치 & 자동 업데이트
              </div>
              <h2 className="text-base font-bold text-slate-900">
                00. 회사 모든 직원용 딸깍 1-클릭 PC 설치 & 최신 버전 자동 업데이트 (Windows)
              </h2>
              <p className="text-xs text-slate-600 max-w-3xl leading-relaxed">
                관리자 권한(UAC) 없이 클릭 한 번으로 설치되며, 실행 즉시 내 PC의 기존 Outlook 계정·서명·PST/OST 파일을 자동 인식해 연결합니다. 재설치·업데이트 시에도 기존 로컬 메일 DB와 설정은 그대로 보존됩니다.
              </p>
            </div>
            <div className="flex flex-wrap items-center gap-2 shrink-0">
              <a
                href="https://mail.seastar.work/releases/Bmail-Outlook-AI-Setup.exe"
                className="inline-flex items-center gap-1.5 px-4 py-2.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg transition-colors whitespace-nowrap"
              >
                <Download className="w-4 h-4" />
                원클릭 설치 파일 (.exe)
              </a>
              <a
                href="https://mail.seastar.work/install.bat"
                className="inline-flex items-center gap-1.5 px-3.5 py-2.5 text-xs font-semibold text-[#0F6CBD] bg-white border border-[#0F6CBD]/40 hover:bg-[#EBF3FC] rounded-lg transition-colors whitespace-nowrap"
              >
                <Laptop className="w-4 h-4" />
                더블클릭 설치/업데이트 런처 (.bat)
              </a>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="bg-white border border-slate-200 rounded-lg p-3.5 space-y-2">
              <div className="text-xs font-semibold text-slate-800">
                사내 공지/채팅 복사 배포용 1줄 설치·업데이트 명령어 (PowerShell)
              </div>
              <div className="flex items-center justify-between gap-2 bg-slate-50 border border-slate-200 rounded px-3 py-2">
                <code className="text-[11px] font-mono text-slate-800 truncate">
                  irm https://mail.seastar.work/install.ps1 | iex
                </code>
                <button
                  type="button"
                  onClick={() => copyText('one-line-ps1', 'irm https://mail.seastar.work/install.ps1 | iex')}
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-100 rounded whitespace-nowrap"
                >
                  {copiedId === 'one-line-ps1' ? (
                    <>
                      <Check className="w-3 h-3 text-emerald-600" />
                      복사됨
                    </>
                  ) : (
                    <>
                      <ClipboardCopy className="w-3 h-3" />
                      명령어 복사
                    </>
                  )}
                </button>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-lg p-3.5 flex flex-col justify-between gap-2">
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs font-semibold text-slate-800">
                  버전 상태 및 인앱 1클릭 업데이트 (현재 v{updateStatus.currentVersion})
                </span>
                <button
                  type="button"
                  disabled={updateStatus.checking}
                  onClick={() => void handleCheckDesktopUpdate()}
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] rounded whitespace-nowrap"
                >
                  <RefreshCw className={`w-3 h-3 ${updateStatus.checking ? 'animate-spin' : ''}`} />
                  업데이트 확인
                </button>
              </div>
              <div className="text-[11px] text-slate-600 flex flex-wrap items-center justify-between gap-2">
                <span>
                  {updateStatus.error
                    ? `확인 오류: ${updateStatus.error}`
                    : updateStatus.message || '최신 배포 버전과 SHA-256 무결성을 검사합니다.'}
                </span>
                {updateStatus.updateAvailable && isDesktopApp && (
                  <button
                    type="button"
                    disabled={updateStatus.installing}
                    onClick={() => void handleRunDesktopInstallUpdate()}
                    className="px-3 py-1 text-[11px] font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded whitespace-nowrap"
                  >
                    {updateStatus.installing
                      ? '다운로드 및 재설치 중...'
                      : `v${updateStatus.latestVersion} 자동 업데이트`}
                  </button>
                )}
              </div>
            </div>
          </div>
        </section>

        {/* Top Configuration & Icon Preview Row */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-stretch">
          <section className="lg:col-span-2 bg-white border border-slate-200 rounded-xl p-6 space-y-4">
            <h2 className="text-sm font-bold text-slate-900">
              01. 플러그인 호스트 URL 및 리본 표시 이름 설정
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label htmlFor="plugin-host-url" className="block text-xs font-medium text-slate-700 mb-1.5">
                  플러그인 호스팅 HTTPS 주소
                </label>
                <input
                  id="plugin-host-url"
                  type="url"
                  value={hostUrl}
                  onChange={(e) => setHostUrl(e.target.value)}
                  className="w-full px-3 py-2 text-xs font-mono bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#0F6CBD]/30 focus:border-[#0F6CBD]"
                />
                <p className="mt-1 text-[11px] text-slate-500">
                  기본값은 현재 배포된 HTTPS 주소입니다. 사내 자체 서버에 올릴 경우 해당 HTTPS 도메인으로 변경하세요.
                </p>
              </div>
              <div>
                <label htmlFor="plugin-display-name" className="block text-xs font-medium text-slate-700 mb-1.5">
                  Outlook 리본 표시 이름
                </label>
                <input
                  id="plugin-display-name"
                  type="text"
                  value={addinName}
                  onChange={(e) => setAddinName(e.target.value)}
                  className="w-full px-3 py-2 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#0F6CBD]/30 focus:border-[#0F6CBD]"
                />
                <p className="mt-1 text-[11px] text-slate-500">
                  Outlook 메일 읽기/작성 리본 상단 탭에 표시되는 버튼 이름입니다.
                </p>
              </div>
            </div>

            <div className="pt-2 flex flex-wrap items-center justify-between gap-3 bg-slate-50 border border-slate-200 rounded-lg px-4 py-3">
              <div className="min-w-0">
                <div className="text-[11px] text-slate-500">URL 사이드로드용 매니페스트 주소</div>
                <div className="text-xs font-mono text-slate-900 truncate">{manifestDirectUrl}</div>
              </div>
              <button
                type="button"
                onClick={() => copyText('manifest-url', manifestDirectUrl)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-100 rounded-md whitespace-nowrap"
              >
                {copiedId === 'manifest-url' ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    복사됨
                  </>
                ) : (
                  <>
                    <ClipboardCopy className="w-3.5 h-3.5" />
                    URL 복사
                  </>
                )}
              </button>
            </div>
          </section>

          {/* Real PNG Ribbon Icon Asset Verification */}
          <section className="bg-white border border-slate-200 rounded-xl p-6 flex flex-col justify-between space-y-4">
            <div className="space-y-1">
              <h2 className="text-sm font-bold text-slate-900">
                02. Outlook 리본 아이콘 에셋 (PNG)
              </h2>
              <p className="text-xs text-slate-500 leading-relaxed">
                Outlook 스키마가 요구하는 16·32·64·80·128px PNG 아이콘을 서버에서 실시간 제공합니다.
              </p>
            </div>
            <div className="flex items-end justify-between gap-2 py-3 px-4 bg-slate-50 border border-slate-200 rounded-lg">
              {[16, 32, 64, 80].map((size) => (
                <div key={size} className="flex flex-col items-center gap-1.5">
                  <img
                    src={`/assets/icon-${size}.png`}
                    alt={`Outlook 리본 ${size}px 아이콘`}
                    width={size > 48 ? 48 : size}
                    height={size > 48 ? 48 : size}
                    className="rounded"
                  />
                  <span className="text-[11px] font-mono tabular-nums text-slate-500">
                    {size}px
                  </span>
                </div>
              ))}
            </div>
            <p className="text-[11px] text-slate-500">
              Mailbox 1.15 기준 작업 창 고정(Pinning), 메일 미선택 실행(<span className="font-mono">SupportsNoItemContext</span>), 다중 메일 선택(<span className="font-mono">SupportsMultiSelect</span>), 공유 사서함 지원이 매니페스트에 기본 포함되어 있습니다.
            </p>
          </section>
        </div>

        {/* 3 Deployment Tracks */}
        <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
            <h2 className="text-sm font-bold text-slate-900">
              03. 환경별 설치 방법 (로컬 PC · 회사 전 직원 · 사내 공유 폴더)
            </h2>
            <div
              role="tablist"
              aria-label="설치 방식 선택"
              className="inline-flex flex-wrap items-center p-1 bg-slate-100 rounded-lg"
            >
              <button
                type="button"
                role="tab"
                aria-selected={deployTrack === 'pst-mailcraft'}
                onClick={() => setDeployTrack('pst-mailcraft')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                  deployTrack === 'pst-mailcraft'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Database className="w-3.5 h-3.5" />
                MailCraft · PST 직접 연동
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={deployTrack === 'local'}
                onClick={() => setDeployTrack('local')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                  deployTrack === 'local'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Laptop className="w-3.5 h-3.5" />
                내 PC Outlook 플러그인
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={deployTrack === 'enterprise-m365'}
                onClick={() => setDeployTrack('enterprise-m365')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                  deployTrack === 'enterprise-m365'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Building2 className="w-3.5 h-3.5" />
                회사 전 직원 배포 (M365)
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={deployTrack === 'enterprise-shared-folder'}
                onClick={() => setDeployTrack('enterprise-shared-folder')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                  deployTrack === 'enterprise-shared-folder'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <FolderCog className="w-3.5 h-3.5" />
                사내 공유 폴더 (.reg)
              </button>
            </div>
          </div>

          <div className="p-6">
            {deployTrack === 'pst-mailcraft' && (
              <div className="space-y-4 text-xs text-slate-700">
                <ol className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  <li className="space-y-2">
                    <div className="font-bold text-slate-900">1단계. 기존 Outlook .pst 그대로 사용</div>
                    <p className="leading-relaxed text-slate-600">
                      회사 보안 정책이나 구버전 Outlook 제약으로 플러그인 설치가 막혀 있어도, PC에 저장된 기존 <strong>.pst / .ost / .eml / .mbox</strong> 파일을 드래그해 즉시 메일함을 열 수 있습니다.
                    </p>
                  </li>
                  <li className="space-y-2">
                    <div className="font-bold text-slate-900">2단계. MailCraft 오픈소스 클라이언트 병행</div>
                    <p className="leading-relaxed text-slate-600">
                      오픈소스 Classic Outlook 대안인 <strong>MailCraft</strong>(<span className="font-mono text-[#0F6CBD]">mailcraft.work</span>)로 로컬 <span className="font-mono">.pst</span> 메일을 가볍게 열람하면서, 동일 <span className="font-mono">.pst/.eml</span>을 본 워크벤치에 연결해 일자별 업무 일지와 블로그를 자동 생성합니다.
                    </p>
                  </li>
                  <li className="space-y-2">
                    <div className="font-bold text-slate-900">3단계. 폴더 자동 분류 &amp; 보안 파싱</div>
                    <p className="leading-relaxed text-slate-600">
                      업로드된 PST 바이너리는 받은 편지함·보낸 편지함·사용자 폴더로 자동 분류되며, 파싱 완료 즉시 임시 파일이 서버에서 자동 삭제됩니다.
                    </p>
                  </li>
                </ol>
                <div className="pt-2 flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={onOpenPstModal}
                    className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg transition-colors"
                  >
                    <Database className="w-3.5 h-3.5" />
                    지금 내 .pst / .eml 파일 불러오기
                  </button>
                  <a
                    href="https://mailcraft.work/personal/"
                    target="_blank"
                    rel="noreferrer"
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
                  >
                    MailCraft 공식 사이트
                    <ExternalLink className="w-3.5 h-3.5" />
                  </a>
                </div>
              </div>
            )}
            {deployTrack === 'local' && (
              <ol className="grid grid-cols-1 md:grid-cols-3 gap-6 text-xs text-slate-700">
                <li className="space-y-2">
                  <div className="font-bold text-slate-900">1단계. manifest.xml 다운로드</div>
                  <p className="leading-relaxed text-slate-600">
                    상단의 <strong>manifest.xml 다운로드</strong> 버튼을 클릭해 매니페스트 파일을 내 PC에 저장합니다. (이미 HTTPS 호스팅 URL이 자동 반영되어 있어 로컬 SSL 인증서 설치가 필요 없습니다.)
                  </p>
                </li>
                <li className="space-y-2">
                  <div className="font-bold text-slate-900">2단계. Outlook 추가 기능 창 열기</div>
                  <p className="leading-relaxed text-slate-600">
                    브라우저에서 <span className="font-mono text-[#0F6CBD]">https://aka.ms/olksideload</span> 에 접속(또는 Outlook Web → 앱 → 추가 기능 관리)한 뒤 <strong>내 추가 기능 → 사용자 지정 추가 기능 → 파일에서 추가</strong>를 선택합니다.
                  </p>
                </li>
                <li className="space-y-2">
                  <div className="font-bold text-slate-900">3단계. 리본 아이콘 클릭 후 바로 사용</div>
                  <p className="leading-relaxed text-slate-600">
                    설치 후 Outlook(웹, New Outlook, 데스크톱 모두 자동 동기화)에서 메일을 열면 상단 리본에 <strong>MailLog AI</strong> 아이콘이 생성됩니다. 아이콘을 클릭하고 우측 상단 핀 아이콘을 누르면 메일을 바꿀 때마다 자동으로 연동됩니다.
                  </p>
                </li>
              </ol>
            )}

            {deployTrack === 'enterprise-m365' && (
              <div className="space-y-4 text-xs text-slate-700">
                <ol className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  <li className="space-y-2">
                    <div className="font-bold text-slate-900">1단계. M365 관리 센터 접속</div>
                    <p className="leading-relaxed text-slate-600">
                      IT 관리자가 <span className="font-mono text-[#0F6CBD]">admin.microsoft.com</span> 에 로그인 후 <strong>설정 → 통합 앱(Integrated apps) → 사용자 지정 앱 업로드</strong>를 클릭합니다.
                    </p>
                  </li>
                  <li className="space-y-2">
                    <div className="font-bold text-slate-900">2단계. 매니페스트 업로드</div>
                    <p className="leading-relaxed text-slate-600">
                      앱 유형으로 <strong>Office 추가 기능</strong>을 선택하고 다운로드한 <span className="font-mono">manifest.xml</span> (또는 통합 앱용 <span className="font-mono">manifest.json</span>)을 업로드합니다.
                    </p>
                  </li>
                  <li className="space-y-2">
                    <div className="font-bold text-slate-900">3단계. 부서 또는 전체 직원 배포</div>
                    <p className="leading-relaxed text-slate-600">
                      배포 대상에 특정 부서 또는 전체 조직을 지정하면 <strong>직원이 개별 설치할 필요 없이</strong> 모든 직원의 Windows/Mac/Web Outlook 리본에 MailLog AI 아이콘이 자동 배포됩니다.
                    </p>
                  </li>
                </ol>
                <div className="pt-2">
                  <a
                    href={liveUnifiedDownloadHref}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-800 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" />
                    M365 통합 앱용 manifest.json 다운로드
                  </a>
                </div>
              </div>
            )}

            {deployTrack === 'enterprise-shared-folder' && (
              <div className="space-y-4 text-xs text-slate-700">
                <p className="text-slate-600 leading-relaxed">
                  M365 관리 센터 권한이 없거나 Classic Outlook(Windows) 사내망 환경에서는 공유 폴더 카탈로그(Trusted Catalog) 방식으로 1분 만에 직원들에게 배포할 수 있습니다.
                </p>
                <div className="flex flex-col sm:flex-row gap-3 items-stretch sm:items-end">
                  <div className="flex-1">
                    <label htmlFor="shared-folder-path" className="block text-xs font-medium text-slate-700 mb-1">
                      사내 공유 폴더 경로 (manifest.xml을 저장할 UNC 또는 로컬 경로)
                    </label>
                    <input
                      id="shared-folder-path"
                      type="text"
                      value={sharedFolderPath}
                      onChange={(e) => setSharedFolderPath(e.target.value)}
                      className="w-full px-3 py-2 text-xs font-mono bg-white border border-slate-300 rounded-lg"
                    />
                  </div>
                  <a
                    href={liveRegDownloadHref}
                    className="inline-flex items-center justify-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg whitespace-nowrap"
                  >
                    <Download className="w-3.5 h-3.5" />
                    신뢰 카탈로그 (.reg) 다운로드
                  </a>
                </div>
                <p className="text-[11px] text-slate-500">
                  사용 방법: 위 경로에 <span className="font-mono">manifest.xml</span>을 복사한 뒤, 직원 PC에서 다운로드한 <span className="font-mono">.reg</span> 파일을 실행하고 Outlook → 추가 기능 → 공유 폴더 탭에서 MailLog AI를 추가합니다.
                </p>
              </div>
            )}
          </div>
        </section>

        {/* Manifest Live Code Inspector + Environment Readiness */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <section className="lg:col-span-2 bg-white border border-slate-200 rounded-xl overflow-hidden flex flex-col">
            <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between gap-3">
              <div className="inline-flex items-center p-0.5 bg-slate-100 rounded-md">
                <button
                  type="button"
                  onClick={() => setManifestFormat('xml')}
                  className={`px-3 py-1 text-xs font-medium rounded transition-colors ${
                    manifestFormat === 'xml' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600'
                  }`}
                >
                  manifest.xml (표준 Outlook)
                </button>
                <button
                  type="button"
                  onClick={() => setManifestFormat('json')}
                  className={`px-3 py-1 text-xs font-medium rounded transition-colors ${
                    manifestFormat === 'json' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600'
                  }`}
                >
                  manifest.json (M365 통합)
                </button>
              </div>
              <button
                type="button"
                onClick={() =>
                  copyText('manifest-code', manifestFormat === 'xml' ? manifestXml : manifestJson)
                }
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md"
              >
                {copiedId === 'manifest-code' ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    복사됨
                  </>
                ) : (
                  <>
                    <ClipboardCopy className="w-3.5 h-3.5" />
                    코드 복사
                  </>
                )}
              </button>
            </div>
            <pre className="p-4 text-[11px] leading-relaxed font-mono text-slate-800 bg-slate-50/60 overflow-auto max-h-[360px]">
              <code>{manifestFormat === 'xml' ? manifestXml : manifestJson}</code>
            </pre>
          </section>

          {/* Server Secrets & Microsoft Graph OAuth Readiness */}
          <section className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-[#0F6CBD]" />
              <h2 className="text-sm font-bold text-slate-900">
                서버 연동 상태 (보안 환경변수)
              </h2>
            </div>
            <p className="text-xs text-slate-600 leading-relaxed">
              모든 AI API 키와 Microsoft OAuth 시크릿은 브라우저에 노출되지 않고 서버 환경변수(Settings &gt; Secrets)에서만 안전하게 호출됩니다.
            </p>

            <div className="divide-y divide-slate-100 border border-slate-200 rounded-lg overflow-hidden text-xs">
              {providers.map((p) => (
                <div key={p.id} className="px-3 py-2.5 flex items-center justify-between gap-2 bg-white">
                  <div>
                    <div className="font-semibold text-slate-900">{p.name}</div>
                    <div className="text-[11px] font-mono text-slate-500">{p.envKey}</div>
                  </div>
                  <span
                    className={`font-medium whitespace-nowrap ${
                      p.configured ? 'text-emerald-700' : 'text-slate-400'
                    }`}
                  >
                    {p.configured ? '활성화됨' : '미등록'}
                  </span>
                </div>
              ))}
              <div className="px-3 py-2.5 flex items-center justify-between gap-2 bg-white">
                <div>
                  <div className="font-semibold text-slate-900">Microsoft Graph OAuth</div>
                  <div className="text-[11px] font-mono text-slate-500">MICROSOFT_CLIENT_ID</div>
                </div>
                <span
                  className={`font-medium whitespace-nowrap ${
                    microsoftConfigured ? 'text-emerald-700' : 'text-slate-400'
                  }`}
                >
                  {microsoftConfigured ? '활성화됨' : '미등록'}
                </span>
              </div>
            </div>

            <div className="space-y-1.5 pt-1">
              <div className="text-[11px] font-medium text-slate-700">
                Azure Entra ID Redirect URI (기간별 실제 메일 동기화 시 등록)
              </div>
              <div className="flex items-center gap-1.5">
                <code className="flex-1 px-2.5 py-1.5 text-[11px] font-mono bg-slate-50 border border-slate-200 rounded truncate text-slate-800">
                  {oauthCallbackUrl}
                </code>
                <button
                  type="button"
                  onClick={() => copyText('oauth-uri', oauthCallbackUrl)}
                  aria-label="OAuth Redirect URI 복사"
                  className="p-1.5 border border-slate-200 rounded hover:bg-slate-50 text-slate-600"
                >
                  {copiedId === 'oauth-uri' ? (
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                  ) : (
                    <ClipboardCopy className="w-3.5 h-3.5" />
                  )}
                </button>
              </div>
            </div>
          </section>
        </div>
      </div>
    </div>
  );
};
