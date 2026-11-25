import React, { useEffect, useState } from 'react';
import {
  Bell,
  Building2,
  Calendar,
  Check,
  Filter,
  Layout,
  Mail,
  PenLine,
  Plus,
  RotateCcw,
  Settings,
  ShieldBan,
  Ship,
  Sliders,
  Smartphone,
  Tag,
  Trash2,
  Umbrella,
  X,
  Zap,
} from 'lucide-react';
import {
  ACCENT_THEME_COLORS,
  DEFAULT_OUTLOOK_PREFERENCES,
  type MailRuleActionType,
  type MailRuleConditionField,
  type OutlookAccentTheme,
  type OutlookCategoryTag,
  type OutlookDensity,
  type OutlookFontSize,
  type OutlookMailRule,
  type OutlookPreferences,
  type OutlookPreviewLines,
  type OutlookQuickStep,
  type OutlookQuickStepAction,
  type OutlookReadingPanePosition,
  type OutlookUndoSendSeconds,
} from '../utils/outlookWorkspaceModel';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';

interface OutlookSettingsModalProps {
  isOpen: boolean;
  preferences: OutlookPreferences;
  projects: ProjectDefinition[];
  onChangePreferences: (next: OutlookPreferences) => void;
  onClose: () => void;
  onOpenServerConfig: () => void;
  onOpenSgwSedms: () => void;
  onDiscoverOutlookSettings?: () => Promise<string | void> | void;
  isDiscoveringOutlookSettings?: boolean;
}

type PrimarySection = 'general' | 'mail' | 'calendar' | 'sedms';
type MailSubSection = 'layout' | 'compose' | 'rules' | 'quicksteps' | 'categories' | 'oof' | 'junk';

const FLUENT_INPUT_CLASS =
  'w-full px-3 py-2 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] placeholder:text-[#707070] transition-colors duration-150 focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD] disabled:bg-[#F5F5F5] disabled:text-[#A19F9D]';

const FLUENT_SELECT_CLASS =
  'w-full px-3 py-2 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] transition-colors duration-150 focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD] disabled:bg-[#F5F5F5] disabled:text-[#A19F9D]';

const FLUENT_TEXTAREA_CLASS =
  'w-full px-3 py-2 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] placeholder:text-[#707070] leading-relaxed transition-colors duration-150 focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD]';

const FLUENT_CHECKBOX_CLASS =
  'w-4 h-4 accent-[#0F6CBD] rounded-[3px] focus:outline-none focus:ring-2 focus:ring-[#0F6CBD] focus:ring-offset-1';

export const OutlookSettingsModal: React.FC<OutlookSettingsModalProps> = ({
  isOpen,
  preferences,
  projects,
  onChangePreferences,
  onClose,
  onOpenServerConfig,
  onOpenSgwSedms,
  onDiscoverOutlookSettings,
  isDiscoveringOutlookSettings = false,
}) => {
  const [section, setSection] = useState<PrimarySection>('mail');
  const [mailSub, setMailSub] = useState<MailSubSection>('layout');

  // New Rule Form State
  const [ruleName, setRuleName] = useState('');
  const [ruleField, setRuleField] = useState<MailRuleConditionField>('drawing');
  const [ruleValue, setRuleValue] = useState('');
  const [ruleAction, setRuleAction] = useState<MailRuleActionType>('assign-project');
  const [ruleTarget, setRuleTarget] = useState('');

  // Junk / Safe Sender Inputs
  const [blockedInput, setBlockedInput] = useState('');
  const [safeInput, setSafeInput] = useState('');

  // Quick Step & Category Form State
  const [qsName, setQsName] = useState('');
  const [qsAction, setQsAction] = useState<OutlookQuickStepAction>('archive');
  const [qsMarkRead, setQsMarkRead] = useState(true);
  const [catName, setCatName] = useState('');
  const [catColor, setCatColor] = useState('#0F6CBD');

  // Backend Proactive Monitor & SGW Mobile Alert State
  const [monitorStatus, setMonitorStatus] = useState<{
    running: boolean;
    settings: {
      enabled: boolean;
      pollIntervalSec: number;
      unviewedGraceSec: number;
      sgwPushEnabled: boolean;
      sgwUrgentOnly: boolean;
      sgwTargetRoomId: string;
    };
    sgwConnected: boolean;
    sgwUserEmail: string | null;
    totalSgwAlertsSent: number;
    pendingUnviewedCount: number;
    digest?: {
      unrepliedIncomingCount: number;
      unreadCount: number;
      totalEmails: number;
    } | null;
  } | null>(null);
  const [monitorNotice, setMonitorNotice] = useState<string | null>(null);
  const [discoverNotice, setDiscoverNotice] = useState<string | null>(null);
  const [isTestingSgw, setIsTestingSgw] = useState(false);

  useEffect(() => {
    if (!isOpen || section !== 'general') return;
    let cancelled = false;
    fetch('/api/local/monitor/status')
      .then((r) => (r.ok ? r.json() : null))
      .then((data) => {
        if (!cancelled && data && data.settings) {
          setMonitorStatus(data);
        }
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, [isOpen, section]);

  const updateMonitorSetting = async (
    patch: Partial<{
      enabled: boolean;
      pollIntervalSec: number;
      unviewedGraceSec: number;
      sgwPushEnabled: boolean;
      sgwUrgentOnly: boolean;
    }>
  ) => {
    setMonitorNotice(null);
    try {
      const res = await fetch('/api/local/monitor/settings', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(patch),
      });
      const data = await res.json();
      if (res.ok && data.status) {
        setMonitorStatus(data.status);
        setMonitorNotice('백그라운드 선행 정리 및 SGW 알림 설정을 저장했습니다.');
      }
    } catch {
      setMonitorNotice('모니터 설정 저장 중 오류가 발생했습니다.');
    }
  };

  const handleTestSgwMobileAlert = async () => {
    setIsTestingSgw(true);
    setMonitorNotice(null);
    try {
      const res = await fetch('/api/local/monitor/test-sgw-alert', { method: 'POST' });
      const data = (await res.json()) as { ok?: boolean; reason?: string };
      setMonitorNotice(data.reason || (data.ok ? '테스트 알림 전송 완료' : '전송 실패'));
    } catch {
      setMonitorNotice('SGW 모바일 알림 테스트 호출에 실패했습니다.');
    } finally {
      setIsTestingSgw(false);
    }
  };

  useEffect(() => {
    if (!isOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  const updateGeneral = (patch: Partial<OutlookPreferences['general']>) => {
    onChangePreferences({
      ...preferences,
      general: { ...preferences.general, ...patch },
    });
  };

  const updateMail = (patch: Partial<OutlookPreferences['mail']>) => {
    onChangePreferences({
      ...preferences,
      mail: { ...preferences.mail, ...patch },
    });
  };

  const updateCalendar = (patch: Partial<OutlookPreferences['calendar']>) => {
    onChangePreferences({
      ...preferences,
      calendar: { ...preferences.calendar, ...patch },
    });
  };

  const updateSedms = (patch: Partial<OutlookPreferences['sedms']>) => {
    onChangePreferences({
      ...preferences,
      sedms: { ...preferences.sedms, ...patch },
    });
  };

  const handleAddRule = () => {
    const cleanName = ruleName.trim() || `규칙 ${preferences.mail.rules.length + 1}`;
    const newRule: OutlookMailRule = {
      id: `rule-${Date.now()}`,
      name: cleanName,
      enabled: true,
      conditionField: ruleField,
      conditionValue: ruleValue.trim(),
      actionType: ruleAction,
      actionTarget: ruleAction === 'assign-project' ? (ruleTarget.trim() || projects[0]?.code || '') : undefined,
    };
    updateMail({ rules: [newRule, ...preferences.mail.rules] });
    setRuleName('');
    setRuleValue('');
  };

  const handleAddBlocked = () => {
    const clean = blockedInput.trim().toLowerCase();
    if (!clean) return;
    if (!preferences.mail.blockedSenders.includes(clean)) {
      updateMail({ blockedSenders: [clean, ...preferences.mail.blockedSenders] });
    }
    setBlockedInput('');
  };

  const handleAddSafe = () => {
    const clean = safeInput.trim().toLowerCase();
    if (!clean) return;
    if (!preferences.mail.safeSenders.includes(clean)) {
      updateMail({ safeSenders: [clean, ...preferences.mail.safeSenders] });
    }
    setSafeInput('');
  };

  const handleAddQuickStep = () => {
    const cleanName = qsName.trim();
    if (!cleanName) return;
    const actions: OutlookQuickStepAction[] = [];
    if (qsMarkRead && qsAction !== 'mark-read') actions.push('mark-read');
    actions.push(qsAction);
    const nextStep: OutlookQuickStep = {
      id: `qs-${Date.now()}`,
      name: cleanName,
      actions,
      categoryId: qsAction === 'assign-category' ? preferences.mail.categories[0]?.id : undefined,
      projectCode: qsAction === 'assign-project' ? projects[0]?.code : undefined,
    };
    updateMail({ quickSteps: [...preferences.mail.quickSteps, nextStep] });
    setQsName('');
  };

  const handleAddCategory = () => {
    const cleanName = catName.trim();
    if (!cleanName) return;
    const nextCat: OutlookCategoryTag = {
      id: `cat-${Date.now()}`,
      name: cleanName,
      color: catColor,
    };
    updateMail({ categories: [...preferences.mail.categories, nextCat] });
    setCatName('');
  };

  const primaryNav: Array<{ id: PrimarySection; label: string; icon: React.ReactNode }> = [
    { id: 'general', label: '일반', icon: <Sliders className="w-4 h-4" /> },
    { id: 'mail', label: '메일', icon: <Mail className="w-4 h-4" /> },
    { id: 'calendar', label: '일정', icon: <Calendar className="w-4 h-4" /> },
    { id: 'sedms', label: '조선·도면 (SEDMS)', icon: <Ship className="w-4 h-4" /> },
  ];

  const mailSubNav: Array<{ id: MailSubSection; label: string; icon: React.ReactNode }> = [
    { id: 'layout', label: '레이아웃', icon: <Layout className="w-3.5 h-3.5" /> },
    { id: 'compose', label: '작성 및 회신', icon: <PenLine className="w-3.5 h-3.5" /> },
    { id: 'rules', label: '규칙', icon: <Filter className="w-3.5 h-3.5" /> },
    { id: 'quicksteps', label: '빠른 단계', icon: <Zap className="w-3.5 h-3.5" /> },
    { id: 'categories', label: '색 범주', icon: <Tag className="w-3.5 h-3.5" /> },
    { id: 'oof', label: '자동 회신', icon: <Umbrella className="w-3.5 h-3.5" /> },
    { id: 'junk', label: '정크 메일', icon: <ShieldBan className="w-3.5 h-3.5" /> },
  ];

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="outlook-settings-modal-title"
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 p-3 sm:p-6"
    >
      <div className="w-full max-w-[920px] h-[640px] max-h-[88vh] bg-white text-[#242424] rounded-[8px] border border-[#D1D1D1] shadow-[0_16px_32px_rgba(0,0,0,0.18)] flex flex-col overflow-hidden">
        {/* Top Outlook Settings Bar */}
        <div className="h-12 px-5 border-b border-[#EDEBE9] flex items-center justify-between shrink-0 bg-[#FAF9F8]">
          <div className="flex items-center gap-2.5">
            <Settings className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
            <h2 id="outlook-settings-modal-title" className="text-sm font-bold text-[#242424]">
              Bmail Outlook 설정
            </h2>
            <span className="text-[11px] text-[#616161] hidden sm:inline">
              변경 사항이 작업 화면에 즉시 반영되며 자동 저장됩니다
            </span>
          </div>
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={() => onChangePreferences(structuredClone(DEFAULT_OUTLOOK_PREFERENCES))}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#424242] hover:bg-black/[0.05] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
              title="기본 Outlook 설정으로 복원"
            >
              <RotateCcw className="w-3.5 h-3.5" strokeWidth={1.75} />
              기본값 복원
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label="Outlook 설정 닫기"
              className="p-1.5 text-[#616161] hover:text-[#242424] hover:bg-black/[0.06] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
            >
              <X className="w-4 h-4" strokeWidth={1.75} />
            </button>
          </div>
        </div>

        {/* Authentic Microsoft New Outlook 3-Column Settings Body */}
        <div className="flex-1 flex min-h-0 overflow-hidden">
          {/* Col 1: Primary Categories */}
          <nav
            aria-label="Outlook 설정 기본 범주"
            className="w-[156px] sm:w-[176px] shrink-0 bg-[#F5F5F5] border-r border-[#E0E0E0] p-2 flex flex-col justify-between"
          >
            <div className="space-y-0.5">
              {primaryNav.map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setSection(item.id)}
                  aria-current={section === item.id ? 'page' : undefined}
                  className={`relative w-full flex items-center gap-2.5 px-3 py-2 rounded-[4px] text-xs transition-colors duration-150 text-left focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] ${
                    section === item.id
                      ? 'bg-white text-[#0F6CBD] font-semibold shadow-xs'
                      : 'text-[#323130] hover:bg-black/[0.04]'
                  }`}
                >
                  {section === item.id && (
                    <span className="absolute left-0 top-1.5 bottom-1.5 w-[3px] rounded-r bg-[#0F6CBD]" />
                  )}
                  {item.icon}
                  <span className="truncate">{item.label}</span>
                </button>
              ))}
            </div>

            <div className="pt-2 border-t border-[#E0E0E0] space-y-1">
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onOpenServerConfig();
                }}
                className="w-full flex items-center gap-2 px-2.5 py-1.5 text-[11px] font-semibold text-[#0F6CBD] bg-[#EBF3FC]/70 hover:bg-[#EBF3FC] rounded-[4px] text-left transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
              >
                <Settings className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" strokeWidth={1.75} />
                <span className="truncate">EMV · API 서버 설정</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onOpenSgwSedms();
                }}
                className="w-full flex items-center gap-2 px-2.5 py-1.5 text-[11px] font-medium text-[#424242] hover:bg-white rounded-[4px] text-left transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
              >
                <Building2 className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" strokeWidth={1.75} />
                <span className="truncate">SGW · SEDMS 허브</span>
              </button>
            </div>
          </nav>

          {/* Col 2: Mail Sub-navigation (just like New Outlook) */}
          {section === 'mail' && (
            <nav
              aria-label="메일 세부 설정"
              className="w-[144px] sm:w-[164px] shrink-0 bg-[#FAF9F8] border-r border-[#EDEBE9] p-2 space-y-0.5"
            >
              {mailSubNav.map((sub) => (
                <button
                  key={sub.id}
                  type="button"
                  onClick={() => setMailSub(sub.id)}
                  aria-current={mailSub === sub.id ? 'page' : undefined}
                  className={`w-full flex items-center gap-2 px-2.5 py-2 rounded-[4px] text-xs text-left transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] ${
                    mailSub === sub.id
                      ? 'bg-[#EBF3FC] text-[#0F6CBD] font-semibold'
                      : 'text-[#424242] hover:bg-black/[0.04]'
                  }`}
                >
                  {sub.icon}
                  <span className="truncate">{sub.label}</span>
                </button>
              ))}
            </nav>
          )}

          {/* Col 3: Detail Settings Canvas */}
          <div className="flex-1 overflow-y-auto p-5 sm:p-6 space-y-6 bg-white">
            {/* ==================== GENERAL ==================== */}
            {section === 'general' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">일반 설정</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    표시 밀도, 테마 색상, 언어 및 알림 동작을 Outlook 데스크톱 기준으로 조정합니다.
                  </p>
                </div>

                {/* Full PC Outlook Discovery & One-Click Apply */}
                <div className="p-3.5 rounded-[8px] border border-[#F0C76E] bg-[#FFFDF6] space-y-2">
                  <div className="flex items-center justify-between gap-2 flex-wrap">
                    <div className="text-xs font-bold text-[#7A4400]">
                      PC Outlook 현재 설정 탐색 · 명함 서명 포함 전체 가져오기
                    </div>
                    <button
                      type="button"
                      disabled={isDiscoveringOutlookSettings}
                      onClick={async () => {
                        setDiscoverNotice(null);
                        if (onDiscoverOutlookSettings) {
                          const msg = await onDiscoverOutlookSettings();
                          if (typeof msg === 'string' && msg) setDiscoverNotice(msg);
                          return;
                        }
                        try {
                          const r = await fetch('/api/local/outlook/discover-settings', { method: 'POST' });
                          const d = (await r.json()) as {
                            signaturesCount?: number;
                            inlinedSignatureImages?: number;
                            accountsCount?: number;
                            favoritesCount?: number;
                            error?: string;
                          };
                          if (!r.ok) {
                            setDiscoverNotice(d.error || 'Outlook 설정 탐색에 실패했습니다.');
                          } else {
                            setDiscoverNotice(
                              `서명 ${d.signaturesCount ?? 0}개(명함 이미지 ${d.inlinedSignatureImages ?? 0}개) · 계정 ${d.accountsCount ?? 0}개 · 즐겨찾기 ${d.favoritesCount ?? 0}개를 반영했습니다.`
                            );
                          }
                        } catch {
                          setDiscoverNotice('Outlook 설정 탐색 중 네트워크 오류가 발생했습니다.');
                        }
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:opacity-60 rounded-[4px]"
                    >
                      아웃룩 현재 설정 탐색·적용
                    </button>
                  </div>
                  <p className="text-[11px] text-[#6E4008] leading-relaxed">
                    설치된 PC Outlook의 서명(.htm + 명함 로고 이미지), 기본 새 메일/회신 서명, 읽기창, 근무시간, 즐겨찾기, 계정 프로필을 한 번에 가져와 적용합니다.
                  </p>
                  {discoverNotice && (
                    <div className="text-[11px] font-semibold text-emerald-700">{discoverNotice}</div>
                  )}
                </div>

                {/* Display Density */}
                <div className="space-y-2 pb-4 border-b border-[#EDEBE9]">
                  <div className="text-xs font-semibold text-[#242424]">표시 밀도 (Display density)</div>
                  <div className="grid grid-cols-3 gap-2">
                    {(
                      [
                        { id: 'roomy', label: '여유롭게', desc: '넓은 행 간격과 여유 공간' },
                        { id: 'comfortable', label: '표준 (권장)', desc: 'New Outlook 기본 간격' },
                        { id: 'compact', label: '간략히', desc: '한 화면에 더 많은 메일 표시' },
                      ] as Array<{ id: OutlookDensity; label: string; desc: string }>
                    ).map((opt) => (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => updateGeneral({ density: opt.id })}
                        aria-pressed={preferences.general.density === opt.id}
                        className={`p-3 rounded-md border text-left transition-colors ${
                          preferences.general.density === opt.id
                            ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD]'
                            : 'bg-white border-[#D1D1D1] text-[#242424] hover:bg-[#FAF9F8]'
                        }`}
                      >
                        <div className="text-xs font-semibold">{opt.label}</div>
                        <div className="text-[11px] text-[#616161] mt-0.5">{opt.desc}</div>
                      </button>
                    ))}
                  </div>
                </div>

                {/* Accent Theme */}
                <div className="space-y-2 pb-4 border-b border-[#EDEBE9]">
                  <div className="text-xs font-semibold text-[#242424]">테마 강조색 (Theme Accent)</div>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                    {(Object.keys(ACCENT_THEME_COLORS) as OutlookAccentTheme[]).map((key) => {
                      const info = ACCENT_THEME_COLORS[key];
                      const active = preferences.general.accentTheme === key;
                      return (
                        <button
                          key={key}
                          type="button"
                          onClick={() => updateGeneral({ accentTheme: key })}
                          aria-pressed={active}
                          className={`flex items-center gap-2 px-3 py-2 rounded-md border text-xs ${
                            active
                              ? 'border-[#0F6CBD] bg-[#EBF3FC] font-semibold'
                              : 'border-[#D1D1D1] hover:bg-[#FAF9F8]'
                          }`}
                        >
                          <span
                            className="w-4 h-4 rounded-full shrink-0 flex items-center justify-center text-white"
                            style={{ backgroundColor: info.primary }}
                          >
                            {active && <Check className="w-2.5 h-2.5" />}
                          </span>
                          <span className="truncate">{info.label}</span>
                        </button>
                      );
                    })}
                  </div>
                </div>

                {/* Language & Timezone */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pb-4 border-b border-[#EDEBE9]">
                  <div className="space-y-1.5">
                    <label htmlFor="pref-language" className="block text-xs font-semibold text-[#242424]">
                      표시 언어
                    </label>
                    <select
                      id="pref-language"
                      value={preferences.general.language}
                      onChange={(e) =>
                        updateGeneral({ language: e.target.value as 'ko-KR' | 'en-US' })
                      }
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="ko-KR">한국어 (대한민국)</option>
                      <option value="en-US">English (United States)</option>
                    </select>
                  </div>
                  <div className="space-y-1.5">
                    <label htmlFor="pref-timezone" className="block text-xs font-semibold text-[#242424]">
                      표준 시간대
                    </label>
                    <select
                      id="pref-timezone"
                      value={preferences.general.timezone}
                      onChange={(e) => updateGeneral({ timezone: e.target.value })}
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="Asia/Seoul (UTC+09:00)">서울 (UTC+09:00)</option>
                      <option value="Asia/Tokyo (UTC+09:00)">도쿄 (UTC+09:00)</option>
                      <option value="Asia/Singapore (UTC+08:00)">싱가포르 (UTC+08:00)</option>
                      <option value="Europe/London (UTC+00:00)">런던 (UTC+00:00)</option>
                      <option value="UTC">협정 세계시 (UTC)</option>
                    </select>
                  </div>
                </div>

                {/* Notifications */}
                <div className="space-y-2.5">
                  <div className="flex items-center gap-1.5 text-xs font-semibold text-[#242424]">
                    <Bell className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
                    알림 설정
                  </div>
                  {[
                    {
                      key: 'desktopNotifications' as const,
                      label: '새 메일 수신 시 데스크톱 알림 배너 표시',
                    },
                    {
                      key: 'notificationSound' as const,
                      label: '새 메일 도착 알림 소리 재생',
                    },
                    {
                      key: 'showUnreadBadge' as const,
                      label: '폴더 및 상단 메뉴에 읽지 않은 메일 수 표시',
                    },
                  ].map((item) => (
                    <label key={item.key} className="flex items-center gap-2.5 text-xs text-[#242424] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={preferences.general[item.key]}
                        onChange={(e) => updateGeneral({ [item.key]: e.target.checked })}
                        className={FLUENT_CHECKBOX_CLASS}
                      />
                      <span>{item.label}</span>
                    </label>
                  ))}
                </div>

                {/* Background Proactive Dashboard Digest & Unviewed SGW Mobile Push */}
                <div className="mt-4 p-4 rounded-[8px] border border-[#CFE4FA] bg-[#F8FBFF] space-y-3">
                  <div className="flex items-center justify-between gap-2 flex-wrap">
                    <div className="flex items-center gap-1.5 text-xs font-bold text-[#0F6CBD]">
                      <Smartphone className="w-4 h-4 shrink-0" />
                      <span>백그라운드 선행 정리 · SGW 모바일(핸드폰) 알림 연동</span>
                    </div>
                    <span
                      className={`px-2 py-0.5 rounded-full text-[10px] font-semibold ${
                        monitorStatus?.sgwConnected
                          ? 'bg-emerald-100 text-emerald-800'
                          : 'bg-amber-100 text-amber-800'
                      }`}
                    >
                      {monitorStatus?.sgwConnected
                        ? `SGW 연결됨 (${monitorStatus.sgwUserEmail || '사내 계정'})`
                        : 'SGW 채팅 미연결 (연결 시 자동 푸시)'}
                    </span>
                  </div>
                  <p className="text-[11px] text-[#424242] leading-relaxed">
                    앱 창이 숨겨지거나 백그라운드 상태여도 새 메일을 먼저 수신·요약해 대시보드에 정리하며, 지정 시간 동안 열람하지 않으면 SGW 사내 채팅(내 계정 쪽지)으로 핸드폰 알림을 보냅니다.
                  </p>
                  {monitorStatus && (
                    <div className="space-y-2.5 pt-1">
                      <label className="flex items-center gap-2 text-xs text-[#242424] cursor-pointer">
                        <input
                          type="checkbox"
                          checked={monitorStatus.settings.enabled}
                          onChange={(e) => void updateMonitorSetting({ enabled: e.target.checked })}
                          className={FLUENT_CHECKBOX_CLASS}
                        />
                        <span>백그라운드 자동 수신확인 및 대시보드 선행 정리 켜기</span>
                      </label>
                      <label className="flex items-center gap-2 text-xs text-[#242424] cursor-pointer">
                        <input
                          type="checkbox"
                          checked={monitorStatus.settings.sgwPushEnabled}
                          onChange={(e) =>
                            void updateMonitorSetting({ sgwPushEnabled: e.target.checked })
                          }
                          className={FLUENT_CHECKBOX_CLASS}
                        />
                        <span>내가 열람하지 않으면 SGW 채팅(모바일)으로 요약 알림 보내기</span>
                      </label>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                        <div className="space-y-1">
                          <label htmlFor="sgw-grace-sec" className="block text-[11px] font-semibold text-[#242424]">
                            미열람 판단 대기 시간
                          </label>
                          <select
                            id="sgw-grace-sec"
                            value={monitorStatus.settings.unviewedGraceSec}
                            onChange={(e) =>
                              void updateMonitorSetting({
                                unviewedGraceSec: Number(e.target.value),
                              })
                            }
                            className={FLUENT_SELECT_CLASS}
                          >
                            <option value={0}>즉시 (0초 · 접수 바로 푸시)</option>
                            <option value={60}>1분 후 미열람 시</option>
                            <option value={120}>2분 후 미열람 시 (권장)</option>
                            <option value={300}>5분 후 미열람 시</option>
                          </select>
                        </div>
                        <div className="flex items-end pb-1">
                          <label className="flex items-center gap-2 text-xs text-[#242424] cursor-pointer">
                            <input
                              type="checkbox"
                              checked={monitorStatus.settings.sgwUrgentOnly}
                              onChange={(e) =>
                                void updateMonitorSetting({ sgwUrgentOnly: e.target.checked })
                              }
                              className={FLUENT_CHECKBOX_CLASS}
                            />
                            <span>긴급 메일만 핸드폰(SGW) 알림 발송</span>
                          </label>
                        </div>
                      </div>
                      <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-[#D8E6F5]">
                        <span className="text-[11px] text-[#616161]">
                          선행 대시보드: 총 {monitorStatus.digest?.totalEmails.toLocaleString() ?? 0}건 · 미회신{' '}
                          {monitorStatus.digest?.unrepliedIncomingCount.toLocaleString() ?? 0}건 · 누적 SGW 알림{' '}
                          {monitorStatus.totalSgwAlertsSent}건
                        </span>
                        <button
                          type="button"
                          disabled={isTestingSgw}
                          onClick={() => void handleTestSgwMobileAlert()}
                          className="px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px] disabled:opacity-50"
                        >
                          {isTestingSgw ? '발송 중...' : 'SGW 모바일 테스트 발송'}
                        </button>
                      </div>
                    </div>
                  )}
                  {monitorNotice && (
                    <div className="text-[11px] font-medium text-[#0F6CBD]">{monitorNotice}</div>
                  )}
                </div>
              </>
            )}

            {/* ==================== MAIL > LAYOUT ==================== */}
            {section === 'mail' && mailSub === 'layout' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">메일 레이아웃</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    내용창(읽기 창) 위치와 크기, 중요 받은 편지함, 목록 미리보기를 세밀하게 제어합니다.
                  </p>
                </div>

                {/* Reading Pane Position */}
                <div className="space-y-2 pb-4 border-b border-[#EDEBE9]">
                  <div className="text-xs font-semibold text-[#242424]">내용창(읽기 창) 표시 위치</div>
                  <div className="grid grid-cols-3 gap-2">
                    {(
                      [
                        { id: 'right', label: '오른쪽에 표시 (기본)', desc: '넓은 내용창 우선 3단 레이아웃' },
                        { id: 'bottom', label: '아래쪽에 표시', desc: '상단 목록 · 하단 전체 폭 내용창' },
                        { id: 'hidden', label: '읽기 창 숨기기', desc: '메일 목록 전체 폭 보기' },
                      ] as Array<{ id: OutlookReadingPanePosition; label: string; desc: string }>
                    ).map((opt) => (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => updateMail({ readingPanePosition: opt.id })}
                        aria-pressed={preferences.mail.readingPanePosition === opt.id}
                        className={`p-3 rounded-md border text-left transition-colors ${
                          preferences.mail.readingPanePosition === opt.id
                            ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD]'
                            : 'bg-white border-[#D1D1D1] text-[#242424] hover:bg-[#FAF9F8]'
                        }`}
                      >
                        <div className="text-xs font-semibold">{opt.label}</div>
                        <div className="text-[11px] text-[#616161] mt-0.5">{opt.desc}</div>
                      </button>
                    ))}
                  </div>
                </div>

                {/* Wide Reading Pane & Reading Font Size */}
                <div className="space-y-3 pb-4 border-b border-[#EDEBE9]">
                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">내용창 최대 비율 우선 모드</div>
                      <div className="text-[11px] text-[#616161]">
                        폴더·목록 폭을 슬림하게 유지하고 메일 본문 내용창 영역을 최대로 확보합니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.mail.wideReadingPane}
                      onChange={(e) => updateMail({ wideReadingPane: e.target.checked })}
                      className="w-4 h-4 accent-[#0F6CBD] rounded"
                    />
                  </label>

                  <div className="flex items-center justify-between gap-3">
                    <span className="text-xs font-semibold text-[#242424]">내용창 본문 글자 크기</span>
                    <div className="inline-flex rounded-[5px] border border-[#D1D1D1] overflow-hidden">
                      {(
                        [
                          { id: 'sm', label: '작게' },
                          { id: 'md', label: '보통' },
                          { id: 'lg', label: '크게' },
                        ] as Array<{ id: OutlookFontSize; label: string }>
                      ).map((fs) => (
                        <button
                          key={fs.id}
                          type="button"
                          onClick={() => updateMail({ readingFontSize: fs.id })}
                          className={`px-3 py-1 text-xs ${
                            preferences.mail.readingFontSize === fs.id
                              ? 'bg-[#0F6CBD] text-white font-semibold'
                              : 'bg-white text-[#424242] hover:bg-[#F5F5F5]'
                          }`}
                        >
                          {fs.label}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Focused Inbox & List Options */}
                <div className="space-y-3">
                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">중요 받은 편지함 (Focused Inbox)</div>
                      <div className="text-[11px] text-[#616161]">
                        받은 편지함을 &ldquo;중요&rdquo;와 &ldquo;기타&rdquo; 탭으로 자동 분리합니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.mail.focusedInbox}
                      onChange={(e) => updateMail({ focusedInbox: e.target.checked })}
                      className="w-4 h-4 accent-[#0F6CBD] rounded"
                    />
                  </label>

                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">대화 스레드로 그룹화</div>
                      <div className="text-[11px] text-[#616161]">
                        동일 제목(RE:/FW:)의 주고받은 메일을 하나의 대화 스레드로 묶어 표시합니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.mail.conversationView}
                      onChange={(e) => updateMail({ conversationView: e.target.checked })}
                      className="w-4 h-4 accent-[#0F6CBD] rounded"
                    />
                  </label>

                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">발신자 이니셜 아바타 표시</div>
                      <div className="text-[11px] text-[#616161]">
                        끄면 목록 행이 더욱 간결해지고 제목 영역이 넓어집니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.mail.showSenderAvatars}
                      onChange={(e) => updateMail({ showSenderAvatars: e.target.checked })}
                      className="w-4 h-4 accent-[#0F6CBD] rounded"
                    />
                  </label>

                  <div className="flex items-center justify-between gap-3 pt-1">
                    <span className="text-xs font-semibold text-[#242424]">목록 미리보기 텍스트</span>
                    <div className="inline-flex rounded-[5px] border border-[#D1D1D1] overflow-hidden">
                      {(
                        [
                          { id: 0, label: '숨기기' },
                          { id: 1, label: '1줄' },
                          { id: 2, label: '2줄' },
                        ] as Array<{ id: OutlookPreviewLines; label: string }>
                      ).map((pl) => (
                        <button
                          key={pl.id}
                          type="button"
                          onClick={() => updateMail({ previewLines: pl.id })}
                          className={`px-3 py-1 text-xs ${
                            preferences.mail.previewLines === pl.id
                              ? 'bg-[#0F6CBD] text-white font-semibold'
                              : 'bg-white text-[#424242] hover:bg-[#F5F5F5]'
                          }`}
                        >
                          {pl.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Single-Line Ribbon & Minimal Header */}
                  <label className="flex items-center justify-between gap-3 cursor-pointer pt-2 border-t border-[#EDEBE9]">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">싱글라인 리본 (EnableSingleLineRibbon)</div>
                      <div className="text-[11px] text-[#616161]">
                        상단 명령 리본을 한 줄 슬림 모드로 압축해 본문 가독성을 높입니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.mail.singleLineRibbon}
                      onChange={(e) => updateMail({ singleLineRibbon: e.target.checked })}
                      className={FLUENT_CHECKBOX_CLASS}
                    />
                  </label>

                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">미니멀 읽기 헤더 (MinimalHeaderOn)</div>
                      <div className="text-[11px] text-[#616161]">
                        본문 상단 발신자/수신자 영역을 간결하게 접어 본문을 더 많이 표시합니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.mail.minimalHeader}
                      onChange={(e) => updateMail({ minimalHeader: e.target.checked })}
                      className={FLUENT_CHECKBOX_CLASS}
                    />
                  </label>

                  {/* Mark as Read Policy */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2 border-t border-[#EDEBE9]">
                    <div className="space-y-1">
                      <label htmlFor="pref-mark-read" className="block text-xs font-semibold text-[#242424]">
                        읽음으로 표시 동작
                      </label>
                      <select
                        id="pref-mark-read"
                        value={preferences.mail.markAsReadMode}
                        onChange={(e) =>
                          updateMail({ markAsReadMode: e.target.value as 'immediate' | 'delay' | 'manual' })
                        }
                        className={FLUENT_SELECT_CLASS}
                      >
                        <option value="immediate">항목 선택 시 즉시 읽음</option>
                        <option value="delay">선택 후 지정 초 경과 시 읽음</option>
                        <option value="manual">직접 읽음 버튼으로만 표시</option>
                      </select>
                    </div>
                    <div className="space-y-1">
                      <label htmlFor="pref-after-delete" className="block text-xs font-semibold text-[#242424]">
                        삭제 · 보관 후 이동 위치
                      </label>
                      <select
                        id="pref-after-delete"
                        value={preferences.mail.afterDeleteAction}
                        onChange={(e) =>
                          updateMail({ afterDeleteAction: e.target.value as 'next' | 'prev' | 'list' })
                        }
                        className={FLUENT_SELECT_CLASS}
                      >
                        <option value="next">다음 항목 열기 (Outlook 기본)</option>
                        <option value="prev">이전 항목 열기</option>
                        <option value="list">목록으로 돌아가기 (선택 해제)</option>
                      </select>
                    </div>
                  </div>
                  {preferences.mail.markAsReadMode === 'delay' && (
                    <div className="flex items-center justify-between gap-3">
                      <label htmlFor="pref-read-delay" className="text-xs font-medium text-[#424242]">
                        읽음 자동 전환 대기 시간 (초)
                      </label>
                      <input
                        id="pref-read-delay"
                        type="number"
                        min={1}
                        max={30}
                        value={preferences.mail.markAsReadDelaySec}
                        onChange={(e) =>
                          updateMail({
                            markAsReadDelaySec: Math.max(1, Math.min(30, Number(e.target.value) || 3)),
                          })
                        }
                        className={`${FLUENT_INPUT_CLASS} w-24 text-right font-mono tabular-nums`}
                      />
                    </div>
                  )}
                </div>
              </>
            )}

            {/* ==================== MAIL > COMPOSE & REPLY ==================== */}
            {section === 'mail' && mailSub === 'compose' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">작성 및 회신</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    이메일 서명, 기본 회신 동작, 발송 취소 대기 시간(Undo Send)을 설정합니다.
                  </p>
                </div>

                {/* Email Signature */}
                <div className="space-y-2 pb-4 border-b border-[#EDEBE9]">
                  <label htmlFor="pref-signature" className="block text-xs font-semibold text-[#242424]">
                    이메일 서명 (Email signature)
                  </label>
                  <textarea
                    id="pref-signature"
                    rows={4}
                    value={preferences.mail.signature}
                    onChange={(e) => updateMail({ signature: e.target.value })}
                    placeholder="예: 홍길동 | 설계팀 선임&#10;Tel: 02-0000-0000 | Bmail · SEDMS"
                    className={FLUENT_TEXTAREA_CLASS}
                  />
                  <div className="flex flex-wrap gap-4 pt-1">
                    <label className="flex items-center gap-2 text-xs text-[#242424] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={preferences.mail.autoSignatureNew}
                        onChange={(e) => updateMail({ autoSignatureNew: e.target.checked })}
                        className={FLUENT_CHECKBOX_CLASS}
                      />
                      새 메일 작성 시 서명 자동 포함
                    </label>
                    <label className="flex items-center gap-2 text-xs text-[#242424] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={preferences.mail.autoSignatureReply}
                        onChange={(e) => updateMail({ autoSignatureReply: e.target.checked })}
                        className={FLUENT_CHECKBOX_CLASS}
                      />
                      회신 및 전달 시 서명 자동 포함
                    </label>
                  </div>
                </div>

                {/* Default Reply & Undo Send */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pb-4 border-b border-[#EDEBE9]">
                  <div className="space-y-1.5">
                    <div className="text-xs font-semibold text-[#242424]">기본 회신 동작</div>
                    <div className="flex gap-2">
                      {(
                        [
                          { id: 'reply', label: '회신' },
                          { id: 'replyAll', label: '전체 회신' },
                        ] as const
                      ).map((opt) => (
                        <button
                          key={opt.id}
                          type="button"
                          onClick={() => updateMail({ defaultReplyAction: opt.id })}
                          className={`flex-1 py-1.5 text-xs rounded-[4px] border ${
                            preferences.mail.defaultReplyAction === opt.id
                              ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-semibold'
                              : 'bg-white border-[#D1D1D1] text-[#424242]'
                          }`}
                        >
                          {opt.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <div className="text-xs font-semibold text-[#242424]">발송 취소 대기 (Undo Send)</div>
                    <div className="flex gap-1.5">
                      {([0, 5, 10] as OutlookUndoSendSeconds[]).map((sec) => (
                        <button
                          key={sec}
                          type="button"
                          onClick={() => updateMail({ undoSendSeconds: sec })}
                          className={`flex-1 py-1.5 text-xs rounded-[4px] border ${
                            preferences.mail.undoSendSeconds === sec
                              ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-semibold'
                              : 'bg-white border-[#D1D1D1] text-[#424242]'
                          }`}
                        >
                          {sec === 0 ? '즉시 발송' : `${sec}초`}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Font Format */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label htmlFor="pref-compose-font" className="block text-xs font-semibold text-[#242424]">
                      작성 글꼴
                    </label>
                    <select
                      id="pref-compose-font"
                      value={preferences.mail.composeFontFamily}
                      onChange={(e) =>
                        updateMail({ composeFontFamily: e.target.value as 'sans' | 'serif' | 'mono' })
                      }
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="sans">맑은 고딕 / Segoe UI (기본)</option>
                      <option value="serif">바탕 / Serif</option>
                      <option value="mono">고정폭 (도면·코드용)</option>
                    </select>
                  </div>
                  <div className="space-y-1.5">
                    <label htmlFor="pref-compose-pt" className="block text-xs font-semibold text-[#242424]">
                      기본 크기
                    </label>
                    <select
                      id="pref-compose-pt"
                      value={preferences.mail.composeFontSizePt}
                      onChange={(e) =>
                        updateMail({ composeFontSizePt: Number(e.target.value) as 10 | 11 | 12 })
                      }
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value={10}>10 pt</option>
                      <option value={11}>11 pt (권장)</option>
                      <option value={12}>12 pt</option>
                    </select>
                  </div>
                </div>

                {/* Default Importance & Read Receipt */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-3 border-t border-[#EDEBE9]">
                  <div className="space-y-1.5">
                    <label htmlFor="pref-importance" className="block text-xs font-semibold text-[#242424]">
                      새 메일 기본 중요도
                    </label>
                    <select
                      id="pref-importance"
                      value={preferences.mail.defaultImportance}
                      onChange={(e) =>
                        updateMail({ defaultImportance: e.target.value as 'low' | 'normal' | 'high' })
                      }
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="high">높음 (긴급 표시)</option>
                      <option value="normal">보통 (기본)</option>
                      <option value="low">낮음</option>
                    </select>
                  </div>
                  <div className="flex items-end pb-1.5">
                    <label className="flex items-center gap-2 text-xs text-[#242424] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={preferences.mail.requestReadReceipt}
                        onChange={(e) => updateMail({ requestReadReceipt: e.target.checked })}
                        className={FLUENT_CHECKBOX_CLASS}
                      />
                      기본 발송 시 읽음 확인 요청 (Read Receipt)
                    </label>
                  </div>
                </div>
              </>
            )}

            {/* ==================== MAIL > RULES ==================== */}
            {section === 'mail' && mailSub === 'rules' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">자동 분류 규칙 (Rules)</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    도면번호, 발신자, 제목 또는 첨부 여부에 따라 프로젝트 자동 배정·플래그·읽음 처리를 실행합니다.
                  </p>
                </div>

                {/* Add Rule Form */}
                <div className="p-4 rounded-[8px] border border-[#CFE4FA] bg-[#F8FBFF] space-y-3">
                  <div className="text-xs font-semibold text-[#0F6CBD]">새 규칙 추가</div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                    <input
                      type="text"
                      value={ruleName}
                      onChange={(e) => setRuleName(e.target.value)}
                      placeholder="규칙 이름 (예: 선체 도면 자동 분류)"
                      aria-label="규칙 이름"
                      className={FLUENT_INPUT_CLASS}
                    />
                    <select
                      aria-label="규칙 조건 항목"
                      value={ruleField}
                      onChange={(e) => setRuleField(e.target.value as MailRuleConditionField)}
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="drawing">조건: 조선·SEDMS 도면번호 포함</option>
                      <option value="from">조건: 보낸 사람 포함</option>
                      <option value="subject">조건: 제목 포함</option>
                      <option value="hasAttachment">조건: 첨부 파일 있음</option>
                    </select>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                    <input
                      type="text"
                      value={ruleValue}
                      disabled={ruleField === 'hasAttachment'}
                      onChange={(e) => setRuleValue(e.target.value)}
                      placeholder={
                        ruleField === 'drawing'
                          ? '도면 키 (비우면 모든 도면)'
                          : ruleField === 'hasAttachment'
                            ? '값 불필요'
                            : '키워드 또는 이메일'
                      }
                      aria-label="규칙 조건 값"
                      className={FLUENT_INPUT_CLASS}
                    />
                    <select
                      aria-label="규칙 실행 동작"
                      value={ruleAction}
                      onChange={(e) => setRuleAction(e.target.value as MailRuleActionType)}
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="assign-project">동작: 프로젝트로 배정</option>
                      <option value="flag">동작: 플래그 지정</option>
                      <option value="mark-read">동작: 읽음 처리</option>
                      <option value="archive">동작: 보관함 이동</option>
                      <option value="move-trash">동작: 지운 편지함 이동</option>
                    </select>
                    {ruleAction === 'assign-project' ? (
                      <input
                        type="text"
                        list="outlook-rule-projects"
                        value={ruleTarget}
                        onChange={(e) => setRuleTarget(e.target.value)}
                        placeholder="대상 프로젝트 코드"
                        aria-label="대상 프로젝트 코드"
                        className={FLUENT_INPUT_CLASS}
                      />
                    ) : (
                      <button
                        type="button"
                        onClick={handleAddRule}
                        className="px-3 py-2 inline-flex items-center justify-center gap-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
                      >
                        <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
                        규칙 저장
                      </button>
                    )}
                    <datalist id="outlook-rule-projects">
                      {projects.map((p) => (
                        <option key={p.id} value={p.code}>
                          {p.name}
                        </option>
                      ))}
                    </datalist>
                  </div>
                  {ruleAction === 'assign-project' && (
                    <div className="flex justify-end">
                      <button
                        type="button"
                        onClick={handleAddRule}
                        className="px-3.5 py-2 inline-flex items-center gap-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
                      >
                        <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
                        규칙 저장
                      </button>
                    </div>
                  )}
                </div>

                {/* Existing Rules List */}
                <div className="space-y-2">
                  {preferences.mail.rules.length === 0 ? (
                    <div className="p-6 text-center text-xs text-[#616161] border border-dashed border-[#D1D1D1] rounded-md">
                      등록된 규칙이 없습니다. 위에서 도면번호·발신자 기준 자동 분류 규칙을 추가하세요.
                    </div>
                  ) : (
                    preferences.mail.rules.map((rule) => (
                      <div
                        key={rule.id}
                        className="flex items-center justify-between gap-3 px-3 py-2.5 rounded-md border border-[#E0E0E0] bg-white"
                      >
                        <label className="flex items-center gap-2.5 min-w-0 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={rule.enabled}
                            onChange={(e) =>
                              updateMail({
                                rules: preferences.mail.rules.map((r) =>
                                  r.id === rule.id ? { ...r, enabled: e.target.checked } : r
                                ),
                              })
                            }
                            className="w-4 h-4 accent-[#0F6CBD] rounded shrink-0"
                          />
                          <div className="min-w-0">
                            <div className="text-xs font-semibold text-[#242424] truncate">
                              {rule.name}
                            </div>
                            <div className="text-[11px] text-[#616161] truncate">
                              조건: {rule.conditionField}
                              {rule.conditionValue ? ` (${rule.conditionValue})` : ''} → 동작:{' '}
                              {rule.actionType}
                              {rule.actionTarget ? ` [${rule.actionTarget}]` : ''}
                            </div>
                          </div>
                        </label>
                        <button
                          type="button"
                          onClick={() =>
                            updateMail({
                              rules: preferences.mail.rules.filter((r) => r.id !== rule.id),
                            })
                          }
                          aria-label={`규칙 '${rule.name}' 삭제`}
                          className="p-1 text-[#616161] hover:text-red-700 hover:bg-red-50 rounded"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ))
                  )}
                </div>
              </>
            )}

            {/* ==================== MAIL > AUTOMATIC REPLIES ==================== */}
            {section === 'mail' && mailSub === 'oof' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">자동 회신 (부재중)</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    출장·휴가 기간 동안 자동 회신 안내를 활성화하고 워크벤치 상단에 상태를 표시합니다.
                  </p>
                </div>

                <label className="flex items-center justify-between p-3.5 rounded-md border border-[#D1D1D1] bg-[#FAF9F8] cursor-pointer">
                  <div>
                    <div className="text-xs font-semibold text-[#242424]">자동 회신 켜기</div>
                    <div className="text-[11px] text-[#616161]">
                      활성화 시 받은 편지함 상단에 부재중 자동 회신 배너가 표시됩니다.
                    </div>
                  </div>
                  <input
                    type="checkbox"
                    checked={preferences.mail.outOfOfficeEnabled}
                    onChange={(e) => updateMail({ outOfOfficeEnabled: e.target.checked })}
                    className="w-4 h-4 accent-[#0F6CBD] rounded"
                  />
                </label>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                  <div className="space-y-1.5">
                    <label htmlFor="oof-start" className="block text-xs font-semibold text-[#242424]">
                      시작 일자
                    </label>
                    <input
                      id="oof-start"
                      type="date"
                      value={preferences.mail.outOfOfficeStart}
                      onChange={(e) => updateMail({ outOfOfficeStart: e.target.value })}
                      className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label htmlFor="oof-end" className="block text-xs font-semibold text-[#242424]">
                      종료 일자
                    </label>
                    <input
                      id="oof-end"
                      type="date"
                      value={preferences.mail.outOfOfficeEnd}
                      onChange={(e) => updateMail({ outOfOfficeEnd: e.target.value })}
                      className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                    />
                  </div>
                </div>

                <div className="space-y-1.5">
                  <label htmlFor="oof-msg" className="block text-xs font-semibold text-[#242424]">
                    자동 회신 메시지
                  </label>
                  <textarea
                    id="oof-msg"
                    rows={5}
                    value={preferences.mail.outOfOfficeMessage}
                    onChange={(e) => updateMail({ outOfOfficeMessage: e.target.value })}
                    className={FLUENT_TEXTAREA_CLASS}
                  />
                </div>
              </>
            )}

            {/* ==================== MAIL > JUNK EMAIL ==================== */}
            {section === 'mail' && mailSub === 'junk' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">정크 메일 · 차단 및 안전 발신자</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    차단된 발신자나 도메인(@domain.com)의 메일은 받은 편지함에서 자동으로 제외됩니다.
                  </p>
                </div>

                <div className="space-y-2 pb-4 border-b border-[#EDEBE9]">
                  <div className="text-xs font-semibold text-[#242424]">차단된 보낸 사람 및 도메인</div>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      value={blockedInput}
                      onChange={(e) => setBlockedInput(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && handleAddBlocked()}
                      placeholder="spam@example.com 또는 @spam-domain.com"
                      aria-label="차단할 이메일 또는 도메인"
                      className={`flex-1 ${FLUENT_INPUT_CLASS}`}
                    />
                    <button
                      type="button"
                      onClick={handleAddBlocked}
                      className="px-3.5 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px] shrink-0 transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
                    >
                      차단 추가
                    </button>
                  </div>
                  <div className="flex flex-wrap gap-1.5 pt-1">
                    {preferences.mail.blockedSenders.length === 0 ? (
                      <span className="text-[11px] text-[#616161]">차단된 발신자가 없습니다.</span>
                    ) : (
                      preferences.mail.blockedSenders.map((item) => (
                        <span
                          key={item}
                          className="inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] bg-red-50 border border-red-200 text-[11px] text-red-800 font-mono"
                        >
                          {item}
                          <button
                            type="button"
                            onClick={() =>
                              updateMail({
                                blockedSenders: preferences.mail.blockedSenders.filter((x) => x !== item),
                              })
                            }
                            aria-label={`${item} 차단 해제`}
                            className="hover:text-red-950"
                          >
                            <X className="w-3 h-3" />
                          </button>
                        </span>
                      ))
                    )}
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="text-xs font-semibold text-[#242424]">안전한 보낸 사람 및 도메인</div>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      value={safeInput}
                      onChange={(e) => setSafeInput(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && handleAddSafe()}
                      placeholder="partner@shipyard.co.kr"
                      aria-label="안전 발신자 이메일 또는 도메인"
                      className={`flex-1 ${FLUENT_INPUT_CLASS}`}
                    />
                    <button
                      type="button"
                      onClick={handleAddSafe}
                      className="px-3.5 py-2 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[4px] shrink-0 transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                    >
                      안전 추가
                    </button>
                  </div>
                  <div className="flex flex-wrap gap-1.5 pt-1">
                    {preferences.mail.safeSenders.length === 0 ? (
                      <span className="text-[11px] text-[#616161]">등록된 안전 발신자가 없습니다.</span>
                    ) : (
                      preferences.mail.safeSenders.map((item) => (
                        <span
                          key={item}
                          className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-emerald-50 border border-emerald-200 text-[11px] text-emerald-800 font-mono"
                        >
                          {item}
                          <button
                            type="button"
                            onClick={() =>
                              updateMail({
                                safeSenders: preferences.mail.safeSenders.filter((x) => x !== item),
                              })
                            }
                            aria-label={`${item} 안전 발신자 제거`}
                            className="hover:text-emerald-950"
                          >
                            <X className="w-3 h-3" />
                          </button>
                        </span>
                      ))
                    )}
                  </div>
                </div>
              </>
            )}

            {/* ==================== MAIL > QUICK STEPS ==================== */}
            {section === 'mail' && mailSub === 'quicksteps' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">빠른 단계 (Quick Steps)</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    여러 작업(읽음 표시·플래그·보관·범주 지정)을 리본 버튼 클릭 1회로 동시 실행합니다.
                  </p>
                </div>

                <div className="p-4 rounded-[8px] border border-[#CFE4FA] bg-[#F8FBFF] space-y-3">
                  <div className="text-xs font-semibold text-[#0F6CBD]">새 빠른 단계 추가</div>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                    <input
                      type="text"
                      value={qsName}
                      onChange={(e) => setQsName(e.target.value)}
                      placeholder="이름 (예: 도면 완료)"
                      aria-label="빠른 단계 이름"
                      className={FLUENT_INPUT_CLASS}
                    />
                    <select
                      value={qsAction}
                      onChange={(e) => setQsAction(e.target.value as OutlookQuickStepAction)}
                      aria-label="빠른 단계 주 동작"
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="archive">동작: 보관함 이동</option>
                      <option value="flag">동작: 플래그 지정</option>
                      <option value="pin">동작: 상단 고정</option>
                      <option value="assign-category">동작: 첫 번째 범주 지정</option>
                      <option value="assign-project">동작: 첫 번째 프로젝트 배정</option>
                      <option value="mark-read">동작: 읽음으로 표시</option>
                      <option value="move-trash">동작: 지운 편지함 이동</option>
                    </select>
                    <button
                      type="button"
                      onClick={handleAddQuickStep}
                      className="px-3.5 py-2 inline-flex items-center justify-center gap-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px]"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      단계 추가
                    </button>
                  </div>
                  <label className="inline-flex items-center gap-2 text-xs text-[#242424] cursor-pointer">
                    <input
                      type="checkbox"
                      checked={qsMarkRead}
                      onChange={(e) => setQsMarkRead(e.target.checked)}
                      className={FLUENT_CHECKBOX_CLASS}
                    />
                    실행 시 자동으로 &ldquo;읽음&rdquo; 함께 적용
                  </label>
                </div>

                <div className="space-y-2">
                  {preferences.mail.quickSteps.map((step) => (
                    <div
                      key={step.id}
                      className="flex items-center justify-between gap-3 px-3 py-2.5 rounded-md border border-[#E0E0E0] bg-white"
                    >
                      <div className="min-w-0">
                        <div className="text-xs font-semibold text-[#242424] flex items-center gap-2">
                          <Zap className="w-3.5 h-3.5 text-[#0F6CBD]" />
                          <span className="truncate">{step.name}</span>
                          {step.shortcut && (
                            <span className="text-[10px] font-mono text-[#616161] bg-[#F5F5F5] px-1.5 py-0.5 rounded">
                              {step.shortcut}
                            </span>
                          )}
                        </div>
                        <div className="text-[11px] text-[#616161] mt-0.5">
                          동작: {step.actions.join(' → ')}
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() =>
                          updateMail({
                            quickSteps: preferences.mail.quickSteps.filter((q) => q.id !== step.id),
                          })
                        }
                        aria-label={`빠른 단계 '${step.name}' 삭제`}
                        className="p-1 text-[#616161] hover:text-red-700 hover:bg-red-50 rounded"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                </div>
              </>
            )}

            {/* ==================== MAIL > COLOR CATEGORIES ==================== */}
            {section === 'mail' && mailSub === 'categories' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">색 범주 (Categories)</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    Outlook 색 범주 태그를 관리하고 메일 목록과 본문 리본에서 즉시 분류합니다.
                  </p>
                </div>

                <div className="p-4 rounded-[8px] border border-[#CFE4FA] bg-[#F8FBFF] flex flex-wrap items-center gap-2.5">
                  <input
                    type="color"
                    value={catColor}
                    onChange={(e) => setCatColor(e.target.value)}
                    aria-label="범주 색상 선택"
                    className="w-9 h-9 p-0.5 rounded border border-[#D1D1D1] bg-white cursor-pointer"
                  />
                  <input
                    type="text"
                    value={catName}
                    onChange={(e) => setCatName(e.target.value)}
                    placeholder="새 범주 이름 (예: 선급 코멘트)"
                    aria-label="범주 이름"
                    className={`flex-1 min-w-[180px] ${FLUENT_INPUT_CLASS}`}
                  />
                  <button
                    type="button"
                    onClick={handleAddCategory}
                    className="px-3.5 py-2 inline-flex items-center gap-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px]"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    범주 추가
                  </button>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {preferences.mail.categories.map((cat) => (
                    <div
                      key={cat.id}
                      className="flex items-center justify-between gap-2 px-3 py-2 rounded-md border border-[#E0E0E0] bg-white"
                    >
                      <div className="flex items-center gap-2.5 min-w-0">
                        <span
                          className="w-3.5 h-3.5 rounded-full shrink-0"
                          style={{ backgroundColor: cat.color }}
                        />
                        <span className="text-xs font-semibold text-[#242424] truncate">{cat.name}</span>
                      </div>
                      <button
                        type="button"
                        onClick={() =>
                          updateMail({
                            categories: preferences.mail.categories.filter((c) => c.id !== cat.id),
                          })
                        }
                        aria-label={`범주 '${cat.name}' 삭제`}
                        className="p-1 text-[#616161] hover:text-red-700 hover:bg-red-50 rounded"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                </div>
              </>
            )}

            {/* ==================== CALENDAR ==================== */}
            {section === 'calendar' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">일정 및 근무 시간</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    Outlook 일정 달력의 주 시작 요일과 기본 회의 길이를 설정합니다.
                  </p>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label htmlFor="cal-week-start" className="block text-xs font-semibold text-[#242424]">
                      주 시작 요일
                    </label>
                    <select
                      id="cal-week-start"
                      value={preferences.calendar.weekStartsOn}
                      onChange={(e) =>
                        updateCalendar({ weekStartsOn: e.target.value as 'sunday' | 'monday' })
                      }
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value="monday">월요일</option>
                      <option value="sunday">일요일</option>
                    </select>
                  </div>
                  <div className="space-y-1.5">
                    <label htmlFor="cal-meeting-min" className="block text-xs font-semibold text-[#242424]">
                      기본 회의 시간
                    </label>
                    <select
                      id="cal-meeting-min"
                      value={preferences.calendar.defaultMeetingMinutes}
                      onChange={(e) =>
                        updateCalendar({ defaultMeetingMinutes: Number(e.target.value) as 30 | 60 })
                      }
                      className={FLUENT_SELECT_CLASS}
                    >
                      <option value={30}>30분</option>
                      <option value={60}>60분</option>
                    </select>
                  </div>
                  <div className="space-y-1.5">
                    <label htmlFor="cal-work-start" className="block text-xs font-semibold text-[#242424]">
                      근무 시작 시간
                    </label>
                    <input
                      id="cal-work-start"
                      type="time"
                      value={preferences.calendar.workStart}
                      onChange={(e) => updateCalendar({ workStart: e.target.value })}
                      className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label htmlFor="cal-work-end" className="block text-xs font-semibold text-[#242424]">
                      근무 종료 시간
                    </label>
                    <input
                      id="cal-work-end"
                      type="time"
                      value={preferences.calendar.workEnd}
                      onChange={(e) => updateCalendar({ workEnd: e.target.value })}
                      className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                    />
                  </div>
                </div>
              </>
            )}

            {/* ==================== SHIPBUILDING / SEDMS ==================== */}
            {section === 'sedms' && (
              <>
                <div>
                  <h3 className="text-base font-bold text-[#242424]">조선 · 도면 (SEDMS) 연동 설정</h3>
                  <p className="text-xs text-[#616161] mt-0.5">
                    선체·배관·전장·의장 도면번호 자동 인식, 프로젝트별 파일함 도면 인덱싱 및 SEDMS 5순위 매칭 연동을 관리합니다.
                  </p>
                </div>

                <div className="space-y-3 pb-4 border-b border-[#EDEBE9]">
                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">
                        메일 본문·제목 조선 도면번호 자동 추출
                      </div>
                      <div className="text-[11px] text-[#616161]">
                        읽기 창 상단에 추출된 도면번호(선체·배관·전장·의장)와 SEDMS 바로가기를 표시합니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.sedms.autoExtractDrawings}
                      onChange={(e) => updateSedms({ autoExtractDrawings: e.target.checked })}
                      className="w-4 h-4 accent-[#0F6CBD] rounded"
                    />
                  </label>

                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">
                        프로젝트 파일함 추가 시 도면번호 자동 인덱싱
                      </div>
                      <div className="text-[11px] text-[#616161]">
                        프로젝트 트리에 PDF·DWG·Excel 파일을 넣을 때 파일명에서 도면번호 키를 자동 연결합니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.sedms.autoIndexProjectFiles}
                      onChange={(e) => updateSedms({ autoIndexProjectFiles: e.target.checked })}
                      className="w-4 h-4 accent-[#0F6CBD] rounded"
                    />
                  </label>

                  <label className="flex items-center justify-between gap-3 cursor-pointer">
                    <div>
                      <div className="text-xs font-semibold text-[#242424]">
                        메일 목록에 도면번호 배지 표시
                      </div>
                      <div className="text-[11px] text-[#616161]">
                        도면번호가 감지된 메일 행에 SEDMS 도면 태그를 간결하게 표시합니다.
                      </div>
                    </div>
                    <input
                      type="checkbox"
                      checked={preferences.sedms.showDrawingBadgeInList}
                      onChange={(e) => updateSedms({ showDrawingBadgeInList: e.target.checked })}
                      className="w-4 h-4 accent-[#0F6CBD] rounded"
                    />
                  </label>

                  <div className="space-y-1.5 pt-1">
                    <label htmlFor="sedms-tenant" className="block text-xs font-semibold text-[#242424]">
                      기본 SEDMS 테넌트 ID
                    </label>
                    <input
                      id="sedms-tenant"
                      type="text"
                      value={preferences.sedms.defaultTenantId}
                      onChange={(e) => updateSedms({ defaultTenantId: e.target.value })}
                      className={`${FLUENT_INPUT_CLASS} max-w-xs font-mono`}
                    />
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      onClose();
                      onOpenSgwSedms();
                    }}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD] focus-visible:ring-offset-1"
                  >
                    <Ship className="w-3.5 h-3.5" strokeWidth={1.75} />
                    SEDMS 5순위 매칭 · SGW 허브 열기
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      onClose();
                      onOpenServerConfig();
                    }}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                  >
                    <Settings className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
                    서버 .env · 외부 PDF/Excel 엔진 설정
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
