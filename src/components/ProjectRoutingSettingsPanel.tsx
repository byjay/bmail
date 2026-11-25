import React, { useMemo, useRef, useState } from 'react';
import {
  AtSign,
  Download,
  FolderCog,
  Globe,
  Mail,
  Plus,
  Sparkles,
  Upload,
  UserCheck,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';
import type { ProjectDefinition, ProjectFolderPreset } from '../utils/deterministicAnalyzer';
import {
  classifyByProjectRoutingRules,
  discoverRoutingSuggestions,
  FOLDER_PRESET_META,
  normalizeRuleList,
  type OrganizedFileEntry,
} from '../utils/projectFileOrganizer';

interface ProjectRoutingSettingsPanelProps {
  projects: ProjectDefinition[];
  selectedProjectCode: string;
  onSelectProjectCode: (code: string) => void;
  emails: MailItem[];
  organizedFiles: OrganizedFileEntry[];
  ownerAccountEmail?: string;
  onUpdateProject: (id: string, patch: Partial<ProjectDefinition>) => void;
  onImportProjectRules: (imported: ProjectDefinition[]) => void;
}

export const ProjectRoutingSettingsPanel: React.FC<ProjectRoutingSettingsPanelProps> = ({
  projects,
  selectedProjectCode,
  onSelectProjectCode,
  emails,
  organizedFiles,
  ownerAccountEmail,
  onUpdateProject,
  onImportProjectRules,
}) => {
  const activeProject = useMemo(() => {
    if (projects.length === 0) return null;
    return projects.find((p) => p.code === selectedProjectCode) ?? projects[0];
  }, [projects, selectedProjectCode]);

  const [senderInput, setSenderInput] = useState('');
  const [domainInput, setDomainInput] = useState('');
  const [accountInput, setAccountInput] = useState('');
  const [keywordInput, setKeywordInput] = useState('');
  const [notice, setNotice] = useState<string | null>(null);
  const importInputRef = useRef<HTMLInputElement | null>(null);

  const suggestions = useMemo(
    () => discoverRoutingSuggestions(emails, ownerAccountEmail),
    [emails, ownerAccountEmail]
  );

  // Live preview of how many emails and files match the active project
  const liveMatchPreview = useMemo(() => {
    if (!activeProject) return null;
    const bySource = { sender: 0, domain: 0, account: 0, 'code-or-keyword': 0, manual: 0 };
    let matchedEmails = 0;

    for (const mail of emails) {
      const res = classifyByProjectRoutingRules({
        senderEmail: mail.fromEmail,
        recipientEmails: [...(mail.to ?? []), ...(mail.cc ?? [])],
        accountEmail: ownerAccountEmail || (mail.folder === 'sent' ? mail.fromEmail : mail.to[0]),
        subject: mail.subject,
        body: mail.body,
        category: mail.folderName || mail.category,
        filenames: (mail.attachments ?? []).map((a) => a.filename),
        projects,
      });
      if (res.projectCode === activeProject.code) {
        matchedEmails += 1;
        if (res.matchSource !== 'unassigned') {
          bySource[res.matchSource] += 1;
        }
      }
    }

    const matchedFiles = organizedFiles.filter((f) => f.projectCode === activeProject.code).length;
    return { matchedEmails, matchedFiles, bySource };
  }, [activeProject, emails, organizedFiles, ownerAccountEmail, projects]);

  const showTempNotice = (msg: string) => {
    setNotice(msg);
    setTimeout(() => setNotice((curr) => (curr === msg ? null : curr)), 3500);
  };

  const addRuleValue = (
    field: 'senderEmails' | 'domains' | 'accounts' | 'keywords',
    rawValue: string
  ) => {
    if (!activeProject) return;
    const parts = rawValue
      .split(/[,;\n]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    if (parts.length === 0) return;

    const existing = activeProject[field] ?? [];
    const merged =
      field === 'keywords'
        ? Array.from(new Set([...existing, ...parts]))
        : normalizeRuleList(
            [...existing, ...parts],
            field === 'senderEmails' ? 'email' : field === 'domains' ? 'domain' : 'account'
          );

    onUpdateProject(activeProject.id, { [field]: merged });
    showTempNotice(`${activeProject.code} 규칙을 저장했습니다.`);
  };

  const removeRuleValue = (
    field: 'senderEmails' | 'domains' | 'accounts' | 'keywords',
    valueToRemove: string
  ) => {
    if (!activeProject) return;
    const existing = activeProject[field] ?? [];
    onUpdateProject(activeProject.id, {
      [field]: existing.filter((v) => v !== valueToRemove),
    });
  };

  const handleExportRulesJson = () => {
    const blob = new Blob([JSON.stringify(projects, null, 2)], {
      type: 'application/json;charset=utf-8',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `bmail-project-routing-rules-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showTempNotice('프로젝트별 송신인·도메인·계정 설정 JSON을 내보냈습니다.');
  };

  const handleImportRulesJson = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as unknown;
      if (!Array.isArray(parsed)) {
        showTempNotice('유효한 프로젝트 설정 배열 JSON이 아닙니다.');
        return;
      }
      const cleaned: ProjectDefinition[] = parsed
        .filter((item): item is Record<string, unknown> => Boolean(item && typeof item === 'object'))
        .map((item, idx) => ({
          id: String(item.id || `prj-import-${Date.now()}-${idx}`),
          code: String(item.code || '').trim().toUpperCase(),
          name: String(item.name || '').trim(),
          keywords: Array.isArray(item.keywords) ? item.keywords.map(String).filter(Boolean) : [],
          color: typeof item.color === 'string' ? item.color : '#0F6CBD',
          isFavorite: Boolean(item.isFavorite),
          createdAt: typeof item.createdAt === 'string' ? item.createdAt : new Date().toISOString(),
          senderEmails: normalizeRuleList(
            Array.isArray(item.senderEmails) ? item.senderEmails.map(String) : [],
            'email'
          ),
          domains: normalizeRuleList(
            Array.isArray(item.domains) ? item.domains.map(String) : [],
            'domain'
          ),
          accounts: normalizeRuleList(
            Array.isArray(item.accounts) ? item.accounts.map(String) : [],
            'account'
          ),
          folderTemplate: item.folderTemplate as ProjectFolderPreset | undefined,
          defaultSubfolder:
            typeof item.defaultSubfolder === 'string' ? item.defaultSubfolder.trim() : undefined,
        }))
        .filter((p) => Boolean(p.code && p.name));

      if (cleaned.length === 0) {
        showTempNotice('가져올 수 있는 유효한 프로젝트 항목이 없습니다.');
        return;
      }
      onImportProjectRules(cleaned);
      showTempNotice(`프로젝트 설정 ${cleaned.length}건을 불러왔습니다.`);
    } catch {
      showTempNotice('JSON 파일을 읽는 중 오류가 발생했습니다.');
    }
  };

  if (!activeProject) {
    return (
      <div className="flex-1 flex items-center justify-center p-8 bg-white rounded-lg border border-[#E0E0E0]">
        <div className="max-w-md text-center space-y-2">
          <FolderCog className="w-8 h-8 text-[#0F6CBD] mx-auto" />
          <h2 className="text-sm font-bold text-[#242424]">
            먼저 좌측 패널에서 프로젝트를 1개 이상 등록해 주세요
          </h2>
          <p className="text-xs text-[#616161] leading-relaxed">
            프로젝트를 등록하면 각 프로젝트별로 송신인 이메일, 거래처 도메인(@domain), 연결 메일 계정, 수신파일 기본 폴더 구조를 세밀하게 세팅할 수 있습니다.
          </p>
        </div>
      </div>
    );
  }

  const senders = activeProject.senderEmails ?? [];
  const domains = activeProject.domains ?? [];
  const accounts = activeProject.accounts ?? [];
  const keywords = activeProject.keywords ?? [];

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-white border border-[#E0E0E0] rounded-lg overflow-hidden">
      {/* Header */}
      <div className="px-4 py-3 border-b border-[#EDEBE9] flex flex-wrap items-center justify-between gap-2 bg-[#FAF9F8]">
        <div className="flex items-center gap-2 min-w-0">
          <span
            className="w-3 h-3 rounded-full shrink-0"
            style={{ backgroundColor: activeProject.color }}
          />
          <span className="font-mono text-xs font-bold text-[#0F6CBD]">{activeProject.code}</span>
          <h2 className="text-sm font-bold text-[#242424] truncate">{activeProject.name}</h2>
          <span className="text-[11px] text-[#616161] hidden sm:inline">
            · 송신인 → 도메인 → 계정 → 코드/키워드 순 자동 분류
          </span>
        </div>

        <div className="flex items-center gap-1.5">
          <select
            aria-label="세팅 대상 프로젝트 선택"
            value={activeProject.code}
            onChange={(e) => onSelectProjectCode(e.target.value)}
            className="px-2 py-1 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
          >
            {projects.map((p) => (
              <option key={p.id} value={p.code}>
                {p.code} · {p.name}
              </option>
            ))}
          </select>

          <button
            type="button"
            onClick={handleExportRulesJson}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F3F2F1] rounded-[4px]"
          >
            <Download className="w-3.5 h-3.5 text-[#0F6CBD]" />
            설정 내보내기
          </button>

          <button
            type="button"
            onClick={() => importInputRef.current?.click()}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F3F2F1] rounded-[4px]"
          >
            <Upload className="w-3.5 h-3.5 text-[#0F6CBD]" />
            가져오기
          </button>
          <input
            ref={importInputRef}
            type="file"
            accept=".json,application/json"
            onChange={(e) => void handleImportRulesJson(e)}
            className="hidden"
          />
        </div>
      </div>

      {notice && (
        <div className="px-4 py-1.5 bg-[#EBF3FC] border-b border-[#B4D6FA] text-xs font-medium text-[#0F6CBD]">
          {notice}
        </div>
      )}

      {/* Live Match Summary Bar */}
      {liveMatchPreview && (
        <div className="px-4 py-2.5 border-b border-[#EDEBE9] bg-white flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex flex-wrap items-center gap-3">
            <span className="inline-flex items-center gap-1 font-semibold text-[#242424]">
              <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD]" />
              실시간 매칭 결과:
            </span>
            <span>
              메일 <strong className="font-mono text-[#0F6CBD]">{liveMatchPreview.matchedEmails}</strong>건
            </span>
            <span>
              정리 파일 <strong className="font-mono text-[#0F6CBD]">{liveMatchPreview.matchedFiles}</strong>건
            </span>
          </div>
          <div className="flex flex-wrap items-center gap-2 text-[11px] text-[#616161] font-mono">
            <span>송신인:{liveMatchPreview.bySource.sender}</span>
            <span>도메인:{liveMatchPreview.bySource.domain}</span>
            <span>계정:{liveMatchPreview.bySource.account}</span>
            <span>키워드:{liveMatchPreview.bySource['code-or-keyword']}</span>
          </div>
        </div>
      )}

      {/* Main Scrollable Settings Body */}
      <div className="flex-1 overflow-y-auto p-4 space-y-5">
        {/* 1. Sender Emails */}
        <section className="space-y-2 pb-4 border-b border-[#EDEBE9]">
          <div className="flex items-center justify-between">
            <label
              htmlFor="routing-sender-input"
              className="text-xs font-bold text-[#242424] flex items-center gap-1.5"
            >
              <UserCheck className="w-4 h-4 text-[#0F6CBD]" />
              1. 프로젝트별 송신인(발신·수신 담당자) 이메일 세팅
              <span className="font-mono text-[11px] text-[#0F6CBD]">({senders.length})</span>
            </label>
            <span className="text-[11px] text-[#616161]">우선순위 1 · 정확 이메일 일치</span>
          </div>

          <div className="flex gap-2">
            <input
              id="routing-sender-input"
              type="text"
              value={senderInput}
              onChange={(e) => setSenderInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  addRuleValue('senderEmails', senderInput);
                  setSenderInput('');
                }
              }}
              placeholder="송신인 이메일 입력 (쉼표 구분, 예: kim@vendor.co.kr, buyer@client.com)"
              className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px] focus:outline-none focus:border-[#0F6CBD]"
            />
            <button
              type="button"
              onClick={() => {
                addRuleValue('senderEmails', senderInput);
                setSenderInput('');
              }}
              disabled={!senderInput.trim()}
              className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[4px]"
            >
              <Plus className="w-3.5 h-3.5" />
              추가
            </button>
          </div>

          <div className="flex flex-wrap gap-1.5 min-h-[28px]">
            {senders.length === 0 ? (
              <span className="text-[11px] text-[#8A8886]">
                등록된 송신인 이메일이 없습니다.
              </span>
            ) : (
              senders.map((email) => (
                <span
                  key={email}
                  className="inline-flex items-center gap-1.5 px-2 py-0.5 bg-[#EBF3FC] border border-[#B4D6FA] text-[#0F6CBD] text-xs font-mono rounded-[4px]"
                >
                  {email}
                  <button
                    type="button"
                    aria-label={`${email} 송신인 제거`}
                    onClick={() => removeRuleValue('senderEmails', email)}
                    className="hover:text-red-700"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))
            )}
          </div>

          {suggestions.senders.length > 0 && (
            <div className="pt-1 flex flex-wrap items-center gap-1.5">
              <span className="text-[11px] text-[#616161]">수신 메일에서 발견된 송신인 빠른 추가:</span>
              {suggestions.senders
                .filter((s) => !senders.includes(s))
                .slice(0, 6)
                .map((s) => (
                  <button
                    key={s}
                    type="button"
                    onClick={() => addRuleValue('senderEmails', s)}
                    className="px-2 py-0.5 text-[11px] font-mono bg-[#FAF9F8] hover:bg-[#EBF3FC] border border-[#D1D1D1] text-[#323130] rounded-[4px]"
                  >
                    + {s}
                  </button>
                ))}
            </div>
          )}
        </section>

        {/* 2. Domains */}
        <section className="space-y-2 pb-4 border-b border-[#EDEBE9]">
          <div className="flex items-center justify-between">
            <label
              htmlFor="routing-domain-input"
              className="text-xs font-bold text-[#242424] flex items-center gap-1.5"
            >
              <Globe className="w-4 h-4 text-[#0F6CBD]" />
              2. 프로젝트별 거래처 · 협력사 도메인 세팅
              <span className="font-mono text-[11px] text-[#0F6CBD]">({domains.length})</span>
            </label>
            <span className="text-[11px] text-[#616161]">우선순위 2 · 도메인 자동 매핑</span>
          </div>

          <div className="flex gap-2">
            <input
              id="routing-domain-input"
              type="text"
              value={domainInput}
              onChange={(e) => setDomainInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  addRuleValue('domains', domainInput);
                  setDomainInput('');
                }
              }}
              placeholder="도메인 입력 (예: vendor-a.co.kr, @shipyard.com)"
              className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px] focus:outline-none focus:border-[#0F6CBD]"
            />
            <button
              type="button"
              onClick={() => {
                addRuleValue('domains', domainInput);
                setDomainInput('');
              }}
              disabled={!domainInput.trim()}
              className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[4px]"
            >
              <Plus className="w-3.5 h-3.5" />
              추가
            </button>
          </div>

          <div className="flex flex-wrap gap-1.5 min-h-[28px]">
            {domains.length === 0 ? (
              <span className="text-[11px] text-[#8A8886]">
                등록된 도메인이 없습니다. 거래처 도메인을 등록하면 해당 회사의 모든 수신 파일이 이 프로젝트로 자동 분류됩니다.
              </span>
            ) : (
              domains.map((dom) => (
                <span
                  key={dom}
                  className="inline-flex items-center gap-1.5 px-2 py-0.5 bg-[#ECFDF5] border border-[#A7F3D0] text-[#065F46] text-xs font-mono rounded-[4px]"
                >
                  @{dom}
                  <button
                    type="button"
                    aria-label={`${dom} 도메인 제거`}
                    onClick={() => removeRuleValue('domains', dom)}
                    className="hover:text-red-700"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))
            )}
          </div>

          {suggestions.domains.length > 0 && (
            <div className="pt-1 flex flex-wrap items-center gap-1.5">
              <span className="text-[11px] text-[#616161]">발견된 도메인 빠른 추가:</span>
              {suggestions.domains
                .filter((d) => !domains.includes(d))
                .slice(0, 6)
                .map((d) => (
                  <button
                    key={d}
                    type="button"
                    onClick={() => addRuleValue('domains', d)}
                    className="px-2 py-0.5 text-[11px] font-mono bg-[#FAF9F8] hover:bg-[#ECFDF5] border border-[#D1D1D1] text-[#323130] rounded-[4px]"
                  >
                    + @{d}
                  </button>
                ))}
            </div>
          )}
        </section>

        {/* 3. Connected Accounts */}
        <section className="space-y-2 pb-4 border-b border-[#EDEBE9]">
          <div className="flex items-center justify-between">
            <label
              htmlFor="routing-account-input"
              className="text-xs font-bold text-[#242424] flex items-center gap-1.5"
            >
              <AtSign className="w-4 h-4 text-[#0F6CBD]" />
              3. 프로젝트별 연결 메일 계정(수신·발신 계정) 세팅
              <span className="font-mono text-[11px] text-[#0F6CBD]">({accounts.length})</span>
            </label>
            <span className="text-[11px] text-[#616161]">우선순위 3 · 멀티 계정 분리 관리</span>
          </div>

          <div className="flex gap-2">
            <input
              id="routing-account-input"
              type="text"
              value={accountInput}
              onChange={(e) => setAccountInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  addRuleValue('accounts', accountInput);
                  setAccountInput('');
                }
              }}
              placeholder="프로젝트 전용 수신/발신 계정 입력 (예: sales@company.co.kr, design-h101@company.co.kr)"
              className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px] focus:outline-none focus:border-[#0F6CBD]"
            />
            <button
              type="button"
              onClick={() => {
                addRuleValue('accounts', accountInput);
                setAccountInput('');
              }}
              disabled={!accountInput.trim()}
              className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[4px]"
            >
              <Plus className="w-3.5 h-3.5" />
              추가
            </button>
          </div>

          <div className="flex flex-wrap gap-1.5 min-h-[28px]">
            {accounts.length === 0 ? (
              <span className="text-[11px] text-[#8A8886]">
                등록된 연결 계정이 없습니다. Outlook·Mailplug·M365 계정을 지정하면 해당 계정으로 송수신된 파일이 자동 분류됩니다.
              </span>
            ) : (
              accounts.map((acc) => (
                <span
                  key={acc}
                  className="inline-flex items-center gap-1.5 px-2 py-0.5 bg-[#FFFBEB] border border-[#FDE68A] text-[#92400E] text-xs font-mono rounded-[4px]"
                >
                  <Mail className="w-3 h-3" />
                  {acc}
                  <button
                    type="button"
                    aria-label={`${acc} 계정 제거`}
                    onClick={() => removeRuleValue('accounts', acc)}
                    className="hover:text-red-700"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))
            )}
          </div>

          {suggestions.accounts.length > 0 && (
            <div className="pt-1 flex flex-wrap items-center gap-1.5">
              <span className="text-[11px] text-[#616161]">연결된 계정 빠른 추가:</span>
              {suggestions.accounts
                .filter((a) => !accounts.includes(a))
                .slice(0, 6)
                .map((a) => (
                  <button
                    key={a}
                    type="button"
                    onClick={() => addRuleValue('accounts', a)}
                    className="px-2 py-0.5 text-[11px] font-mono bg-[#FAF9F8] hover:bg-[#FFFBEB] border border-[#D1D1D1] text-[#323130] rounded-[4px]"
                  >
                    + {a}
                  </button>
                ))}
            </div>
          )}
        </section>

        {/* 4. Keywords + Folder Filing Template */}
        <section className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="space-y-2">
            <label
              htmlFor="routing-keyword-input"
              className="text-xs font-bold text-[#242424] block"
            >
              4. 보조 키워드 (우선순위 4 · 제목/본문/첨부명)
            </label>
            <div className="flex gap-2">
              <input
                id="routing-keyword-input"
                type="text"
                value={keywordInput}
                onChange={(e) => setKeywordInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    addRuleValue('keywords', keywordInput);
                    setKeywordInput('');
                  }
                }}
                placeholder="키워드 추가 (쉼표 구분)"
                className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px] focus:outline-none focus:border-[#0F6CBD]"
              />
              <button
                type="button"
                onClick={() => {
                  addRuleValue('keywords', keywordInput);
                  setKeywordInput('');
                }}
                disabled={!keywordInput.trim()}
                className="px-3 py-1.5 text-xs font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] disabled:opacity-40 rounded-[4px]"
              >
                추가
              </button>
            </div>
            <div className="flex flex-wrap gap-1.5">
              {keywords.map((kw) => (
                <span
                  key={kw}
                  className="inline-flex items-center gap-1 px-2 py-0.5 bg-[#F3F2F1] border border-[#D1D1D1] text-[#323130] text-xs rounded-[4px]"
                >
                  {kw}
                  <button
                    type="button"
                    aria-label={`${kw} 키워드 제거`}
                    onClick={() => removeRuleValue('keywords', kw)}
                    className="hover:text-red-700"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))}
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-xs font-bold text-[#242424] block">
              5. 이 프로젝트의 수신파일 기본 폴더 구조
            </label>
            <select
              aria-label="프로젝트 기본 폴더 템플릿"
              value={activeProject.folderTemplate ?? ''}
              onChange={(e) =>
                onUpdateProject(activeProject.id, {
                  folderTemplate: (e.target.value || undefined) as ProjectFolderPreset | undefined,
                })
              }
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            >
              <option value="">전체 기본 폴더 구조 따름</option>
              {(Object.keys(FOLDER_PRESET_META) as ProjectFolderPreset[]).map((presetKey) => (
                <option key={presetKey} value={presetKey}>
                  {FOLDER_PRESET_META[presetKey].label} ({FOLDER_PRESET_META[presetKey].patternLabel})
                </option>
              ))}
            </select>

            <div>
              <label
                htmlFor="routing-default-subfolder"
                className="text-[11px] font-medium text-[#616161] block mb-1"
              >
                프로젝트 기본 하위 폴더명 (선택, 예: 01_수신자료)
              </label>
              <input
                id="routing-default-subfolder"
                type="text"
                value={activeProject.defaultSubfolder ?? ''}
                onChange={(e) =>
                  onUpdateProject(activeProject.id, {
                    defaultSubfolder: e.target.value,
                  })
                }
                placeholder="예: 01_수신자료"
                className="w-full px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px] focus:outline-none focus:border-[#0F6CBD]"
              />
            </div>
          </div>
        </section>
      </div>
    </div>
  );
};
