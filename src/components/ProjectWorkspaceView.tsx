import React, { useMemo, useState } from 'react';
import {
  Briefcase,
  Download,
  FileBarChart2,
  FileSpreadsheet,
  FolderCog,
  FolderPlus,
  FolderTree,
  Mail,
  Paperclip,
  Sparkles,
  Star,
  Trash2,
} from 'lucide-react';
import type { AnalysisResult, MailItem, ParsedDocumentSummary } from '../types/mail';
import {
  buildDeterministicWorkLog,
  type ProjectDefinition,
  type ProjectFolderPreset,
} from '../utils/deterministicAnalyzer';
import {
  classifyByProjectRoutingRules,
  collectOrganizedFiles,
  normalizeRuleList,
  type FileOrganizerFilter,
  type ManualOrganizedFile,
} from '../utils/projectFileOrganizer';
import { ProjectFileFolderOrganizer } from './ProjectFileFolderOrganizer';
import { ProjectFileReportPanel } from './ProjectFileReportPanel';
import { ProjectRoutingSettingsPanel } from './ProjectRoutingSettingsPanel';

interface ProjectWorkspaceViewProps {
  emails: MailItem[];
  documents: ParsedDocumentSummary[];
  projects: ProjectDefinition[];
  ownerAccountEmail?: string;
  onAddProject: (project: Omit<ProjectDefinition, 'id' | 'createdAt'>) => void;
  onUpdateProject: (id: string, patch: Partial<ProjectDefinition>) => void;
  onImportProjectRules?: (imported: ProjectDefinition[]) => void;
  onToggleFavoriteProject: (id: string) => void;
  onDeleteProject: (id: string) => void;
  onApplyDeterministicWorkLog: (result: AnalysisResult) => void;
  onExportMbox: (subset: MailItem[], filename: string) => void;
  onSaveDraftToOutlook?: (subject: string, body: string) => void;
}

const PROJECT_COLORS = ['#0F6CBD', '#107C10', '#B45309', '#6B21A8', '#C50F1F', '#0369A1'];
const MANUAL_FILES_STORAGE_KEY = 'bmail_project_manual_files_v1';
const FILE_OVERRIDES_STORAGE_KEY = 'bmail_project_file_overrides_v1';
const FOLDER_PRESET_STORAGE_KEY = 'bmail_project_folder_preset_v1';

type WorkspaceTab = 'report' | 'folders' | 'routing' | 'ledger';

function loadStorage<T>(key: string, fallback: T): T {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return fallback;
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

export const ProjectWorkspaceView: React.FC<ProjectWorkspaceViewProps> = ({
  emails,
  documents,
  projects,
  ownerAccountEmail,
  onAddProject,
  onUpdateProject,
  onImportProjectRules,
  onToggleFavoriteProject,
  onDeleteProject,
  onApplyDeterministicWorkLog,
  onExportMbox,
  onSaveDraftToOutlook,
}) => {
  const [activeTab, setActiveTab] = useState<WorkspaceTab>('report');
  const [selectedCode, setSelectedCode] = useState<string>('ALL');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [templateName, setTemplateName] = useState<'standard' | 'project-case' | 'executive'>('project-case');

  // Create Project Form
  const [newCode, setNewCode] = useState('');
  const [newName, setNewName] = useState('');
  const [newKeywords, setNewKeywords] = useState('');

  // File Organizer & Routing Persistence
  const [folderPreset, setFolderPreset] = useState<ProjectFolderPreset>(() =>
    loadStorage<ProjectFolderPreset>(FOLDER_PRESET_STORAGE_KEY, 'project-month-sender')
  );
  const [manualFiles, setManualFiles] = useState<ManualOrganizedFile[]>(() =>
    loadStorage<ManualOrganizedFile[]>(MANUAL_FILES_STORAGE_KEY, [])
  );
  const [fileProjectOverrides, setFileProjectOverrides] = useState<Record<string, string>>(() =>
    loadStorage<Record<string, string>>(FILE_OVERRIDES_STORAGE_KEY, {})
  );
  const [reportFilter, setReportFilter] = useState<FileOrganizerFilter>({
    direction: 'all',
    fileKind: 'all',
  });

  const handleChangePreset = (next: ProjectFolderPreset) => {
    setFolderPreset(next);
    try {
      window.localStorage.setItem(FOLDER_PRESET_STORAGE_KEY, JSON.stringify(next));
    } catch {
      // ignore storage quota errors
    }
  };

  const handleAddManualFiles = (items: ManualOrganizedFile[]) => {
    setManualFiles((prev) => {
      const next = [...items, ...prev];
      try {
        window.localStorage.setItem(MANUAL_FILES_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  };

  const handleRemoveManualFile = (fileId: string) => {
    setManualFiles((prev) => {
      const next = prev.filter((f) => f.id !== fileId);
      try {
        window.localStorage.setItem(MANUAL_FILES_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  };

  const handleReassignFileProject = (fileId: string, projectCode: string) => {
    setFileProjectOverrides((prev) => {
      const next = { ...prev, [fileId]: projectCode };
      try {
        window.localStorage.setItem(FILE_OVERRIDES_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  };

  const handleQuickBindRule = (params: {
    projectCode: string;
    senderEmail?: string;
    domain?: string;
  }) => {
    const target = projects.find((p) => p.code === params.projectCode);
    if (!target) return;
    const patch: Partial<ProjectDefinition> = {};
    if (params.senderEmail) {
      patch.senderEmails = normalizeRuleList(
        [...(target.senderEmails ?? []), params.senderEmail],
        'email'
      );
    }
    if (params.domain) {
      patch.domains = normalizeRuleList([...(target.domains ?? []), params.domain], 'domain');
    }
    onUpdateProject(target.id, patch);
  };

  // Classify every email using the full Sender -> Domain -> Account -> Code/Keyword routing engine
  const classifiedEmails = useMemo(() => {
    return emails.map((mail) => {
      const classification = classifyByProjectRoutingRules({
        senderEmail: mail.fromEmail,
        recipientEmails: [...(mail.to ?? []), ...(mail.cc ?? [])],
        accountEmail: ownerAccountEmail || (mail.folder === 'sent' ? mail.fromEmail : mail.to[0]),
        subject: mail.subject,
        body: mail.body,
        category: mail.folderName || mail.category,
        filenames: (mail.attachments ?? []).map((a) => a.filename),
        projects,
      });
      return {
        mail,
        code: classification.projectCode === 'UNASSIGNED' ? 'UNCLASSIFIED' : classification.projectCode,
        matchSource: classification.matchSource,
        matchEvidence: classification.matchEvidence,
      };
    });
  }, [emails, ownerAccountEmail, projects]);

  // All organized files (unfiltered for counts, filtered for active view)
  const allOrganizedFiles = useMemo(
    () =>
      collectOrganizedFiles({
        emails,
        manualFiles,
        projects,
        preset: folderPreset,
        ownerAccountEmail,
        fileProjectOverrides,
      }),
    [emails, manualFiles, projects, folderPreset, ownerAccountEmail, fileProjectOverrides]
  );

  const filteredOrganizedFiles = useMemo(
    () =>
      collectOrganizedFiles({
        emails,
        manualFiles,
        projects,
        preset: folderPreset,
        ownerAccountEmail,
        fileProjectOverrides,
        filter: {
          ...reportFilter,
          projectCode: selectedCode,
          dateFrom: startDate || undefined,
          dateTo: endDate || undefined,
        },
      }),
    [
      emails,
      manualFiles,
      projects,
      folderPreset,
      ownerAccountEmail,
      fileProjectOverrides,
      reportFilter,
      selectedCode,
      startDate,
      endDate,
    ]
  );

  const discoveredCounts = useMemo(() => {
    const mailCounts = new Map<string, number>();
    const fileCounts = new Map<string, number>();
    for (const item of classifiedEmails) {
      mailCounts.set(item.code, (mailCounts.get(item.code) ?? 0) + 1);
    }
    for (const f of allOrganizedFiles) {
      fileCounts.set(f.projectCode, (fileCounts.get(f.projectCode) ?? 0) + 1);
    }
    return { mailCounts, fileCounts };
  }, [classifiedEmails, allOrganizedFiles]);

  const filteredLedger = useMemo(() => {
    return classifiedEmails
      .filter(({ mail, code }) => {
        if (selectedCode !== 'ALL' && code !== selectedCode) return false;
        if (startDate && mail.date < startDate) return false;
        if (endDate && mail.date > endDate) return false;
        return true;
      })
      .sort((a, b) => b.mail.timestamp.localeCompare(a.mail.timestamp));
  }, [classifiedEmails, selectedCode, startDate, endDate]);

  const handleCreateProject = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanCode = newCode.trim().toUpperCase();
    const cleanName = newName.trim();
    if (!cleanCode || !cleanName) return;
    onAddProject({
      code: cleanCode,
      name: cleanName,
      keywords: newKeywords
        .split(',')
        .map((k) => k.trim())
        .filter(Boolean),
      color: PROJECT_COLORS[projects.length % PROJECT_COLORS.length],
      isFavorite: true,
      senderEmails: [],
      domains: [],
      accounts: [],
    });
    setNewCode('');
    setNewName('');
    setNewKeywords('');
    setSelectedCode(cleanCode);
  };

  const handleGenerateOfflineWorkLog = () => {
    const subset = filteredLedger.map((i) => i.mail);
    if (subset.length === 0) return;
    const periodLabel =
      startDate || endDate
        ? `${startDate || '전체'} ~ ${endDate || '현재'} (${selectedCode})`
        : `${selectedCode === 'ALL' ? '전체 프로젝트' : selectedCode}`;
    const result = buildDeterministicWorkLog({
      emails: subset,
      documents,
      projects,
      periodLabel,
      templateName,
    });
    onApplyDeterministicWorkLog(result);
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-[#F5F5F5] overflow-hidden">
      {/* Top Workspace Header + 4 Pivot Tabs */}
      <div className="bg-white border-b border-[#E0E0E0] px-5 pt-3 shrink-0">
        <div className="flex flex-wrap items-center justify-between gap-3 pb-2.5">
          <div className="space-y-0.5">
            <h1 className="text-base font-bold text-[#242424]">
              프로젝트별 · 일자별 · 송수신인별 파일 정리 보고서 & 규칙 세팅
            </h1>
            <p className="text-xs text-[#616161]">
              프로젝트별 송신인·도메인·계정 자동 분류 규칙 → 수신파일 폴더 트리 자동 정리 → 보고서(.md/.csv) 및 사건 기록부 추출
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 text-xs">
            <input
              type="date"
              aria-label="분석 시작일"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              className="px-2 py-1 border border-[#D1D1D1] rounded-[4px] text-xs bg-white"
            />
            <span className="text-[#616161]">~</span>
            <input
              type="date"
              aria-label="분석 종료일"
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              className="px-2 py-1 border border-[#D1D1D1] rounded-[4px] text-xs bg-white"
            />
            {(startDate || endDate) && (
              <button
                type="button"
                onClick={() => {
                  setStartDate('');
                  setEndDate('');
                }}
                className="text-[#0F6CBD] hover:underline font-medium"
              >
                기간 해제
              </button>
            )}
          </div>
        </div>

        {/* Outlook Fluent 2 Pivot Tabs */}
        <div role="tablist" aria-label="프로젝트 파일 정리 워크스페이스 탭" className="flex items-center gap-5 overflow-x-auto">
          {[
            {
              id: 'report' as const,
              label: '종합 분석 보고서',
              badge: allOrganizedFiles.length,
              icon: <FileBarChart2 className="w-4 h-4" />,
            },
            {
              id: 'folders' as const,
              label: '수신파일 폴더 정리',
              badge: allOrganizedFiles.length,
              icon: <FolderTree className="w-4 h-4" />,
            },
            {
              id: 'routing' as const,
              label: '송신인·도메인·계정 세팅',
              badge: projects.length,
              icon: <FolderCog className="w-4 h-4" />,
            },
            {
              id: 'ledger' as const,
              label: '사건 기록부 · 업무 일지',
              badge: emails.length,
              icon: <Briefcase className="w-4 h-4" />,
            },
          ].map((tab) => {
            const active = activeTab === tab.id;
            return (
              <button
                key={tab.id}
                type="button"
                role="tab"
                aria-selected={active}
                onClick={() => setActiveTab(tab.id)}
                className={`inline-flex items-center gap-1.5 pb-2.5 pt-1 text-xs border-b-2 transition-colors whitespace-nowrap shrink-0 ${
                  active
                    ? 'border-[#0F6CBD] text-[#0F6CBD] font-bold'
                    : 'border-transparent text-[#616161] hover:text-[#242424] font-medium'
                }`}
              >
                {tab.icon}
                <span>{tab.label}</span>
                <span className="font-mono text-[10px] text-[#616161]">({tab.badge})</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* Main Grid: Left Project Registry + Right Active Tab Panel */}
      <div className="flex-1 min-h-0 p-3 grid grid-cols-1 lg:grid-cols-12 gap-3 overflow-hidden">
        {/* Left: Project Registry & Quick Create */}
        <aside className="lg:col-span-3 bg-white border border-[#E0E0E0] rounded-lg flex flex-col overflow-hidden">
          <div className="px-3.5 py-2.5 border-b border-[#EDEBE9] font-bold text-xs text-[#242424] flex items-center justify-between">
            <span>등록 프로젝트 ({projects.length})</span>
            <button
              type="button"
              onClick={() => setSelectedCode('ALL')}
              className={`text-[11px] font-semibold ${
                selectedCode === 'ALL' ? 'text-[#0F6CBD] underline' : 'text-[#616161]'
              }`}
            >
              전체 ({allOrganizedFiles.length}파일)
            </button>
          </div>

          <div className="flex-1 overflow-y-auto p-2.5 space-y-1.5">
            {projects.length === 0 ? (
              <div className="p-3.5 rounded-lg bg-[#FAF9F8] border border-[#EDEBE9] text-xs text-[#616161] leading-relaxed">
                아직 등록된 프로젝트가 없습니다. 아래에서 프로젝트(예: <code className="font-mono">PRJ-H101</code>)를 추가한 뒤 송신인·도메인·계정 규칙을 세팅하세요.
              </div>
            ) : (
              projects.map((p) => {
                const mailCnt = discoveredCounts.mailCounts.get(p.code) ?? 0;
                const fileCnt = discoveredCounts.fileCounts.get(p.code) ?? 0;
                const ruleCount =
                  (p.senderEmails?.length ?? 0) +
                  (p.domains?.length ?? 0) +
                  (p.accounts?.length ?? 0);
                const active = selectedCode === p.code;
                return (
                  <div
                    key={p.id}
                    className={`flex items-start justify-between gap-1.5 px-2.5 py-2 rounded-md border transition-colors ${
                      active
                        ? 'bg-[#EBF3FC] border-[#0F6CBD]'
                        : 'bg-white border-[#EDEBE9] hover:bg-[#FAF9F8]'
                    }`}
                  >
                    <button
                      type="button"
                      onClick={() => setSelectedCode(p.code)}
                      className="flex-1 text-left min-w-0"
                    >
                      <div className="flex items-center gap-1.5">
                        <span
                          className="w-2.5 h-2.5 rounded-full shrink-0"
                          style={{ backgroundColor: p.color }}
                        />
                        <span className="text-xs font-bold font-mono text-[#0F6CBD] shrink-0">
                          {p.code}
                        </span>
                        <span className="text-xs font-semibold text-[#242424] truncate">
                          {p.name}
                        </span>
                      </div>
                      <div className="text-[10px] font-mono text-[#616161] pl-4 mt-0.5 flex flex-wrap gap-x-2">
                        <span>파일 {fileCnt}</span>
                        <span>메일 {mailCnt}</span>
                        <span>규칙 {ruleCount}</span>
                      </div>
                    </button>

                    <div className="flex items-center gap-0.5 shrink-0">
                      <button
                        type="button"
                        onClick={() => {
                          setSelectedCode(p.code);
                          setActiveTab('routing');
                        }}
                        title="이 프로젝트의 송신인·도메인·계정 세팅"
                        aria-label={`${p.code} 규칙 세팅`}
                        className="p-1 rounded text-[#616161] hover:text-[#0F6CBD] hover:bg-white"
                      >
                        <FolderCog className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => onToggleFavoriteProject(p.id)}
                        aria-label="즐겨찾기 전환"
                        className={`p-1 rounded ${
                          p.isFavorite ? 'text-amber-500' : 'text-[#9E9E9E] hover:text-amber-500'
                        }`}
                      >
                        <Star className="w-3.5 h-3.5" fill={p.isFavorite ? 'currentColor' : 'none'} />
                      </button>
                      <button
                        type="button"
                        onClick={() => onDeleteProject(p.id)}
                        aria-label="프로젝트 삭제"
                        className="p-1 rounded text-[#9E9E9E] hover:text-red-700"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                );
              })
            )}

            {(discoveredCounts.fileCounts.get('UNASSIGNED') ?? 0) > 0 && (
              <button
                type="button"
                onClick={() => setSelectedCode('UNASSIGNED')}
                className={`w-full flex items-center justify-between px-2.5 py-2 rounded-md border text-xs transition-colors ${
                  selectedCode === 'UNASSIGNED'
                    ? 'bg-amber-50 border-amber-400 text-amber-900 font-semibold'
                    : 'bg-[#FAF9F8] border-[#EDEBE9] text-[#616161] hover:bg-amber-50/50'
                }`}
              >
                <span>미분류 파일/메일</span>
                <span className="font-mono text-[11px]">
                  {discoveredCounts.fileCounts.get('UNASSIGNED') ?? 0}파일
                </span>
              </button>
            )}
          </div>

          {/* Add New Project Form */}
          <form
            onSubmit={handleCreateProject}
            className="p-3 border-t border-[#EDEBE9] bg-[#FAF9F8] space-y-2"
          >
            <div className="text-xs font-bold text-[#242424] flex items-center gap-1.5">
              <FolderPlus className="w-3.5 h-3.5 text-[#0F6CBD]" />새 프로젝트 등록
            </div>
            <div className="grid grid-cols-5 gap-1.5">
              <input
                type="text"
                aria-label="새 프로젝트 코드"
                value={newCode}
                onChange={(e) => setNewCode(e.target.value)}
                placeholder="PRJ-01"
                className="col-span-2 px-2 py-1.5 text-xs font-mono bg-white border border-[#D1D1D1] rounded-[4px]"
              />
              <input
                type="text"
                aria-label="새 프로젝트 이름"
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                placeholder="프로젝트명"
                className="col-span-3 px-2 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
              />
            </div>
            <input
              type="text"
              aria-label="새 프로젝트 키워드"
              value={newKeywords}
              onChange={(e) => setNewKeywords(e.target.value)}
              placeholder="보조 키워드 (쉼표 구분, 선택)"
              className="w-full px-2 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            />
            <button
              type="submit"
              disabled={!newCode.trim() || !newName.trim()}
              className="w-full py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[4px]"
            >
              프로젝트 추가
            </button>
          </form>
        </aside>

        {/* Right: Active Workspace Tab */}
        <div className="lg:col-span-9 flex flex-col min-h-0 overflow-hidden">
          {activeTab === 'report' && (
            <ProjectFileReportPanel
              files={filteredOrganizedFiles}
              projects={projects}
              preset={folderPreset}
              filter={reportFilter}
              onChangeFilter={setReportFilter}
              onSaveDraftToOutlook={onSaveDraftToOutlook}
            />
          )}

          {activeTab === 'folders' && (
            <ProjectFileFolderOrganizer
              files={filteredOrganizedFiles}
              projects={projects}
              preset={folderPreset}
              onChangePreset={handleChangePreset}
              onReassignFileProject={handleReassignFileProject}
              onAddManualFiles={handleAddManualFiles}
              onRemoveManualFile={handleRemoveManualFile}
              onQuickBindRule={handleQuickBindRule}
              ownerAccountEmail={ownerAccountEmail}
            />
          )}

          {activeTab === 'routing' && (
            <ProjectRoutingSettingsPanel
              projects={projects}
              selectedProjectCode={selectedCode === 'ALL' || selectedCode === 'UNASSIGNED' ? projects[0]?.code || '' : selectedCode}
              onSelectProjectCode={setSelectedCode}
              emails={emails}
              organizedFiles={allOrganizedFiles}
              ownerAccountEmail={ownerAccountEmail}
              onUpdateProject={onUpdateProject}
              onImportProjectRules={(imported) => {
                if (onImportProjectRules) {
                  onImportProjectRules(imported);
                } else {
                  for (const item of imported) {
                    const existing = projects.find((p) => p.code === item.code);
                    if (existing) {
                      onUpdateProject(existing.id, item);
                    } else {
                      onAddProject(item);
                    }
                  }
                }
              }}
            />
          )}

          {activeTab === 'ledger' && (
            <section className="flex-1 bg-white border border-[#E0E0E0] rounded-lg flex flex-col overflow-hidden">
              <div className="px-4 py-2.5 border-b border-[#EDEBE9] flex flex-wrap items-center justify-between gap-2 bg-[#FAF9F8]">
                <div className="flex items-center gap-2">
                  <Briefcase className="w-4 h-4 text-[#0F6CBD]" />
                  <h2 className="text-sm font-bold text-[#242424]">
                    사건 기록부 ({selectedCode === 'ALL' ? '전체' : selectedCode}) · {filteredLedger.length}건
                  </h2>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <select
                    aria-label="업무 일지 양식"
                    value={templateName}
                    onChange={(e) =>
                      setTemplateName(e.target.value as 'standard' | 'project-case' | 'executive')
                    }
                    className="px-2 py-1 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
                  >
                    <option value="project-case">프로젝트 사건 기록 양식</option>
                    <option value="standard">표준 일일 업무 일지 양식</option>
                    <option value="executive">임원 보고 요약 양식</option>
                  </select>

                  <button
                    type="button"
                    disabled={filteredLedger.length === 0}
                    onClick={handleGenerateOfflineWorkLog}
                    className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[4px] whitespace-nowrap"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    코드 기반 업무 일지 즉시 생성
                  </button>

                  <button
                    type="button"
                    disabled={filteredLedger.length === 0}
                    onClick={() =>
                      onExportMbox(
                        filteredLedger.map((i) => i.mail),
                        `project-${selectedCode.toLowerCase()}.mbox`
                      )
                    }
                    className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] disabled:text-slate-400 rounded-[4px] whitespace-nowrap"
                  >
                    <Download className="w-3.5 h-3.5" />
                    .mbox 내보내기
                  </button>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto p-4 space-y-2.5">
                {filteredLedger.length === 0 ? (
                  <div className="p-8 text-center rounded-lg bg-[#FAF9F8] border border-[#EDEBE9] text-xs text-[#616161] space-y-1">
                    <FileSpreadsheet className="w-7 h-7 text-[#0F6CBD] mx-auto" />
                    <p>현재 프로젝트/기간 조건에 해당하는 메일 사건 기록이 없습니다.</p>
                    <p>Outlook 메일함에서 원본 .pst/.eml을 불러오거나 기간 필터를 조정해 주세요.</p>
                  </div>
                ) : (
                  filteredLedger.map(({ mail, code, matchEvidence }) => (
                    <div
                      key={mail.id}
                      className="p-3.5 rounded-lg border border-[#E0E0E0] hover:border-[#B4D6FA] transition-colors space-y-1.5"
                    >
                      <div className="flex items-center justify-between gap-2 text-xs">
                        <div className="flex items-center gap-2 min-w-0">
                          <span className="px-1.5 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[10px] font-bold shrink-0">
                            {code === 'UNCLASSIFIED' ? '미분류' : code}
                          </span>
                          <span
                            className={`px-1.5 py-0.5 rounded text-[10px] font-semibold shrink-0 ${
                              mail.folder === 'sent'
                                ? 'bg-emerald-50 text-emerald-700'
                                : 'bg-slate-100 text-[#424242]'
                            }`}
                          >
                            {mail.folder === 'sent' ? '발신' : '수신'}
                          </span>
                          <strong className="text-[#242424] truncate">{mail.subject}</strong>
                        </div>
                        <span className="text-[11px] font-mono text-[#616161] shrink-0">
                          {mail.date} {mail.timeLabel}
                        </span>
                      </div>
                      <div className="flex items-center justify-between gap-2 text-[11px] text-[#616161]">
                        <span className="inline-flex items-center gap-1.5 truncate">
                          <Mail className="w-3 h-3 shrink-0" />
                          {mail.fromName} &lt;{mail.fromEmail}&gt; → {mail.to.join(', ')}
                        </span>
                        <span className="text-[10px] text-[#0F6CBD] shrink-0">{matchEvidence}</span>
                      </div>
                      <p className="text-xs text-[#424242] line-clamp-2">{mail.snippet}</p>
                      {mail.attachments && mail.attachments.length > 0 && (
                        <div className="flex flex-wrap items-center gap-1.5 pt-1">
                          {mail.attachments.map((att) => (
                            <span
                              key={att.id}
                              className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#FAF9F8] border border-[#D1D1D1] text-[11px] text-[#242424]"
                            >
                              <Paperclip className="w-3 h-3 text-[#0F6CBD]" />
                              {att.filename}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  ))
                )}
              </div>
            </section>
          )}
        </div>
      </div>
    </div>
  );
};
