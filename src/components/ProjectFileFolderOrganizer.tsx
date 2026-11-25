import React, { useMemo, useRef, useState } from 'react';
import {
  ChevronDown,
  ChevronRight,
  Download,
  FileSpreadsheet,
  FileText,
  Folder,
  FolderOpen,
  FolderTree,
  Paperclip,
  Plus,
  Terminal,
  Trash2,
  Upload,
  UserPlus,
} from 'lucide-react';
import type { ProjectDefinition, ProjectFolderPreset } from '../utils/deterministicAnalyzer';
import {
  buildFolderOrganizationTree,
  FOLDER_PRESET_META,
  formatFileSize,
  generateFolderOrganizerScript,
  MATCH_SOURCE_LABELS,
  type FolderTreeNode,
  type ManualOrganizedFile,
  type OrganizedFileEntry,
} from '../utils/projectFileOrganizer';

interface ProjectFileFolderOrganizerProps {
  files: OrganizedFileEntry[];
  projects: ProjectDefinition[];
  preset: ProjectFolderPreset;
  onChangePreset: (preset: ProjectFolderPreset) => void;
  onReassignFileProject: (fileId: string, projectCode: string) => void;
  onAddManualFiles: (items: ManualOrganizedFile[]) => void;
  onRemoveManualFile: (fileId: string) => void;
  onQuickBindRule: (params: {
    projectCode: string;
    senderEmail?: string;
    domain?: string;
  }) => void;
  ownerAccountEmail?: string;
}

export const ProjectFileFolderOrganizer: React.FC<ProjectFileFolderOrganizerProps> = ({
  files,
  projects,
  preset,
  onChangePreset,
  onReassignFileProject,
  onAddManualFiles,
  onRemoveManualFile,
  onQuickBindRule,
  ownerAccountEmail,
}) => {
  const [selectedFolderPath, setSelectedFolderPath] = useState<string>('ALL');
  const [collapsedPaths, setCollapsedPaths] = useState<Set<string>>(() => new Set());
  const [showAddForm, setShowAddForm] = useState(false);
  const [uploadSenderName, setUploadSenderName] = useState('');
  const [uploadSenderEmail, setUploadSenderEmail] = useState('');
  const [uploadAccountEmail, setUploadAccountEmail] = useState(ownerAccountEmail || '');
  const [uploadDate, setUploadDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [uploadDirection, setUploadDirection] = useState<'received' | 'sent'>('received');
  const [uploadManualProject, setUploadManualProject] = useState('');
  const [uploadCustomName, setUploadCustomName] = useState('');
  const [notice, setNotice] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const tree = useMemo(() => buildFolderOrganizationTree(files), [files]);

  const visibleFiles = useMemo(() => {
    if (selectedFolderPath === 'ALL') return files;
    return files.filter(
      (f) => f.folderPath === selectedFolderPath || f.folderPath.startsWith(`${selectedFolderPath}/`)
    );
  }, [files, selectedFolderPath]);

  const showTempNotice = (msg: string) => {
    setNotice(msg);
    setTimeout(() => setNotice((curr) => (curr === msg ? null : curr)), 3500);
  };

  const toggleCollapse = (pathKey: string) => {
    setCollapsedPaths((prev) => {
      const next = new Set(prev);
      if (next.has(pathKey)) next.delete(pathKey);
      else next.add(pathKey);
      return next;
    });
  };

  const handleLocalFilePick = (e: React.ChangeEvent<HTMLInputElement>) => {
    const picked = Array.from(e.target.files ?? []);
    e.target.value = '';
    if (picked.length === 0) return;

    const now = new Date().toISOString();
    const created: ManualOrganizedFile[] = picked.map((f, idx) => ({
      id: `man-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 6)}`,
      filename: f.name,
      sizeBytes: f.size,
      date: uploadDate || now.slice(0, 10),
      direction: uploadDirection,
      senderName: uploadSenderName.trim() || uploadSenderEmail.trim() || '로컬 수신 파일',
      senderEmail: uploadSenderEmail.trim(),
      recipientEmails: uploadAccountEmail.trim() ? [uploadAccountEmail.trim()] : [],
      accountEmail: uploadAccountEmail.trim() || ownerAccountEmail,
      manualProjectCode: uploadManualProject || undefined,
      addedAt: now,
    }));
    onAddManualFiles(created);
    setShowAddForm(false);
    showTempNotice(`수신 파일 ${created.length}건을 규칙에 따라 폴더 트리에 자동 정리했습니다.`);
  };

  const handleQuickRegisterEntry = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanFilename = uploadCustomName.trim();
    if (!cleanFilename) return;
    const now = new Date().toISOString();
    onAddManualFiles([
      {
        id: `man-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        filename: cleanFilename,
        sizeBytes: 0,
        date: uploadDate || now.slice(0, 10),
        direction: uploadDirection,
        senderName: uploadSenderName.trim() || uploadSenderEmail.trim() || '직접 등록',
        senderEmail: uploadSenderEmail.trim(),
        recipientEmails: uploadAccountEmail.trim() ? [uploadAccountEmail.trim()] : [],
        accountEmail: uploadAccountEmail.trim() || ownerAccountEmail,
        manualProjectCode: uploadManualProject || undefined,
        addedAt: now,
      },
    ]);
    setUploadCustomName('');
    setShowAddForm(false);
    showTempNotice(`"${cleanFilename}" 파일을 폴더 트리에 정리했습니다.`);
  };

  const downloadTextFile = (content: string, filename: string, mime: string) => {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const renderTreeNode = (node: FolderTreeNode): React.ReactNode => {
    const isCollapsed = collapsedPaths.has(node.path);
    const isSelected = selectedFolderPath === node.path;
    const hasChildren = node.children.length > 0;

    return (
      <div key={node.id}>
        <div
          className={`flex items-center gap-1 px-2 py-1.5 rounded-[4px] text-xs cursor-pointer transition-colors ${
            isSelected
              ? 'bg-[#EBF3FC] text-[#0F6CBD] font-semibold'
              : 'text-[#323130] hover:bg-[#F3F2F1]'
          }`}
          style={{ paddingLeft: `${8 + node.depth * 14}px` }}
          onClick={() => setSelectedFolderPath(node.path)}
        >
          <button
            type="button"
            aria-label={isCollapsed ? `${node.name} 펼치기` : `${node.name} 접기`}
            onClick={(ev) => {
              ev.stopPropagation();
              if (hasChildren) toggleCollapse(node.path);
            }}
            className="w-4 h-4 flex items-center justify-center text-[#616161] shrink-0"
          >
            {hasChildren ? (
              isCollapsed ? (
                <ChevronRight className="w-3.5 h-3.5" />
              ) : (
                <ChevronDown className="w-3.5 h-3.5" />
              )
            ) : (
              <span className="w-3" />
            )}
          </button>
          {isSelected ? (
            <FolderOpen className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />
          ) : (
            <Folder className="w-3.5 h-3.5 text-[#C57B00] shrink-0" />
          )}
          <span className="truncate flex-1">{node.name}</span>
          <span className="font-mono text-[10px] text-[#616161] shrink-0">{node.fileCount}</span>
        </div>

        {hasChildren && !isCollapsed && (
          <div>{node.children.map((child) => renderTreeNode(child))}</div>
        )}
      </div>
    );
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-white border border-[#E0E0E0] rounded-lg overflow-hidden">
      {/* Top Folder Organizer Controls */}
      <div className="px-4 py-2.5 border-b border-[#EDEBE9] bg-[#FAF9F8] flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2">
          <FolderTree className="w-4 h-4 text-[#0F6CBD]" />
          <span className="text-xs font-bold text-[#242424]">폴더 구조 기준:</span>
          <select
            aria-label="수신파일 폴더 구조 프리셋"
            value={preset}
            onChange={(e) => {
              onChangePreset(e.target.value as ProjectFolderPreset);
              setSelectedFolderPath('ALL');
            }}
            className="px-2.5 py-1 text-xs bg-white border border-[#D1D1D1] rounded-[4px] font-medium"
          >
            {(Object.keys(FOLDER_PRESET_META) as ProjectFolderPreset[]).map((key) => (
              <option key={key} value={key}>
                {FOLDER_PRESET_META[key].label}
              </option>
            ))}
          </select>
          <span className="text-[11px] font-mono text-[#616161] hidden xl:inline">
            {FOLDER_PRESET_META[preset].patternLabel}
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            onClick={() => setShowAddForm((v) => !v)}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px]"
          >
            <Plus className="w-3.5 h-3.5" />
            수신 파일 추가
          </button>

          <button
            type="button"
            disabled={files.length === 0}
            onClick={() => {
              const script = generateFolderOrganizerScript(files, 'ps1');
              downloadTextFile(
                script,
                `bmail-organize-folders-${new Date().toISOString().slice(0, 10)}.ps1`,
                'text/plain;charset=utf-8'
              );
              showTempNotice('Windows PowerShell (.ps1) 자동 폴더 정리 스크립트를 다운로드했습니다.');
            }}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F3F2F1] disabled:opacity-40 rounded-[4px]"
          >
            <Terminal className="w-3.5 h-3.5 text-[#0F6CBD]" />
            PowerShell (.ps1)
          </button>

          <button
            type="button"
            disabled={files.length === 0}
            onClick={() => {
              const script = generateFolderOrganizerScript(files, 'sh');
              downloadTextFile(
                script,
                `bmail-organize-folders-${new Date().toISOString().slice(0, 10)}.sh`,
                'text/plain;charset=utf-8'
              );
              showTempNotice('macOS/Linux Shell (.sh) 자동 폴더 정리 스크립트를 다운로드했습니다.');
            }}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F3F2F1] disabled:opacity-40 rounded-[4px]"
          >
            <Download className="w-3.5 h-3.5 text-[#0F6CBD]" />
            Shell (.sh)
          </button>

          <button
            type="button"
            disabled={files.length === 0}
            onClick={() => {
              downloadTextFile(
                JSON.stringify({ preset, generatedAt: new Date().toISOString(), files }, null, 2),
                `bmail-folder-manifest-${new Date().toISOString().slice(0, 10)}.json`,
                'application/json;charset=utf-8'
              );
              showTempNotice('폴더 정리 JSON 매니페스트를 내보냈습니다.');
            }}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F3F2F1] disabled:opacity-40 rounded-[4px]"
          >
            JSON
          </button>
        </div>
      </div>

      {notice && (
        <div className="px-4 py-1.5 bg-[#EBF3FC] border-b border-[#B4D6FA] text-xs font-medium text-[#0F6CBD]">
          {notice}
        </div>
      )}

      {/* Optional Add Received File Drawer */}
      {showAddForm && (
        <form
          onSubmit={handleQuickRegisterEntry}
          className="px-4 py-3 bg-[#F3F8FE] border-b border-[#B4D6FA] space-y-2.5"
        >
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-[#0F6CBD]">
              수신/발신 파일 직접 등록 → 송신인·도메인·계정 규칙으로 폴더 자동 분류
            </span>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-[#0F6CBD] bg-white border border-[#0F6CBD] hover:bg-[#EBF3FC] rounded-[4px]"
            >
              <Upload className="w-3.5 h-3.5" />
              로컬 파일 선택 (다중)
            </button>
            <input
              ref={fileInputRef}
              type="file"
              multiple
              onChange={handleLocalFilePick}
              className="hidden"
            />
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-2">
            <input
              type="text"
              value={uploadCustomName}
              onChange={(e) => setUploadCustomName(e.target.value)}
              placeholder="파일명 (예: DWG-H101-S02.pdf)"
              aria-label="등록할 수신 파일명"
              className="lg:col-span-2 px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            />
            <input
              type="email"
              value={uploadSenderEmail}
              onChange={(e) => setUploadSenderEmail(e.target.value)}
              placeholder="송신인 이메일 (규칙 매칭용)"
              aria-label="송신인 이메일"
              className="px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            />
            <input
              type="text"
              value={uploadSenderName}
              onChange={(e) => setUploadSenderName(e.target.value)}
              placeholder="송신인 이름/업체"
              aria-label="송신인 이름"
              className="px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            />
            <input
              type="date"
              value={uploadDate}
              onChange={(e) => setUploadDate(e.target.value)}
              aria-label="수신 일자"
              className="px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            />
            <div className="flex gap-1.5">
              <select
                aria-label="수신 또는 발신 구분"
                value={uploadDirection}
                onChange={(e) => setUploadDirection(e.target.value as 'received' | 'sent')}
                className="px-2 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
              >
                <option value="received">수신</option>
                <option value="sent">발신</option>
              </select>
              <button
                type="submit"
                disabled={!uploadCustomName.trim()}
                className="flex-1 px-2.5 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[4px]"
              >
                등록
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            <input
              type="text"
              value={uploadAccountEmail}
              onChange={(e) => setUploadAccountEmail(e.target.value)}
              placeholder="연결 수신 계정 (선택, 예: sales@company.co.kr)"
              aria-label="연결 수신 계정"
              className="px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            />
            <select
              aria-label="수동 프로젝트 지정 (선택)"
              value={uploadManualProject}
              onChange={(e) => setUploadManualProject(e.target.value)}
              className="px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            >
              <option value="">프로젝트 자동 판별 (송신인·도메인·계정·키워드 규칙)</option>
              {projects.map((p) => (
                <option key={p.id} value={p.code}>
                  수동 지정: {p.code} ({p.name})
                </option>
              ))}
            </select>
          </div>
        </form>
      )}

      {/* Split Explorer: Left Folder Tree + Right File Table */}
      <div className="flex-1 min-h-0 grid grid-cols-1 lg:grid-cols-12 overflow-hidden">
        {/* Folder Tree */}
        <aside className="lg:col-span-4 border-b lg:border-b-0 lg:border-r border-[#EDEBE9] flex flex-col min-h-0 bg-[#FAF9F8]/60">
          <div className="px-3 py-2 border-b border-[#EDEBE9] flex items-center justify-between text-xs">
            <button
              type="button"
              onClick={() => setSelectedFolderPath('ALL')}
              className={`font-bold ${
                selectedFolderPath === 'ALL' ? 'text-[#0F6CBD]' : 'text-[#323130] hover:text-[#0F6CBD]'
              }`}
            >
              전체 폴더 트리 ({files.length}건)
            </button>
            <span className="text-[11px] font-mono text-[#616161]">
              {formatFileSize(files.reduce((sum, f) => sum + f.sizeBytes, 0))}
            </span>
          </div>

          <div className="flex-1 overflow-y-auto p-2 space-y-0.5">
            {tree.length === 0 ? (
              <div className="p-5 text-center text-xs text-[#616161] leading-relaxed">
                정리할 첨부파일이나 수신 파일이 없습니다. 메일함(.pst/.eml)을 불러오거나 상단{' '}
                <strong>수신 파일 추가</strong>로 파일을 등록하세요.
              </div>
            ) : (
              tree.map((node) => renderTreeNode(node))
            )}
          </div>
        </aside>

        {/* File Table */}
        <section className="lg:col-span-8 flex flex-col min-h-0 overflow-hidden">
          <div className="px-4 py-2 border-b border-[#EDEBE9] flex items-center justify-between gap-2 text-xs bg-white">
            <span className="font-mono text-[#0F6CBD] font-semibold truncate">
              {selectedFolderPath === 'ALL' ? '/전체_수신파일_폴더' : selectedFolderPath}
            </span>
            <span className="text-[#616161] shrink-0">{visibleFiles.length}건 표시</span>
          </div>

          <div className="flex-1 overflow-auto">
            {visibleFiles.length === 0 ? (
              <div className="p-10 text-center text-xs text-[#616161] space-y-1.5">
                <Paperclip className="w-6 h-6 text-[#0F6CBD] mx-auto" />
                <p>선택한 폴더에 정리된 파일이 없습니다.</p>
              </div>
            ) : (
              <table className="w-full text-left border-collapse text-xs">
                <thead className="sticky top-0 bg-[#FAF9F8] border-b border-[#E0E0E0] text-[11px] text-[#616161]">
                  <tr>
                    <th className="py-2 px-3 font-semibold">파일명 · 도면</th>
                    <th className="py-2 px-2 font-semibold">프로젝트 배정</th>
                    <th className="py-2 px-2 font-semibold">분류 근거</th>
                    <th className="py-2 px-2 font-semibold">일자 · 구분</th>
                    <th className="py-2 px-2 font-semibold">송신인 · 도메인</th>
                    <th className="py-2 px-2 font-semibold text-right">작업</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#EDEBE9]">
                  {visibleFiles.map((file) => (
                    <tr key={file.id} className="hover:bg-[#FAF9F8] align-top">
                      <td className="py-2.5 px-3 max-w-[240px]">
                        <div className="flex items-start gap-1.5">
                          {file.kind === 'excel' ? (
                            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700 shrink-0 mt-0.5" />
                          ) : (
                            <FileText className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0 mt-0.5" />
                          )}
                          <div className="min-w-0">
                            <div className="font-semibold text-[#242424] break-all">
                              {file.filename}
                            </div>
                            <div className="text-[10px] font-mono text-[#616161] truncate" title={file.folderPath}>
                              {file.folderPath}
                            </div>
                            {file.drawingKeys.length > 0 && (
                              <div className="flex flex-wrap gap-1 mt-1">
                                {file.drawingKeys.slice(0, 2).map((dk) => (
                                  <span
                                    key={dk}
                                    className="px-1.5 py-0.2 bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[10px] rounded"
                                  >
                                    {dk}
                                  </span>
                                ))}
                              </div>
                            )}
                          </div>
                        </div>
                      </td>

                      <td className="py-2.5 px-2 whitespace-nowrap">
                        <select
                          aria-label={`${file.filename} 프로젝트 이동`}
                          value={file.projectCode}
                          onChange={(e) => onReassignFileProject(file.id, e.target.value)}
                          className="px-1.5 py-1 text-[11px] font-mono bg-white border border-[#D1D1D1] rounded-[4px]"
                        >
                          <option value="UNASSIGNED">미분류</option>
                          {projects.map((p) => (
                            <option key={p.id} value={p.code}>
                              {p.code}
                            </option>
                          ))}
                        </select>
                      </td>

                      <td className="py-2.5 px-2 whitespace-nowrap">
                        <span
                          title={file.matchEvidence}
                          className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold ${
                            file.matchSource === 'unassigned'
                              ? 'bg-amber-50 text-amber-800 border border-amber-200'
                              : file.matchSource === 'sender' || file.matchSource === 'domain'
                                ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                : 'bg-[#EBF3FC] text-[#0F6CBD] border border-[#B4D6FA]'
                          }`}
                        >
                          {MATCH_SOURCE_LABELS[file.matchSource]}
                        </span>
                      </td>

                      <td className="py-2.5 px-2 whitespace-nowrap">
                        <div className="font-mono text-[11px] text-[#242424]">{file.date}</div>
                        <div className="text-[10px] text-[#616161]">
                          {file.direction === 'received' ? '수신' : '발신'} · {formatFileSize(file.sizeBytes)}
                        </div>
                      </td>

                      <td className="py-2.5 px-2 max-w-[180px]">
                        <div className="text-[#242424] truncate font-medium">{file.senderName}</div>
                        <div className="text-[10px] font-mono text-[#616161] truncate">
                          {file.senderEmail || '-'}
                        </div>
                      </td>

                      <td className="py-2.5 px-2 text-right whitespace-nowrap">
                        <div className="inline-flex items-center gap-1">
                          {file.projectCode !== 'UNASSIGNED' &&
                            file.matchSource !== 'sender' &&
                            file.senderEmail && (
                              <button
                                type="button"
                                title={`${file.senderEmail} 송신인을 ${file.projectCode} 규칙에 저장`}
                                onClick={() => {
                                  onQuickBindRule({
                                    projectCode: file.projectCode,
                                    senderEmail: file.senderEmail,
                                  });
                                  showTempNotice(
                                    `${file.senderEmail} 송신인을 ${file.projectCode} 규칙에 등록했습니다.`
                                  );
                                }}
                                className="p-1 text-[#0F6CBD] hover:bg-[#EBF3FC] rounded"
                              >
                                <UserPlus className="w-3.5 h-3.5" />
                              </button>
                            )}
                          {file.sourceType === 'manual-upload' && (
                            <button
                              type="button"
                              aria-label={`${file.filename} 삭제`}
                              onClick={() => onRemoveManualFile(file.id)}
                              className="p-1 text-[#616161] hover:text-red-700 rounded"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </section>
      </div>
    </div>
  );
};
