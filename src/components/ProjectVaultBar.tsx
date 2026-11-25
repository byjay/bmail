import React, { useRef, useState } from 'react';
import {
  FileSpreadsheet,
  FileText,
  FolderPlus,
  Paperclip,
  Plus,
  Ship,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import type {
  ProjectTreeSubfolder,
  ProjectVaultFile,
} from '../utils/outlookWorkspaceModel';

interface ProjectVaultBarProps {
  project: ProjectDefinition;
  activeSubfolderId: string | null;
  subfolders: ProjectTreeSubfolder[];
  files: ProjectVaultFile[];
  linkedEmailCount: number;
  onSelectSubfolder: (subfolderId: string | null) => void;
  onAddSubfolder: (projectCode: string, name: string, parentId: string | null) => void;
  onDeleteSubfolder: (subfolderId: string) => void;
  onAddFiles: (projectCode: string, subfolderId: string | null, fileList: File[]) => void;
  onRemoveFile: (fileId: string) => void;
  onOpenSedmsHub?: () => void;
}

function formatBytes(bytes: number): string {
  if (!bytes || bytes < 1024) return `${bytes || 0} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export const ProjectVaultBar: React.FC<ProjectVaultBarProps> = ({
  project,
  activeSubfolderId,
  subfolders,
  files,
  linkedEmailCount,
  onSelectSubfolder,
  onAddSubfolder,
  onDeleteSubfolder,
  onAddFiles,
  onRemoveFile,
  onOpenSedmsHub,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [isAddingFolder, setIsAddingFolder] = useState(false);
  const [newFolderName, setNewFolderName] = useState('');

  const projectSubfolders = subfolders.filter((s) => s.projectCode === project.code);
  const scopedFiles = files.filter(
    (f) =>
      f.projectCode === project.code &&
      (activeSubfolderId === null ? true : f.subfolderId === activeSubfolderId)
  );
  const sedmsLinkedCount = scopedFiles.filter((f) => Boolean(f.drawingNoKey)).length;

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const list = e.target.files ? Array.from(e.target.files) : [];
    if (list.length > 0) {
      onAddFiles(project.code, activeSubfolderId, list);
    }
    e.target.value = '';
  };

  const handleDropFiles = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const list = e.dataTransfer.files ? Array.from(e.dataTransfer.files) : [];
    if (list.length > 0) {
      onAddFiles(project.code, activeSubfolderId, list);
    }
  };

  const handleCreateSubfolder = (e: React.FormEvent) => {
    e.preventDefault();
    const clean = newFolderName.trim();
    if (!clean) return;
    onAddSubfolder(project.code, clean, activeSubfolderId);
    setNewFolderName('');
    setIsAddingFolder(false);
  };

  return (
    <section
      aria-label={`프로젝트 ${project.code} 파일 및 조선 도면함`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault();
          setIsDragOver(true);
        }
      }}
      onDragLeave={() => setIsDragOver(false)}
      onDrop={handleDropFiles}
      className={`border-b border-[#E0E0E0] px-5 py-3 transition-colors shrink-0 ${
        isDragOver ? 'bg-[#EBF3FC] ring-1 ring-inset ring-[#0F6CBD]' : 'bg-[#FAF9F8]'
      }`}
    >
      <input
        ref={fileInputRef}
        type="file"
        multiple
        aria-label={`${project.code} 프로젝트 파일 선택`}
        accept=".pdf,.xlsx,.xls,.csv,.dwg,.dxf,.docx,.doc,.pptx,.png,.jpg,.jpeg,.eml,.zip,.txt"
        onChange={handleFileChange}
        className="hidden"
      />

      {/* Row 1: Project Header + File Add / Subfolder / SEDMS Actions */}
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-2 min-w-0">
          <span
            className="w-2.5 h-2.5 rounded-full shrink-0"
            style={{ backgroundColor: project.color || '#0F6CBD' }}
          />
          <span className="font-mono text-xs font-bold text-[#0F6CBD]">{project.code}</span>
          <span className="text-xs font-semibold text-[#242424] truncate">{project.name}</span>
          <span className="text-[11px] text-[#616161] font-mono tabular-nums whitespace-nowrap">
            · 메일 {linkedEmailCount}건 · 파일 {scopedFiles.length}건
            {sedmsLinkedCount > 0 ? ` · 도면 ${sedmsLinkedCount}건` : ''}
          </span>
        </div>

        <div className="flex items-center gap-1.5 shrink-0">
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px] whitespace-nowrap"
          >
            <Upload className="w-3 h-3" />
            파일 넣기
          </button>
          <button
            type="button"
            onClick={() => setIsAddingFolder((v) => !v)}
            className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[4px] whitespace-nowrap"
          >
            <FolderPlus className="w-3 h-3 text-[#0F6CBD]" />
            하위 폴더
          </button>
          {onOpenSedmsHub && (
            <button
              type="button"
              onClick={onOpenSedmsHub}
              className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[4px] whitespace-nowrap"
            >
              <Ship className="w-3 h-3" />
              SEDMS 도면 허브
            </button>
          )}
        </div>
      </div>

      {/* Row 2: Subfolder Breadcrumb & Inline Subfolder Creator */}
      {(projectSubfolders.length > 0 || isAddingFolder) && (
        <div className="mt-2 flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            onClick={() => onSelectSubfolder(null)}
            className={`px-2 py-0.5 text-[11px] rounded-[4px] border ${
              activeSubfolderId === null
                ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-semibold'
                : 'bg-white border-[#D1D1D1] text-[#424242] hover:bg-[#F5F5F5]'
            }`}
          >
            전체 ({files.filter((f) => f.projectCode === project.code).length})
          </button>
          {projectSubfolders.map((sf) => {
            const count = files.filter((f) => f.subfolderId === sf.id).length;
            const active = activeSubfolderId === sf.id;
            return (
              <span
                key={sf.id}
                className={`inline-flex items-center gap-1 pl-2 pr-1 py-0.5 text-[11px] rounded-[4px] border ${
                  active
                    ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-semibold'
                    : 'bg-white border-[#D1D1D1] text-[#424242]'
                }`}
              >
                <button type="button" onClick={() => onSelectSubfolder(sf.id)} className="hover:underline">
                  {sf.name} ({count})
                </button>
                <button
                  type="button"
                  onClick={() => onDeleteSubfolder(sf.id)}
                  aria-label={`하위 폴더 '${sf.name}' 삭제`}
                  className="p-0.5 text-[#616161] hover:text-red-700"
                >
                  <X className="w-2.5 h-2.5" />
                </button>
              </span>
            );
          })}

          {isAddingFolder && (
            <form onSubmit={handleCreateSubfolder} className="inline-flex items-center gap-1">
              <input
                type="text"
                value={newFolderName}
                onChange={(e) => setNewFolderName(e.target.value)}
                placeholder="폴더명 (예: 선체도면, 견적서)"
                aria-label="새 프로젝트 하위 폴더 이름"
                autoFocus
                className="h-6 px-2 text-[11px] bg-white border border-[#0F6CBD] rounded-[4px] focus:outline-none w-40"
              />
              <button
                type="submit"
                className="h-6 px-2 text-[11px] font-semibold text-white bg-[#0F6CBD] rounded-[4px]"
              >
                <Plus className="w-3 h-3" />
              </button>
              <button
                type="button"
                onClick={() => setIsAddingFolder(false)}
                aria-label="하위 폴더 추가 취소"
                className="h-6 px-1.5 text-[11px] text-[#616161] hover:text-[#242424]"
              >
                <X className="w-3 h-3" />
              </button>
            </form>
          )}
        </div>
      )}

      {/* Row 3: Project Files Strip (Add / Drag-Drop / Remove) */}
      <div className="mt-2.5">
        {scopedFiles.length === 0 ? (
          <div className="px-3 py-2 rounded-[5px] border border-dashed border-[#D1D1D1] bg-white text-[11px] text-[#616161] flex items-center justify-between gap-2">
            <span>
              이 프로젝트에 PDF·DWG·Excel 도면 및 문서를 끌어다 놓거나 <strong>파일 넣기</strong>로 보관·제거할 수 있습니다.
            </span>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="text-[#0F6CBD] font-semibold hover:underline shrink-0"
            >
              + 파일 선택
            </button>
          </div>
        ) : (
          <div className="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto">
            {scopedFiles.map((file) => (
              <div
                key={file.id}
                className="inline-flex items-center gap-1.5 pl-2 pr-1.5 py-1 rounded-[5px] bg-white border border-[#D1D1D1] text-[11px] text-[#242424] max-w-full"
              >
                {file.kind === 'excel' ? (
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
                ) : file.kind === 'pdf' ? (
                  <FileText className="w-3.5 h-3.5 text-red-600 shrink-0" />
                ) : (
                  <Paperclip className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />
                )}
                <span className="font-medium truncate max-w-[180px]" title={file.filename}>
                  {file.filename}
                </span>
                <span className="text-[10px] font-mono text-[#616161] shrink-0">
                  {formatBytes(file.sizeBytes)}
                </span>
                {file.drawingNoKey && (
                  <span
                    className="px-1.5 py-0.2 rounded bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[10px] font-semibold shrink-0"
                    title={`SEDMS 도면 자동 인덱싱 (${file.discipline || '도면'})`}
                  >
                    SEDMS:{file.drawingNoKey}
                  </span>
                )}
                <button
                  type="button"
                  onClick={() => onRemoveFile(file.id)}
                  aria-label={`프로젝트 파일 '${file.filename}' 빼기`}
                  title="프로젝트에서 파일 빼기"
                  className="p-0.5 text-[#616161] hover:text-red-700 hover:bg-red-50 rounded shrink-0"
                >
                  <Trash2 className="w-3 h-3" />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </section>
  );
};
