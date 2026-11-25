import React, { useEffect, useRef, useState } from 'react';
import {
  AlertCircle,
  CheckCircle2,
  Database,
  FileArchive,
  FolderOpen,
  HardDrive,
  Loader2,
  Sparkles,
  Upload,
  X,
} from 'lucide-react';
import type { InheritedOutlookProfile, MailItem } from '../types/mail';

export interface PstImportResult {
  storeName: string;
  fileType: 'pst' | 'eml' | 'mbox' | 'msg';
  folders: string[];
  emails: MailItem[];
  skippedCount: number;
  inheritedProfile?: InheritedOutlookProfile;
}

interface PstImportModalProps {
  isOpen: boolean;
  onClose: () => void;
  onImportComplete: (
    payload: PstImportResult,
    mode: 'replace' | 'merge',
    autoAnalyze: boolean
  ) => void;
  localDbAvailable?: boolean;
  localTotalEmails?: number;
  localPstPaths?: string[];
  onLocalSyncComplete?: () => void;
}

export const PstImportModal: React.FC<PstImportModalProps> = ({
  isOpen,
  onClose,
  onImportComplete,
  localDbAvailable = false,
  localTotalEmails = 0,
  localPstPaths = [],
  onLocalSyncComplete,
}) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const pollTimerRef = useRef<number | null>(null);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [importMode, setImportMode] = useState<'replace' | 'merge'>('replace');
  const [autoAnalyze, setAutoAnalyze] = useState(true);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [lastSuccess, setLastSuccess] = useState<PstImportResult | null>(null);
  const [localFolderPath, setLocalFolderPath] = useState('F:\\MAIL');
  const [isLocalMigrating, setIsLocalMigrating] = useState(false);
  const [localSyncNote, setLocalSyncNote] = useState<string | null>(null);

  useEffect(() => {
    return () => {
      if (pollTimerRef.current !== null) {
        window.clearTimeout(pollTimerRef.current);
      }
    };
  }, []);

  if (!isOpen) return null;

  const isLocalHost =
    localDbAvailable ||
    (typeof window !== 'undefined' &&
      (window.location.hostname === '127.0.0.1' || window.location.hostname === 'localhost'));

  const handleRecomputeProjects = async () => {
    setError(null);
    setLocalSyncNote(null);
    setIsLocalMigrating(true);
    try {
      const res = await fetch('/api/local/outlook/recompute-projects', { method: 'POST' });
      const data = (await res.json()) as { ok?: boolean; projectsCount?: number; error?: string };
      if (!res.ok) throw new Error(data.error || '프로젝트 재계산 실패');
      setLocalSyncNote(
        `로컬 SQLite 폴더 기반 프로젝트 ${Number(data.projectsCount || 0)}건을 즉시 갱신했습니다.`
      );
      onLocalSyncComplete?.();
    } catch (err) {
      setError(err instanceof Error ? err.message : '프로젝트 재계산 중 오류');
    } finally {
      setIsLocalMigrating(false);
    }
  };

  const handleLocalFolderMigrate = async () => {
    const targetPath = localFolderPath.trim();
    if (!targetPath) {
      setError('스캔할 로컬 PC 폴더 경로를 입력해 주세요.');
      return;
    }
    setError(null);
    setLocalSyncNote(null);
    setIsLocalMigrating(true);
    try {
      const res = await fetch('/api/local/outlook/migrate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folderPath: targetPath }),
      });
      const data = (await res.json()) as {
        started?: boolean;
        alreadyRunning?: boolean;
        scannedFiles?: number;
        error?: string;
        progress?: {
          status: string;
          totalEmailsIndexed: number;
          totalFoldersIndexed: number;
          error?: string | null;
        };
      };
      if (!res.ok) {
        throw new Error(data.error || '로컬 폴더 PST 스캔 시작 실패');
      }
      setLocalSyncNote(
        `${targetPath} 폴더(${data.scannedFiles || localPstPaths.length}개 PST) 백그라운드 동기화 시작됨...`
      );

      const pollStatus = async () => {
        try {
          const stRes = await fetch('/api/local/outlook/migrate/status');
          if (!stRes.ok) return;
          const stData = (await stRes.json()) as {
            progress?: {
              status: string;
              totalEmailsIndexed: number;
              totalContactsIndexed: number;
              totalFoldersIndexed: number;
              discoveredProjectsCount: number;
              error?: string | null;
            };
          };
          const prog = stData.progress;
          if (!prog) return;
          if (prog.status === 'running') {
            setLocalSyncNote(
              `동기화 중 · 메일 ${prog.totalEmailsIndexed.toLocaleString()}건 · 폴더 ${prog.totalFoldersIndexed}개`
            );
            pollTimerRef.current = window.setTimeout(() => void pollStatus(), 1500);
          } else if (prog.status === 'completed') {
            setIsLocalMigrating(false);
            setLocalSyncNote(
              `완료 · 메일 ${prog.totalEmailsIndexed.toLocaleString()}건 · 폴더 ${prog.totalFoldersIndexed}개 · 프로젝트 ${prog.discoveredProjectsCount}건`
            );
            onLocalSyncComplete?.();
          } else if (prog.status === 'failed') {
            setIsLocalMigrating(false);
            setError(prog.error || '로컬 PST 동기화 실패');
          } else {
            setIsLocalMigrating(false);
          }
        } catch {
          setIsLocalMigrating(false);
        }
      };

      pollTimerRef.current = window.setTimeout(() => void pollStatus(), 1200);
    } catch (err) {
      setIsLocalMigrating(false);
      setError(err instanceof Error ? err.message : '로컬 폴더 PST 스캔 중 오류');
    }
  };

  const acceptFiles = (list: FileList | null) => {
    if (!list || list.length === 0) return;
    setError(null);
    setLastSuccess(null);
    setSelectedFiles(Array.from(list));
  };

  const handleUpload = async () => {
    if (selectedFiles.length === 0) {
      setError('연동할 .pst, .ost, .eml, .mbox 또는 .msg 파일을 먼저 선택해 주세요.');
      return;
    }

    setIsUploading(true);
    setError(null);

    try {
      const formData = new FormData();
      for (const file of selectedFiles) {
        formData.append('files', file);
      }

      const res = await fetch('/api/pst/import', {
        method: 'POST',
        body: formData,
      });

      const data = (await res.json()) as PstImportResult & { error?: string };
      if (!res.ok) {
        throw new Error(data.error || '메일 파일 파싱 중 오류가 발생했습니다.');
      }

      if (!data.emails || data.emails.length === 0) {
        throw new Error('파일에서 추출할 수 있는 메일 항목을 찾지 못했습니다.');
      }

      setLastSuccess(data);
      onImportComplete(data, importMode, autoAnalyze);
    } catch (err) {
      setError(err instanceof Error ? err.message : '파일 연동 중 알 수 없는 오류가 발생했습니다.');
    } finally {
      setIsUploading(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="pst-modal-title"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/45 p-4"
    >
      <div className="w-full max-w-xl bg-white rounded-xl border border-slate-200 shadow-xl overflow-hidden">
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200">
          <div className="flex items-center gap-2.5">
            <Database className="w-4 h-4 text-[#0F6CBD]" />
            <h3 id="pst-modal-title" className="text-sm font-bold text-slate-900">
              Outlook 원본 PST 등록 · 설정 자동 인수 & AI 분석
            </h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="PST 파일 연동 창 닫기"
            className="p-1 text-slate-400 hover:text-slate-700 rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-5 space-y-4">
          <p className="text-xs text-slate-600 leading-relaxed">
            기존 Outlook <strong>.pst / .ost</strong>(또는 <strong>.eml / .mbox</strong>) 파일을 등록하면
            계정 주소, 폴더 계층, 사용자 폴더 기반 프로젝트 코드, 첨부파일, 일정까지 Outlook 설정을 그대로 인수하고
            주고받은 메일에서 <strong>내가 한 일과 프로젝트별 전체 내용</strong>을 AI가 즉시 정리합니다.
          </p>

          {isLocalHost && (
            <div className="p-3.5 rounded-xl border border-[#0F6CBD]/30 bg-[#EBF3FC]/60 space-y-2">
              <div className="flex items-center justify-between gap-2">
                <div className="inline-flex items-center gap-1.5 text-xs font-bold text-[#0F6CBD]">
                  <HardDrive className="w-3.5 h-3.5" />
                  <span>로컬 PC 대용량 PST 폴더 직접 연결 (업로드 불필요)</span>
                </div>
                {localDbAvailable && (
                  <span className="text-[11px] font-mono font-semibold text-[#0F6CBD]">
                    연결 {localPstPaths.length}개 · {localTotalEmails.toLocaleString()}건 보관
                  </span>
                )}
              </div>
              <div className="flex flex-col sm:flex-row gap-2">
                <input
                  type="text"
                  value={localFolderPath}
                  onChange={(e) => setLocalFolderPath(e.target.value)}
                  placeholder="F:\MAIL"
                  aria-label="로컬 PST 폴더 경로"
                  className="flex-1 px-2.5 py-1.5 text-xs font-mono bg-white border border-slate-300 rounded-md focus:outline-none focus:border-[#0F6CBD]"
                />
                {typeof window !== 'undefined' &&
                  (window as { bmailDesktop?: { pickLocalFolder?: () => Promise<string> } })
                    .bmailDesktop?.pickLocalFolder && (
                    <button
                      type="button"
                      onClick={async () => {
                        const picked = await (
                          window as { bmailDesktop?: { pickLocalFolder?: () => Promise<string> } }
                        ).bmailDesktop?.pickLocalFolder?.();
                        if (picked) setLocalFolderPath(picked);
                      }}
                      className="px-2.5 py-1.5 text-xs font-medium text-[#0F6CBD] bg-white hover:bg-[#EBF3FC] border border-slate-300 rounded-md whitespace-nowrap"
                    >
                      폴더 선택
                    </button>
                  )}
                <button
                  type="button"
                  disabled={isLocalMigrating}
                  onClick={() => void handleLocalFolderMigrate()}
                  className="inline-flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-400 rounded-md transition-colors whitespace-nowrap"
                >
                  {isLocalMigrating && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
                  로컬 폴더 스캔 · 동기화
                </button>
                <button
                  type="button"
                  disabled={isLocalMigrating || !localDbAvailable}
                  onClick={() => void handleRecomputeProjects()}
                  className="px-2.5 py-1.5 text-xs font-medium text-[#0F6CBD] bg-white border border-[#0F6CBD]/40 hover:bg-[#EBF3FC] disabled:opacity-50 rounded-md transition-colors whitespace-nowrap"
                >
                  프로젝트 즉시 재추론
                </button>
              </div>
              {localSyncNote && (
                <div className="text-[11px] text-[#0B5A08] font-medium">{localSyncNote}</div>
              )}
            </div>
          )}

          {/* Drag & Drop File Zone */}
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setIsDragging(false);
              acceptFiles(e.dataTransfer.files);
            }}
            className={`border-2 border-dashed rounded-xl p-6 text-center transition-colors ${
              isDragging
                ? 'border-[#0F6CBD] bg-[#0F6CBD]/5'
                : 'border-slate-300 bg-slate-50/70 hover:border-slate-400'
            }`}
          >
            <input
              ref={fileInputRef}
              id="pst-file-input"
              type="file"
              multiple
              accept=".pst,.ost,.eml,.mbox,.msg"
              onChange={(e) => acceptFiles(e.target.files)}
              className="sr-only"
            />
            <FileArchive className="w-8 h-8 text-[#0F6CBD] mx-auto mb-2.5" />
            <div className="text-sm font-semibold text-slate-900">
              Outlook .pst / .ost / .eml / .mbox / .msg 파일을 여기에 놓으세요
            </div>
            <p className="mt-1 text-xs text-slate-500">
              계정·폴더·프로젝트·첨부파일·일정 자동 인수 · 최대 150MB · 파싱 직후 임시파일 즉시 삭제
            </p>
            <div className="mt-4 flex items-center justify-center gap-2">
              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg transition-colors"
              >
                <FolderOpen className="w-3.5 h-3.5" />
                내 PC에서 파일 선택
              </button>
            </div>
          </div>

          {selectedFiles.length > 0 && (
            <div className="px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-lg space-y-1">
              <div className="text-xs font-semibold text-slate-800">
                선택된 파일 ({selectedFiles.length}개)
              </div>
              <ul className="text-xs text-slate-600 space-y-0.5 font-mono">
                {selectedFiles.map((f, idx) => (
                  <li key={`${f.name}-${idx}`} className="flex items-center justify-between gap-2">
                    <span className="truncate">{f.name}</span>
                    <span className="text-slate-400 tabular-nums shrink-0">
                      {(f.size / 1024).toFixed(1)} KB
                    </span>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Replace vs Merge Option */}
          <div className="flex items-center justify-between gap-3 pt-1">
            <span className="text-xs font-medium text-slate-700">가져오기 방식</span>
            <div className="inline-flex items-center p-0.5 bg-slate-100 rounded-lg">
              <button
                type="button"
                aria-pressed={importMode === 'replace'}
                onClick={() => setImportMode('replace')}
                className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
                  importMode === 'replace' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600'
                }`}
              >
                기존 목록 대체 (내 PST만 보기)
              </button>
              <button
                type="button"
                aria-pressed={importMode === 'merge'}
                onClick={() => setImportMode('merge')}
                className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
                  importMode === 'merge' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600'
                }`}
              >
                현재 메일함에 병합
              </button>
            </div>
          </div>

          {/* Core AI Immediate Analysis Toggle */}
          <label className="flex items-start gap-2.5 p-3 rounded-lg border border-[#0F6CBD]/25 bg-[#EBF3FC]/50 cursor-pointer">
            <input
              type="checkbox"
              checked={autoAnalyze}
              onChange={(e) => setAutoAnalyze(e.target.checked)}
              className="mt-0.5 rounded border-slate-300 text-[#0F6CBD] focus:ring-[#0F6CBD]"
            />
            <div className="text-xs space-y-0.5">
              <div className="font-semibold text-slate-900 inline-flex items-center gap-1.5">
                <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD]" />
                등록 즉시 AI 핵심 분석 실행 (내가 한 일 · 프로젝트별 정리 · 업무 일지)
              </div>
              <p className="text-[11px] text-slate-600">
                Outlook 폴더·계정·프로젝트를 자동 인수한 직후, 주고받은 메일 전체를 AI 에이전트가 분석합니다.
              </p>
            </div>
          </label>

          {/* Where to find local Outlook PST files */}
          <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-1 text-[11px] text-slate-600">
            <div className="font-semibold text-slate-800">내 PC에서 Outlook 원본 .pst 파일 찾는 위치</div>
            <div className="font-mono text-slate-700">
              Windows: C:\Users\&lt;사용자&gt;\Documents\Outlook 파일\*.pst
            </div>
            <div className="font-mono text-slate-700">
              캐시 파일: %LOCALAPPDATA%\Microsoft\Outlook\*.pst / *.ost
            </div>
          </div>

          {error && (
            <div role="alert" className="flex items-start gap-2 p-3 rounded-lg bg-red-50 border border-red-200 text-xs text-red-800">
              <AlertCircle className="w-4 h-4 text-red-600 shrink-0 mt-0.5" />
              <span>{error}</span>
            </div>
          )}

          {lastSuccess && (
            <div className="p-3 rounded-lg bg-emerald-50 border border-emerald-200 text-xs text-emerald-900 space-y-1.5">
              <div className="flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <div>
                  <strong>{lastSuccess.storeName}</strong> Outlook 설정 인수 완료 — 메일{' '}
                  <span className="font-mono tabular-nums">{lastSuccess.emails.length}건</span>, 폴더{' '}
                  <span className="font-mono tabular-nums">{lastSuccess.folders.length}개</span>
                </div>
              </div>
              {lastSuccess.inheritedProfile && (
                <div className="pl-6 text-[11px] text-emerald-800 space-y-1">
                  {lastSuccess.inheritedProfile.ownerEmail && (
                    <div>
                      감지된 Outlook 계정:{' '}
                      <span className="font-mono font-semibold">
                        {lastSuccess.inheritedProfile.ownerEmail}
                      </span>
                    </div>
                  )}
                  <div>
                    자동 등록 프로젝트{' '}
                    <span className="font-mono font-semibold">
                      {lastSuccess.inheritedProfile.discoveredProjects.length}개
                    </span>{' '}
                    · 첨부 메일{' '}
                    <span className="font-mono font-semibold">
                      {lastSuccess.inheritedProfile.attachmentMailCount}건
                    </span>{' '}
                    · 추출 일정{' '}
                    <span className="font-mono font-semibold">
                      {lastSuccess.inheritedProfile.calendarHintCount}건
                    </span>
                  </div>
                </div>
              )}
            </div>
          )}

          <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-200">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg"
            >
              닫기
            </button>
            <button
              type="button"
              disabled={selectedFiles.length === 0 || isUploading}
              onClick={handleUpload}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 disabled:text-slate-500 rounded-lg transition-colors"
            >
              {isUploading ? (
                <>
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  PST 설정 인수 및 분석 중...
                </>
              ) : (
                <>
                  <Upload className="w-3.5 h-3.5" />
                  {autoAnalyze ? '설정 인수 + AI 즉시 분석' : 'Outlook 설정 및 메일 불러오기'}
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
