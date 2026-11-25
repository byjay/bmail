import React, { useEffect, useMemo, useState } from 'react';
import {
  CheckCircle2,
  Download,
  Eye,
  FileSpreadsheet,
  FileText,
  FolderCheck,
  FolderOpen,
  Image as ImageIcon,
  Loader2,
  Paperclip,
  Sparkles,
  X,
} from 'lucide-react';
import type {
  AttachmentKind,
  DocumentAnalysisResult,
  MailAttachmentMeta,
  MailItem,
  ParsedDocumentSummary,
} from '../types/mail';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import { isInlineBusinessCardImage } from '../utils/outlookWorkspaceModel';
import type {
  AttachmentConflictPolicy,
  AttachmentExportPlan,
  AttachmentExportResult,
  AttachmentExportScope,
  AttachmentFilenamePattern,
  AttachmentFolderStructure,
  AttachmentReportFormat,
} from '../services/attachmentBatchExporter';

export interface AttachmentPreviewDockProps {
  email: MailItem | null;
  selectedAttachmentIdx?: number | null;
  onSelectAttachmentIdx?: (idx: number) => void;
  activeFolderFilter?: string;
  activeStoreFilter?: string;
  projects: ProjectDefinition[];
  onClose?: () => void;
}

const IMAGE_EXT_RE = /\.(png|jpe?g|gif|bmp|webp|svg)$/i;

function isLikelyInlineSignature(att: MailAttachmentMeta): boolean {
  return isInlineBusinessCardImage(att);
}

function formatBytes(bytes?: number): string {
  const b = Number(bytes || 0);
  if (!b || b <= 0) return '-';
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / (1024 * 1024)).toFixed(2)} MB`;
}

export const AttachmentPreviewDock: React.FC<AttachmentPreviewDockProps> = ({
  email,
  selectedAttachmentIdx,
  onSelectAttachmentIdx,
  activeFolderFilter = '',
  activeStoreFilter = '',
  projects,
  onClose,
}) => {
  const [dockTab, setDockTab] = useState<'preview' | 'export'>('preview');
  const [showInlineSignatures, setShowInlineSignatures] = useState(false);
  const [activeIdx, setActiveIdx] = useState<number>(0);

  // Parsed preview & AI analysis state
  const [parsing, setParsing] = useState(false);
  const [parsedDoc, setParsedDoc] = useState<ParsedDocumentSummary | null>(null);
  const [analyzingDoc, setAnalyzingDoc] = useState(false);
  const [docAnalysis, setDocAnalysis] = useState<DocumentAnalysisResult | null>(null);
  const [previewNotice, setPreviewNotice] = useState<string | null>(null);

  // Batch Folder Exporter state
  const [destDir, setDestDir] = useState('D:\\Bmail_Attachments');
  const [scope, setScope] = useState<AttachmentExportScope>('all');
  const [projectCode, setProjectCode] = useState<string>(projects[0]?.code || '');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [senderQuery, setSenderQuery] = useState('');
  const [keyword, setKeyword] = useState('');
  const [selectedKinds, setSelectedKinds] = useState<AttachmentKind[]>([]);
  const [excludeInlineImages, setExcludeInlineImages] = useState(true);
  const [folderStructure, setFolderStructure] = useState<AttachmentFolderStructure>('by-project');
  const [filenamePattern, setFilenamePattern] = useState<AttachmentFilenamePattern>('date-sender-original');
  const [conflictPolicy, setConflictPolicy] = useState<AttachmentConflictPolicy>('skip-identical');
  const [reportFormats, setReportFormats] = useState<AttachmentReportFormat[]>(['md', 'csv', 'html']);

  const [planning, setPlanning] = useState(false);
  const [exportPlan, setExportPlan] = useState<AttachmentExportPlan | null>(null);
  const [executing, setExecuting] = useState(false);
  const [exportResult, setExportResult] = useState<AttachmentExportResult | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);

  const allAttachments = useMemo(() => {
    const list = email?.attachments || [];
    return list.map((att, idx) => {
      const m = String(att.id || '').match(/-att-(\d+)$/);
      const realIdx = m ? Number(m[1]) : idx;
      return {
        att,
        realIdx,
        isInlineSig: isLikelyInlineSignature(att),
      };
    });
  }, [email]);

  const visibleAttachments = useMemo(() => {
    if (showInlineSignatures) return allAttachments;
    return allAttachments.filter((item) => !item.isInlineSig);
  }, [allAttachments, showInlineSignatures]);

  const hiddenInlineCount = allAttachments.length - visibleAttachments.length;

  useEffect(() => {
    if (typeof selectedAttachmentIdx === 'number' && selectedAttachmentIdx >= 0) {
      setActiveIdx(selectedAttachmentIdx);
      setDockTab('preview');
    } else if (visibleAttachments.length > 0) {
      setActiveIdx(visibleAttachments[0].realIdx);
    } else {
      setActiveIdx(-1);
    }
    setParsedDoc(null);
    setDocAnalysis(null);
    setPreviewNotice(null);
  }, [email?.id, selectedAttachmentIdx, visibleAttachments.length]);

  const currentAttachment = useMemo(() => {
    if (visibleAttachments.length === 0) return null;
    if (typeof selectedAttachmentIdx === 'number' && selectedAttachmentIdx >= 0) {
      const explicit = visibleAttachments.find((x) => x.realIdx === selectedAttachmentIdx);
      if (explicit) return explicit;
    }
    return visibleAttachments.find((x) => x.realIdx === activeIdx) || visibleAttachments[0] || null;
  }, [visibleAttachments, activeIdx, selectedAttachmentIdx]);

  // Auto-parse when selecting Excel/Word/Text/PDF attachment
  useEffect(() => {
    if (!email || !currentAttachment || dockTab !== 'preview') return;
    const { att, realIdx } = currentAttachment;
    if (IMAGE_EXT_RE.test(att.filename)) {
      setParsedDoc(null);
      return;
    }
    let cancelled = false;
    setParsing(true);
    setPreviewNotice(null);
    fetch('/api/local/pst/attachment/parse', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ emailId: email.id, idx: realIdx }),
    })
      .then(async (res) => {
        if (!res.ok) return null;
        const data = (await res.json()) as { document?: ParsedDocumentSummary };
        return data.document || null;
      })
      .then((doc) => {
        if (!cancelled) setParsedDoc(doc);
      })
      .catch(() => {
        if (!cancelled) setParsedDoc(null);
      })
      .finally(() => {
        if (!cancelled) setParsing(false);
      });
    return () => {
      cancelled = true;
    };
  }, [email?.id, currentAttachment?.realIdx, dockTab]);

  const handleAnalyzeDocumentWithAi = async () => {
    if (!email || !currentAttachment) return;
    setAnalyzingDoc(true);
    setPreviewNotice(null);
    const docToAnalyze: ParsedDocumentSummary = parsedDoc || {
      id: currentAttachment.att.id,
      filename: currentAttachment.att.filename,
      kind: currentAttachment.att.kind,
      mimeType: currentAttachment.att.mimeType,
      sizeBytes: currentAttachment.att.sizeBytes || 0,
      extractedText: `${email.subject}\n${email.body}`,
      pageCount: 1,
      numericHighlights: [],
      engineUsed: 'built-in',
    };
    try {
      const res = await fetch('/api/analyze/documents', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          documents: [docToAnalyze],
          emails: [email],
          provider: 'claude',
          maskSensitiveInfo: false,
        }),
      });
      if (res.ok) {
        const data = (await res.json()) as DocumentAnalysisResult;
        setDocAnalysis(data);
        return;
      }
    } catch {
      // Fall through to deterministic summary
    } finally {
      setAnalyzingDoc(false);
    }
    setDocAnalysis({
      id: `doc-local-${Date.now()}`,
      createdAt: new Date().toISOString(),
      providerUsed: 'builtin',
      providerModel: '로컬 문서 분석 엔진',
      files: [
        {
          filename: currentAttachment.att.filename,
          kind: currentAttachment.att.kind,
          engineUsed: 'built-in',
          statsLabel: formatBytes(currentAttachment.att.sizeBytes),
        },
      ],
      linkedEmailCount: 1,
      overview: `"${currentAttachment.att.filename}" (${formatBytes(
        currentAttachment.att.sizeBytes
      )}) 문서가 "${email.subject}" 메일에 첨부되었습니다. ${
        parsedDoc?.extractedText
          ? parsedDoc.extractedText.slice(0, 320)
          : '상단 미리보기 또는 원본 다운로드를 통해 세부 도면·사양을 확인하세요.'
      }`,
      keyFindings: [
        {
          sourceFile: currentAttachment.att.filename,
          category: '첨부 문서 요약',
          summary: `발신자: ${email.fromName || email.fromEmail} (${email.date})`,
          metrics: parsedDoc?.numericHighlights || [],
        },
      ],
      risksAndGaps: [],
      actionItems: [
        {
          task: '도면·POS 리비전 및 변경 사양 대조 검토',
          owner: '담당자',
          priority: '높음',
          basis: email.subject,
        },
      ],
      emailCrossInsights: [],
      markdownReport: `# ${currentAttachment.att.filename} 분석 요약`,
    });
  };

  const handlePickNativeFolder = async () => {
    const picker = (
      window as unknown as { bmailDesktop?: { pickLocalFolder?: () => Promise<string> } }
    ).bmailDesktop?.pickLocalFolder;
    if (picker) {
      const picked = await picker();
      if (picked) setDestDir(picked);
    }
  };

  const buildFilterPayload = () => ({
    destDir: destDir.trim(),
    scope,
    storeName: scope === 'current-folder' && activeStoreFilter ? activeStoreFilter : undefined,
    folderPaths:
      scope === 'current-folder' && activeFolderFilter && activeFolderFilter !== 'all'
        ? [activeFolderFilter]
        : undefined,
    emailIds: scope === 'selected-emails' && email ? [email.id] : undefined,
    projectCode: scope === 'selected-project' ? projectCode : undefined,
    dateFrom: dateFrom || undefined,
    dateTo: dateTo || undefined,
    senderQuery: senderQuery.trim() || undefined,
    keyword: keyword.trim() || undefined,
    kinds: selectedKinds.length > 0 ? selectedKinds : undefined,
    excludeInlineImages,
    folderStructure,
    filenamePattern,
    conflictPolicy,
    reportFormats,
  });

  const handleRunExportPlan = async () => {
    setPlanning(true);
    setExportError(null);
    setExportResult(null);
    try {
      const res = await fetch('/api/local/attachments/export-plan', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(buildFilterPayload()),
      });
      const data = await res.json();
      if (!res.ok) {
        setExportError(data.error || '첨부파일 저장 계획 조회에 실패했습니다.');
        return;
      }
      setExportPlan(data as AttachmentExportPlan);
    } catch {
      setExportError('로컬 서버와 통신할 수 없습니다. 데스크톱 앱 실행 상태를 확인해 주세요.');
    } finally {
      setPlanning(false);
    }
  };

  const handleRunExportExecute = async () => {
    setExecuting(true);
    setExportError(null);
    try {
      const res = await fetch('/api/local/attachments/export-run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(buildFilterPayload()),
      });
      const data = await res.json();
      if (!res.ok) {
        setExportError(data.error || '첨부파일 일괄 저장에 실패했습니다.');
        return;
      }
      setExportResult(data as AttachmentExportResult);
    } catch {
      setExportError('첨부파일 저장 실행 중 통신 오류가 발생했습니다.');
    } finally {
      setExecuting(false);
    }
  };

  const toggleKind = (k: AttachmentKind) => {
    setSelectedKinds((prev) => (prev.includes(k) ? prev.filter((x) => x !== k) : [...prev, k]));
  };

  const toggleReportFormat = (fmt: AttachmentReportFormat) => {
    setReportFormats((prev) =>
      prev.includes(fmt) ? (prev.length > 1 ? prev.filter((x) => x !== fmt) : prev) : [...prev, fmt]
    );
  };

  const inlinePreviewUrl =
    email && currentAttachment
      ? `/api/local/pst/attachment?emailId=${encodeURIComponent(email.id)}&idx=${currentAttachment.realIdx}&inline=1`
      : '';
  const downloadUrl =
    email && currentAttachment
      ? `/api/local/pst/attachment?emailId=${encodeURIComponent(email.id)}&idx=${currentAttachment.realIdx}`
      : '';

  return (
    <aside
      aria-label="세로 2분할 첨부 미리보기 및 폴더 일괄 저장 도크"
      className="w-[288px] xl:w-[316px] bg-white border-l border-[#E0E0E0] flex flex-col min-h-0 shrink-0 select-none"
    >
      {/* Top Mode Switcher: 첨부 미리보기·분석 vs 폴더 저장·리포트 */}
      <div className="px-2.5 py-2 bg-[#FAF9F8] border-b border-[#E0E0E0] flex items-center justify-between gap-1.5 shrink-0">
        <div className="grid grid-cols-2 flex-1 bg-[#E5E5E5] p-0.5 rounded-md text-[11px] font-bold">
          <button
            type="button"
            onClick={() => setDockTab('preview')}
            className={`py-1 px-2 rounded flex items-center justify-center gap-1 transition-colors ${
              dockTab === 'preview'
                ? 'bg-white text-[#0F6CBD] shadow-sm'
                : 'text-[#616161] hover:text-[#242424]'
            }`}
          >
            <Eye className="w-3.5 h-3.5 shrink-0" />
            <span>첨부 미리보기·분석</span>
          </button>
          <button
            type="button"
            onClick={() => setDockTab('export')}
            className={`py-1 px-2 rounded flex items-center justify-center gap-1 transition-colors ${
              dockTab === 'export'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-[#616161] hover:text-[#242424]'
            }`}
          >
            <FolderCheck className="w-3.5 h-3.5 shrink-0" />
            <span>폴더 저장·리포트</span>
          </button>
        </div>

        {onClose && (
          <button
            type="button"
            onClick={onClose}
            title="첨부 도크 닫기"
            className="p-1 rounded hover:bg-slate-200 text-slate-500 shrink-0"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        )}
      </div>

      {/* TAB 1: Inline Attachment Preview & AI Document Analysis */}
      {dockTab === 'preview' ? (
        <div className="flex-1 overflow-y-auto p-3 space-y-3 text-xs select-text">
          {!email || allAttachments.length === 0 ? (
            <div className="py-12 text-center text-slate-500 space-y-2">
              <Paperclip className="w-6 h-6 mx-auto text-slate-400" />
              <p className="font-semibold text-slate-700">선택된 메일에 첨부파일이 없습니다</p>
              <p className="text-[11px]">
                전체 메일함의 첨부파일을 기간·보낸사람·프로젝트별로 일괄 추출하려면 상단{' '}
                <button
                  type="button"
                  onClick={() => setDockTab('export')}
                  className="text-emerald-700 font-bold underline"
                >
                  [폴더 저장·리포트]
                </button>{' '}
                탭을 클릭하세요.
              </p>
            </div>
          ) : (
            <>
              {/* Attachment Selector Strip */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between text-[11px]">
                  <span className="font-bold text-slate-700">
                    실첨부 파일 ({visibleAttachments.length}개)
                  </span>
                  {hiddenInlineCount > 0 && (
                    <button
                      type="button"
                      onClick={() => setShowInlineSignatures((v) => !v)}
                      className="text-[10px] text-[#0F6CBD] hover:underline"
                    >
                      {showInlineSignatures
                        ? '서명 이미지 숨기기'
                        : `서명 이미지 ${hiddenInlineCount}개 포함 보기`}
                    </button>
                  )}
                </div>

                <div className="flex flex-col gap-1 max-h-32 overflow-y-auto">
                  {visibleAttachments.length === 0 ? (
                    <div className="px-2.5 py-3 rounded border border-slate-200 bg-slate-50 text-[11px] text-slate-600 text-center">
                      실제 첨부 문서가 없습니다.
                      {hiddenInlineCount > 0 && (
                        <div className="mt-1 text-[10px] text-slate-500">
                          (메일 서명·명함·네임태그 이미지 {hiddenInlineCount}개 자동 무시됨)
                        </div>
                      )}
                    </div>
                  ) : (
                    visibleAttachments.map(({ att, realIdx, isInlineSig }) => {
                      const isSelected = currentAttachment?.realIdx === realIdx;
                      return (
                        <div
                          key={att.id || realIdx}
                          onClick={() => {
                            setActiveIdx(realIdx);
                            onSelectAttachmentIdx?.(realIdx);
                          }}
                          className={`px-2.5 py-1.5 rounded border text-[11px] cursor-pointer flex items-center justify-between gap-2 transition-colors ${
                            isSelected
                              ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-bold'
                              : 'bg-slate-50 border-slate-200 text-slate-800 hover:bg-slate-100'
                          }`}
                        >
                          <div className="flex items-center gap-1.5 min-w-0">
                            {att.kind === 'pdf' ? (
                              <FileText className="w-3.5 h-3.5 text-rose-600 shrink-0" />
                            ) : att.kind === 'excel' ? (
                              <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                            ) : IMAGE_EXT_RE.test(att.filename) ? (
                              <ImageIcon className="w-3.5 h-3.5 text-sky-600 shrink-0" />
                            ) : (
                              <Paperclip className="w-3.5 h-3.5 text-slate-500 shrink-0" />
                            )}
                            <span className="truncate">{att.filename}</span>
                            {isInlineSig && (
                              <span className="px-1 py-0.2 rounded bg-slate-200 text-slate-600 text-[9px]">
                                인라인
                              </span>
                            )}
                          </div>
                          <span className="text-[10px] text-slate-500 shrink-0">
                            {formatBytes(att.sizeBytes)}
                          </span>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>

              {/* Active File Action Bar */}
              {currentAttachment && (
                <div className="flex items-center justify-between gap-1.5 pt-1 border-t border-slate-200">
                  <button
                    type="button"
                    disabled={analyzingDoc}
                    onClick={() => void handleAnalyzeDocumentWithAi()}
                    className="flex-1 py-1.5 px-2.5 rounded bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-[11px] font-semibold flex items-center justify-center gap-1.5"
                  >
                    {analyzingDoc ? (
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <Sparkles className="w-3.5 h-3.5" />
                    )}
                    <span>AI 첨부 문서 정밀 분석</span>
                  </button>
                  <a
                    href={downloadUrl}
                    download={currentAttachment.att.filename}
                    className="py-1.5 px-2.5 rounded bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] font-semibold inline-flex items-center gap-1 shrink-0"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>저장</span>
                  </a>
                </div>
              )}

              {/* Live Visual Preview Surface */}
              {currentAttachment && (
                <div className="rounded-lg border border-[#D1D1D1] bg-slate-50 overflow-hidden">
                  <div className="px-2.5 py-1.5 bg-slate-100 border-b border-slate-200 flex items-center justify-between text-[11px]">
                    <span className="font-bold text-slate-700 truncate">
                      미리보기: {currentAttachment.att.filename}
                    </span>
                    <span className="text-[10px] uppercase font-mono text-slate-500">
                      {currentAttachment.att.kind}
                    </span>
                  </div>

                  {currentAttachment.att.kind === 'pdf' ? (
                    <iframe
                      title={currentAttachment.att.filename}
                      src={inlinePreviewUrl}
                      className="w-full h-[320px] bg-white border-0"
                    />
                  ) : IMAGE_EXT_RE.test(currentAttachment.att.filename) ? (
                    <div className="p-2 bg-white flex items-center justify-center max-h-[320px] overflow-auto">
                      <img
                        src={inlinePreviewUrl}
                        alt={currentAttachment.att.filename}
                        className="max-w-full max-h-[290px] object-contain rounded"
                      />
                    </div>
                  ) : parsing ? (
                    <div className="p-6 text-center text-slate-500 flex flex-col items-center gap-2">
                      <Loader2 className="w-5 h-5 animate-spin text-[#0F6CBD]" />
                      <span>첨부 문서 내용을 추출하는 중...</span>
                    </div>
                  ) : parsedDoc ? (
                    <div className="p-2.5 bg-white space-y-2 max-h-[300px] overflow-y-auto">
                      {parsedDoc.numericHighlights && parsedDoc.numericHighlights.length > 0 && (
                        <div className="flex flex-wrap gap-1">
                          {parsedDoc.numericHighlights.map((m, i) => (
                            <span
                              key={i}
                              className="px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 text-[10px] font-medium"
                            >
                              {m}
                            </span>
                          ))}
                        </div>
                      )}
                      <pre className="whitespace-pre-wrap font-sans text-[11px] text-slate-800 leading-relaxed">
                        {parsedDoc.extractedText || '(추출 가능한 텍스트가 없습니다)'}
                      </pre>
                    </div>
                  ) : (
                    <div className="p-4 text-center text-slate-500 text-[11px]">
                      바이너리 첨부파일입니다. 상단 [저장] 또는 [AI 첨부 문서 정밀 분석]을 이용하세요.
                    </div>
                  )}
                </div>
              )}

              {/* AI Document Analysis Result Card */}
              {docAnalysis && (
                <div className="p-2.5 rounded-lg bg-[#F4F8FD] border border-[#C7E0F4] space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-bold text-[#0F6CBD]">
                      AI 첨부 문서 분석 리포트
                    </span>
                    <span className="text-[10px] text-slate-500">{docAnalysis.providerModel}</span>
                  </div>
                  <p className="text-[11px] text-slate-800 leading-relaxed whitespace-pre-wrap">
                    {docAnalysis.overview}
                  </p>
                  {docAnalysis.keyFindings.length > 0 && (
                    <ul className="space-y-1 text-[11px] text-slate-700">
                      {docAnalysis.keyFindings.map((k, i) => (
                        <li key={i}>• [{k.category}] {k.summary}</li>
                      ))}
                    </ul>
                  )}
                </div>
              )}
            </>
          )}
        </div>
      ) : (
        /* TAB 2: Designated-Folder Batch Attachment Exporter & 3-Format Report Generator */
        <div className="flex-1 overflow-y-auto p-3 space-y-3 text-xs select-text">
          <div className="p-2.5 rounded-lg bg-emerald-50/70 border border-emerald-200 space-y-2">
            <div className="text-[11px] font-bold text-emerald-900">
              1. 첨부파일 저장 폴더 지정
            </div>
            <div className="flex gap-1.5">
              <input
                type="text"
                value={destDir}
                onChange={(e) => setDestDir(e.target.value)}
                placeholder="예: D:\Bmail_Attachments"
                className="flex-1 px-2 py-1.5 text-xs font-mono bg-white border border-emerald-300 rounded focus:outline-none focus:border-emerald-600"
              />
              <button
                type="button"
                onClick={() => void handlePickNativeFolder()}
                className="px-2.5 py-1.5 rounded bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-semibold flex items-center gap-1 shrink-0"
              >
                <FolderOpen className="w-3.5 h-3.5" />
                <span>폴더 선택</span>
              </button>
            </div>
          </div>

          {/* Scope & Multi-Option Filters */}
          <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2.5">
            <div className="text-[11px] font-bold text-slate-800">
              2. 추출 대상 범위 · 기간 · 발신자 옵션
            </div>

            <div className="grid grid-cols-3 gap-1">
              {(
                [
                  { id: 'all', label: '전체 메일함' },
                  { id: 'current-folder', label: '현재 폴더' },
                  { id: 'selected-emails', label: '현재 선택메일' },
                  { id: 'selected-project', label: '특정 프로젝트' },
                  { id: 'date-range', label: '기간 지정' },
                ] as const
              ).map((sc) => (
                <button
                  key={sc.id}
                  type="button"
                  onClick={() => setScope(sc.id)}
                  className={`py-1 px-1.5 rounded text-[11px] font-semibold border ${
                    scope === sc.id
                      ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD]'
                      : 'bg-slate-50 border-slate-200 text-slate-700'
                  }`}
                >
                  {sc.label}
                </button>
              ))}
            </div>

            {scope === 'selected-project' && (
              <div>
                <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                  대상 프로젝트 선택
                </label>
                <select
                  value={projectCode}
                  onChange={(e) => setProjectCode(e.target.value)}
                  className="w-full px-2 py-1 text-xs border border-[#D1D1D1] rounded"
                >
                  {projects.map((p) => (
                    <option key={p.code} value={p.code}>
                      [{p.code}] {p.name}
                    </option>
                  ))}
                </select>
              </div>
            )}

            <div className="grid grid-cols-2 gap-1.5">
              <div>
                <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                  시작일 (전체/일부)
                </label>
                <input
                  type="date"
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                  className="w-full px-2 py-1 text-[11px] border border-[#D1D1D1] rounded"
                />
              </div>
              <div>
                <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                  종료일
                </label>
                <input
                  type="date"
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                  className="w-full px-2 py-1 text-[11px] border border-[#D1D1D1] rounded"
                />
              </div>
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                보낸 사람 필터 (이름·이메일·도메인, 쉼표 구분)
              </label>
              <input
                type="text"
                value={senderQuery}
                onChange={(e) => setSenderQuery(e.target.value)}
                placeholder="예: 최길용, 김정인, mnceng.co.kr (비우면 전체)"
                className="w-full px-2 py-1 text-xs border border-[#D1D1D1] rounded"
              />
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                제목 · 파일명 키워드 (비우면 전체)
              </label>
              <input
                type="text"
                value={keyword}
                onChange={(e) => setKeyword(e.target.value)}
                placeholder="예: POS, KYTS-1037, 승인도, DWG"
                className="w-full px-2 py-1 text-xs border border-[#D1D1D1] rounded"
              />
            </div>

            <div>
              <label className="block text-[10px] font-semibold text-slate-600 mb-1">
                파일 유형 필터 (선택 없으면 모든 유형 저장)
              </label>
              <div className="flex flex-wrap gap-1">
                {(
                  [
                    { id: 'pdf', label: 'PDF 도면·POS' },
                    { id: 'excel', label: 'Excel·CSV' },
                    { id: 'word', label: 'Word' },
                    { id: 'other', label: 'CAD·ZIP·기타' },
                  ] as const
                ).map((k) => (
                  <button
                    key={k.id}
                    type="button"
                    onClick={() => toggleKind(k.id)}
                    className={`px-2 py-0.5 rounded text-[10px] font-semibold border ${
                      selectedKinds.includes(k.id)
                        ? 'bg-[#0F6CBD] text-white border-[#0F6CBD]'
                        : 'bg-slate-100 text-slate-700 border-slate-200'
                    }`}
                  >
                    {k.label}
                  </button>
                ))}
              </div>
            </div>

            <label className="flex items-center gap-1.5 text-[11px] text-slate-700 cursor-pointer">
              <input
                type="checkbox"
                checked={excludeInlineImages}
                onChange={(e) => setExcludeInlineImages(e.target.checked)}
              />
              <span>서명 로고·인라인 작은 이미지(image001.png 등) 자동 제외</span>
            </label>
          </div>

          {/* Folder Structure, Naming Rule & Report Formats */}
          <div className="p-2.5 rounded-lg bg-white border border-[#E0E0E0] space-y-2">
            <div className="text-[11px] font-bold text-slate-800">
              3. 하위 폴더 분류 · 파일명 · 리포트 양식
            </div>

            <div className="grid grid-cols-2 gap-1.5">
              <div>
                <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                  하위 폴더 자동 구성
                </label>
                <select
                  value={folderStructure}
                  onChange={(e) => setFolderStructure(e.target.value as AttachmentFolderStructure)}
                  className="w-full px-1.5 py-1 text-[11px] border border-[#D1D1D1] rounded"
                >
                  <option value="by-project">프로젝트(호선)별 폴더</option>
                  <option value="by-sender">보낸 사람별 폴더</option>
                  <option value="by-date">년-월(YYYY-MM)별 폴더</option>
                  <option value="by-outlook-folder">Outlook 폴더 구조 그대로</option>
                  <option value="by-kind">파일 종류(PDF/EXCEL)별</option>
                  <option value="flat">단일 폴더에 모으기</option>
                </select>
              </div>
              <div>
                <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                  파일명 규칙
                </label>
                <select
                  value={filenamePattern}
                  onChange={(e) => setFilenamePattern(e.target.value as AttachmentFilenamePattern)}
                  className="w-full px-1.5 py-1 text-[11px] border border-[#D1D1D1] rounded"
                >
                  <option value="date-sender-original">날짜_발신자_원본파일명</option>
                  <option value="project-date-original">프로젝트_날짜_원본파일명</option>
                  <option value="original">원본 파일명 그대로</option>
                </select>
              </div>
            </div>

            <div className="flex items-center justify-between pt-1">
              <span className="text-[10px] font-semibold text-slate-600">동시 생성 리포트:</span>
              <div className="flex gap-1.5">
                {(
                  [
                    { id: 'md', label: 'Markdown (.md)' },
                    { id: 'csv', label: 'Excel (.csv)' },
                    { id: 'html', label: '웹 보고서 (.html)' },
                  ] as const
                ).map((fmt) => (
                  <label key={fmt.id} className="flex items-center gap-1 text-[10px] cursor-pointer">
                    <input
                      type="checkbox"
                      checked={reportFormats.includes(fmt.id)}
                      onChange={() => toggleReportFormat(fmt.id)}
                    />
                    <span>{fmt.label}</span>
                  </label>
                ))}
              </div>
            </div>
          </div>

          {/* 2-Step Action Buttons */}
          <div className="flex gap-1.5">
            <button
              type="button"
              disabled={planning || executing}
              onClick={() => void handleRunExportPlan()}
              className="flex-1 py-2 px-2.5 rounded bg-slate-800 hover:bg-slate-900 text-white text-[11px] font-bold flex items-center justify-center gap-1"
            >
              {planning ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Eye className="w-3.5 h-3.5" />}
              <span>1. 대상 미리보기</span>
            </button>
            <button
              type="button"
              disabled={executing || planning}
              onClick={() => void handleRunExportExecute()}
              className="flex-1 py-2 px-2.5 rounded bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-bold flex items-center justify-center gap-1 shadow-sm"
            >
              {executing ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <FolderCheck className="w-3.5 h-3.5" />
              )}
              <span>2. 폴더 저장 + 리포트</span>
            </button>
          </div>

          {exportError && (
            <div className="p-2 rounded bg-rose-50 border border-rose-200 text-rose-700 text-[11px]">
              {exportError}
            </div>
          )}

          {/* Plan Summary Card */}
          {exportPlan && (
            <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 space-y-1.5 text-[11px]">
              <div className="font-bold text-slate-800">
                추출 대상 확인: 메일 {exportPlan.matchedEmails}건 · 파일 {exportPlan.matchedFiles}개 (
                {formatBytes(exportPlan.totalBytes)})
              </div>
              {exportPlan.skippedInlineCount > 0 && (
                <div className="text-[10px] text-slate-500">
                  • 서명·인라인 이미지 {exportPlan.skippedInlineCount}개 자동 제외됨
                </div>
              )}
              <div className="max-h-36 overflow-y-auto divide-y divide-slate-200 bg-white rounded border p-1.5 text-[10px]">
                {exportPlan.items.slice(0, 40).map((it, idx) => (
                  <div key={idx} className="py-1 flex items-center justify-between gap-2">
                    <span className="truncate font-medium text-slate-800">
                      [{it.projectCode}] {it.filename}
                    </span>
                    <span className="text-slate-400 shrink-0">{it.date}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Execution & Report Result Card */}
          {exportResult && (
            <div className="p-2.5 rounded-lg bg-emerald-50 border border-emerald-300 space-y-2 text-[11px]">
              <div className="flex items-center gap-1.5 font-bold text-emerald-900">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                <span>
                  저장 완료 {exportResult.exportedCount}개 · 동일 스킵{' '}
                  {exportResult.skippedIdenticalCount}개 ({formatBytes(exportResult.totalBytesWritten)})
                </span>
              </div>
              <div className="text-[10px] font-mono text-emerald-950 bg-white p-1.5 rounded border border-emerald-200 space-y-0.5">
                <div>폴더: {exportResult.destDir}</div>
                {exportResult.reportPaths.map((rp) => (
                  <div key={rp.format} className="truncate">
                    리포트({rp.format.toUpperCase()}): {rp.path}
                  </div>
                ))}
              </div>
              <details className="bg-white rounded border border-emerald-200 p-2">
                <summary className="font-bold text-emerald-800 cursor-pointer text-[11px]">
                  생성된 첨부 분석 리포트 본문 미리보기
                </summary>
                <pre className="mt-2 whitespace-pre-wrap font-sans text-[10px] text-slate-700 max-h-56 overflow-y-auto leading-relaxed">
                  {exportResult.reportMarkdown}
                </pre>
              </details>
            </div>
          )}
        </div>
      )}
    </aside>
  );
};
