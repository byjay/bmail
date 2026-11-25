import React, { useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  CalendarPlus,
  CheckCircle2,
  Copy,
  FileAudio,
  FileSpreadsheet,
  FileText,
  FolderCog,
  Loader2,
  Mail,
  Mic,
  Search,
  Sparkles,
  Trash2,
  Upload,
} from 'lucide-react';
import type {
  AiProviderId,
  DocumentAnalysisResult,
  MailItem,
  MeetingMinutesResult,
  ParsedDocumentSummary,
  ProviderStatus,
} from '../types/mail';
import {
  type ProjectDefinition,
  runDeterministicConversationalSearch,
} from '../utils/deterministicAnalyzer';

interface IntelligenceStudioProps {
  emails: MailItem[];
  providers: ProviderStatus[];
  selectedProvider: AiProviderId;
  maskSensitiveInfo: boolean;
  projects: ProjectDefinition[];
  parsedDocuments: ParsedDocumentSummary[];
  onAddParsedDocuments: (docs: ParsedDocumentSummary[]) => void;
  onRemoveParsedDocument: (id: string) => void;
  meetingMinutesHistory: MeetingMinutesResult[];
  onAddMeetingMinutes: (result: MeetingMinutesResult) => void;
  onSyncMinutesToCalendar: (minutes: MeetingMinutesResult) => void;
  onCreateOutlookDraft: (subject: string, body: string) => void;
  onOpenConfigExternalEngines: () => void;
  initialTab?: 'documents' | 'meeting' | 'search';
  prefillMeetingTitle?: string;
  prefillMeetingDate?: string;
}

type StudioTab = 'documents' | 'meeting' | 'search';

export const IntelligenceStudio: React.FC<IntelligenceStudioProps> = ({
  emails,
  providers,
  selectedProvider,
  maskSensitiveInfo,
  projects,
  parsedDocuments,
  onAddParsedDocuments,
  onRemoveParsedDocument,
  meetingMinutesHistory,
  onAddMeetingMinutes,
  onSyncMinutesToCalendar,
  onCreateOutlookDraft,
  onOpenConfigExternalEngines,
  initialTab = 'documents',
  prefillMeetingTitle = '',
  prefillMeetingDate = '',
}) => {
  const [activeTab, setActiveTab] = useState<StudioTab>(initialTab);
  const [provider, setProvider] = useState<AiProviderId>(selectedProvider);

  // Document Parsing & AI Analysis State
  const docInputRef = useRef<HTMLInputElement | null>(null);
  const [isParsingDocs, setIsParsingDocs] = useState(false);
  const [isAnalyzingDocs, setIsAnalyzingDocs] = useState(false);
  const [docCustomFocus, setDocCustomFocus] = useState('');
  const [docError, setDocError] = useState<string | null>(null);
  const [docAnalysis, setDocAnalysis] = useState<DocumentAnalysisResult | null>(null);

  // Meeting & Audio Analysis State
  const audioInputRef = useRef<HTMLInputElement | null>(null);
  const meetingDocsInputRef = useRef<HTMLInputElement | null>(null);
  const [meetingTitle, setMeetingTitle] = useState(prefillMeetingTitle);
  const [meetingDate, setMeetingDate] = useState(
    prefillMeetingDate || new Date().toISOString().slice(0, 10)
  );
  const [audioFile, setAudioFile] = useState<File | null>(null);
  const [meetingRefFiles, setMeetingRefFiles] = useState<File[]>([]);
  const [meetingNotesText, setMeetingNotesText] = useState('');
  const [meetingCustomFocus, setMeetingCustomFocus] = useState('');
  const [isAnalyzingMeeting, setIsAnalyzingMeeting] = useState(false);
  const [meetingError, setMeetingError] = useState<string | null>(null);
  const [activeMinutes, setActiveMinutes] = useState<MeetingMinutesResult | null>(
    meetingMinutesHistory[0] ?? null
  );
  const [calendarSyncedId, setCalendarSyncedId] = useState<string | null>(null);
  const [copiedKey, setCopiedKey] = useState<string | null>(null);

  // Conversational Fast Deterministic Search State
  const [chatQuery, setChatQuery] = useState('');

  const searchResult = useMemo(() => {
    if (!chatQuery.trim()) return null;
    return runDeterministicConversationalSearch(chatQuery, emails, parsedDocuments, projects);
  }, [chatQuery, emails, parsedDocuments, projects]);

  const handleCopyText = async (key: string, text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedKey(key);
      setTimeout(() => setCopiedKey(null), 1800);
    } catch {
      // ignore clipboard errors
    }
  };

  const handleUploadDocuments = async (fileList: FileList | null) => {
    if (!fileList || fileList.length === 0) return;
    setIsParsingDocs(true);
    setDocError(null);
    try {
      const form = new FormData();
      Array.from(fileList).forEach((file) => form.append('files', file));
      const res = await fetch('/api/documents/parse', {
        method: 'POST',
        body: form,
      });
      const data = (await res.json()) as {
        documents?: ParsedDocumentSummary[];
        error?: string;
        errors?: Array<{ filename: string; error: string }>;
      };
      if (!res.ok || !data.documents) {
        throw new Error(data.error || '문서 파싱에 실패했습니다.');
      }
      onAddParsedDocuments(data.documents);
      if (data.errors && data.errors.length > 0) {
        setDocError(`일부 파일 경고: ${data.errors.map((e) => `${e.filename}(${e.error})`).join(', ')}`);
      }
    } catch (err) {
      setDocError(err instanceof Error ? err.message : '문서 업로드 중 오류가 발생했습니다.');
    } finally {
      setIsParsingDocs(false);
      if (docInputRef.current) docInputRef.current.value = '';
    }
  };

  const handleAnalyzeDocuments = async () => {
    if (parsedDocuments.length === 0) {
      setDocError('먼저 PDF, Excel(.xlsx/.csv), 또는 Word(.docx) 파일을 업로드해 주세요.');
      return;
    }
    setIsAnalyzingDocs(true);
    setDocError(null);
    try {
      const res = await fetch('/api/analyze/documents', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          documents: parsedDocuments,
          emails: emails.slice(0, 12),
          provider,
          maskSensitiveInfo,
          customFocus: docCustomFocus,
        }),
      });
      const data = (await res.json()) as DocumentAnalysisResult & { error?: string };
      if (!res.ok || data.error) {
        throw new Error(data.error || 'AI 문서 분석에 실패했습니다.');
      }
      setDocAnalysis(data);
    } catch (err) {
      setDocError(err instanceof Error ? err.message : '문서 분석 중 오류가 발생했습니다.');
    } finally {
      setIsAnalyzingDocs(false);
    }
  };

  const handleAnalyzeMeeting = async () => {
    if (!audioFile && !meetingNotesText.trim() && meetingRefFiles.length === 0) {
      setMeetingError('음성 파일(.mp3/.wav/.m4a), 회의 메모, 또는 참고 문서 중 하나 이상을 입력해 주세요.');
      return;
    }
    setIsAnalyzingMeeting(true);
    setMeetingError(null);
    try {
      const form = new FormData();
      if (audioFile) form.append('audio', audioFile);
      meetingRefFiles.forEach((f) => form.append('docs', f));
      form.append('title', meetingTitle);
      form.append('meetingDate', meetingDate);
      form.append('notesText', meetingNotesText);
      form.append('customFocus', meetingCustomFocus);
      form.append('provider', provider);
      form.append('maskSensitiveInfo', String(maskSensitiveInfo));
      if (emails.length > 0) {
        form.append('emailsJson', JSON.stringify(emails.slice(0, 10)));
      }

      const res = await fetch('/api/analyze/meeting', {
        method: 'POST',
        body: form,
      });
      const data = (await res.json()) as MeetingMinutesResult & { error?: string };
      if (!res.ok || data.error) {
        throw new Error(data.error || '회의록 및 음성 파일 분석에 실패했습니다.');
      }
      setActiveMinutes(data);
      onAddMeetingMinutes(data);
    } catch (err) {
      setMeetingError(err instanceof Error ? err.message : '회의록 분석 중 오류가 발생했습니다.');
    } finally {
      setIsAnalyzingMeeting(false);
    }
  };

  const handlePickDesktopFiles = async () => {
    const picker = (window as { bmailDesktop?: { pickLocalFiles?: () => Promise<string[]> } })
      .bmailDesktop?.pickLocalFiles;
    if (!picker) return;
    try {
      const picked = await picker();
      if (!Array.isArray(picked) || picked.length === 0) return;
      setIsParsingDocs(true);
      setDocError(null);
      const res = await fetch('/api/local/files/parse-paths', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ paths: picked }),
      });
      const data = (await res.json()) as {
        documents?: ParsedDocumentSummary[];
        errors?: { path?: string; error: string }[];
        error?: string;
      };
      if (!res.ok || data.error) {
        throw new Error(data.error || '로컬 PC 파일 추출에 실패했습니다.');
      }
      if (data.documents && data.documents.length > 0) {
        onAddParsedDocuments(data.documents);
      }
      if (data.errors && data.errors.length > 0) {
        setDocError(
          `일부 파일 경고: ${data.errors.map((e) => `${e.path || '파일'}(${e.error})`).join(', ')}`
        );
      }
    } catch (err) {
      setDocError(err instanceof Error ? err.message : '로컬 파일 읽기 중 오류가 발생했습니다.');
    } finally {
      setIsParsingDocs(false);
    }
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-[#F5F5F5] overflow-hidden">
      {/* Top Fluent 2 Sub-Header & Tabs */}
      <div className="bg-white border-b border-[#E0E0E0] px-6 py-3.5 flex flex-wrap items-center justify-between gap-3 shrink-0">
        <div className="space-y-0.5">
          <h1 className="text-base font-bold text-[#242424] tracking-tight">
            문서(PDF · Excel · Word) · 회의록 · 음성 AI 스튜디오
          </h1>
          <p className="text-xs text-[#616161]">
            코드 기반 자동 추출·집계 및 외부 PDF/Excel 프로젝트 폴더 연동 + AI 정밀 요약·음성 전사(STT)
          </p>
        </div>

        <div className="flex items-center gap-2">
          <select
            aria-label="스튜디오 AI 엔진 선택"
            value={provider}
            onChange={(e) => setProvider(e.target.value as AiProviderId)}
            className="px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[5px] text-[#242424]"
          >
            {providers.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name} {p.configured ? '✓' : '(키 미등록)'}
              </option>
            ))}
          </select>

          <button
            type="button"
            onClick={onOpenConfigExternalEngines}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] rounded-[5px] whitespace-nowrap"
          >
            <FolderCog className="w-3.5 h-3.5" />
            외부 PDF·Excel 폴더 연동
          </button>
        </div>
      </div>

      {/* Mode Switch Tabs */}
      <div className="px-6 pt-3 bg-[#F5F5F5] flex items-center gap-2 shrink-0">
        {(
          [
            { id: 'documents', label: `PDF · Excel · Word 분석 (${parsedDocuments.length})`, icon: FileSpreadsheet },
            { id: 'meeting', label: `회의록 · 음성 파일 분석 (${meetingMinutesHistory.length})`, icon: Mic },
            { id: 'search', label: '대화형 고속 검색 · 검증 (LLM 불필요)', icon: Search },
          ] as const
        ).map((tab) => {
          const Icon = tab.icon;
          const active = activeTab === tab.id;
          return (
            <button
              key={tab.id}
              type="button"
              onClick={() => setActiveTab(tab.id)}
              className={`inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-t-lg border-x border-t transition-colors whitespace-nowrap ${
                active
                  ? 'bg-white border-[#E0E0E0] text-[#0F6CBD]'
                  : 'bg-[#EBEBEB] border-transparent text-[#616161] hover:text-[#242424]'
              }`}
            >
              <Icon className="w-3.5 h-3.5" />
              {tab.label}
            </button>
          );
        })}
      </div>

      {/* Active Workspace Card */}
      <div className="flex-1 min-h-0 px-6 pb-5 overflow-hidden flex">
        <div className="flex-1 bg-white border border-[#E0E0E0] rounded-b-lg rounded-tr-lg shadow-[0_1px_2px_rgba(0,0,0,0.04)] overflow-y-auto p-6">
          {/* TAB 1: PDF / EXCEL / WORD */}
          {activeTab === 'documents' && (
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              <div className="lg:col-span-5 space-y-4">
                <div className="p-5 border-2 border-dashed border-[#B4D6FA] bg-[#FAFBFD] rounded-lg text-center space-y-3">
                  <Upload className="w-7 h-7 text-[#0F6CBD] mx-auto" />
                  <div className="space-y-1">
                    <h2 className="text-sm font-bold text-[#242424]">
                      PDF · Excel(.xlsx/.csv) · Word(.docx) 첨부파일 업로드
                    </h2>
                    <p className="text-xs text-[#616161] leading-relaxed break-keep">
                      설정에 외부 PDF/Excel 프로젝트 폴더가 지정되어 있으면 해당 엔진을 우선 실행하고, 미설정 시 내장 MIT/BSD 파서(pdf-parse·exceljs·mammoth)로 자동 추출합니다.
                    </p>
                  </div>
                  <input
                    ref={docInputRef}
                    type="file"
                    multiple
                    accept=".pdf,.xlsx,.xls,.csv,.docx,.txt,.md,.json"
                    onChange={(e) => void handleUploadDocuments(e.target.files)}
                    className="hidden"
                  />
                  {typeof window !== 'undefined' &&
                    (window as { bmailDesktop?: { pickLocalFiles?: () => Promise<string[]> } })
                      .bmailDesktop?.pickLocalFiles && (
                      <button
                        type="button"
                        disabled={isParsingDocs}
                        onClick={() => void handlePickDesktopFiles()}
                        className="inline-flex items-center gap-1.5 px-4 py-2 mr-2 mb-2 text-xs font-semibold text-[#0F6CBD] bg-white hover:bg-[#EBF3FC] border border-[#0F6CBD]/40 disabled:bg-slate-100 rounded-[5px]"
                      >
                        <FolderCog className="w-3.5 h-3.5" />
                        내 PC 파일 직접 열기 (로컬 분석)
                      </button>
                    )}
                  <button
                    type="button"
                    disabled={isParsingDocs}
                    onClick={() => docInputRef.current?.click()}
                    className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[5px]"
                  >
                    {isParsingDocs ? (
                      <>
                        <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        문서 추출 중...
                      </>
                    ) : (
                      <>
                        <Upload className="w-3.5 h-3.5" />
                        문서 파일 선택 (다중 가능)
                      </>
                    )}
                  </button>
                </div>

                {docError && (
                  <div
                    role="alert"
                    className="flex items-start gap-2 p-3 rounded-lg bg-red-50 border border-red-200 text-xs text-red-800"
                  >
                    <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                    <span>{docError}</span>
                  </div>
                )}

                {/* Parsed Files List (Zero-LLM deterministic stats visible immediately!) */}
                <div className="space-y-2">
                  <div className="flex items-center justify-between text-xs font-semibold text-[#424242]">
                    <span>추출 완료 문서 ({parsedDocuments.length}건)</span>
                    <span className="text-[11px] font-normal text-[#616161]">
                      연결된 메일 {emails.length}건 교차 분석 가능
                    </span>
                  </div>
                  {parsedDocuments.length === 0 ? (
                    <div className="p-4 rounded-lg bg-[#FAF9F8] border border-[#EDEBE9] text-xs text-[#616161]">
                      업로드된 문서가 없습니다. 메일에 첨부된 PDF 계약서, Excel 견적표, Word 기획서를 올리면 즉시 수치·텍스트가 추출됩니다.
                    </div>
                  ) : (
                    <div className="space-y-2 max-h-[320px] overflow-y-auto pr-1">
                      {parsedDocuments.map((doc) => (
                        <div
                          key={doc.id}
                          className="p-3 rounded-lg border border-[#E0E0E0] bg-white space-y-1.5"
                        >
                          <div className="flex items-start justify-between gap-2">
                            <div className="min-w-0">
                              <div className="text-xs font-bold text-[#242424] truncate">
                                {doc.filename}
                              </div>
                              <div className="text-[11px] text-[#0F6CBD] font-medium">
                                {doc.engineUsed}
                              </div>
                            </div>
                            <button
                              type="button"
                              onClick={() => onRemoveParsedDocument(doc.id)}
                              aria-label={`${doc.filename} 제거`}
                              className="p-1 text-[#616161] hover:text-red-700 rounded"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>

                          <div className="flex flex-wrap gap-2 text-[11px] font-mono text-[#616161]">
                            <span>{doc.kind.toUpperCase()}</span>
                            {doc.pageCount && <span>{doc.pageCount}쪽</span>}
                            {doc.sheetNames?.length && <span>시트 {doc.sheetNames.length}개</span>}
                            {doc.rowCount && <span>{doc.rowCount}행</span>}
                            {doc.wordCount && <span>{doc.wordCount}단어</span>}
                          </div>

                          {doc.numericHighlights && doc.numericHighlights.length > 0 && (
                            <div className="p-2 rounded bg-[#FAF9F8] border border-[#EDEBE9] text-[11px] text-[#242424] space-y-0.5">
                              <div className="font-semibold text-[#0F6CBD]">자동 코드 수치 집계 (LLM 미사용):</div>
                              {doc.numericHighlights.slice(0, 3).map((hl, idx) => (
                                <div key={idx} className="truncate font-mono">
                                  • {hl}
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* AI Deep Analysis Trigger */}
                <div className="pt-2 border-t border-[#EDEBE9] space-y-2.5">
                  <label htmlFor="doc-custom-focus" className="block text-xs font-semibold text-[#424242]">
                    AI 정밀 분석 요청사항 (선택)
                  </label>
                  <input
                    id="doc-custom-focus"
                    type="text"
                    value={docCustomFocus}
                    onChange={(e) => setDocCustomFocus(e.target.value)}
                    placeholder="예: 견적서 단가 변동과 계약서 위약금 조항을 메일 내용과 비교해 줘"
                    className="w-full px-3 py-2 text-xs border border-[#D1D1D1] rounded-[5px]"
                  />
                  <button
                    type="button"
                    disabled={isAnalyzingDocs || parsedDocuments.length === 0}
                    onClick={() => void handleAnalyzeDocuments()}
                    className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[5px]"
                  >
                    {isAnalyzingDocs ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        문서·메일 교차 AI 분석 중...
                      </>
                    ) : (
                      <>
                        <Sparkles className="w-4 h-4" />
                        PDF · Excel · Word + 메일 교차 분석 실행 (키 미설정 시 100% 로컬 분석)
                      </>
                    )}
                  </button>
                </div>
              </div>

              {/* Right Column: Document AI Analysis Output */}
              <div className="lg:col-span-7 space-y-4">
                {!docAnalysis ? (
                  <div className="h-full min-h-[340px] flex flex-col items-center justify-center p-8 border border-[#EDEBE9] rounded-lg bg-[#FAF9F8] text-center space-y-2">
                    <FileText className="w-8 h-8 text-[#0F6CBD]" />
                    <h3 className="text-sm font-bold text-[#242424]">문서 분석 결과가 여기에 표시됩니다</h3>
                    <p className="text-xs text-[#616161] max-w-md leading-relaxed break-keep">
                      1단계로 파일 업로드 즉시 코드 기반 텍스트·표·합계 추출이 완료되며, 2단계로 AI 정밀 분석을 누르면 파일별 핵심 수치·리스크·메일 교차 검증 보고서가 생성됩니다.
                    </p>
                  </div>
                ) : (
                  <div className="space-y-4">
                    <div className="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-[#EDEBE9]">
                      <div>
                        <span className="text-[11px] font-mono text-[#0F6CBD] font-semibold">
                          {docAnalysis.providerModel}
                        </span>
                        <h3 className="text-base font-bold text-[#242424]">
                          문서 · 첨부파일 정밀 분석 보고서
                        </h3>
                      </div>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => void handleCopyText('doc-md', docAnalysis.markdownReport)}
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px]"
                        >
                          <Copy className="w-3.5 h-3.5" />
                          {copiedKey === 'doc-md' ? '복사됨' : 'Markdown 복사'}
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            onCreateOutlookDraft(
                              `[문서 분석 보고] ${docAnalysis.files.map((f) => f.filename).join(', ')}`,
                              docAnalysis.markdownReport
                            )
                          }
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[5px]"
                        >
                          <Mail className="w-3.5 h-3.5" />
                          Outlook 메일 초안
                        </button>
                      </div>
                    </div>

                    <div className="p-4 rounded-lg bg-[#EBF3FC] border border-[#B4D6FA] text-xs text-[#242424] leading-relaxed">
                      {docAnalysis.overview}
                    </div>

                    <div className="space-y-2.5">
                      <h4 className="text-xs font-bold text-[#242424]">파일별 핵심 인사이트</h4>
                      {docAnalysis.keyFindings.map((item, idx) => (
                        <div key={idx} className="p-3.5 rounded-lg border border-[#E0E0E0] space-y-1.5">
                          <div className="flex items-center justify-between text-xs">
                            <strong className="text-[#242424]">{item.sourceFile}</strong>
                            <span className="text-[11px] font-medium text-[#0F6CBD]">{item.category}</span>
                          </div>
                          <p className="text-xs text-[#424242] leading-relaxed">{item.summary}</p>
                          {item.metrics.length > 0 && (
                            <ul className="list-disc pl-4 text-[11px] text-[#616161] space-y-0.5">
                              {item.metrics.map((m, mi) => (
                                <li key={mi}>{m}</li>
                              ))}
                            </ul>
                          )}
                        </div>
                      ))}
                    </div>

                    {docAnalysis.risksAndGaps.length > 0 && (
                      <div className="p-3.5 rounded-lg bg-amber-50 border border-amber-200 space-y-1">
                        <h4 className="text-xs font-bold text-amber-900">리스크 및 누락 점검</h4>
                        <ul className="list-disc pl-4 text-xs text-amber-900 space-y-0.5">
                          {docAnalysis.risksAndGaps.map((r, i) => (
                            <li key={i}>{r}</li>
                          ))}
                        </ul>
                      </div>
                    )}

                    {docAnalysis.actionItems.length > 0 && (
                      <div className="space-y-1.5">
                        <h4 className="text-xs font-bold text-[#242424]">후속 실행 항목</h4>
                        <div className="divide-y divide-[#EDEBE9] border border-[#E0E0E0] rounded-lg">
                          {docAnalysis.actionItems.map((act, i) => (
                            <div key={i} className="p-3 text-xs flex items-start justify-between gap-3">
                              <div className="space-y-0.5">
                                <div className="font-semibold text-[#242424]">{act.task}</div>
                                <div className="text-[11px] text-[#616161]">
                                  담당: {act.owner} · 근거: {act.basis}
                                </div>
                              </div>
                              <span className="text-[11px] font-semibold text-[#0F6CBD] shrink-0">
                                {act.priority}
                              </span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 2: MEETING MINUTES + AUDIO TRANSCRIPTION */}
          {activeTab === 'meeting' && (
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              <div className="lg:col-span-5 space-y-4">
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label htmlFor="mtg-title" className="block text-xs font-semibold text-[#424242] mb-1">
                      회의 제목
                    </label>
                    <input
                      id="mtg-title"
                      type="text"
                      value={meetingTitle}
                      onChange={(e) => setMeetingTitle(e.target.value)}
                      placeholder="예: 주간 프로젝트 점검 회의"
                      className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded-[5px]"
                    />
                  </div>
                  <div>
                    <label htmlFor="mtg-date" className="block text-xs font-semibold text-[#424242] mb-1">
                      회의 일자
                    </label>
                    <input
                      id="mtg-date"
                      type="date"
                      value={meetingDate}
                      onChange={(e) => setMeetingDate(e.target.value)}
                      className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded-[5px]"
                    />
                  </div>
                </div>

                {/* Audio File Upload */}
                <div className="p-4 border border-[#B4D6FA] bg-[#FAFBFD] rounded-lg space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="inline-flex items-center gap-1.5 text-xs font-bold text-[#242424]">
                      <FileAudio className="w-4 h-4 text-[#0F6CBD]" />
                      회의 음성 파일 업로드 (MP3 · WAV · M4A · WEBM · OGG)
                    </span>
                    {audioFile && (
                      <button
                        type="button"
                        onClick={() => setAudioFile(null)}
                        className="text-[11px] text-red-700 hover:underline"
                      >
                        제거
                      </button>
                    )}
                  </div>
                  <p className="text-[11px] text-[#616161]">
                    기본 내장 <code className="font-mono">gemini-3.5-transcribe</code>, Groq{' '}
                    <code className="font-mono">whisper-large-v3-turbo</code>, 또는 OpenAI Whisper로 음성을 자동 전사합니다.
                  </p>
                  <input
                    ref={audioInputRef}
                    type="file"
                    accept="audio/*,.mp3,.wav,.m4a,.webm,.ogg,.flac"
                    onChange={(e) => setAudioFile(e.target.files?.[0] ?? null)}
                    className="hidden"
                  />
                  <button
                    type="button"
                    onClick={() => audioInputRef.current?.click()}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-white border border-[#B4D6FA] hover:bg-[#EBF3FC] rounded-[5px]"
                  >
                    <Mic className="w-3.5 h-3.5" />
                    {audioFile ? `선택됨: ${audioFile.name}` : '음성 파일 선택'}
                  </button>
                </div>

                {/* Optional Reference PDF/Excel/Word Docs */}
                <div className="space-y-1.5">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-[#424242]">
                      회의 참고 자료 (PDF · Excel · Word, 선택)
                    </span>
                    <button
                      type="button"
                      onClick={() => meetingDocsInputRef.current?.click()}
                      className="text-xs font-medium text-[#0F6CBD] hover:underline"
                    >
                      + 자료 첨부 ({meetingRefFiles.length})
                    </button>
                    <input
                      ref={meetingDocsInputRef}
                      type="file"
                      multiple
                      accept=".pdf,.xlsx,.csv,.docx,.txt,.md"
                      onChange={(e) => setMeetingRefFiles(Array.from(e.target.files ?? []))}
                      className="hidden"
                    />
                  </div>
                  {meetingRefFiles.length > 0 && (
                    <div className="text-[11px] text-[#616161] font-mono">
                      {meetingRefFiles.map((f) => f.name).join(', ')}
                    </div>
                  )}
                </div>

                {/* Pasted Meeting Notes / Transcript */}
                <div>
                  <label htmlFor="mtg-notes" className="block text-xs font-semibold text-[#424242] mb-1">
                    회의 메모 또는 기존 녹취 텍스트 (선택)
                  </label>
                  <textarea
                    id="mtg-notes"
                    rows={6}
                    value={meetingNotesText}
                    onChange={(e) => setMeetingNotesText(e.target.value)}
                    placeholder="음성 파일 없이 회의 메모나 녹취록 텍스트만 붙여넣어도 즉시 정식 회의록과 달력 Action Item을 생성합니다."
                    className="w-full p-3 text-xs border border-[#D1D1D1] rounded-[5px] leading-relaxed"
                  />
                </div>

                <div>
                  <label htmlFor="mtg-focus" className="block text-xs font-semibold text-[#424242] mb-1">
                    회의록 추가 강조 지시 (선택)
                  </label>
                  <input
                    id="mtg-focus"
                    type="text"
                    value={meetingCustomFocus}
                    onChange={(e) => setMeetingCustomFocus(e.target.value)}
                    placeholder="예: 담당자별 마감일을 YYYY-MM-DD 형식으로 명확히 정리해 줘"
                    className="w-full px-3 py-1.5 text-xs border border-[#D1D1D1] rounded-[5px]"
                  />
                </div>

                {meetingError && (
                  <div
                    role="alert"
                    className="flex items-start gap-2 p-3 rounded-lg bg-red-50 border border-red-200 text-xs text-red-800"
                  >
                    <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                    <span>{meetingError}</span>
                  </div>
                )}

                <button
                  type="button"
                  disabled={isAnalyzingMeeting}
                  onClick={() => void handleAnalyzeMeeting()}
                  className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[5px]"
                >
                  {isAnalyzingMeeting ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      음성 전사 · 회의록 자동 생성 중...
                    </>
                  ) : (
                    <>
                      <Sparkles className="w-4 h-4" />
                      음성 전사 & 정식 회의록 AI 생성
                    </>
                  )}
                </button>
              </div>

              {/* Right Column: Structured Meeting Minutes Output */}
              <div className="lg:col-span-7 space-y-4">
                {!activeMinutes ? (
                  <div className="h-full min-h-[340px] flex flex-col items-center justify-center p-8 border border-[#EDEBE9] rounded-lg bg-[#FAF9F8] text-center space-y-2">
                    <Mic className="w-8 h-8 text-[#0F6CBD]" />
                    <h3 className="text-sm font-bold text-[#242424]">생성된 회의록이 여기에 표시됩니다</h3>
                    <p className="text-xs text-[#616161] max-w-md leading-relaxed break-keep">
                      회의 음성 파일이나 메모를 분석하면 안건별 결론, 확정 의사결정, 담당자별 실행 항목이 추출되며 원클릭으로 Outlook 일정 달력에 자동 등록할 수 있습니다.
                    </p>
                  </div>
                ) : (
                  <div className="space-y-4">
                    <div className="flex flex-wrap items-start justify-between gap-2 pb-3 border-b border-[#EDEBE9]">
                      <div>
                        <div className="text-[11px] font-mono text-[#0F6CBD]">
                          {activeMinutes.meetingDate} · {activeMinutes.providerModel}
                          {activeMinutes.transcriptionEngine ? ` · ${activeMinutes.transcriptionEngine}` : ''}
                        </div>
                        <h3 className="text-base font-bold text-[#242424]">{activeMinutes.title}</h3>
                      </div>
                      <div className="flex flex-wrap items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() => {
                            onSyncMinutesToCalendar(activeMinutes);
                            setCalendarSyncedId(activeMinutes.id);
                          }}
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 hover:bg-emerald-100 rounded-[5px]"
                        >
                          {calendarSyncedId === activeMinutes.id ? (
                            <>
                              <CheckCircle2 className="w-3.5 h-3.5" />
                              달력 반영 완료
                            </>
                          ) : (
                            <>
                              <CalendarPlus className="w-3.5 h-3.5" />
                              일정 달력에 반영 ({activeMinutes.actionItems.length})
                            </>
                          )}
                        </button>
                        <button
                          type="button"
                          onClick={() => void handleCopyText('mtg-md', activeMinutes.markdownMinutes)}
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px]"
                        >
                          <Copy className="w-3.5 h-3.5" />
                          {copiedKey === 'mtg-md' ? '복사됨' : '복사'}
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            onCreateOutlookDraft(
                              `[회의록 공유] ${activeMinutes.title} (${activeMinutes.meetingDate})`,
                              activeMinutes.markdownMinutes
                            )
                          }
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[5px]"
                        >
                          <Mail className="w-3.5 h-3.5" />
                          Outlook 공유 초안
                        </button>
                      </div>
                    </div>

                    <div className="p-3.5 rounded-lg bg-[#EBF3FC] border border-[#B4D6FA] text-xs text-[#242424] leading-relaxed">
                      {activeMinutes.executiveSummary}
                    </div>

                    {activeMinutes.actionItems.length > 0 && (
                      <div className="space-y-1.5">
                        <h4 className="text-xs font-bold text-[#242424]">담당자별 실행 과제 (Action Items)</h4>
                        <div className="divide-y divide-[#EDEBE9] border border-[#E0E0E0] rounded-lg">
                          {activeMinutes.actionItems.map((it, idx) => (
                            <div key={idx} className="p-3 text-xs flex items-center justify-between gap-2">
                              <div>
                                <div className="font-semibold text-[#242424]">{it.task}</div>
                                <div className="text-[11px] text-[#616161]">
                                  담당: {it.owner} · 기한: <span className="font-mono">{it.deadline}</span>
                                </div>
                              </div>
                              <span className="text-[11px] font-semibold text-[#0F6CBD] shrink-0">
                                {it.priority}
                              </span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    {activeMinutes.agendaTopics.length > 0 && (
                      <div className="space-y-2">
                        <h4 className="text-xs font-bold text-[#242424]">안건별 논의 및 합의</h4>
                        {activeMinutes.agendaTopics.map((ag, idx) => (
                          <div key={idx} className="p-3 rounded-lg border border-[#E0E0E0] space-y-1 text-xs">
                            <div className="font-bold text-[#242424]">
                              {idx + 1}. {ag.topic}
                            </div>
                            <p className="text-[#424242]">{ag.discussion}</p>
                            <p className="text-[#0F6CBD] font-medium">결론: {ag.conclusion}</p>
                          </div>
                        ))}
                      </div>
                    )}

                    <details className="border border-[#E0E0E0] rounded-lg p-3 text-xs">
                      <summary className="cursor-pointer font-semibold text-[#424242]">
                        원본 음성 전사록 / 입력 텍스트 보기
                      </summary>
                      <pre className="mt-2 whitespace-pre-wrap font-sans text-[11px] text-[#616161] leading-relaxed max-h-[220px] overflow-y-auto">
                        {activeMinutes.transcript}
                      </pre>
                    </details>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 3: CONVERSATIONAL FAST DETERMINISTIC SEARCH & VERIFICATION */}
          {activeTab === 'search' && (
            <div className="max-w-4xl mx-auto space-y-5">
              <div className="space-y-1.5">
                <h2 className="text-sm font-bold text-[#242424]">
                  대화형 고속 메일 · 문서 검색 및 검증 (LLM 호출 0건 · 즉시 응답)
                </h2>
                <p className="text-xs text-[#616161]">
                  날짜 범위(예: <code className="font-mono">11월 1일부터 11월 19일까지</code>), 프로젝트 코드(예:{' '}
                  <code className="font-mono">PRJ-01</code>), 파일 형식(<code className="font-mono">PDF / 엑셀 / 워드 / 음성</code>), 수신·발신 조건을 문장으로 입력하면 코드 엔진이 5ms 이내로 정확하게 필터링·검증합니다.
                </p>
              </div>

              <div className="flex gap-2">
                <input
                  type="text"
                  value={chatQuery}
                  onChange={(e) => setChatQuery(e.target.value)}
                  placeholder="예: 11월 1일부터 11월 19일까지 PRJ-01 보낸 메일 PDF 견적서 찾아줘"
                  className="flex-1 px-3.5 py-2.5 text-xs border border-[#D1D1D1] rounded-lg focus:outline-none focus:border-[#0F6CBD]"
                />
                {chatQuery && (
                  <button
                    type="button"
                    onClick={() => setChatQuery('')}
                    className="px-3 py-2 text-xs font-medium text-[#616161] bg-[#F5F5F5] hover:bg-[#E0E0E0] rounded-lg"
                  >
                    초기화
                  </button>
                )}
              </div>

              {searchResult && (
                <div className="space-y-3">
                  {/* Parsed Query Verification Bar */}
                  <div className="p-3 rounded-lg bg-[#EBF3FC] border border-[#B4D6FA] flex flex-wrap items-center gap-2 text-[11px] text-[#0F6CBD]">
                    <strong className="font-semibold">자동 파싱된 검증 조건:</strong>
                    {searchResult.parsed.startDate && (
                      <span className="px-2 py-0.5 bg-white rounded border border-[#B4D6FA] font-mono">
                        기간: {searchResult.parsed.startDate} ~ {searchResult.parsed.endDate}
                      </span>
                    )}
                    {searchResult.parsed.projectCode && (
                      <span className="px-2 py-0.5 bg-white rounded border border-[#B4D6FA] font-mono">
                        프로젝트: {searchResult.parsed.projectCode}
                      </span>
                    )}
                    {searchResult.parsed.attachmentKind && (
                      <span className="px-2 py-0.5 bg-white rounded border border-[#B4D6FA] font-mono">
                        파일유형: {searchResult.parsed.attachmentKind.toUpperCase()}
                      </span>
                    )}
                    {searchResult.parsed.folderHint && (
                      <span className="px-2 py-0.5 bg-white rounded border border-[#B4D6FA]">
                        방향: {searchResult.parsed.folderHint === 'sent' ? '보낸 메일' : '받은 메일'}
                      </span>
                    )}
                    {searchResult.parsed.keywords.length > 0 && (
                      <span className="px-2 py-0.5 bg-white rounded border border-[#B4D6FA]">
                        키워드: {searchResult.parsed.keywords.join(', ')}
                      </span>
                    )}
                    <span className="ml-auto font-bold font-mono">
                      총 {searchResult.hits.length}건 검증됨
                    </span>
                  </div>

                  {searchResult.hits.length === 0 ? (
                    <div className="p-8 text-center rounded-lg border border-[#EDEBE9] bg-[#FAF9F8] text-xs text-[#616161]">
                      조건에 일치하는 메일 또는 업로드 문서가 없습니다. (현재 로드된 메일 {emails.length}건 · 문서{' '}
                      {parsedDocuments.length}건)
                    </div>
                  ) : (
                    <div className="divide-y divide-[#EDEBE9] border border-[#E0E0E0] rounded-lg overflow-hidden">
                      {searchResult.hits.map((hit) => (
                        <div key={`${hit.type}-${hit.id}`} className="p-3.5 bg-white space-y-1">
                          <div className="flex items-center justify-between gap-2 text-xs">
                            <div className="flex items-center gap-2 min-w-0">
                              <span
                                className={`px-1.5 py-0.5 rounded text-[10px] font-bold shrink-0 ${
                                  hit.type === 'email'
                                    ? 'bg-[#EBF3FC] text-[#0F6CBD]'
                                    : 'bg-emerald-50 text-emerald-700'
                                }`}
                              >
                                {hit.type === 'email' ? '메일' : '문서'}
                              </span>
                              <strong className="text-[#242424] truncate">{hit.title}</strong>
                            </div>
                            <span className="text-[11px] font-mono text-[#616161] shrink-0">{hit.date}</span>
                          </div>
                          <div className="text-[11px] text-[#616161]">{hit.subtitle}</div>
                          <p className="text-xs text-[#424242] line-clamp-2">{hit.snippet}</p>
                          <div className="flex flex-wrap gap-1.5 pt-1">
                            {hit.matchedReasons.map((r, ri) => (
                              <span
                                key={ri}
                                className="text-[10px] px-1.5 py-0.5 bg-[#FAF9F8] border border-[#E0E0E0] rounded text-[#424242]"
                              >
                                {r}
                              </span>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
