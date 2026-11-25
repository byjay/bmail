import React, { useMemo, useState } from 'react';
import {
  Briefcase,
  Calendar,
  Check,
  Copy,
  Download,
  FileSpreadsheet,
  FileText,
  Mail,
  Printer,
  Search,
  Users,
} from 'lucide-react';
import type { AttachmentKind } from '../types/mail';
import type { ProjectDefinition, ProjectFolderPreset } from '../utils/deterministicAnalyzer';
import {
  buildFileOrganizationReport,
  FOLDER_PRESET_META,
  formatFileSize,
  formatOrganizationReportCsv,
  formatOrganizationReportMarkdown,
  type FileOrganizerFilter,
  type OrganizedFileEntry,
} from '../utils/projectFileOrganizer';

interface ProjectFileReportPanelProps {
  files: OrganizedFileEntry[];
  projects: ProjectDefinition[];
  preset: ProjectFolderPreset;
  filter: FileOrganizerFilter;
  onChangeFilter: (next: FileOrganizerFilter) => void;
  onSaveDraftToOutlook?: (subject: string, body: string) => void;
  onQuickAssignDomainToProject?: (projectCode: string, domain: string) => void;
}

type ReportSubView = 'by-project' | 'by-date' | 'by-participant' | 'markdown';

export const ProjectFileReportPanel: React.FC<ProjectFileReportPanelProps> = ({
  files,
  projects,
  preset,
  filter,
  onChangeFilter,
  onSaveDraftToOutlook,
}) => {
  const [subView, setSubView] = useState<ReportSubView>('by-project');
  const [copied, setCopied] = useState(false);

  const report = useMemo(() => buildFileOrganizationReport(files, preset), [files, preset]);
  const markdownText = useMemo(() => formatOrganizationReportMarkdown(report), [report]);

  const downloadBlob = (content: string, filename: string, mime: string) => {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleCopyMarkdown = async () => {
    try {
      await navigator.clipboard.writeText(markdownText);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // Clipboard API may be blocked in restricted contexts
    }
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-white border border-[#E0E0E0] rounded-lg overflow-hidden">
      {/* Executive KPI Summary Strip */}
      <div className="grid grid-cols-2 sm:grid-cols-5 border-b border-[#EDEBE9] bg-[#FAF9F8] divide-x divide-[#EDEBE9]">
        <div className="px-4 py-2.5">
          <div className="text-[11px] text-[#616161]">분류 프로젝트</div>
          <div className="text-lg font-bold font-mono text-[#0F6CBD] tabular-nums">
            {report.summary.projectCount}
            <span className="text-[11px] font-normal text-[#616161] ml-1">개</span>
          </div>
        </div>
        <div className="px-4 py-2.5">
          <div className="text-[11px] text-[#616161]">정리 파일 (수신/발신)</div>
          <div className="text-lg font-bold font-mono text-[#242424] tabular-nums">
            {report.summary.totalFiles}
            <span className="text-[11px] font-normal text-[#616161] ml-1">
              ({report.summary.receivedFileCount}/{report.summary.sentFileCount})
            </span>
          </div>
        </div>
        <div className="px-4 py-2.5">
          <div className="text-[11px] text-[#616161]">총 파일 용량</div>
          <div className="text-lg font-bold font-mono text-[#242424] tabular-nums">
            {formatFileSize(report.summary.totalBytes)}
          </div>
        </div>
        <div className="px-4 py-2.5">
          <div className="text-[11px] text-[#616161]">송신인 · 도메인</div>
          <div className="text-lg font-bold font-mono text-[#242424] tabular-nums">
            {report.summary.uniqueSendersCount}
            <span className="text-[11px] font-normal text-[#616161] ml-1">
              명 · {report.summary.uniqueDomainsCount}도메인
            </span>
          </div>
        </div>
        <div className="px-4 py-2.5 col-span-2 sm:col-span-1">
          <div className="text-[11px] text-[#616161]">미분류 · 식별 도면</div>
          <div className="text-lg font-bold font-mono tabular-nums">
            <span className={report.summary.unassignedFileCount > 0 ? 'text-amber-700' : 'text-emerald-700'}>
              {report.summary.unassignedFileCount}
            </span>
            <span className="text-[11px] font-normal text-[#616161] ml-1">
              미분류 · 도면 {report.summary.totalDrawingsDetected}
            </span>
          </div>
        </div>
      </div>

      {/* Filter & Export Toolbar */}
      <div className="px-4 py-2.5 border-b border-[#EDEBE9] flex flex-wrap items-center justify-between gap-2 bg-white">
        <div className="flex flex-wrap items-center gap-1.5">
          <div className="relative flex items-center">
            <Search className="w-3.5 h-3.5 text-[#616161] absolute left-2 pointer-events-none" />
            <input
              type="search"
              aria-label="보고서 내 검색"
              value={filter.query ?? ''}
              onChange={(e) => onChangeFilter({ ...filter, query: e.target.value })}
              placeholder="파일·송신인·도면 검색"
              className="pl-7 pr-2 py-1 text-xs bg-[#FAF9F8] border border-[#D1D1D1] rounded-[4px] w-44 focus:outline-none focus:border-[#0F6CBD]"
            />
          </div>

          <select
            aria-label="수발신 구분 필터"
            value={filter.direction ?? 'all'}
            onChange={(e) =>
              onChangeFilter({
                ...filter,
                direction: e.target.value as 'all' | 'received' | 'sent',
              })
            }
            className="px-2 py-1 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
          >
            <option value="all">수신+발신</option>
            <option value="received">수신 파일만</option>
            <option value="sent">발신 파일만</option>
          </select>

          <select
            aria-label="파일 유형 필터"
            value={filter.fileKind ?? 'all'}
            onChange={(e) =>
              onChangeFilter({
                ...filter,
                fileKind: e.target.value as 'all' | AttachmentKind,
              })
            }
            className="px-2 py-1 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
          >
            <option value="all">모든 파일 유형</option>
            <option value="pdf">PDF</option>
            <option value="excel">Excel</option>
            <option value="word">Word</option>
            <option value="audio">음성</option>
            <option value="text">텍스트</option>
            <option value="other">기타</option>
          </select>

          <input
            type="text"
            aria-label="송신인 또는 도메인 필터"
            value={filter.senderOrDomain ?? ''}
            onChange={(e) => onChangeFilter({ ...filter, senderOrDomain: e.target.value })}
            placeholder="송신인/도메인 (@vendor)"
            className="px-2 py-1 text-xs bg-white border border-[#D1D1D1] rounded-[4px] w-36"
          />

          <input
            type="text"
            aria-label="계정 필터"
            value={filter.account ?? ''}
            onChange={(e) => onChangeFilter({ ...filter, account: e.target.value })}
            placeholder="연결 계정 필터"
            className="px-2 py-1 text-xs bg-white border border-[#D1D1D1] rounded-[4px] w-32"
          />
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            disabled={files.length === 0}
            onClick={() =>
              downloadBlob(
                markdownText,
                `bmail-file-organization-report-${new Date().toISOString().slice(0, 10)}.md`,
                'text/markdown;charset=utf-8'
              )
            }
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[4px]"
          >
            <Download className="w-3.5 h-3.5" />
            보고서 (.md)
          </button>

          <button
            type="button"
            disabled={files.length === 0}
            onClick={() =>
              downloadBlob(
                formatOrganizationReportCsv(files),
                `bmail-file-organization-${new Date().toISOString().slice(0, 10)}.csv`,
                'text/csv;charset=utf-8'
              )
            }
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F3F2F1] disabled:opacity-40 rounded-[4px]"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
            CSV 엑셀
          </button>

          <button
            type="button"
            onClick={() => void handleCopyMarkdown()}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F3F2F1] rounded-[4px]"
          >
            {copied ? <Check className="w-3.5 h-3.5 text-emerald-700" /> : <Copy className="w-3.5 h-3.5" />}
            {copied ? '복사됨' : '복사'}
          </button>

          {onSaveDraftToOutlook && (
            <button
              type="button"
              disabled={files.length === 0}
              onClick={() =>
                onSaveDraftToOutlook(
                  `[파일 정리 보고서] 프로젝트·일자·송수신인별 집계 (${new Date().toISOString().slice(0, 10)})`,
                  markdownText
                )
              }
              className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] disabled:opacity-40 rounded-[4px]"
            >
              <Mail className="w-3.5 h-3.5" />
              메일 초안 저장
            </button>
          )}

          <button
            type="button"
            onClick={() => window.print()}
            className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-[#616161] hover:text-[#242424] hover:bg-[#F3F2F1] rounded-[4px]"
            title="보고서 인쇄 / PDF 저장"
          >
            <Printer className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Sub-View Pivot Tabs */}
      <div className="px-4 border-b border-[#EDEBE9] flex items-center gap-4 bg-[#FAF9F8]/50 text-xs">
        {[
          { id: 'by-project' as const, label: `프로젝트별 (${report.byProject.length})`, icon: <Briefcase className="w-3.5 h-3.5" /> },
          { id: 'by-date' as const, label: `일자별 (${report.byDate.length})`, icon: <Calendar className="w-3.5 h-3.5" /> },
          { id: 'by-participant' as const, label: `송수신인·도메인별 (${report.byParticipant.length})`, icon: <Users className="w-3.5 h-3.5" /> },
          { id: 'markdown' as const, label: '통합 보고서 원문', icon: <FileText className="w-3.5 h-3.5" /> },
        ].map((tab) => (
          <button
            key={tab.id}
            type="button"
            onClick={() => setSubView(tab.id)}
            className={`inline-flex items-center gap-1.5 py-2.5 border-b-2 transition-colors ${
              subView === tab.id
                ? 'border-[#0F6CBD] text-[#0F6CBD] font-bold'
                : 'border-transparent text-[#616161] hover:text-[#242424]'
            }`}
          >
            {tab.icon}
            {tab.label}
          </button>
        ))}
      </div>

      {/* Report Body */}
      <div className="flex-1 overflow-auto p-4">
        {subView === 'by-project' && (
          <div className="space-y-3">
            {report.byProject.length === 0 ? (
              <div className="p-10 text-center rounded-lg bg-[#FAF9F8] border border-dashed border-[#D1D1D1] text-xs text-[#616161] space-y-1">
                <p className="font-semibold text-[#242424]">정리 분석할 파일 데이터가 아직 없습니다.</p>
                <p>
                  Outlook .pst/.eml 메일함을 불러오거나 <strong>수신파일 폴더 정리</strong> 탭에서 파일을 추가하면 프로젝트·일자·송수신인별 보고서가 즉시 생성됩니다.
                </p>
              </div>
            ) : (
              <div className="overflow-x-auto border border-[#E0E0E0] rounded-lg">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-[#FAF9F8] border-b border-[#E0E0E0] text-[11px] text-[#616161]">
                    <tr>
                      <th className="py-2.5 px-3 font-semibold">프로젝트</th>
                      <th className="py-2.5 px-2 font-semibold text-right">파일 수</th>
                      <th className="py-2.5 px-2 font-semibold text-center">수신 / 발신</th>
                      <th className="py-2.5 px-2 font-semibold text-right">총 용량</th>
                      <th className="py-2.5 px-2 font-semibold">매칭 근거</th>
                      <th className="py-2.5 px-2 font-semibold">관련 도메인 · 계정</th>
                      <th className="py-2.5 px-2 font-semibold">식별 도면</th>
                      <th className="py-2.5 px-3 font-semibold text-right">기간</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#EDEBE9]">
                    {report.byProject.map((row) => {
                      const projMeta = projects.find((p) => p.code === row.projectCode);
                      return (
                        <tr key={row.projectCode} className="hover:bg-[#FAF9F8] align-top">
                          <td className="py-2.5 px-3">
                            <div className="flex items-center gap-1.5">
                              <span
                                className="w-2.5 h-2.5 rounded-full shrink-0"
                                style={{ backgroundColor: projMeta?.color || '#8A8886' }}
                              />
                              <span className="font-mono font-bold text-[#0F6CBD]">
                                {row.projectCode === 'UNASSIGNED' ? '미분류' : row.projectCode}
                              </span>
                            </div>
                            <div className="text-[11px] text-[#424242] pl-4">{row.projectName}</div>
                          </td>
                          <td className="py-2.5 px-2 text-right font-mono font-bold text-[#242424]">
                            {row.fileCount}건
                          </td>
                          <td className="py-2.5 px-2 text-center font-mono">
                            <span className="text-[#0F6CBD]">{row.receivedCount}</span>
                            <span className="text-[#8A8886] mx-1">/</span>
                            <span className="text-emerald-700">{row.sentCount}</span>
                          </td>
                          <td className="py-2.5 px-2 text-right font-mono text-[#424242]">
                            {formatFileSize(row.totalBytes)}
                          </td>
                          <td className="py-2.5 px-2 text-[11px] font-mono text-[#616161]">
                            {[
                              row.matchBreakdown.sender ? `송신인:${row.matchBreakdown.sender}` : '',
                              row.matchBreakdown.domain ? `도메인:${row.matchBreakdown.domain}` : '',
                              row.matchBreakdown.account ? `계정:${row.matchBreakdown.account}` : '',
                              row.matchBreakdown['code-or-keyword']
                                ? `키워드:${row.matchBreakdown['code-or-keyword']}`
                                : '',
                              row.matchBreakdown.manual ? `수동:${row.matchBreakdown.manual}` : '',
                              row.matchBreakdown.unassigned ? `미분류:${row.matchBreakdown.unassigned}` : '',
                            ]
                              .filter(Boolean)
                              .join(' · ')}
                          </td>
                          <td className="py-2.5 px-2 max-w-[200px]">
                            <div className="text-[11px] font-mono text-[#242424] truncate">
                              {row.domains.length > 0 ? row.domains.map((d) => `@${d}`).join(', ') : '-'}
                            </div>
                            {row.accounts.length > 0 && (
                              <div className="text-[10px] font-mono text-[#616161] truncate">
                                계정: {row.accounts.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="py-2.5 px-2">
                            <div className="flex flex-wrap gap-1">
                              {row.drawingKeys.length === 0 ? (
                                <span className="text-[#8A8886]">-</span>
                              ) : (
                                row.drawingKeys.slice(0, 3).map((dk) => (
                                  <span
                                    key={dk}
                                    className="px-1.5 py-0.5 bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[10px] rounded"
                                  >
                                    {dk}
                                  </span>
                                ))
                              )}
                            </div>
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono text-[11px] text-[#616161] whitespace-nowrap">
                            {row.earliestDate === row.latestDate
                              ? row.latestDate
                              : `${row.earliestDate} ~ ${row.latestDate}`}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {subView === 'by-date' && (
          <div className="space-y-3">
            {report.byDate.length === 0 ? (
              <div className="p-8 text-center text-xs text-[#616161]">일자별 파일 내역이 없습니다.</div>
            ) : (
              report.byDate.map((dateRow) => (
                <div key={dateRow.date} className="border border-[#E0E0E0] rounded-lg overflow-hidden">
                  <div className="px-3.5 py-2 bg-[#FAF9F8] border-b border-[#EDEBE9] flex flex-wrap items-center justify-between gap-2 text-xs">
                    <div className="flex items-center gap-2">
                      <Calendar className="w-3.5 h-3.5 text-[#0F6CBD]" />
                      <span className="font-mono font-bold text-[#242424]">{dateRow.date}</span>
                      <span className="text-[#616161]">
                        파일 {dateRow.fileCount}건 (수신 {dateRow.receivedCount} · 발신 {dateRow.sentCount}) ·{' '}
                        {formatFileSize(dateRow.totalBytes)}
                      </span>
                    </div>
                    <div className="flex flex-wrap gap-1">
                      {dateRow.projectCodes.map((code) => (
                        <span
                          key={code}
                          className="px-1.5 py-0.5 bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[10px] font-semibold rounded"
                        >
                          {code === 'UNASSIGNED' ? '미분류' : code}
                        </span>
                      ))}
                    </div>
                  </div>
                  <div className="divide-y divide-[#EDEBE9]">
                    {dateRow.files.map((f) => (
                      <div
                        key={f.id}
                        className="px-3.5 py-2 flex flex-wrap items-center justify-between gap-2 text-xs hover:bg-[#FAF9F8]"
                      >
                        <div className="flex items-center gap-2 min-w-0">
                          <span className="font-mono text-[10px] px-1.5 py-0.5 rounded bg-[#F3F2F1] text-[#323130] shrink-0">
                            {f.projectCode === 'UNASSIGNED' ? '미분류' : f.projectCode}
                          </span>
                          <span className="font-semibold text-[#242424] truncate">{f.filename}</span>
                          <span className="text-[11px] text-[#616161] truncate">
                            ({f.senderName} · {f.senderEmail || '-'})
                          </span>
                        </div>
                        <div className="font-mono text-[11px] text-[#616161] shrink-0">
                          {f.folderPath}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))
            )}
          </div>
        )}

        {subView === 'by-participant' && (
          <div className="overflow-x-auto border border-[#E0E0E0] rounded-lg">
            {report.byParticipant.length === 0 ? (
              <div className="p-8 text-center text-xs text-[#616161]">
                송수신인별 파일 내역이 없습니다.
              </div>
            ) : (
              <table className="w-full text-left border-collapse text-xs">
                <thead className="bg-[#FAF9F8] border-b border-[#E0E0E0] text-[11px] text-[#616161]">
                  <tr>
                    <th className="py-2.5 px-3 font-semibold">송수신인</th>
                    <th className="py-2.5 px-2 font-semibold">이메일 · 도메인</th>
                    <th className="py-2.5 px-2 font-semibold text-center">역할</th>
                    <th className="py-2.5 px-2 font-semibold text-right">파일 수</th>
                    <th className="py-2.5 px-2 font-semibold text-right">용량</th>
                    <th className="py-2.5 px-2 font-semibold">관련 프로젝트</th>
                    <th className="py-2.5 px-3 font-semibold text-right">최근 일자</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#EDEBE9]">
                  {report.byParticipant.map((u) => (
                    <tr key={u.email} className="hover:bg-[#FAF9F8]">
                      <td className="py-2 px-3 font-semibold text-[#242424]">{u.displayName}</td>
                      <td className="py-2 px-2">
                        <div className="font-mono text-[11px] text-[#242424]">{u.email}</div>
                        {u.domain && (
                          <div className="font-mono text-[10px] text-[#0F6CBD]">@{u.domain}</div>
                        )}
                      </td>
                      <td className="py-2 px-2 text-center">
                        <span className="px-1.5 py-0.5 rounded bg-[#F3F2F1] text-[#323130] text-[10px] font-semibold">
                          {u.role === 'both' ? '송·수신' : u.role === 'sender' ? '송신' : '수신'}
                        </span>
                      </td>
                      <td className="py-2 px-2 text-right font-mono font-bold">{u.fileCount}건</td>
                      <td className="py-2 px-2 text-right font-mono text-[#424242]">
                        {formatFileSize(u.totalBytes)}
                      </td>
                      <td className="py-2 px-2">
                        <div className="flex flex-wrap gap-1">
                          {u.projectCodes.map((code) => (
                            <span
                              key={code}
                              className="px-1.5 py-0.5 bg-[#EBF3FC] text-[#0F6CBD] font-mono text-[10px] rounded"
                            >
                              {code === 'UNASSIGNED' ? '미분류' : code}
                            </span>
                          ))}
                        </div>
                      </td>
                      <td className="py-2 px-3 text-right font-mono text-[11px] text-[#616161]">
                        {u.latestDate}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}

        {subView === 'markdown' && (
          <div className="space-y-2">
            <div className="flex items-center justify-between text-xs text-[#616161]">
              <span>
                폴더 구조 기준: <strong className="text-[#242424]">{FOLDER_PRESET_META[preset].label}</strong>
              </span>
              <span className="font-mono text-[11px]">{markdownText.length.toLocaleString()} chars</span>
            </div>
            <pre className="p-4 bg-[#FAF9F8] border border-[#E0E0E0] rounded-lg text-xs font-mono text-[#242424] whitespace-pre-wrap leading-relaxed overflow-x-auto">
              {markdownText}
            </pre>
          </div>
        )}
      </div>
    </div>
  );
};
