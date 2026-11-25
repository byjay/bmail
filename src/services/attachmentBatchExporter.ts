import fs from 'node:fs';
import path from 'node:path';
import type { AttachmentKind, MailAttachmentMeta, MailItem } from '../types/mail';
import {
  EMPTY_PROJECT_MAIL_OVERRIDES,
  isEmailInProject,
  isInlineBusinessCardImage,
  type ProjectMailOverrides,
} from '../utils/outlookWorkspaceModel';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import type { LocalMailDb } from './localMailDb';
import { readPstAttachmentBytes } from './localOutlookMigrator';

export type AttachmentExportScope =
  | 'all'
  | 'current-folder'
  | 'selected-emails'
  | 'selected-project'
  | 'date-range';

export type AttachmentFolderStructure =
  | 'flat'
  | 'by-date'
  | 'by-sender'
  | 'by-project'
  | 'by-outlook-folder'
  | 'by-kind';

export type AttachmentFilenamePattern =
  | 'original'
  | 'date-sender-original'
  | 'project-date-original';

export type AttachmentConflictPolicy = 'skip-identical' | 'rename' | 'overwrite';
export type AttachmentReportFormat = 'md' | 'csv' | 'html';

export interface AttachmentExportFilterOptions {
  destDir: string;
  scope?: AttachmentExportScope;
  storeName?: string;
  folderPaths?: string[];
  emailIds?: string[];
  projectCode?: string;
  dateFrom?: string;
  dateTo?: string;
  senderQuery?: string;
  keyword?: string;
  kinds?: AttachmentKind[];
  extensions?: string[];
  excludeInlineImages?: boolean;
  minSizeBytes?: number;
  folderStructure?: AttachmentFolderStructure;
  filenamePattern?: AttachmentFilenamePattern;
  conflictPolicy?: AttachmentConflictPolicy;
  reportFormats?: AttachmentReportFormat[];
  maxFiles?: number;
}

export interface AttachmentPlanItem {
  emailId: string;
  subject: string;
  fromName: string;
  fromEmail: string;
  date: string;
  storeName: string;
  folderName: string;
  projectCode: string;
  pstPath: string | null;
  descriptorNodeId: number | null;
  attachmentIndex: number;
  filename: string;
  mimeType: string;
  sizeBytes: number;
  kind: AttachmentKind;
  isInline: boolean;
  targetRelativePath: string;
  targetFullPath: string;
}

export interface AttachmentExportPlan {
  destDir: string;
  scope: AttachmentExportScope;
  matchedEmails: number;
  matchedFiles: number;
  skippedInlineCount: number;
  totalBytes: number;
  summaryByKind: Record<string, number>;
  summaryBySender: Array<{ sender: string; files: number; bytes: number }>;
  summaryByFolder: Array<{ folder: string; files: number; bytes: number }>;
  items: AttachmentPlanItem[];
}

export interface AttachmentExportResult {
  destDir: string;
  exportedCount: number;
  skippedIdenticalCount: number;
  failedCount: number;
  totalBytesWritten: number;
  reportPaths: Array<{ format: AttachmentReportFormat; path: string }>;
  reportMarkdown: string;
  failures: Array<{ filename: string; subject: string; error: string }>;
  sampleExportedFiles: string[];
}

export function classifyAttachmentInline(att: MailAttachmentMeta, bodyHtml?: string): boolean {
  return isInlineBusinessCardImage(att, bodyHtml);
}

export function sanitizeWindowsSegment(raw: string, fallback = 'item', maxLen = 80): string {
  const cleaned = String(raw || '')
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
    .trim();
  const reserved = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;
  const safe = !cleaned || reserved.test(cleaned) ? fallback : cleaned;
  return safe.slice(0, maxLen).trim() || fallback;
}

export function validateExportDestDir(rawDir: string): string {
  const trimmed = String(rawDir || '').trim();
  if (!trimmed || trimmed.includes('\0')) {
    throw new Error('저장할 대상 폴더 경로를 정확히 지정해 주세요.');
  }
  if (/(?:^|[\\/])\.\.(?:$|[\\/])/.test(trimmed)) {
    throw new Error('상대 상위 경로(..)가 포함된 폴더 경로는 허용되지 않습니다.');
  }
  const resolved = path.resolve(trimmed);
  const lower = resolved.toLowerCase();
  const blockedPrefixes = [
    'c:\\windows',
    'c:\\program files',
    'c:\\program files (x86)',
    'c:\\programdata',
  ];
  if (blockedPrefixes.some((p) => lower === p || lower.startsWith(`${p}\\`))) {
    throw new Error(`Windows 시스템 보호 경로(${resolved})에는 저장할 수 없습니다.`);
  }
  if (/^[a-z]:\\$/i.test(resolved)) {
    throw new Error('드라이브 루트 대신 하위 폴더(예: D:\\Bmail_Export)를 지정해 주세요.');
  }
  return resolved;
}

function resolveAttachmentIndex(att: MailAttachmentMeta, arrayIdx: number): number {
  const m = String(att.id || '').match(/-att-(\d+)$/);
  if (m) {
    const parsed = Number(m[1]);
    if (Number.isInteger(parsed) && parsed >= 0) return parsed;
  }
  return arrayIdx;
}

function inferPrimaryProjectCode(
  email: MailItem,
  projects: ProjectDefinition[],
  overrides: ProjectMailOverrides = EMPTY_PROJECT_MAIL_OVERRIDES
): string {
  for (const p of projects) {
    if (isEmailInProject(email, p.code, projects, overrides)) {
      return p.code;
    }
  }
  const codeMatch = `${email.subject} ${email.folderName || ''}`.match(/\b([A-Z]{2,6}[-_]\d{2,6}|\d{2}K[-_\s]?FD|[A-Z]\d{3,4})\b/i);
  if (codeMatch) return codeMatch[1].toUpperCase().replace(/\s+/g, '');
  return 'GENERAL';
}

function buildTargetRelativePath(
  email: MailItem,
  filename: string,
  kind: AttachmentKind,
  projectCode: string,
  structure: AttachmentFolderStructure,
  pattern: AttachmentFilenamePattern
): string {
  const safeOrig = sanitizeWindowsSegment(filename, 'attachment.bin', 140);
  const safeSender = sanitizeWindowsSegment(email.fromName || email.fromEmail || 'unknown', 'unknown', 40);
  const safeDate = sanitizeWindowsSegment(email.date || 'undated', 'undated', 20);
  const safeProject = sanitizeWindowsSegment(projectCode || 'GENERAL', 'GENERAL', 32);
  const folderSegs = String(email.folderName || '받은 편지함')
    .split('/')
    .map((s) => sanitizeWindowsSegment(s, 'folder', 40))
    .filter(Boolean);

  let finalFile = safeOrig;
  if (pattern === 'date-sender-original') {
    finalFile = `${safeDate}_${safeSender}_${safeOrig}`;
  } else if (pattern === 'project-date-original') {
    finalFile = `${safeProject}_${safeDate}_${safeOrig}`;
  }

  let dirParts: string[] = [];
  if (structure === 'by-date') {
    dirParts = [safeDate.slice(0, 7) || safeDate];
  } else if (structure === 'by-sender') {
    dirParts = [safeSender];
  } else if (structure === 'by-project') {
    dirParts = [safeProject];
  } else if (structure === 'by-outlook-folder') {
    dirParts = folderSegs.slice(0, 4);
  } else if (structure === 'by-kind') {
    dirParts = [kind.toUpperCase()];
  }

  return dirParts.length > 0 ? path.join(...dirParts, finalFile) : finalFile;
}

export function buildAttachmentExportPlan(
  db: LocalMailDb,
  options: AttachmentExportFilterOptions,
  projects: ProjectDefinition[] = [],
  overrides: ProjectMailOverrides = EMPTY_PROJECT_MAIL_OVERRIDES
): AttachmentExportPlan {
  const destDir = validateExportDestDir(options.destDir);
  const scope: AttachmentExportScope = options.scope || 'all';
  const structure: AttachmentFolderStructure = options.folderStructure || 'by-project';
  const pattern: AttachmentFilenamePattern = options.filenamePattern || 'original';
  const excludeInline = options.excludeInlineImages !== false;
  const maxFiles = Math.max(1, Math.min(5000, Number(options.maxFiles) || 1500));

  const senderTokens = (options.senderQuery || '')
    .split(/[;,]/)
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);
  const keywordLower = (options.keyword || '').trim().toLowerCase();
  const allowedKinds =
    Array.isArray(options.kinds) && options.kinds.length > 0 ? new Set(options.kinds) : null;
  const allowedExts =
    Array.isArray(options.extensions) && options.extensions.length > 0
      ? new Set(options.extensions.map((e) => (e.startsWith('.') ? e.toLowerCase() : `.${e.toLowerCase()}`)))
      : null;

  const candidateEmails = db.queryEmailsWithAttachments({
    storeName: options.storeName,
    folderPaths: scope === 'current-folder' ? options.folderPaths : undefined,
    emailIds: scope === 'selected-emails' ? options.emailIds : undefined,
    fromDate: options.dateFrom,
    toDate: options.dateTo,
    limit: 5000,
  });

  const items: AttachmentPlanItem[] = [];
  const matchedEmailIds = new Set<string>();
  const summaryByKind: Record<string, number> = {};
  const senderMap = new Map<string, { files: number; bytes: number }>();
  const folderMap = new Map<string, { files: number; bytes: number }>();
  let skippedInlineCount = 0;
  let totalBytes = 0;

  for (const email of candidateEmails) {
    if (items.length >= maxFiles) break;

    if (scope === 'selected-project' && options.projectCode) {
      if (!isEmailInProject(email, options.projectCode, projects, overrides)) continue;
    }

    if (senderTokens.length > 0) {
      const hay = `${email.fromName} ${email.fromEmail}`.toLowerCase();
      if (!senderTokens.some((tok) => hay.includes(tok))) continue;
    }

    const atts = Array.isArray(email.attachments) ? email.attachments : [];
    if (atts.length === 0) continue;

    const projectCode =
      scope === 'selected-project' && options.projectCode
        ? options.projectCode
        : inferPrimaryProjectCode(email, projects, overrides);

    for (let i = 0; i < atts.length; i += 1) {
      if (items.length >= maxFiles) break;
      const att = atts[i];
      const isInline = classifyAttachmentInline(att);
      if (excludeInline && isInline) {
        skippedInlineCount += 1;
        continue;
      }
      if (allowedKinds && !allowedKinds.has(att.kind)) continue;
      const ext = path.extname(att.filename || '').toLowerCase();
      if (allowedExts && (!ext || !allowedExts.has(ext))) continue;
      const sizeBytes = Number(att.sizeBytes || 0);
      if (typeof options.minSizeBytes === 'number' && options.minSizeBytes > 0 && sizeBytes < options.minSizeBytes) {
        continue;
      }

      if (keywordLower) {
        const hay = `${email.subject} ${email.body.slice(0, 600)} ${att.filename}`.toLowerCase();
        if (!hay.includes(keywordLower)) continue;
      }

      const relPath = buildTargetRelativePath(
        email,
        att.filename,
        att.kind,
        projectCode,
        structure,
        pattern
      );
      const targetFullPath = path.join(destDir, relPath);
      const attIndex = resolveAttachmentIndex(att, i);

      matchedEmailIds.add(email.id);
      items.push({
        emailId: email.id,
        subject: email.subject,
        fromName: email.fromName,
        fromEmail: email.fromEmail,
        date: email.date,
        storeName: email.storeName || 'Outlook',
        folderName: email.folderName || email.folder,
        projectCode,
        pstPath: email.pstPath,
        descriptorNodeId: email.descriptorNodeId,
        attachmentIndex: attIndex,
        filename: att.filename,
        mimeType: att.mimeType,
        sizeBytes,
        kind: att.kind,
        isInline,
        targetRelativePath: relPath,
        targetFullPath,
      });

      totalBytes += sizeBytes;
      summaryByKind[att.kind] = (summaryByKind[att.kind] || 0) + 1;

      const senderKey = email.fromName ? `${email.fromName} <${email.fromEmail}>` : email.fromEmail;
      const prevS = senderMap.get(senderKey) || { files: 0, bytes: 0 };
      senderMap.set(senderKey, { files: prevS.files + 1, bytes: prevS.bytes + sizeBytes });

      const folderKey = email.folderName || email.folder;
      const prevF = folderMap.get(folderKey) || { files: 0, bytes: 0 };
      folderMap.set(folderKey, { files: prevF.files + 1, bytes: prevF.bytes + sizeBytes });
    }
  }

  const summaryBySender = Array.from(senderMap.entries())
    .map(([sender, stat]) => ({ sender, ...stat }))
    .sort((a, b) => b.files - a.files)
    .slice(0, 30);

  const summaryByFolder = Array.from(folderMap.entries())
    .map(([folder, stat]) => ({ folder, ...stat }))
    .sort((a, b) => b.files - a.files)
    .slice(0, 30);

  return {
    destDir,
    scope,
    matchedEmails: matchedEmailIds.size,
    matchedFiles: items.length,
    skippedInlineCount,
    totalBytes,
    summaryByKind,
    summaryBySender,
    summaryByFolder,
    items,
  };
}

function escapeCsvCell(val: string | number): string {
  const s = String(val ?? '');
  if (/[",\r\n]/.test(s)) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function escapeHtml(raw: string): string {
  return String(raw || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatBytesHuman(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function makeUniquePath(targetPath: string): string {
  const ext = path.extname(targetPath);
  const stem = ext ? targetPath.slice(0, -ext.length) : targetPath;
  for (let n = 2; n < 500; n += 1) {
    const candidate = `${stem}_(${n})${ext}`;
    if (!fs.existsSync(candidate)) return candidate;
  }
  return `${stem}_${Date.now()}${ext}`;
}

export function buildExportReportMarkdown(
  plan: AttachmentExportPlan,
  exportedRows: Array<AttachmentPlanItem & { status: 'saved' | 'skipped-identical' | 'failed'; savedPath: string; error?: string }>,
  options: AttachmentExportFilterOptions
): string {
  const now = new Date().toISOString();
  const saved = exportedRows.filter((r) => r.status === 'saved');
  const skipped = exportedRows.filter((r) => r.status === 'skipped-identical');
  const failed = exportedRows.filter((r) => r.status === 'failed');

  const lines: string[] = [
    '# Bmail 첨부파일 일괄 저장 및 분석 리포트',
    '',
    `- **생성 시각**: ${now}`,
    `- **저장 폴더**: \`${plan.destDir}\``,
    `- **추출 범위**: \`${plan.scope}\`${options.dateFrom || options.dateTo ? ` (${options.dateFrom || '~'} ~ ${options.dateTo || '~'})` : ''}`,
    options.senderQuery ? `- **발신자 필터**: \`${options.senderQuery}\`` : '',
    options.projectCode ? `- **프로젝트 필터**: \`${options.projectCode}\`` : '',
    options.keyword ? `- **검색어 필터**: \`${options.keyword}\`` : '',
    `- **대상 메일 / 첨부파일**: ${plan.matchedEmails}건 / ${plan.matchedFiles}개 (${formatBytesHuman(plan.totalBytes)})`,
    `- **실제 저장 / 동일 파일 건너뜀 / 실패**: ${saved.length}개 / ${skipped.length}개 / ${failed.length}건`,
    plan.skippedInlineCount > 0
      ? `- **서명·인라인 이미지 자동 제외**: ${plan.skippedInlineCount}개`
      : '',
    '',
    '## 1. 파일 유형별 집계',
    '',
    '| 유형 | 파일 수 |',
    '| --- | ---: |',
    ...Object.entries(plan.summaryByKind).map(([kind, count]) => `| ${kind.toUpperCase()} | ${count} |`),
    '',
    '## 2. 발신자별 첨부 통계 (상위 15)',
    '',
    '| 발신자 | 파일 수 | 용량 |',
    '| --- | ---: | ---: |',
    ...plan.summaryBySender
      .slice(0, 15)
      .map((s) => `| ${s.sender.replace(/\|/g, '/')} | ${s.files} | ${formatBytesHuman(s.bytes)} |`),
    '',
    '## 3. 저장된 첨부파일 상세 목록',
    '',
    '| 일자 | 발신자 | 프로젝트 | 메일 제목 | 파일명 | 크기 | 상태 | 저장 경로 |',
    '| --- | --- | --- | --- | --- | ---: | --- | --- |',
    ...exportedRows.slice(0, 300).map((r) => {
      const statusKo = r.status === 'saved' ? '저장 완료' : r.status === 'skipped-identical' ? '동일 스킵' : '실패';
      return `| ${r.date} | ${r.fromName.replace(/\|/g, '/')} | ${r.projectCode} | ${r.subject
        .slice(0, 48)
        .replace(/\|/g, '/')} | ${r.filename.replace(/\|/g, '/')} | ${formatBytesHuman(
        r.sizeBytes
      )} | ${statusKo} | \`${r.targetRelativePath}\` |`;
    }),
  ].filter(Boolean);

  if (failed.length > 0) {
    lines.push('', '## 4. 추출 실패 항목', '');
    for (const f of failed.slice(0, 50)) {
      lines.push(`- **${f.filename}** (${f.subject}): ${f.error || '추출 오류'}`);
    }
  }

  return lines.join('\n');
}

function buildExportReportCsv(
  exportedRows: Array<AttachmentPlanItem & { status: string; savedPath: string; error?: string }>
): string {
  const header = [
    'Date',
    'FromName',
    'FromEmail',
    'Project',
    'OutlookFolder',
    'Subject',
    'Filename',
    'Kind',
    'SizeBytes',
    'Status',
    'SavedPath',
    'Error',
  ];
  const rows = exportedRows.map((r) =>
    [
      r.date,
      r.fromName,
      r.fromEmail,
      r.projectCode,
      r.folderName,
      r.subject,
      r.filename,
      r.kind,
      r.sizeBytes,
      r.status,
      r.savedPath,
      r.error || '',
    ]
      .map(escapeCsvCell)
      .join(',')
  );
  return `\uFEFF${[header.join(','), ...rows].join('\r\n')}`;
}

function buildExportReportHtml(
  plan: AttachmentExportPlan,
  exportedRows: Array<AttachmentPlanItem & { status: string; savedPath: string; error?: string }>
): string {
  const rowsHtml = exportedRows
    .slice(0, 500)
    .map(
      (r) => `<tr>
        <td>${escapeHtml(r.date)}</td>
        <td>${escapeHtml(r.fromName)}</td>
        <td>${escapeHtml(r.projectCode)}</td>
        <td>${escapeHtml(r.subject)}</td>
        <td><strong>${escapeHtml(r.filename)}</strong></td>
        <td style="text-align:right">${escapeHtml(formatBytesHuman(r.sizeBytes))}</td>
        <td>${escapeHtml(r.status)}</td>
        <td><code>${escapeHtml(r.targetRelativePath)}</code></td>
      </tr>`
    )
    .join('\n');

  return `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>Bmail 첨부파일 추출 리포트</title>
<style>
  body { font-family: 'Malgun Gothic', 'Pretendard', sans-serif; margin: 28px; color: #1b1b1b; }
  h1 { font-size: 20px; color: #0F6CBD; margin-bottom: 8px; }
  .meta { background: #F4F8FD; border: 1px solid #C7E0F4; padding: 12px 16px; border-radius: 6px; margin-bottom: 18px; font-size: 13px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th, td { border: 1px solid #E1E1E1; padding: 6px 8px; text-align: left; }
  th { background: #F3F2F1; color: #323130; }
</style>
</head>
<body>
  <h1>Bmail 첨부파일 일괄 저장 및 분석 리포트</h1>
  <div class="meta">
    <div><strong>저장 위치:</strong> ${escapeHtml(plan.destDir)}</div>
    <div><strong>대상 메일:</strong> ${plan.matchedEmails}건 · <strong>첨부파일:</strong> ${plan.matchedFiles}개 (${escapeHtml(formatBytesHuman(plan.totalBytes))}) · <strong>인라인 제외:</strong> ${plan.skippedInlineCount}개</div>
  </div>
  <table>
    <thead>
      <tr><th>일자</th><th>발신자</th><th>프로젝트</th><th>메일 제목</th><th>파일명</th><th>크기</th><th>상태</th><th>상대 경로</th></tr>
    </thead>
    <tbody>${rowsHtml}</tbody>
  </table>
</body>
</html>`;
}

export function executeAttachmentExportPlan(
  db: LocalMailDb,
  options: AttachmentExportFilterOptions,
  projects: ProjectDefinition[] = [],
  overrides: ProjectMailOverrides = EMPTY_PROJECT_MAIL_OVERRIDES,
  readAttachmentFn: (
    pstPath: string,
    descriptorNodeId: number,
    attachmentIndex: number
  ) => { filename: string; mimeType: string; buffer: Buffer } = readPstAttachmentBytes
): AttachmentExportResult {
  const plan = buildAttachmentExportPlan(db, options, projects, overrides);
  const policy: AttachmentConflictPolicy = options.conflictPolicy || 'skip-identical';
  const formats: AttachmentReportFormat[] =
    Array.isArray(options.reportFormats) && options.reportFormats.length > 0
      ? options.reportFormats
      : ['md', 'csv', 'html'];

  fs.mkdirSync(plan.destDir, { recursive: true });

  const exportedRows: Array<
    AttachmentPlanItem & {
      status: 'saved' | 'skipped-identical' | 'failed';
      savedPath: string;
      error?: string;
    }
  > = [];

  let exportedCount = 0;
  let skippedIdenticalCount = 0;
  let failedCount = 0;
  let totalBytesWritten = 0;
  const sampleExportedFiles: string[] = [];
  const failures: Array<{ filename: string; subject: string; error: string }> = [];

  for (const item of plan.items) {
    let destFile = item.targetFullPath;
    try {
      if (!item.pstPath || item.descriptorNodeId === null) {
        throw new Error('원본 PST 노드 정보가 없어 바이너리를 추출할 수 없습니다.');
      }

      if (fs.existsSync(destFile)) {
        const existingStat = fs.statSync(destFile);
        if (policy === 'skip-identical' && (item.sizeBytes === 0 || existingStat.size === item.sizeBytes)) {
          skippedIdenticalCount += 1;
          exportedRows.push({
            ...item,
            status: 'skipped-identical',
            savedPath: destFile,
          });
          continue;
        }
        if (policy === 'rename') {
          destFile = makeUniquePath(destFile);
        }
      }

      const att = readAttachmentFn(item.pstPath, item.descriptorNodeId, item.attachmentIndex);
      fs.mkdirSync(path.dirname(destFile), { recursive: true });
      fs.writeFileSync(destFile, att.buffer);

      exportedCount += 1;
      totalBytesWritten += att.buffer.length;
      if (sampleExportedFiles.length < 30) {
        sampleExportedFiles.push(destFile);
      }
      exportedRows.push({
        ...item,
        sizeBytes: att.buffer.length,
        status: 'saved',
        savedPath: destFile,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      failedCount += 1;
      failures.push({
        filename: item.filename,
        subject: item.subject,
        error: msg,
      });
      exportedRows.push({
        ...item,
        status: 'failed',
        savedPath: destFile,
        error: msg,
      });
    }
  }

  const reportMarkdown = buildExportReportMarkdown(plan, exportedRows, options);
  const reportPaths: Array<{ format: AttachmentReportFormat; path: string }> = [];
  const stamp = new Date().toISOString().slice(0, 10);

  if (formats.includes('md')) {
    const mdPath = path.join(plan.destDir, `bmail-attachment-report_${stamp}.md`);
    fs.writeFileSync(mdPath, reportMarkdown, 'utf8');
    reportPaths.push({ format: 'md', path: mdPath });
  }
  if (formats.includes('csv')) {
    const csvPath = path.join(plan.destDir, `bmail-attachment-report_${stamp}.csv`);
    fs.writeFileSync(csvPath, buildExportReportCsv(exportedRows), 'utf8');
    reportPaths.push({ format: 'csv', path: csvPath });
  }
  if (formats.includes('html')) {
    const htmlPath = path.join(plan.destDir, `bmail-attachment-report_${stamp}.html`);
    fs.writeFileSync(htmlPath, buildExportReportHtml(plan, exportedRows), 'utf8');
    reportPaths.push({ format: 'html', path: htmlPath });
  }

  return {
    destDir: plan.destDir,
    exportedCount,
    skippedIdenticalCount,
    failedCount,
    totalBytesWritten,
    reportPaths,
    reportMarkdown,
    failures,
    sampleExportedFiles,
  };
}
