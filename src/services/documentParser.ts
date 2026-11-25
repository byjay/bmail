import ExcelJS from 'exceljs';
import mammoth from 'mammoth';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type {
  AttachmentKind,
  ExternalEnginesConfig,
  ParsedDocumentSummary,
} from '../types/mail';
import { tryRunExternalDocumentEngine } from './externalEngineBridge';

export function detectAttachmentKind(filename: string, mimeType = ''): AttachmentKind {
  const ext = path.extname(filename).toLowerCase();
  const mime = mimeType.toLowerCase();

  if (ext === '.pdf' || mime.includes('pdf')) return 'pdf';
  if (
    ext === '.xlsx' ||
    ext === '.xls' ||
    ext === '.xlsm' ||
    ext === '.csv' ||
    mime.includes('spreadsheet') ||
    mime.includes('excel') ||
    mime === 'text/csv'
  ) {
    return 'excel';
  }
  if (
    ext === '.docx' ||
    ext === '.doc' ||
    mime.includes('wordprocessingml') ||
    mime.includes('msword')
  ) {
    return 'word';
  }
  if (
    ['.mp3', '.wav', '.m4a', '.webm', '.ogg', '.flac', '.aac', '.mp4', '.mpeg'].includes(ext) ||
    mime.startsWith('audio/')
  ) {
    return 'audio';
  }
  if (['.txt', '.md', '.json', '.xml', '.log', '.tsv', '.html'].includes(ext) || mime.startsWith('text/')) {
    return 'text';
  }
  return 'other';
}

function countWords(text: string): number {
  return text
    .trim()
    .split(/\s+/)
    .filter(Boolean).length;
}

async function parsePdfBuiltIn(buffer: Buffer): Promise<{ text: string; pageCount: number }> {
  const mod = (await import('pdf-parse')) as Record<string, unknown>;
  const PDFParseCtor = mod.PDFParse as
    | (new (opts: { data: Buffer | Uint8Array }) => {
        getText: () => Promise<{ text?: string; total?: number; pages?: unknown[] }>;
        destroy?: () => Promise<void>;
      })
    | undefined;

  if (typeof PDFParseCtor === 'function') {
    const parser = new PDFParseCtor({ data: new Uint8Array(buffer) });
    try {
      const result = await parser.getText();
      const pageCount =
        typeof result.total === 'number'
          ? result.total
          : Array.isArray(result.pages)
            ? result.pages.length
            : 1;
      return { text: (result.text || '').trim(), pageCount };
    } finally {
      await parser.destroy?.().catch(() => undefined);
    }
  }

  // Fallback for v1 default function shape
  const legacyFn = (mod.default ?? mod) as unknown as (
    buf: Buffer
  ) => Promise<{ text?: string; numpages?: number }>;
  if (typeof legacyFn === 'function') {
    const parsed = await legacyFn(buffer);
    return {
      text: (parsed.text || '').trim(),
      pageCount: parsed.numpages || 1,
    };
  }

  throw new Error('내장 pdf-parse 모듈을 초기화하지 못했습니다.');
}

function formatCellValue(raw: unknown): string {
  if (raw === null || raw === undefined) return '';
  if (raw instanceof Date) return raw.toISOString().slice(0, 10);
  if (typeof raw === 'object') {
    const obj = raw as Record<string, unknown>;
    if ('result' in obj && obj.result !== undefined) return formatCellValue(obj.result);
    if ('text' in obj && typeof obj.text === 'string') return obj.text;
    if ('richText' in obj && Array.isArray(obj.richText)) {
      return obj.richText.map((r: { text?: string }) => r.text || '').join('');
    }
    return JSON.stringify(raw);
  }
  return String(raw).trim();
}

async function parseExcelBuiltIn(
  filePath: string,
  filename: string
): Promise<{
  text: string;
  sheetNames: string[];
  rowCount: number;
  numericHighlights: string[];
}> {
  const ext = path.extname(filename).toLowerCase();
  const workbook = new ExcelJS.Workbook();

  if (ext === '.csv') {
    await workbook.csv.readFile(filePath);
  } else {
    await workbook.xlsx.readFile(filePath);
  }

  const sheetNames: string[] = [];
  const markdownParts: string[] = [];
  const numericHighlights: string[] = [];
  let totalRows = 0;

  workbook.eachSheet((worksheet) => {
    sheetNames.push(worksheet.name);
    const rows: string[][] = [];
    const columnNumericStats = new Map<number, { header: string; values: number[] }>();

    let headerRow: string[] = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowIdx) => {
      totalRows += 1;
      const values = (row.values as unknown[]).slice(1).map((v) => formatCellValue(v));
      if (rowIdx === 1) {
        headerRow = values;
      } else {
        values.forEach((cellText, colIdx) => {
          const cleanNum = cellText.replace(/[,\s₩$%원]/g, '');
          if (cleanNum && /^-?\d+(\.\d+)?$/.test(cleanNum)) {
            const num = Number(cleanNum);
            if (Number.isFinite(num)) {
              const entry = columnNumericStats.get(colIdx) ?? {
                header: headerRow[colIdx] || `열 ${colIdx + 1}`,
                values: [],
              };
              entry.values.push(num);
              columnNumericStats.set(colIdx, entry);
            }
          }
        });
      }
      if (rows.length < 80) {
        rows.push(values.slice(0, 16));
      }
    });

    // Deterministic numeric aggregations (works 100% without LLM)
    for (const stat of columnNumericStats.values()) {
      if (stat.values.length >= 2) {
        const sum = stat.values.reduce((acc, v) => acc + v, 0);
        const avg = sum / stat.values.length;
        const max = Math.max(...stat.values);
        numericHighlights.push(
          `[${worksheet.name}] ${stat.header}: 합계 ${sum.toLocaleString('ko-KR')} / 평균 ${Number(
            avg.toFixed(1)
          ).toLocaleString('ko-KR')} / 최대 ${max.toLocaleString('ko-KR')} (${stat.values.length}건)`
        );
      }
    }

    if (rows.length > 0) {
      const header = rows[0].map((h) => h || '-');
      const sep = header.map(() => '---');
      const bodyLines = rows.slice(1, 50).map((r) => {
        const padded = header.map((_, i) => (r[i] || '').replace(/\|/g, '/').replace(/\n/g, ' '));
        return `| ${padded.join(' | ')} |`;
      });
      markdownParts.push(
        [
          `## 시트: ${worksheet.name} (총 ${worksheet.rowCount}행)`,
          `| ${header.join(' | ')} |`,
          `| ${sep.join(' | ')} |`,
          ...bodyLines,
        ].join('\n')
      );
    }
  });

  return {
    text: markdownParts.join('\n\n').slice(0, 28_000),
    sheetNames,
    rowCount: totalRows,
    numericHighlights: numericHighlights.slice(0, 12),
  };
}

async function parseWordBuiltIn(buffer: Buffer): Promise<{ text: string; warnings: string[] }> {
  const result = await mammoth.extractRawText({ buffer });
  return {
    text: (result.value || '').replace(/\n{3,}/g, '\n\n').trim().slice(0, 28_000),
    warnings: (result.messages ?? []).map((m) => m.message).slice(0, 5),
  };
}

/**
 * Parses uploaded PDF, Excel (.xlsx/.csv), Word (.docx), or text documents.
 * Automatically delegates PDF/Excel to user's configured external project folder when present,
 * with seamless fallback to verified MIT/BSD built-in parsers.
 */
export async function parseUploadedDocument(
  file: { path: string; originalname: string; mimetype?: string; size?: number },
  externalConfig?: ExternalEnginesConfig
): Promise<ParsedDocumentSummary> {
  const kind = detectAttachmentKind(file.originalname, file.mimetype || '');
  const stat = fs.statSync(file.path);
  const sizeBytes = file.size ?? stat.size;
  const warnings: string[] = [];

  if (kind === 'pdf') {
    try {
      const externalOut = await tryRunExternalDocumentEngine('pdf', file.path, externalConfig);
      if (externalOut) {
        return {
          id: crypto.randomUUID(),
          filename: file.originalname,
          kind,
          mimeType: file.mimetype || 'application/pdf',
          sizeBytes,
          engineUsed: externalOut.engineLabel,
          pageCount: externalOut.pageCount,
          wordCount: countWords(externalOut.text),
          extractedText: externalOut.text.slice(0, 32_000),
          warnings: externalOut.warnings,
        };
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : '외부 PDF 엔진 호출 실패';
      if (externalConfig?.fallbackToBuiltIn === false) {
        throw new Error(`외부 PDF 엔진 오류 (폴백 비활성화됨): ${msg}`);
      }
      warnings.push(`외부 PDF 엔진 폴백됨: ${msg}`);
    }

    const buf = fs.readFileSync(file.path);
    const parsed = await parsePdfBuiltIn(buf);
    return {
      id: crypto.randomUUID(),
      filename: file.originalname,
      kind,
      mimeType: file.mimetype || 'application/pdf',
      sizeBytes,
      engineUsed: '내장 PDF 엔진 (pdf-parse · MIT)',
      pageCount: parsed.pageCount,
      wordCount: countWords(parsed.text),
      extractedText: parsed.text.slice(0, 32_000),
      warnings,
    };
  }

  if (kind === 'excel') {
    try {
      const externalOut = await tryRunExternalDocumentEngine('excel', file.path, externalConfig);
      if (externalOut) {
        return {
          id: crypto.randomUUID(),
          filename: file.originalname,
          kind,
          mimeType: file.mimetype || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          sizeBytes,
          engineUsed: externalOut.engineLabel,
          sheetNames: externalOut.sheetNames,
          rowCount: externalOut.rowCount,
          numericHighlights: externalOut.numericHighlights,
          wordCount: countWords(externalOut.text),
          extractedText: externalOut.text.slice(0, 32_000),
          warnings: externalOut.warnings,
        };
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : '외부 Excel 엔진 호출 실패';
      if (externalConfig?.fallbackToBuiltIn === false) {
        throw new Error(`외부 Excel 엔진 오류 (폴백 비활성화됨): ${msg}`);
      }
      warnings.push(`외부 Excel 엔진 폴백됨: ${msg}`);
    }

    const parsed = await parseExcelBuiltIn(file.path, file.originalname);
    return {
      id: crypto.randomUUID(),
      filename: file.originalname,
      kind,
      mimeType: file.mimetype || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      sizeBytes,
      engineUsed: '내장 Excel 엔진 (exceljs · MIT)',
      sheetNames: parsed.sheetNames,
      rowCount: parsed.rowCount,
      numericHighlights: parsed.numericHighlights,
      wordCount: countWords(parsed.text),
      extractedText: parsed.text,
      warnings,
    };
  }

  if (kind === 'word') {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.doc') {
      throw new Error('구형 바이너리 .doc 파일은 LibreOffice로 .docx 변환 후 업로드하거나 로컬 에이전트에서 변환해 주세요.');
    }
    const buf = fs.readFileSync(file.path);
    const parsed = await parseWordBuiltIn(buf);
    return {
      id: crypto.randomUUID(),
      filename: file.originalname,
      kind,
      mimeType:
        file.mimetype || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      sizeBytes,
      engineUsed: '내장 Word 엔진 (mammoth · BSD-2)',
      wordCount: countWords(parsed.text),
      extractedText: parsed.text,
      warnings: [...warnings, ...parsed.warnings],
    };
  }

  if (kind === 'text') {
    const text = fs.readFileSync(file.path, 'utf-8').trim().slice(0, 32_000);
    return {
      id: crypto.randomUUID(),
      filename: file.originalname,
      kind,
      mimeType: file.mimetype || 'text/plain',
      sizeBytes,
      engineUsed: 'UTF-8 텍스트 파서',
      wordCount: countWords(text),
      extractedText: text,
      warnings,
    };
  }

  throw new Error(
    `지원하지 않는 문서 형식입니다: ${file.originalname} (.pdf, .xlsx, .csv, .docx, .txt, .md 지원)`
  );
}
