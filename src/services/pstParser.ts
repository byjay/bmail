import fs from 'node:fs';
import path from 'node:path';
import { PSTFile, PSTFolder, PSTMessage } from 'pst-extractor';
import type { AttachmentKind, MailAttachmentMeta, MailItem } from '../types/mail';
import {
  classifyOutlookFolder,
  parseEmlMessage,
  parseMboxArchive,
  stripHtmlToText,
  type ParsedMailboxImport,
} from '../utils/mailFileParsers';
import { normalizeConversationSubject } from '../utils/deterministicAnalyzer';
import { buildInheritedOutlookProfile } from '../utils/pstInheritance';
import { parseMsgFile } from './msgFileParser';

const MAX_EXTRACTED_MESSAGES = 500;

function extractRecipients(displayTo?: string | null): string[] {
  if (!displayTo) return [];
  return displayTo
    .split(/[;,]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function inferAttachmentKind(filename: string, mimeType: string): AttachmentKind {
  const lower = filename.toLowerCase();
  const mime = mimeType.toLowerCase();
  if (lower.endsWith('.pdf') || mime.includes('pdf')) return 'pdf';
  if (/\.(xlsx?|csv)$/i.test(lower) || mime.includes('spreadsheet') || mime.includes('excel')) {
    return 'excel';
  }
  if (/\.(docx?)$/i.test(lower) || mime.includes('wordprocessingml') || mime.includes('msword')) {
    return 'word';
  }
  if (/\.(mp3|wav|m4a|webm|ogg|flac)$/i.test(lower) || mime.startsWith('audio/')) return 'audio';
  if (/\.(txt|md|json)$/i.test(lower) || mime.startsWith('text/')) return 'text';
  return 'other';
}

function extractPstAttachments(msg: PSTMessage, msgKey: string): MailAttachmentMeta[] {
  const count = Number(msg.numberOfAttachments) || 0;
  if (count <= 0) return [];
  const list: MailAttachmentMeta[] = [];
  const limit = Math.min(count, 20);
  for (let i = 0; i < limit; i += 1) {
    try {
      const att = msg.getAttachment(i);
      if (!att) continue;
      const filename = (att.longFilename || att.filename || `attachment-${i + 1}`).trim();
      const mimeType = (att.mimeTag || 'application/octet-stream').trim();
      const sizeBytes = Number(att.filesize || att.size || 0);
      list.push({
        id: `${msgKey}-att-${i}`,
        filename,
        mimeType,
        sizeBytes: Number.isFinite(sizeBytes) && sizeBytes > 0 ? sizeBytes : 0,
        kind: inferAttachmentKind(filename, mimeType),
      });
    } catch (err) {
      const reason = err instanceof Error ? err.message : String(err);
      console.warn(`[pstParser] Skipped attachment #${i} in ${msgKey}: ${reason}`);
    }
  }
  return list;
}

/**
 * Parses a real Microsoft Outlook .pst or .ost binary archive using pst-extractor.
 * Traverses all folders recursively, classifies Inbox vs Sent Items, extracts attachment metadata,
 * and builds an InheritedOutlookProfile (owner email, custom folders, auto-discovered projects, schedules).
 */
export function parsePstBinaryFile(filePath: string, originalFileName: string): ParsedMailboxImport {
  const pstFile = new PSTFile(filePath);
  const storeName = pstFile.getMessageStore()?.displayName || path.basename(originalFileName);
  const rootFolder = pstFile.getRootFolder();

  const foldersFound = new Set<string>();
  const emails: MailItem[] = [];
  let skippedCount = 0;

  const walkFolder = (folder: PSTFolder, trail: string[]) => {
    const folderTitle = (folder.displayName || '').trim();
    const nextTrail = folderTitle ? [...trail, folderTitle] : trail;
    const displayPath =
      nextTrail.filter((seg) => !seg.toLowerCase().startsWith('top of')).join(' / ') || '기본 폴더';

    if (folder.contentCount > 0) {
      if (folderTitle) foldersFound.add(displayPath);
      const classifiedFolder = classifyOutlookFolder(displayPath);

      let child = folder.getNextChild();
      while (child != null) {
        if (emails.length >= MAX_EXTRACTED_MESSAGES) {
          skippedCount += 1;
        } else if (child instanceof PSTMessage) {
          const msg = child as PSTMessage;
          try {
            const rawDate =
              msg.messageDeliveryTime || msg.clientSubmitTime || msg.creationTime || new Date();
            const dt = rawDate instanceof Date && !Number.isNaN(rawDate.getTime()) ? rawDate : new Date();
            const dateStr = dt.toISOString().slice(0, 10);
            const timeLabel = dt.toTimeString().slice(0, 5);

            const plainBody = (msg.body || '').trim();
            const htmlBody = (msg.bodyHTML || '').trim();
            const finalBody = plainBody || (htmlBody ? stripHtmlToText(htmlBody) : '') || '(본문 없음)';

            const senderEmail =
              msg.senderEmailAddress ||
              msg.sentRepresentingEmailAddress ||
              msg.senderName ||
              'unknown@outlook.pst';
            const senderName =
              msg.senderName || msg.sentRepresentingName || senderEmail || '알 수 없음';

            const recipients = extractRecipients(msg.displayTo);
            const subject = (msg.subject || '').trim() || '(제목 없음)';
            const msgId = `pst-${msg.descriptorNodeId || emails.length}-${dt.getTime()}`;
            const attachments = extractPstAttachments(msg, msgId);

            emails.push({
              id: msgId,
              folder: classifiedFolder,
              subject,
              fromName: senderName,
              fromEmail: senderEmail,
              to: recipients.length > 0 ? recipients : ['me@outlook.local'],
              date: dateStr,
              timestamp: dt.toISOString(),
              timeLabel,
              snippet: finalBody.replace(/\s+/g, ' ').slice(0, 130),
              body: finalBody.slice(0, 8000),
              hasAttachment: Boolean(msg.hasAttachments) || attachments.length > 0,
              attachments: attachments.length > 0 ? attachments : undefined,
              conversationKey: normalizeConversationSubject(subject),
              importance: msg.importance >= 2 ? 'high' : 'normal',
              category: displayPath,
              folderName: displayPath,
              isRead: Boolean(msg.isRead),
              source: 'pst-import',
            });
          } catch (err) {
            skippedCount += 1;
            const reason = err instanceof Error ? err.message : String(err);
            console.warn(`[pstParser] Skipped malformed message in "${displayPath}": ${reason}`);
          }
        }
        child = folder.getNextChild();
      }
    }

    if (folder.hasSubfolders) {
      const subFolders = folder.getSubFolders();
      for (const sub of subFolders) {
        walkFolder(sub, nextTrail);
      }
    }
  };

  walkFolder(rootFolder, []);
  emails.sort((a, b) => b.timestamp.localeCompare(a.timestamp));

  const folders = Array.from(foldersFound);
  const inheritedProfile = buildInheritedOutlookProfile({
    storeName,
    folders,
    emails,
  });

  return {
    storeName,
    fileType: 'pst',
    folders,
    emails,
    skippedCount,
    inheritedProfile,
  };
}

export function parseUploadedMailFiles(
  files: Array<{ path: string; originalname: string }>
): ParsedMailboxImport {
  if (files.length === 0) {
    throw new Error('업로드된 메일 파일이 없습니다.');
  }

  const allEmails: MailItem[] = [];
  const allFolders = new Set<string>();
  let totalSkipped = 0;
  let primaryStore = files[0].originalname;
  let primaryType: 'pst' | 'eml' | 'mbox' | 'msg' = 'pst';

  for (const file of files) {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.pst' || ext === '.ost') {
      const parsed = parsePstBinaryFile(file.path, file.originalname);
      primaryStore = parsed.storeName;
      primaryType = 'pst';
      parsed.folders.forEach((f) => allFolders.add(f));
      allEmails.push(...parsed.emails);
      totalSkipped += parsed.skippedCount;
    } else if (ext === '.mbox') {
      const raw = fs.readFileSync(file.path, 'utf8');
      const parsed = parseMboxArchive(raw, file.originalname);
      primaryType = 'mbox';
      parsed.folders.forEach((f) => allFolders.add(f));
      allEmails.push(...parsed.emails);
      totalSkipped += parsed.skippedCount;
    } else if (ext === '.eml') {
      const raw = fs.readFileSync(file.path, 'utf8');
      const item = parseEmlMessage(raw, {
        idPrefix: 'eml',
        category: 'EML 파일',
      });
      primaryType = 'eml';
      const folderLabel = item.folderName || item.category || (item.folder === 'sent' ? '보낸 편지함 (EML)' : '받은 편지함 (EML)');
      allFolders.add(folderLabel);
      allEmails.push(item);
    } else if (ext === '.msg') {
      const item = parseMsgFile(file.path, file.originalname);
      primaryType = 'msg';
      const folderLabel = item.folderName || item.category || 'MSG 불러오기';
      allFolders.add(folderLabel);
      allEmails.push(item);
    } else {
      throw new Error(
        `지원하지 않는 파일 형식(${ext || '없음'})입니다. Outlook .pst, .ost, .eml, .mbox, .msg 파일을 업로드해 주세요.`
      );
    }
  }

  allEmails.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  const cappedEmails = allEmails.slice(0, MAX_EXTRACTED_MESSAGES);
  const folders = Array.from(allFolders);
  const inheritedProfile = buildInheritedOutlookProfile({
    storeName: primaryStore,
    folders,
    emails: cappedEmails,
  });

  return {
    storeName: primaryStore,
    fileType: primaryType,
    folders,
    emails: cappedEmails,
    skippedCount: totalSkipped,
    inheritedProfile,
  };
}
