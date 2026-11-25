import type { EmailFolder, InheritedOutlookProfile, MailItem } from '../types/mail';

export interface ParsedMailboxImport {
  storeName: string;
  fileType: 'pst' | 'eml' | 'mbox' | 'msg';
  folders: string[];
  emails: MailItem[];
  skippedCount: number;
  inheritedProfile?: InheritedOutlookProfile;
}

const SENT_FOLDER_PATTERNS = [
  'sent',
  '보낸',
  '발신',
  'outbox',
  '보낼',
];

export function classifyOutlookFolder(folderName: string): EmailFolder {
  const lower = folderName.toLowerCase();
  return SENT_FOLDER_PATTERNS.some((kw) => lower.includes(kw)) ? 'sent' : 'inbox';
}

export function stripHtmlToText(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Decodes RFC 2047 MIME encoded-words (e.g. =?UTF-8?B?...?= or =?UTF-8?Q?...?=)
 * commonly found in Korean Outlook .eml Subject and From headers.
 */
export function decodeMimeHeader(raw: string): string {
  if (!raw) return '';
  return raw
    .replace(/\r?\n[ \t]+/g, ' ')
    .replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (_match, charset: string, enc: string, text: string) => {
      try {
        const normalizedCharset = charset.toLowerCase().trim();
        let bytes: Buffer;
        if (enc.toUpperCase() === 'B') {
          bytes = Buffer.from(text, 'base64');
        } else {
          const qp = text.replace(/_/g, ' ').replace(/=([0-9A-Fa-f]{2})/g, (_m, hex: string) =>
            String.fromCharCode(parseInt(hex, 16))
          );
          bytes = Buffer.from(qp, 'latin1');
        }
        if (normalizedCharset === 'utf-8' || normalizedCharset === 'utf8') {
          return bytes.toString('utf8');
        }
        const decoder = new TextDecoder(normalizedCharset);
        return decoder.decode(bytes);
      } catch {
        return text;
      }
    })
    .trim();
}

function decodeQuotedPrintable(input: string): string {
  const cleaned = input.replace(/=\r?\n/g, '');
  const bytes: number[] = [];
  for (let i = 0; i < cleaned.length; i += 1) {
    if (cleaned[i] === '=' && /^[0-9A-Fa-f]{2}$/.test(cleaned.slice(i + 1, i + 3))) {
      bytes.push(parseInt(cleaned.slice(i + 1, i + 3), 16));
      i += 2;
    } else {
      bytes.push(cleaned.charCodeAt(i));
    }
  }
  return Buffer.from(bytes).toString('utf8');
}

function extractAddressParts(rawHeader: string): { name: string; email: string } {
  const decoded = decodeMimeHeader(rawHeader);
  const angleMatch = decoded.match(/^(.*?)<([^>]+)>/);
  if (angleMatch) {
    const name = angleMatch[1].replace(/^["'\s]+|["'\s]+$/g, '').trim();
    const email = angleMatch[2].trim();
    return { name: name || email, email };
  }
  return { name: decoded || '알 수 없음', email: decoded };
}

function parseDateSafe(dateHeader: string): { date: string; timestamp: string; timeLabel: string } {
  const parsed = dateHeader ? new Date(dateHeader) : new Date();
  const valid = Number.isNaN(parsed.getTime()) ? new Date() : parsed;
  return {
    date: valid.toISOString().slice(0, 10),
    timestamp: valid.toISOString(),
    timeLabel: valid.toTimeString().slice(0, 5),
  };
}

/**
 * Parses a single RFC 2822 .eml message text into a normalized MailItem.
 */
export function parseEmlMessage(rawEml: string, options?: { idPrefix?: string; folderHint?: EmailFolder; category?: string }): MailItem {
  const normalized = rawEml.replace(/\r\n/g, '\n');
  const splitIdx = normalized.indexOf('\n\n');
  const headerBlock = splitIdx === -1 ? normalized : normalized.slice(0, splitIdx);
  let bodyBlock = splitIdx === -1 ? '' : normalized.slice(splitIdx + 2);

  // Unfold multiline headers
  const unfolded = headerBlock.replace(/\n[ \t]+/g, ' ');
  const headers: Record<string, string> = {};
  for (const line of unfolded.split('\n')) {
    const colon = line.indexOf(':');
    if (colon === -1) continue;
    const key = line.slice(0, colon).trim().toLowerCase();
    const val = line.slice(colon + 1).trim();
    if (!headers[key]) headers[key] = val;
  }

  const contentType = headers['content-type'] || 'text/plain';
  const transferEncoding = (headers['content-transfer-encoding'] || '').toLowerCase();

  // Extract multipart text/plain, text/html, and attachment metadata if boundary exists
  const extractedAttachments: NonNullable<MailItem['attachments']> = [];
  const boundaryMatch = contentType.match(/boundary="?([^";\s]+)"?/i);
  if (boundaryMatch) {
    const boundary = boundaryMatch[1];
    const parts = bodyBlock.split(`--${boundary}`);
    let plainPart = '';
    let htmlPart = '';
    for (const part of parts) {
      if (!part.trim() || part.trim() === '--') continue;
      const pSplit = part.indexOf('\n\n');
      if (pSplit === -1) continue;
      const pRawHeaders = part.slice(0, pSplit);
      const pHeaders = pRawHeaders.toLowerCase();
      let pBody = part.slice(pSplit + 2).trim();

      const filenameMatch = pRawHeaders.match(/filename\*?=(?:UTF-8''|")?([^";\r\n]+)"?/i);
      if (filenameMatch || pHeaders.includes('content-disposition: attachment')) {
        const rawFilename = filenameMatch ? decodeMimeHeader(filenameMatch[1].trim()) : 'attachment.bin';
        const mimeMatch = pRawHeaders.match(/content-type:\s*([^;\r\n]+)/i);
        const mime = mimeMatch ? mimeMatch[1].trim().toLowerCase() : 'application/octet-stream';
        const lowerName = rawFilename.toLowerCase();
        const kind = lowerName.endsWith('.pdf')
          ? 'pdf'
          : /\.(xlsx?|csv)$/i.test(lowerName)
            ? 'excel'
            : /\.(docx?)$/i.test(lowerName)
              ? 'word'
              : /\.(mp3|wav|m4a|webm|ogg|flac)$/i.test(lowerName)
                ? 'audio'
                : /\.(txt|md|json)$/i.test(lowerName)
                  ? 'text'
                  : 'other';
        extractedAttachments.push({
          id: `att-${extractedAttachments.length + 1}`,
          filename: rawFilename,
          mimeType: mime,
          sizeBytes: Math.round((pBody.length * 3) / 4),
          kind,
        });
        continue;
      }

      if (pHeaders.includes('content-transfer-encoding: base64')) {
        try {
          pBody = Buffer.from(pBody.replace(/\s+/g, ''), 'base64').toString('utf8');
        } catch {
          // keep raw
        }
      } else if (pHeaders.includes('content-transfer-encoding: quoted-printable')) {
        pBody = decodeQuotedPrintable(pBody);
      }

      if (pHeaders.includes('text/plain') && !plainPart) {
        plainPart = pBody;
      } else if (pHeaders.includes('text/html') && !htmlPart) {
        htmlPart = stripHtmlToText(pBody);
      }
    }
    bodyBlock = plainPart || htmlPart || bodyBlock;
  } else {
    if (transferEncoding.includes('base64')) {
      try {
        bodyBlock = Buffer.from(bodyBlock.replace(/\s+/g, ''), 'base64').toString('utf8');
      } catch {
        // keep raw
      }
    } else if (transferEncoding.includes('quoted-printable')) {
      bodyBlock = decodeQuotedPrintable(bodyBlock);
    }
    if (contentType.toLowerCase().includes('text/html')) {
      bodyBlock = stripHtmlToText(bodyBlock);
    }
  }

  const subject = decodeMimeHeader(headers.subject || '(제목 없음)');
  const from = extractAddressParts(headers.from || '');
  const toList = (headers.to || '')
    .split(',')
    .map((entry) => extractAddressParts(entry).email)
    .filter(Boolean);
  const dt = parseDateSafe(headers.date || '');
  const cleanBody = bodyBlock.trim();

  const xFolder = headers['x-folder'] || headers['x-outlook-folder'] || '';
  const folder: EmailFolder =
    options?.folderHint ?? (xFolder ? classifyOutlookFolder(xFolder) : 'inbox');

  const conversationKey = subject
    .replace(/^((\s*(re|fw|fwd|답장|회신|전달)\s*:\s*)|(\s*\[(회신|전달|답장|외부)\]\s*))+/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  return {
    id: `${options?.idPrefix || 'eml'}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    folder,
    subject,
    fromName: from.name,
    fromEmail: from.email,
    to: toList.length > 0 ? toList : ['me@workspace.kr'],
    date: dt.date,
    timestamp: dt.timestamp,
    timeLabel: dt.timeLabel,
    snippet: cleanBody.replace(/\s+/g, ' ').slice(0, 130),
    body: cleanBody || '(본문 없음)',
    hasAttachment: extractedAttachments.length > 0,
    attachments: extractedAttachments.length > 0 ? extractedAttachments : undefined,
    conversationKey,
    category: options?.category || (xFolder || 'EML 가져오기'),
    isRead: true,
    source: 'pst-import',
  };
}

/**
 * Parses an .mbox archive containing multiple RFC 2822 messages separated by "From " lines.
 */
export function parseMboxArchive(rawMbox: string, storeName = 'archive.mbox'): ParsedMailboxImport {
  const normalized = rawMbox.replace(/\r\n/g, '\n');
  const chunks = normalized
    .split(/^From [^\n]*\n/m)
    .map((chunk) => chunk.trim())
    .filter((chunk) => chunk.length > 0 && chunk.includes(':'));

  const emails: MailItem[] = [];
  let skippedCount = 0;
  for (let i = 0; i < chunks.length; i += 1) {
    try {
      emails.push(parseEmlMessage(chunks[i], { idPrefix: `mbox-${i}`, category: storeName }));
    } catch {
      skippedCount += 1;
    }
  }

  return {
    storeName,
    fileType: 'mbox',
    folders: [storeName],
    emails,
    skippedCount,
  };
}
