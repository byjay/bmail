import type { InheritedOutlookProfile, MailItem } from '../types/mail';
import { extractDeterministicSchedules, type ProjectDefinition } from './deterministicAnalyzer';

const SYSTEM_FOLDER_PATTERNS = [
  /^inbox$/i,
  /^sent(\s*items)?$/i,
  /^drafts?$/i,
  /^deleted(\s*items)?$/i,
  /^junk(\s*e-?mail)?$/i,
  /^outbox$/i,
  /^calendar$/i,
  /^contacts$/i,
  /^tasks$/i,
  /^notes$/i,
  /^journal$/i,
  /^rss\s*feeds?$/i,
  /^sync\s*issues/i,
  /^conversation\s*history/i,
  /^clutter$/i,
  /^archive$/i,
  /^기본\s*폴더$/,
  /^받은\s*편지함(\s*\(eml\))?$/i,
  /^보낸\s*편지함(\s*\(eml\))?$/i,
  /^임시\s*보관함$/,
  /^지운\s*편지함$/,
  /^정크\s*메일$/,
  /^보낼\s*편지함$/,
  /^일정$/,
  /^연락처$/,
  /^작업$/,
  /^메모$/,
  /^대화\s*기록$/,
  /^동기화\s*문제/,
  /^eml\s*파일$/i,
];

const PROJECT_COLORS = [
  '#0F6CBD',
  '#0B7A75',
  '#6B4C9A',
  '#B45309',
  '#BE123C',
  '#1D4ED8',
  '#047857',
];

export function isSystemOutlookFolder(folderPath: string): boolean {
  const leaf = folderPath
    .split('/')
    .map((s) => s.trim())
    .filter(Boolean)
    .pop();
  if (!leaf) return true;
  return SYSTEM_FOLDER_PATTERNS.some((rx) => rx.test(leaf));
}

function makeFolderProjectCode(folderLeaf: string, index: number): string {
  const explicit = folderLeaf.match(/\b([A-Z]{2,6}[-_]\d{2,6}(?:[-_][A-Z0-9]{1,6})?)\b/i);
  if (explicit) return explicit[1].toUpperCase();
  const asciiTag = folderLeaf
    .replace(/[^a-zA-Z0-9]/g, '')
    .toUpperCase()
    .slice(0, 4);
  if (asciiTag.length >= 2) {
    return `${asciiTag}-${String(index + 1).padStart(2, '0')}`;
  }
  return `PRJ-${String(index + 1).padStart(2, '0')}`;
}

/**
 * Extracts a full Outlook profile (owner account, folder hierarchy stats, auto-discovered projects,
 * date range, and schedule hints) directly from imported PST/OST/EML/MBOX items.
 */
export function buildInheritedOutlookProfile(params: {
  storeName: string;
  folders: string[];
  emails: MailItem[];
}): InheritedOutlookProfile {
  const safeEmails = Array.isArray(params.emails) ? params.emails : [];
  const safeFolders = Array.isArray(params.folders) ? params.folders : [];

  // 1. Infer Outlook owner email/name from Sent Items frequency (fallback: most frequent recipient)
  const sentSenderCounts = new Map<string, { count: number; name: string }>();
  const recipientCounts = new Map<string, number>();

  for (const mail of safeEmails) {
    if (mail.folder === 'sent' && mail.fromEmail && !mail.fromEmail.includes('unknown@')) {
      const key = mail.fromEmail.trim().toLowerCase();
      const prev = sentSenderCounts.get(key);
      sentSenderCounts.set(key, {
        count: (prev?.count ?? 0) + 1,
        name: mail.fromName || prev?.name || key,
      });
    } else if (mail.folder === 'inbox' && Array.isArray(mail.to)) {
      for (const addr of mail.to) {
        const clean = addr.trim().toLowerCase();
        if (clean && clean.includes('@') && !clean.endsWith('.local')) {
          recipientCounts.set(clean, (recipientCounts.get(clean) ?? 0) + 1);
        }
      }
    }
  }

  let ownerEmail = '';
  let ownerName = '';
  let maxSentCount = 0;
  for (const [email, info] of sentSenderCounts.entries()) {
    if (info.count > maxSentCount) {
      maxSentCount = info.count;
      ownerEmail = email;
      ownerName = info.name;
    }
  }
  if (!ownerEmail) {
    let maxRecip = 0;
    for (const [email, count] of recipientCounts.entries()) {
      if (count > maxRecip) {
        maxRecip = count;
        ownerEmail = email;
        ownerName = email.split('@')[0] || email;
      }
    }
  }

  // 2. Folder stats
  const folderCounts = new Map<string, number>();
  for (const f of safeFolders) {
    if (f.trim()) folderCounts.set(f.trim(), 0);
  }
  for (const mail of safeEmails) {
    const label = (mail.folderName || mail.category || '').trim();
    if (label) {
      folderCounts.set(label, (folderCounts.get(label) ?? 0) + 1);
    }
  }

  const folderStats = Array.from(folderCounts.entries())
    .map(([name, count]) => ({
      name,
      count,
      isSystem: isSystemOutlookFolder(name),
    }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));

  // 3. Auto-discover projects from custom Outlook folders + bracketed subject codes ([PRJ-01])
  const projectMap = new Map<
    string,
    {
      code: string;
      name: string;
      keywords: Set<string>;
      source: 'folder' | 'subject-code';
      mailCount: number;
    }
  >();

  let folderIdx = 0;
  for (const stat of folderStats) {
    if (stat.isSystem || stat.count === 0) continue;
    const leaf = stat.name
      .split('/')
      .map((s) => s.trim())
      .filter(Boolean)
      .pop();
    if (!leaf) continue;
    const code = makeFolderProjectCode(leaf, folderIdx);
    folderIdx += 1;
    const existing = projectMap.get(code);
    if (existing) {
      existing.keywords.add(leaf);
      existing.mailCount += stat.count;
    } else {
      projectMap.set(code, {
        code,
        name: leaf,
        keywords: new Set([leaf, code]),
        source: 'folder',
        mailCount: stat.count,
      });
    }
  }

  const codeRegex = /(?:\[|#)([A-Z]{2,6}[-_]\d{2,6}(?:[-_][A-Z0-9]{1,6})?)\]?/gi;
  for (const mail of safeEmails) {
    const combined = `${mail.subject} ${mail.folderName || ''}`;
    const matches = Array.from(combined.matchAll(codeRegex));
    for (const m of matches) {
      const code = m[1].toUpperCase();
      const existing = projectMap.get(code);
      if (existing) {
        existing.mailCount += 1;
      } else {
        const cleanSubj = mail.subject
          .replace(/^(re|fw|fwd|회신|전달)\s*:\s*/gi, '')
          .replace(codeRegex, '')
          .trim()
          .slice(0, 28);
        projectMap.set(code, {
          code,
          name: cleanSubj || `${code} 프로젝트`,
          keywords: new Set([code]),
          source: 'subject-code',
          mailCount: 1,
        });
      }
    }
  }

  const discoveredProjects = Array.from(projectMap.values())
    .sort((a, b) => b.mailCount - a.mailCount)
    .slice(0, 12)
    .map((p, idx) => ({
      code: p.code,
      name: p.name,
      keywords: Array.from(p.keywords).filter(Boolean).slice(0, 6),
      color: PROJECT_COLORS[idx % PROJECT_COLORS.length],
      isFavorite: idx < 5,
      source: p.source,
      mailCount: p.mailCount,
    }));

  // 4. Date range & counters
  const validDates = safeEmails
    .map((m) => m.date)
    .filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d))
    .sort();
  const earliest = validDates[0] || '';
  const latest = validDates[validDates.length - 1] || '';

  const receivedCount = safeEmails.filter((m) => m.folder === 'inbox').length;
  const sentCount = safeEmails.filter((m) => m.folder === 'sent').length;
  const attachmentMailCount = safeEmails.filter(
    (m) => m.hasAttachment || (m.attachments && m.attachments.length > 0)
  ).length;

  const tempProjects: ProjectDefinition[] = discoveredProjects.map((p, idx) => ({
    id: `tmp-${idx}`,
    code: p.code,
    name: p.name,
    keywords: p.keywords,
    color: p.color,
    isFavorite: p.isFavorite,
    createdAt: new Date().toISOString(),
  }));
  const calendarHintCount = extractDeterministicSchedules(safeEmails, tempProjects).length;

  return {
    storeName: params.storeName || 'Outlook PST',
    ownerEmail,
    ownerName,
    folderStats,
    discoveredProjects,
    dateRange: { earliest, latest },
    receivedCount,
    sentCount,
    attachmentMailCount,
    calendarHintCount,
  };
}

/**
 * Merges auto-discovered PST projects into existing user projects without duplicates or overwriting manual edits.
 */
export function mergeInheritedProjects(
  existing: ProjectDefinition[],
  discovered: InheritedOutlookProfile['discoveredProjects']
): { merged: ProjectDefinition[]; addedCount: number } {
  const existingCodes = new Set(existing.map((p) => p.code.trim().toLowerCase()));
  const additions: ProjectDefinition[] = [];

  for (const item of discovered) {
    const key = item.code.trim().toLowerCase();
    if (!key || existingCodes.has(key)) continue;
    existingCodes.add(key);
    additions.push({
      id: `pst-prj-${Date.now()}-${additions.length}`,
      code: item.code,
      name: item.name,
      keywords: item.keywords,
      color: item.color,
      isFavorite: item.isFavorite,
      createdAt: new Date().toISOString(),
    });
  }

  return {
    merged: [...additions, ...existing],
    addedCount: additions.length,
  };
}
