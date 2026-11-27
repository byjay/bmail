import { execFile, execFileSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import Long from 'long';
import { PSTFile, PSTFolder, PSTMessage } from 'pst-extractor';
import { PSTContact } from 'pst-extractor/dist/PSTContact.class.js';
import { PSTUtil } from 'pst-extractor/dist/PSTUtil.class.js';
import type { AttachmentKind, InheritedOutlookProfile, MailAttachmentMeta, MailItem } from '../types/mail';
import { normalizeConversationSubject, type ProjectDefinition } from '../utils/deterministicAnalyzer';
import { classifyOutlookFolder, stripHtmlToText } from '../utils/mailFileParsers';
import { buildInheritedOutlookProfile, isSystemOutlookFolder, mergeInheritedProjects } from '../utils/pstInheritance';
import { classifyAttachmentInline } from './attachmentBatchExporter';
import type { OutlookContactRecord } from './outlookLocalBridge';
import type {
  LocalAccountRecord,
  LocalEmailRecord,
  LocalFolderRecord,
  LocalMailDb,
  LocalSignatureRecord,
} from './localMailDb';

const MAX_BODY_CHARS = 16000;
const BATCH_SIZE = 500;

const INTERNAL_ROOT_FOLDERS = new Set([
  '루트 찾기',
  'spams search folder 2',
  'spam search folder 2',
  'ipm_views',
  'ipm_common_views',
  '미리 알림',
  'reminders',
  'itemprocsearch',
  'contact search',
  '할 일 검색',
  'to-do search',
  '메일 추적 처리',
  'freebusy data',
  'subscriptions',
  'finder',
  'search root',
]);

export interface OutlookFavoriteLink {
  name: string;
  ltype: string;
  folderHint: string;
}

export interface OutlookWindowSettings {
  enableSingleLineRibbon?: boolean;
  pinFolderTree?: boolean;
  pinMail?: boolean;
  minimalHeaderOn?: boolean;
  editorPreference?: number;
  moduleVisible15?: string;
  calWorkStart?: string;
  calWorkEnd?: string;
  calWorkDayMask?: number;
  reopenWindowsOption?: number;
  disableTnef?: boolean;
  defaultImportance?: 'low' | 'normal' | 'high';
  requestReadReceipt?: boolean;
  markAsReadWaitSec?: number;
  conversationView?: boolean;
  newestOnTop?: boolean;
  readingPanePosition?: 'right' | 'bottom' | 'hidden';
}

export interface OutlookCategoryEntry {
  name: string;
  color: string;
}

export interface OutlookRuleEntry {
  name: string;
  enabled: boolean;
}

export interface InspectedOutlookSettings {
  outlookVersion: string;
  defaultProfileName: string;
  accounts: LocalAccountRecord[];
  pstPaths: string[];
  signatures: LocalSignatureRecord[];
  outlookFavorites: OutlookFavoriteLink[];
  windowSettings: OutlookWindowSettings;
}

export interface MigrationProgressSummary {
  status: 'idle' | 'running' | 'completed' | 'failed';
  startedAt: string | null;
  completedAt: string | null;
  storesProcessed: number;
  totalEmailsIndexed: number;
  totalContactsIndexed: number;
  totalFoldersIndexed: number;
  totalSignaturesIndexed: number;
  discoveredProjectsCount: number;
  error: string | null;
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

function extractRecipients(displayTo?: string | null): string[] {
  if (!displayTo) return [];
  return displayTo
    .split(/[;,]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function decodeWindowsTextBuffer(buf: Buffer): string {
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) {
    return buf.subarray(2).toString('utf16le');
  }
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return buf.subarray(3).toString('utf8');
  }
  const utf8Text = buf.toString('utf8');
  const declaresKoreanLegacy = /charset\s*=\s*(?:ks_c_5601-1987|euc-kr|cp949)/i.test(utf8Text);
  if (declaresKoreanLegacy || utf8Text.includes('\uFFFD')) {
    try {
      return new TextDecoder('euc-kr').decode(buf);
    } catch {
      return utf8Text;
    }
  }
  return utf8Text;
}

function inferSignatureImageMime(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.png') return 'image/png';
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
  if (ext === '.gif') return 'image/gif';
  if (ext === '.bmp') return 'image/bmp';
  if (ext === '.webp') return 'image/webp';
  if (ext === '.svg') return 'image/svg+xml';
  return 'image/png';
}

/**
 * Inlines relative business-card / logo images referenced by Outlook `.htm` signatures
 * (e.g. `a%20(designsir@seastargo.com).files/image001.png`) as base64 `data:image/...` URIs,
 * strips MSO VML fallback comments while keeping `<![if !vml]>` standard `<img>` tags,
 * and returns clean `<body>` HTML ready for rich-text composition & SMTP MIME sending.
 */
export function inlineSignatureHtmlImages(
  rawHtml: string,
  sigDir: string
): { html: string; inlinedImages: number } {
  if (!rawHtml) return { html: '', inlinedImages: 0 };
  const resolvedBase = path.resolve(sigDir);
  const bodyMatch = rawHtml.match(/<body\b[^>]*>([\s\S]*?)<\/body>/i);
  let cleaned = bodyMatch ? bodyMatch[1] : rawHtml;

  // Strip Word/Outlook VML conditional blocks while preserving downlevel `![if !vml]` HTML <img>
  cleaned = cleaned
    .replace(/<!--\[if\s+gte\s+vml\s+1\]>[\s\S]*?<!\[endif\]-->/gi, '')
    .replace(/<!\[if\s+!?vml\]>/gi, '')
    .replace(/<!\[endif\]>/gi, '')
    .replace(/<o:p>\s*<\/o:p>/gi, '');

  let inlinedImages = 0;
  const imgTagRe = /<img\b([^>]*?)\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))([^>]*)>/gi;
  cleaned = cleaned.replace(
    imgTagRe,
    (full, pre: string, dq: string, sq: string, uq: string, post: string) => {
      const rawSrc = (dq || sq || uq || '').trim();
      if (!rawSrc || /^(?:data:|https?:|cid:)/i.test(rawSrc)) {
        return full;
      }
      let decodedRel = rawSrc;
      try {
        decodedRel = decodeURIComponent(rawSrc);
      } catch {
        // Keep raw relative path if malformed percent-encoding
      }
      const normalizedRel = decodedRel.replace(/^[\\/]+/, '').replace(/\//g, path.sep);
      const targetPath = path.resolve(resolvedBase, normalizedRel);
      const baseWithSep = resolvedBase.endsWith(path.sep)
        ? resolvedBase.toLowerCase()
        : `${resolvedBase.toLowerCase()}${path.sep}`;
      if (!targetPath.toLowerCase().startsWith(baseWithSep)) {
        return full;
      }
      try {
        if (!fs.existsSync(targetPath) || !fs.statSync(targetPath).isFile()) {
          return full;
        }
        const buf = fs.readFileSync(targetPath);
        const mime = inferSignatureImageMime(targetPath);
        const dataUri = `data:${mime};base64,${buf.toString('base64')}`;
        inlinedImages += 1;
        return `<img${pre}src="${dataUri}"${post}>`;
      } catch {
        return full;
      }
    }
  );

  return {
    html: sanitizeEmailHtml(cleaned).trim(),
    inlinedImages,
  };
}

export function readSignaturesFromDisk(
  activeNewSig: string,
  activeReplySig: string,
  sigDirOverride?: string,
  primaryEmailHint?: string
): LocalSignatureRecord[] {
  const appData = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
  const sigDir = sigDirOverride || path.join(appData, 'Microsoft', 'Signatures');
  if (!fs.existsSync(sigDir)) return [];

  const byName = new Map<string, { html: string; text: string }>();
  const entries = fs.readdirSync(sigDir, { withFileTypes: true });
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    const ext = path.extname(entry.name).toLowerCase();
    if (ext !== '.htm' && ext !== '.html' && ext !== '.txt') continue;
    const baseName = entry.name.slice(0, entry.name.length - ext.length);
    if (!baseName) continue;
    const prev = byName.get(baseName) ?? { html: '', text: '' };
    try {
      const buf = fs.readFileSync(path.join(sigDir, entry.name));
      const decoded = decodeWindowsTextBuffer(buf);
      if (ext === '.txt') {
        prev.text = decoded.trim();
      } else {
        const inlined = inlineSignatureHtmlImages(decoded, sigDir);
        prev.html = inlined.html || decoded;
        if (!prev.text) {
          prev.text = stripHtmlToText(prev.html);
        }
      }
      byName.set(baseName, prev);
    } catch {
      // ignore unreadable signature file
    }
  }

  const emailHintLower = String(primaryEmailHint || '').trim().toLowerCase();
  const names = Array.from(byName.keys());
  const hintDefault =
    (!activeNewSig && emailHintLower
      ? names.find((n) => n.toLowerCase().includes(emailHintLower))
      : undefined) ||
    (!activeNewSig && names.length === 1 ? names[0] : undefined);

  return Array.from(byName.entries())
    .map(([name, content]) => ({
      name,
      htmlContent: content.html,
      textContent: content.text || stripHtmlToText(content.html),
      isDefaultNew: Boolean(
        (activeNewSig && name.toLowerCase() === activeNewSig.toLowerCase()) ||
          (hintDefault && name === hintDefault)
      ),
      isDefaultReply: Boolean(
        (activeReplySig && name.toLowerCase() === activeReplySig.toLowerCase()) ||
          (activeNewSig && !activeReplySig && name.toLowerCase() === activeNewSig.toLowerCase()) ||
          (hintDefault && name === hintDefault)
      ),
    }))
    .sort((a, b) => Number(b.isDefaultNew) - Number(a.isDefaultNew) || a.name.localeCompare(b.name, 'ko'));
}

function inferProviderPreset(hostOrEmail: string): string {
  const lower = hostOrEmail.toLowerCase();
  if (lower.includes('mailplug')) return 'mailplug';
  if (lower.includes('hiworks')) return 'hiworks';
  if (lower.includes('gabia')) return 'gabia';
  if (lower.includes('worksmobile')) return 'naverworks';
  if (lower.includes('daouoffice')) return 'daouoffice';
  if (lower.includes('dooray')) return 'dooray';
  if (lower.includes('gmail') || lower.includes('google')) return 'google';
  if (lower.includes('office365') || lower.includes('outlook')) return 'microsoft';
  if (lower.includes('naver')) return 'naver';
  if (lower.includes('daum') || lower.includes('kakao')) return 'daum';
  return 'auto';
}

function decodeBasicXmlEntities(raw: string): string {
  return raw
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .trim();
}

export function parseOutlookFavoritesXmlText(xmlText: string): OutlookFavoriteLink[] {
  if (!xmlText || !xmlText.includes('<wdLnk')) return [];
  const initMailMatch = xmlText.match(/<initMail>([\s\S]*?)<\/initMail>/i);
  const scope = initMailMatch ? initMailMatch[1] : xmlText;
  const results: OutlookFavoriteLink[] = [];
  const seen = new Set<string>();
  const linkRe = /<wdLnk\b[^>]*>([\s\S]*?)<\/wdLnk>/gi;
  let m: RegExpExecArray | null = linkRe.exec(scope);
  while (m !== null) {
    const inner = m[1] || '';
    const nameMatch = inner.match(/<name>([\s\S]*?)<\/name>/i);
    const ltypeMatch = inner.match(/<ltype>([\s\S]*?)<\/ltype>/i);
    const name = nameMatch ? decodeBasicXmlEntities(nameMatch[1]) : '';
    const ltype = ltypeMatch ? decodeBasicXmlEntities(ltypeMatch[1]) : 'folder';
    if (name && !seen.has(name.toLowerCase())) {
      seen.add(name.toLowerCase());
      results.push({
        name,
        ltype,
        folderHint: name,
      });
    }
    m = linkRe.exec(scope);
  }
  return results;
}

function readOutlookFavoritesFromDisk(): OutlookFavoriteLink[] {
  const appData = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
  const outlookDir = path.join(appData, 'Microsoft', 'Outlook');
  if (!fs.existsSync(outlookDir)) return [];

  try {
    const xmlFiles = fs
      .readdirSync(outlookDir, { withFileTypes: true })
      .filter((e) => e.isFile() && /\.xml$/i.test(e.name))
      .map((e) => {
        const fullPath = path.join(outlookDir, e.name);
        const st = fs.statSync(fullPath);
        return { fullPath, mtimeMs: st.mtimeMs };
      })
      .sort((a, b) => b.mtimeMs - a.mtimeMs);

    const merged: OutlookFavoriteLink[] = [];
    const seen = new Set<string>();
    for (const file of xmlFiles.slice(0, 6)) {
      const buf = fs.readFileSync(file.fullPath);
      const xmlText = decodeWindowsTextBuffer(buf);
      for (const fav of parseOutlookFavoritesXmlText(xmlText)) {
        const key = fav.name.toLowerCase();
        if (!seen.has(key)) {
          seen.add(key);
          merged.push(fav);
        }
      }
    }
    return merged;
  } catch {
    return [];
  }
}

/**
 * Safely reads non-secret Outlook registry settings across Office 16.0 / 15.0 / 14.0
 * and Windows Messaging Subsystem (all profiles), PST/OST paths, XML favorites, window prefs,
 * and signature files on the local Windows PC. NEVER reads or exports account passwords.
 */
export function inspectLocalOutlookSettings(): InspectedOutlookSettings {
  if (process.platform !== 'win32') {
    return {
      outlookVersion: 'Non-Windows',
      defaultProfileName: '',
      accounts: [],
      pstPaths: [],
      signatures: [],
      outlookFavorites: [],
      windowSettings: {},
    };
  }

  const psScript = `
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$roots = @(
  @{ Ver = '16.0'; Base = 'HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook'; MailSettings = 'HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings' },
  @{ Ver = '15.0'; Base = 'HKCU:\\Software\\Microsoft\\Office\\15.0\\Outlook'; MailSettings = 'HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\MailSettings' },
  @{ Ver = '14.0'; Base = 'HKCU:\\Software\\Microsoft\\Office\\14.0\\Outlook'; MailSettings = 'HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\MailSettings' },
  @{ Ver = 'Legacy'; Base = 'HKCU:\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows Messaging Subsystem'; MailSettings = '' }
)

$decode = {
  param($val)
  if ($val -is [byte[]]) { return [System.Text.Encoding]::Unicode.GetString($val).TrimEnd([char]0) }
  if ($null -eq $val) { return '' }
  return [string]$val
}

$accountsByEmail = @{}
$pstPaths = New-Object System.Collections.Generic.List[string]
$detectedVersion = ''
$defaultProfile = ''
$activeNewSig = ''
$activeReplySig = ''

foreach ($root in $roots) {
  $base = $root.Base
  if (Test-Path $base) {
    if (-not $detectedVersion) { $detectedVersion = $root.Ver }
    $defProf = [string](Get-ItemProperty -Path $base -Name 'DefaultProfile').DefaultProfile
    if ($defProf -and -not $defaultProfile) { $defaultProfile = $defProf }

    $profilesPath = "$base\\Profiles"
    if (Test-Path $profilesPath) {
      $allProfItems = Get-ChildItem $profilesPath
      $orderedNames = New-Object System.Collections.Generic.List[string]
      if ($defProf) { $orderedNames.Add($defProf) }
      foreach ($item in $allProfItems) {
        if (-not $orderedNames.Contains($item.PSChildName)) {
          $orderedNames.Add($item.PSChildName)
        }
      }
      if (-not $defaultProfile -and $orderedNames.Count -gt 0) {
        $defaultProfile = $orderedNames[0]
      }

      foreach ($profName in $orderedNames) {
        $profPath = "$profilesPath\\$profName"
        if (-not (Test-Path $profPath)) { continue }

        $accRoot = "$profPath\\9375CFF0413111d3B88A00104B2A6676"
        if (Test-Path $accRoot) {
          Get-ChildItem $accRoot | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath
            $email = (& $decode $p.Email).Trim().ToLower()
            if ($email -and $email.Contains('@') -and -not $accountsByEmail.ContainsKey($email)) {
              $newSig = (& $decode $p.'New Signature').Trim()
              $repSig = (& $decode $p.'Reply-Forward Signature').Trim()
              if ($newSig -and -not $activeNewSig) { $activeNewSig = $newSig }
              if ($repSig -and -not $activeReplySig) { $activeReplySig = $repSig }
              $popHost = (& $decode $p.'POP3 Server').Trim()
              $imapHost = (& $decode $p.'IMAP Server').Trim()
              $smtpHost = (& $decode $p.'SMTP Server').Trim()
              $accountsByEmail[$email] = [PSCustomObject]@{
                email = $email
                displayName = (& $decode $p.'Display Name').Trim()
                protocol = if ($imapHost) { 'imap' } else { 'pop3' }
                pop3Host = $popHost
                pop3Port = [int]($p.'POP3 Port')
                pop3Tls = ([int]$p.'POP3 Use SSL' -ne 0)
                imapHost = $imapHost
                imapPort = [int]($p.'IMAP Port')
                imapTls = ([int]$p.'IMAP Use SSL' -ne 0)
                smtpHost = $smtpHost
                smtpPort = [int]($p.'SMTP Port')
                smtpTls = ([int]$p.'SMTP Use SSL' -ne 0)
                leaveOnServer = ([int]$p.'Leave on Server' -ne 0)
              }
            }
          }
        }

        Get-ChildItem $profPath -Recurse | ForEach-Object {
          $props = Get-ItemProperty $_.PSPath
          foreach ($pname in @('001f6700','001e6700','001f6701')) {
            $v = $props.$pname
            if ($null -ne $v) {
              $s = (& $decode $v).Trim()
              if ($s -match '\\.(pst|ost)$' -and ($s -notmatch '\\\\~') -and (Test-Path $s)) {
                $pstPaths.Add($s)
              }
            }
          }
        }
      }
    }
  }

  if ($root.MailSettings -and (Test-Path $root.MailSettings)) {
    $ms = Get-ItemProperty $root.MailSettings
    if (-not $activeNewSig -and $ms.NewSignature) { $activeNewSig = [string]$ms.NewSignature }
    if (-not $activeReplySig -and $ms.ReplySignature) { $activeReplySig = [string]$ms.ReplySignature }
  }
}

$docsDir = [Environment]::GetFolderPath('MyDocuments')
$candidateDirs = @(
  (Join-Path $docsDir 'Outlook Files'),
  (Join-Path $docsDir 'Outlook 파일'),
  (Join-Path $env:USERPROFILE 'OneDrive\\Documents\\Outlook Files'),
  (Join-Path $env:USERPROFILE 'OneDrive\\Documents\\Outlook 파일'),
  (Join-Path $env:LOCALAPPDATA 'Microsoft\\Outlook'),
  'F:\\MAIL',
  'D:\\MAIL',
  'E:\\MAIL'
)

foreach ($dir in $candidateDirs) {
  if ($dir -and (Test-Path $dir)) {
    Get-ChildItem -Path $dir -File | Where-Object {
      ($_.Extension -match '^\\.(pst|ost)$') -and ($_.Name -notlike '~*') -and ($_.Name -notlike '*.tmp')
    } | ForEach-Object {
      $pstPaths.Add($_.FullName)
    }
  }
}

$winSettings = @{}
foreach ($ver in @('16.0','15.0','14.0')) {
  $prefPath = "HKCU:\\Software\\Microsoft\\Office\\$ver\\Outlook\\Preferences"
  if (Test-Path $prefPath) {
    $pr = Get-ItemProperty $prefPath
    if (-not $winSettings.ContainsKey('enableSingleLineRibbon') -and ($null -ne $pr.EnableSingleLineRibbon)) {
      $winSettings['enableSingleLineRibbon'] = ([int]$pr.EnableSingleLineRibbon -ne 0)
    }
    if (-not $winSettings.ContainsKey('pinFolderTree') -and ($null -ne $pr.PinFTree)) {
      $winSettings['pinFolderTree'] = ([int]$pr.PinFTree -ne 0)
    }
    if (-not $winSettings.ContainsKey('pinMail') -and ($null -ne $pr.PinMail)) {
      $winSettings['pinMail'] = ([int]$pr.PinMail -ne 0)
    }
    if (-not $winSettings.ContainsKey('moduleVisible15') -and ($null -ne $pr.ModuleVisible15)) {
      $winSettings['moduleVisible15'] = (& $decode $pr.ModuleVisible15).Trim()
    }
    if (-not $winSettings.ContainsKey('reopenWindowsOption') -and ($null -ne $pr.ReopenWindowsOption)) {
      $winSettings['reopenWindowsOption'] = [int]$pr.ReopenWindowsOption
    }
    if (-not $winSettings.ContainsKey('disableTnef') -and ($null -ne $pr.DisableTNEF)) {
      $winSettings['disableTnef'] = ([int]$pr.DisableTNEF -ne 0)
    }
    if (-not $winSettings.ContainsKey('defaultImportance') -and ($null -ne $pr.Importance)) {
      $imp = [int]$pr.Importance
      $winSettings['defaultImportance'] = if ($imp -eq 2) { 'high' } elseif ($imp -eq 0) { 'low' } else { 'normal' }
    }
    if (-not $winSettings.ContainsKey('requestReadReceipt') -and ($null -ne $pr.ReadReceipt)) {
      $winSettings['requestReadReceipt'] = ([int]$pr.ReadReceipt -ne 0)
    }
  }
  $mailOptPath = "HKCU:\\Software\\Microsoft\\Office\\$ver\\Outlook\\Options\\Mail"
  if (Test-Path $mailOptPath) {
    $mo = Get-ItemProperty $mailOptPath
    if (-not $winSettings.ContainsKey('minimalHeaderOn') -and ($null -ne $mo.MinimalHeaderOn)) {
      $winSettings['minimalHeaderOn'] = ([int]$mo.MinimalHeaderOn -ne 0)
    }
    if (-not $winSettings.ContainsKey('editorPreference') -and ($null -ne $mo.EditorPreference)) {
      $winSettings['editorPreference'] = [int]$mo.EditorPreference
    }
    if (-not $winSettings.ContainsKey('markAsReadWaitSec') -and ($null -ne $mo.'MarkAsReadWait')) {
      $winSettings['markAsReadWaitSec'] = [int]$mo.'MarkAsReadWait'
    }
    if (-not $winSettings.ContainsKey('readingPanePosition') -and ($null -ne $mo.ReadPanes)) {
      $rp = [int]$mo.ReadPanes
      $winSettings['readingPanePosition'] = if ($rp -eq 2) { 'bottom' } elseif ($rp -eq 0) { 'hidden' } else { 'right' }
    }
  }
  $calOptPath = "HKCU:\\Software\\Microsoft\\Office\\$ver\\Outlook\\Options\\Calendar"
  if (Test-Path $calOptPath) {
    $co = Get-ItemProperty $calOptPath
    if (-not $winSettings.ContainsKey('calWorkStart') -and ($null -ne $co.CalDefStart)) {
      $mins = [int]$co.CalDefStart
      if ($mins -ge 0 -and $mins -lt 1440) {
        $winSettings['calWorkStart'] = ('{0:D2}:{1:D2}' -f [int][math]::Floor($mins / 60), ($mins % 60))
      }
    }
    if (-not $winSettings.ContainsKey('calWorkEnd') -and ($null -ne $co.CalDefEnd)) {
      $mins = [int]$co.CalDefEnd
      if ($mins -ge 0 -and $mins -le 1440) {
        $h = [int][math]::Floor($mins / 60)
        $m = ($mins % 60)
        if ($h -eq 24 -and $m -eq 0) { $winSettings['calWorkEnd'] = '23:59' } else { $winSettings['calWorkEnd'] = ('{0:D2}:{1:D2}' -f $h, $m) }
      }
    }
    if (-not $winSettings.ContainsKey('calWorkDayMask') -and ($null -ne $co.WorkDay)) {
      $winSettings['calWorkDayMask'] = [int]$co.WorkDay
    }
  }
}

[PSCustomObject]@{
  outlookVersion = if ($detectedVersion) { $detectedVersion } else { '16.0' }
  defaultProfileName = $defaultProfile
  activeNewSignature = $activeNewSig
  activeReplySignature = $activeReplySig
  accounts = @($accountsByEmail.Values)
  pstPaths = @($pstPaths | Select-Object -Unique)
  windowSettings = $winSettings
} | ConvertTo-Json -Depth 5
`.trim();

  const rawJson = execFileSync(
    'powershell.exe',
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', psScript],
    { encoding: 'utf8', timeout: 25_000, windowsHide: true }
  );

  const parsed = JSON.parse(rawJson.trim() || '{}') as {
    outlookVersion?: string;
    defaultProfileName?: string;
    activeNewSignature?: string;
    activeReplySignature?: string;
    accounts?: Array<Partial<LocalAccountRecord>>;
    pstPaths?: string[] | string;
    windowSettings?: OutlookWindowSettings;
  };

  const pstPathsList = Array.isArray(parsed.pstPaths)
    ? parsed.pstPaths
    : typeof parsed.pstPaths === 'string' && parsed.pstPaths.trim()
      ? [parsed.pstPaths.trim()]
      : [];

  // Prioritize .pst over .ost and filter out temp lock files
  const uniquePstPaths = Array.from(
    new Set(
      pstPathsList
        .map((p) => path.resolve(p))
        .filter((p) => {
          const base = path.basename(p);
          return !base.startsWith('~') && !base.toLowerCase().endsWith('.tmp') && fs.existsSync(p);
        })
    )
  ).sort((a, b) => {
    const aIsPst = a.toLowerCase().endsWith('.pst') ? 0 : 1;
    const bIsPst = b.toLowerCase().endsWith('.pst') ? 0 : 1;
    return aIsPst - bIsPst;
  });

  const rawAccounts = Array.isArray(parsed.accounts) ? parsed.accounts : [];
  const accounts: LocalAccountRecord[] = rawAccounts
    .filter((a) => typeof a?.email === 'string' && a.email.includes('@'))
    .map((a) => ({
      email: String(a.email).trim().toLowerCase(),
      displayName: String(a.displayName || a.email).trim(),
      providerPreset: inferProviderPreset(`${a.pop3Host || ''} ${a.imapHost || ''} ${a.email || ''}`),
      protocol: a.protocol === 'imap' ? 'imap' : 'pop3',
      pop3Host: String(a.pop3Host || ''),
      pop3Port: Number(a.pop3Port) || 995,
      pop3Tls: a.pop3Tls !== false,
      imapHost: String(a.imapHost || ''),
      imapPort: Number(a.imapPort) || 993,
      imapTls: a.imapTls !== false,
      smtpHost: String(a.smtpHost || ''),
      smtpPort: Number(a.smtpPort) || 465,
      smtpTls: a.smtpTls !== false,
      leaveOnServer: a.leaveOnServer !== false,
      deliveryStorePath: uniquePstPaths[0] || '',
    }));

  const signatures = readSignaturesFromDisk(
    String(parsed.activeNewSignature || ''),
    String(parsed.activeReplySignature || ''),
    undefined,
    accounts[0]?.email
  );
  const outlookFavorites = readOutlookFavoritesFromDisk();
  const windowSettings: OutlookWindowSettings =
    parsed.windowSettings && typeof parsed.windowSettings === 'object' ? parsed.windowSettings : {};

  return {
    outlookVersion: String(parsed.outlookVersion || '16.0'),
    defaultProfileName: String(parsed.defaultProfileName || ''),
    accounts,
    pstPaths: uniquePstPaths,
    signatures,
    outlookFavorites,
    windowSettings,
  };
}

export interface LocalPstFileState {
  path: string;
  size: number;
  mtimeMs: number;
}

export function inspectPstFileStates(pstPaths: string[]): LocalPstFileState[] {
  const states: LocalPstFileState[] = [];
  for (const rawPath of pstPaths) {
    try {
      const resolved = path.resolve(rawPath);
      const st = fs.statSync(resolved);
      if (st.isFile()) {
        states.push({
          path: resolved,
          size: st.size,
          mtimeMs: Math.round(st.mtimeMs),
        });
      }
    } catch {
      // ignore missing/inaccessible store path
    }
  }
  return states;
}

export interface DiscoveredOutlookSettingsReport {
  ok: true;
  outlookVersion: string;
  defaultProfileName: string;
  signaturesCount: number;
  inlinedSignatureImages: number;
  defaultNewSignature: string | null;
  defaultReplySignature: string | null;
  accountsCount: number;
  favoritesCount: number;
  pstCount: number;
  windowSettings: OutlookWindowSettings;
  signatures: LocalSignatureRecord[];
  accounts: LocalAccountRecord[];
  outlookFavorites: OutlookFavoriteLink[];
  appliedAt: string;
}

/**
 * Discovers and applies ALL current Windows Outlook settings (accounts, HTML signatures
 * with embedded business card images, favorites, reading pane, calendar work hours, ribbon,
 * TNEF/importance/read-receipt options) without touching heavy PST mail indexing.
 */
export function discoverAndApplyAllOutlookSettings(
  db: LocalMailDb
): DiscoveredOutlookSettingsReport {
  const inspected = syncOutlookSettingsOnly(db);
  const inlinedSignatureImages = inspected.signatures.reduce((sum, s) => {
    const matches = s.htmlContent.match(/data:image\/[a-z0-9+.-]+;base64,/gi);
    return sum + (matches ? matches.length : 0);
  }, 0);
  const defNewSig =
    inspected.signatures.find((s) => s.isDefaultNew) || inspected.signatures[0] || null;
  const defRepSig =
    inspected.signatures.find((s) => s.isDefaultReply) || defNewSig;

  return {
    ok: true,
    outlookVersion: inspected.outlookVersion,
    defaultProfileName: inspected.defaultProfileName,
    signaturesCount: inspected.signatures.length,
    inlinedSignatureImages,
    defaultNewSignature: defNewSig ? defNewSig.name : null,
    defaultReplySignature: defRepSig ? defRepSig.name : null,
    accountsCount: inspected.accounts.length,
    favoritesCount: inspected.outlookFavorites.length,
    pstCount: inspected.pstPaths.length,
    windowSettings: inspected.windowSettings,
    signatures: inspected.signatures,
    accounts: inspected.accounts,
    outlookFavorites: inspected.outlookFavorites,
    appliedAt: new Date().toISOString(),
  };
}

export function syncOutlookSettingsOnly(db: LocalMailDb): InspectedOutlookSettings {
  const inspected = inspectLocalOutlookSettings();
  const existingPaths = (db.getState<string[]>('pstPaths') || [])
    .map((p) => path.resolve(p))
    .filter((p) => fs.existsSync(p));
  const mergedPstPaths = Array.from(new Set([...inspected.pstPaths, ...existingPaths]));

  if (inspected.accounts.length > 0) {
    db.upsertAccounts(inspected.accounts);
  }
  if (inspected.signatures.length > 0) {
    db.upsertSignatures(inspected.signatures);
  }
  db.setState('pstPaths', mergedPstPaths);
  if (inspected.outlookFavorites.length > 0) {
    db.setState('outlook_favorites', inspected.outlookFavorites);
  }
  db.setState('outlook_window_settings', inspected.windowSettings);

  return {
    ...inspected,
    pstPaths: mergedPstPaths,
  };
}

function extractPstAttachmentsMeta(
  msg: PSTMessage,
  msgKey: string,
  bodyHtmlHint?: string
): MailAttachmentMeta[] {
  const count = Number(msg.numberOfAttachments) || 0;
  if (count <= 0) return [];
  const list: MailAttachmentMeta[] = [];
  const limit = Math.min(count, 25);
  for (let i = 0; i < limit; i += 1) {
    try {
      const att = msg.getAttachment(i);
      if (!att) continue;
      const filename = (att.longFilename || att.filename || `attachment-${i + 1}`).trim();
      const mimeType = (att.mimeTag || 'application/octet-stream').trim();
      const sizeBytes = Number(att.filesize || att.size || 0);
      const cleanCid = (att.contentId || '').trim().replace(/^<|>$/g, '') || undefined;
      const meta: MailAttachmentMeta = {
        id: `${msgKey}-att-${i}`,
        filename,
        mimeType,
        sizeBytes: Number.isFinite(sizeBytes) && sizeBytes > 0 ? sizeBytes : 0,
        kind: inferAttachmentKind(filename, mimeType),
        contentId: cleanCid,
      };
      meta.isInline =
        Boolean(att.isAttachmentInvisibleInHtml) ||
        Boolean(cleanCid && bodyHtmlHint && bodyHtmlHint.toLowerCase().includes(`cid:${cleanCid.toLowerCase()}`)) ||
        classifyAttachmentInline(meta, bodyHtmlHint);
      list.push(meta);
    } catch {
      // skip unreadable attachment metadata cleanly
    }
  }
  return list;
}

function isTopLevelMailTreeRoot(folderName: string): boolean {
  const lower = folderName.trim().toLowerCase();
  return (
    lower.includes('최상위') ||
    lower.startsWith('top of') ||
    lower === 'ipm_subtree'
  );
}

function isTrashFolderPath(displayPath: string): boolean {
  const first = (displayPath.split('/')[0] || '').trim().toLowerCase();
  return first === '지운 편지함' || first === 'deleted items' || first === 'trash';
}

function isContactsFolderPath(displayPath: string): boolean {
  const leaf = (displayPath.split('/').pop() || '').trim().toLowerCase();
  return leaf === '연락처' || leaf === 'contacts' || leaf.includes('연락처');
}

function buildPstPathHash(pstPath: string): string {
  return crypto.createHash('sha1').update(path.resolve(pstPath).toLowerCase()).digest('hex').slice(0, 6);
}

function indexSinglePstFile(
  pstPath: string,
  db: LocalMailDb,
  onProgress?: (stats: { emails: number; contacts: number; folders: number }) => void,
  skipExistingIds?: Set<string>
): {
  storeName: string;
  emailsIndexed: number;
  contactsIndexed: number;
  folders: LocalFolderRecord[];
  sampleForProjects: MailItem[];
} {
  const pathHash = buildPstPathHash(pstPath);
  const pstFile = new PSTFile(pstPath);
  try {
    const storeName = (
      pstFile.getMessageStore()?.displayName || path.basename(pstPath, path.extname(pstPath))
    ).trim();
    const rootFolder = pstFile.getRootFolder();

    const folders: LocalFolderRecord[] = [];
    const contactsBatch: OutlookContactRecord[] = [];
    const sampleForProjects: MailItem[] = [];
    let batch: LocalEmailRecord[] = [];
    let emailsIndexed = 0;

    const flushBatch = () => {
      if (batch.length === 0) return;
      db.upsertEmailsBatch(batch);
      emailsIndexed += batch.length;
      batch = [];
      onProgress?.({
        emails: emailsIndexed,
        contacts: contactsBatch.length,
        folders: folders.length,
      });
    };

    const walkFolder = (folder: PSTFolder, trail: string[], depth: number) => {
      const folderTitle = (folder.displayName || '').trim();
      const lowerTitle = folderTitle.toLowerCase();

      if (depth === 1) {
        if (INTERNAL_ROOT_FOLDERS.has(lowerTitle)) return;
        const rootChildren = rootFolder.getSubFolders();
        const hasMailRoot = rootChildren.some((sf) => isTopLevelMailTreeRoot(sf.displayName || ''));
        if (hasMailRoot && !isTopLevelMailTreeRoot(folderTitle)) {
          return;
        }
      }

      const nextTrail =
        folderTitle && !isTopLevelMailTreeRoot(folderTitle) ? [...trail, folderTitle] : trail;
      const displayPath = nextTrail.join(' / ') || '기본 폴더';

      if (folder.contentCount > 0 && displayPath !== '기본 폴더') {
        const classifiedFolder = classifyOutlookFolder(displayPath);
        folders.push({
          storeName,
          folderPath: displayPath,
          classifiedFolder,
          itemCount: folder.contentCount,
          isSystem: isSystemOutlookFolder(displayPath),
        });

        const contactFolder = isContactsFolderPath(displayPath);
        const trashed = isTrashFolderPath(displayPath);

        let child = folder.getNextChild();
        while (child != null) {
          try {
            if (child instanceof PSTContact || (contactFolder && child instanceof PSTMessage)) {
              const c = child as PSTContact;
              const emailAddr = (
                c.email1EmailAddress ||
                c.email2EmailAddress ||
                c.email3EmailAddress ||
                ''
              )
                .trim()
                .toLowerCase();
              if (emailAddr && emailAddr.includes('@')) {
                const fullName = (
                  c.displayName ||
                  c.email1DisplayName ||
                  `${c.surname || ''}${c.givenName || ''}`.trim() ||
                  emailAddr
                ).trim();
                const cNodeId = c.descriptorNodeId ? c.descriptorNodeId.toNumber() : contactsBatch.length;
                contactsBatch.push({
                  id: `pst-contact-${pathHash}-${cNodeId}`,
                  fullName,
                  email: emailAddr,
                  company: (c.companyName || '').trim() || undefined,
                  department: (c.departmentName || '').trim() || undefined,
                  jobTitle: (c.title || '').trim() || undefined,
                  phone:
                    (c.mobileTelephoneNumber || c.businessTelephoneNumber || c.primaryTelephoneNumber || '').trim() ||
                    undefined,
                  source: 'outlook-pc',
                });
              }
            } else if (child instanceof PSTMessage) {
              const msg = child as PSTMessage;
              const descId = msg.descriptorNodeId ? msg.descriptorNodeId.toNumber() : emailsIndexed + batch.length;
              const msgId = `pst-${pathHash}-${descId}`;
              if (skipExistingIds && skipExistingIds.has(msgId)) {
                child = folder.getNextChild();
                continue;
              }
              const rawDate =
                msg.messageDeliveryTime || msg.clientSubmitTime || msg.creationTime || new Date();
              const dt =
                rawDate instanceof Date && !Number.isNaN(rawDate.getTime()) ? rawDate : new Date();
              const dateStr = dt.toISOString().slice(0, 10);
              const timeLabel = dt.toTimeString().slice(0, 5);

              const plainBody = (msg.body || '').trim();
              const htmlBody = (msg.bodyHTML || '').trim();
              const finalBody =
                plainBody || (htmlBody ? stripHtmlToText(htmlBody) : '') || '(본문 없음)';

              const senderEmail = (
                msg.senderEmailAddress ||
                msg.sentRepresentingEmailAddress ||
                msg.senderName ||
                'unknown@outlook.pst'
              ).trim();
              const senderName = (
                msg.senderName ||
                msg.sentRepresentingName ||
                senderEmail ||
                '알 수 없음'
              ).trim();
              const recipients = extractRecipients(msg.displayTo);
              const subject = (msg.subject || '').trim() || '(제목 없음)';
              const attachments = extractPstAttachmentsMeta(msg, msgId, htmlBody);

              const record: LocalEmailRecord = {
                id: msgId,
                storeName,
                pstPath,
                descriptorNodeId: descId,
                folder: classifiedFolder,
                subject,
                fromName: senderName,
                fromEmail: senderEmail,
                to: recipients.length > 0 ? recipients : ['me@outlook.local'],
                date: dateStr,
                timestamp: dt.toISOString(),
                timeLabel,
                snippet: finalBody.replace(/\s+/g, ' ').slice(0, 140),
                body: finalBody.slice(0, MAX_BODY_CHARS),
                hasAttachment: Boolean(msg.hasAttachments) || attachments.length > 0,
                attachments: attachments.length > 0 ? attachments : undefined,
                conversationKey: normalizeConversationSubject(subject),
                importance: msg.importance >= 2 ? 'high' : 'normal',
                category: displayPath,
                folderName: displayPath,
                isRead: Boolean(msg.isRead),
                isFlagged: false,
                isDeleted: trashed,
                source: 'pst-import',
              };

              batch.push(record);
              if (sampleForProjects.length < 1200 && !trashed) {
                sampleForProjects.push({
                  ...record,
                  body: record.body.slice(0, 1200),
                });
              }

              if (batch.length >= BATCH_SIZE) {
                flushBatch();
              }
            }
          } catch {
            // Skip malformed individual PST item without aborting migration
          }
          child = folder.getNextChild();
        }
      }

      if (folder.hasSubfolders && depth < 12) {
        const subFolders = folder.getSubFolders();
        for (const sub of subFolders) {
          walkFolder(sub, nextTrail, depth + 1);
        }
      }
    };

    walkFolder(rootFolder, [], 0);
    flushBatch();

    if (folders.length > 0) {
      db.upsertFolders(folders);
    }
    if (contactsBatch.length > 0) {
      db.upsertContacts(contactsBatch);
    }

    return {
      storeName,
      emailsIndexed,
      contactsIndexed: contactsBatch.length,
      folders,
      sampleForProjects,
    };
  } finally {
    pstFile.close();
  }
}

const RECOMPUTED_PROJECT_COLORS = [
  '#0F6CBD',
  '#0B7A75',
  '#6B4C9A',
  '#B45309',
  '#BE123C',
  '#1D4ED8',
  '#047857',
];

function deriveLocalFolderProjectCode(folderLeaf: string, index: number): string {
  const explicit = folderLeaf.match(/\b([A-Z]{2,6}[-_]\d{2,6}(?:[-_][A-Z0-9]{1,6})?)\b/i);
  if (explicit) return explicit[1].toUpperCase();
  const hullMatch = folderLeaf.match(/\b([A-Z]{1,4}\d{3,5})\b/i);
  if (hullMatch) return hullMatch[1].toUpperCase();
  const asciiTag = folderLeaf
    .replace(/[^a-zA-Z0-9]/g, '')
    .toUpperCase()
    .slice(0, 4);
  if (asciiTag.length >= 2) {
    return `${asciiTag}-${String(index + 1).padStart(2, '0')}`;
  }
  return `PRJ-${String(index + 1).padStart(2, '0')}`;
}

export function recomputeProjectsFromLocalDb(db: LocalMailDb): ProjectDefinition[] {
  const nonSystemFolders = db.getFolders({ nonSystemOnly: true, minItemCount: 1, limit: 24 });
  const existingProjects = db.getState<ProjectDefinition[]>('projects') || [];
  const discoveredMap = new Map<string, InheritedOutlookProfile['discoveredProjects'][number]>();

  let seq = 0;
  for (const folder of nonSystemFolders) {
    const parts = folder.folderPath
      .split('/')
      .map((s) => s.trim())
      .filter(Boolean);
    const leaf = parts[parts.length - 1] || folder.folderPath.trim();
    if (!leaf || isSystemOutlookFolder(leaf)) continue;

    let code = deriveLocalFolderProjectCode(leaf, seq);
    seq += 1;

    const cleanFolder = folder.folderPath.replace(/^[^/]+\s*\/\s*/, '').trim();
    const keywords = Array.from(new Set([leaf, code, cleanFolder].filter(Boolean))).slice(0, 6);

    const existingHit = discoveredMap.get(code.toLowerCase());
    if (existingHit) {
      existingHit.mailCount += folder.itemCount;
      existingHit.keywords = Array.from(new Set([...existingHit.keywords, ...keywords])).slice(0, 6);
    } else {
      const idx = discoveredMap.size;
      discoveredMap.set(code.toLowerCase(), {
        code,
        name: leaf,
        keywords,
        color: RECOMPUTED_PROJECT_COLORS[idx % RECOMPUTED_PROJECT_COLORS.length],
        isFavorite: idx < 5,
        source: 'folder',
        mailCount: folder.itemCount,
      });
    }
    if (discoveredMap.size >= 12) break;
  }

  const sampleEmails = db.queryEmails({ limit: 300 }).items;
  const inherited = buildInheritedOutlookProfile({
    storeName: 'Outlook',
    folders: nonSystemFolders.map((f) => f.folderPath),
    emails: sampleEmails,
  });
  for (const subjProj of inherited.discoveredProjects) {
    if (discoveredMap.size >= 12) break;
    const key = subjProj.code.trim().toLowerCase();
    if (!discoveredMap.has(key)) {
      discoveredMap.set(key, subjProj);
    }
  }

  const discoveredList = Array.from(discoveredMap.values())
    .sort((a, b) => b.mailCount - a.mailCount)
    .slice(0, 12);

  const { merged } = mergeInheritedProjects(existingProjects, discoveredList);
  db.setState('projects', merged);
  return merged;
}

export function migrateLocalOutlookToDb(
  db: LocalMailDb,
  options?: { extraPstPaths?: string[]; onProgress?: (summary: MigrationProgressSummary) => void }
): MigrationProgressSummary {
  const startedAt = new Date().toISOString();
  const summary: MigrationProgressSummary = {
    status: 'running',
    startedAt,
    completedAt: null,
    storesProcessed: 0,
    totalEmailsIndexed: 0,
    totalContactsIndexed: 0,
    totalFoldersIndexed: 0,
    totalSignaturesIndexed: 0,
    discoveredProjectsCount: 0,
    error: null,
  };

  try {
    const inspected = inspectLocalOutlookSettings();
    const extraPaths = (options?.extraPstPaths || [])
      .map((p) => path.resolve(p))
      .filter((p) => fs.existsSync(p));
    const allPstPaths = Array.from(new Set([...inspected.pstPaths, ...extraPaths]));

    if (inspected.accounts.length > 0) {
      db.upsertAccounts(inspected.accounts);
    }
    if (inspected.signatures.length > 0) {
      db.upsertSignatures(inspected.signatures);
    }
    db.setState('pstPaths', allPstPaths);
    if (inspected.outlookFavorites.length > 0) {
      db.setState('outlook_favorites', inspected.outlookFavorites);
    }
    db.setState('outlook_window_settings', inspected.windowSettings);
    summary.totalSignaturesIndexed = inspected.signatures.length;

    const watermarks: Record<string, { size: number; mtimeMs: number; lastSyncedAt: string }> = {};
    for (const pstPath of allPstPaths) {
      try {
        const storeResult = indexSinglePstFile(pstPath, db, (st) => {
          options?.onProgress?.({
            ...summary,
            totalEmailsIndexed: summary.totalEmailsIndexed + st.emails,
            totalContactsIndexed: summary.totalContactsIndexed + st.contacts,
            totalFoldersIndexed: summary.totalFoldersIndexed + st.folders,
          });
        });
        summary.storesProcessed += 1;
        summary.totalEmailsIndexed += storeResult.emailsIndexed;
        summary.totalContactsIndexed += storeResult.contactsIndexed;
        summary.totalFoldersIndexed += storeResult.folders.length;
        try {
          const stFile = fs.statSync(pstPath);
          watermarks[path.resolve(pstPath)] = {
            size: stFile.size,
            mtimeMs: Math.round(stFile.mtimeMs),
            lastSyncedAt: new Date().toISOString(),
          };
        } catch {
          // ignore stat errors
        }
      } catch (storeErr) {
        // Never abort valid .pst indexing if a locked .ost (Outlook open) or secondary store cannot open
        if (!pstPath.toLowerCase().endsWith('.ost') && allPstPaths.length === 1) {
          throw storeErr;
        }
      }
    }

    const merged = recomputeProjectsFromLocalDb(db);
    summary.discoveredProjectsCount = merged.length;

    const completedAt = new Date().toISOString();
    summary.status = 'completed';
    summary.completedAt = completedAt;
    db.setState('pst_file_states', inspectPstFileStates(allPstPaths));
    db.setState('pst_sync_watermarks', watermarks);
    db.setState('migration_meta', {
      migratedAt: completedAt,
      defaultProfileName: inspected.defaultProfileName,
      outlookVersion: inspected.outlookVersion,
      storesProcessed: summary.storesProcessed,
      totalEmailsIndexed: summary.totalEmailsIndexed,
      totalContactsIndexed: summary.totalContactsIndexed,
      totalFoldersIndexed: summary.totalFoldersIndexed,
      totalSignaturesIndexed: summary.totalSignaturesIndexed,
    });

    options?.onProgress?.(summary);
    return summary;
  } catch (err) {
    summary.status = 'failed';
    summary.completedAt = new Date().toISOString();
    summary.error = err instanceof Error ? err.message : String(err);
    options?.onProgress?.(summary);
    throw err;
  }
}

export function readPstAttachmentBytes(
  pstPath: string,
  descriptorNodeId: number,
  attachmentIndex: number
): { filename: string; mimeType: string; buffer: Buffer } {
  if (!fs.existsSync(pstPath)) {
    throw new Error('원본 PST 파일을 찾을 수 없습니다.');
  }
  const pstFile = new PSTFile(pstPath);
  try {
    const obj = PSTUtil.detectAndLoadPSTObject(pstFile, Long.fromNumber(descriptorNodeId));
    if (!(obj instanceof PSTMessage)) {
      throw new Error('해당 메일 항목을 PST에서 찾을 수 없습니다.');
    }
    const count = Number(obj.numberOfAttachments) || 0;
    if (attachmentIndex < 0 || attachmentIndex >= count) {
      throw new Error('요청한 첨부파일 인덱스가 범위를 벗어났습니다.');
    }
    const att = obj.getAttachment(attachmentIndex);
    if (!att) {
      throw new Error('첨부파일 데이터를 읽을 수 없습니다.');
    }
    const filename = (att.longFilename || att.filename || `attachment-${attachmentIndex + 1}`).trim();
    const mimeType = (att.mimeTag || 'application/octet-stream').trim();
    const stream = att.fileInputStream;
    if (!stream) {
      throw new Error('포함(embedded) 또는 참조형 첨부파일은 스트림 추출을 지원하지 않습니다.');
    }

    const expectedLen = Math.max(0, stream.length.toNumber());
    const chunks: Buffer[] = [];
    let totalRead = 0;
    while (totalRead < expectedLen) {
      const toRead = Math.min(8176, expectedLen - totalRead);
      const buf = Buffer.alloc(toRead);
      const n = stream.read(buf);
      if (n <= 0) break;
      chunks.push(n === toRead ? buf : buf.subarray(0, n));
      totalRead += n;
    }

    return {
      filename,
      mimeType,
      buffer: Buffer.concat(chunks),
    };
  } finally {
    pstFile.close();
  }
}

export function sanitizeEmailHtml(html: string): string {
  if (!html) return '';
  let out = String(html);
  // Strip executable/active tags completely
  out = out
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, '')
    .replace(/<object[\s\S]*?<\/object>/gi, '')
    .replace(/<embed\b[^>]*>/gi, '')
    .replace(/<applet[\s\S]*?<\/applet>/gi, '')
    .replace(/<meta\b[^>]*http-equiv\s*=\s*["']?refresh["']?[^>]*>/gi, '')
    .replace(/<\/?form\b[^>]*>/gi, '');

  // Strip inline DOM event attributes (onclick=, onload=, onerror=, ...)
  out = out.replace(/\s+on[a-z]{2,24}\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '');

  // Neutralize javascript: / vbscript: URIs in href, src, action
  out = out.replace(
    /\b(href|src|action)\s*=\s*(["'])\s*(?:javascript|vbscript|data:text\/html):[\s\S]*?\2/gi,
    '$1=$2#$2'
  );

  // Ensure external links open safely in new tab without opener leakage
  out = out.replace(/<a\b((?:(?!target\s*=)[^>])*)>/gi, '<a$1 target="_blank" rel="noopener noreferrer">');

  return out;
}

function escapePlainTextForHtml(text: string): string {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export function readPstMessageRichView(
  pstPath: string,
  descriptorNodeId: number,
  emailId: string
): { bodyHtml: string; attachments: MailAttachmentMeta[] } {
  if (!fs.existsSync(pstPath)) {
    throw new Error('원본 PST 파일을 찾을 수 없습니다.');
  }
  const pstFile = new PSTFile(pstPath);
  try {
    const obj = PSTUtil.detectAndLoadPSTObject(pstFile, Long.fromNumber(descriptorNodeId));
    if (!(obj instanceof PSTMessage)) {
      throw new Error('해당 메일 항목을 PST에서 찾을 수 없습니다.');
    }
    const rawHtml = (obj.bodyHTML || '').trim();
    const rawText = (obj.body || '').trim();
    const attachments = extractPstAttachmentsMeta(obj, emailId, rawHtml);

    let html = rawHtml;
    if (!html) {
      html = `<div style="white-space:pre-wrap;font-family:'Malgun Gothic','Segoe UI',sans-serif;line-height:1.65;color:#242424;">${escapePlainTextForHtml(
        rawText || '(본문 없음)'
      )}</div>`;
    } else {
      // Rewrite cid:<id> references to local inline endpoint and flag matching attachments as inline
      const cidToIndex = new Map<string, number>();
      attachments.forEach((att, idx) => {
        if (att.contentId) {
          cidToIndex.set(att.contentId.replace(/^<|>$/g, '').toLowerCase(), idx);
        }
      });

      html = html.replace(
        /(src\s*=\s*["']?)cid:([^"'\s>]+)/gi,
        (full, prefix: string, rawCid: string) => {
          const clean = String(rawCid || '').replace(/^<|>$/g, '').trim().toLowerCase();
          let matchedIdx = cidToIndex.get(clean);
          if (matchedIdx === undefined && clean.includes('@')) {
            const left = clean.split('@')[0];
            for (const [cidKey, idx] of cidToIndex.entries()) {
              if (cidKey === left || cidKey.startsWith(`${left}@`)) {
                matchedIdx = idx;
                break;
              }
            }
          }
          if (matchedIdx !== undefined) {
            attachments[matchedIdx] = { ...attachments[matchedIdx], isInline: true };
            return `${prefix}/api/local/pst/attachment?emailId=${encodeURIComponent(
              emailId
            )}&idx=${matchedIdx}&inline=1`;
          }
          return full;
        }
      );
      html = sanitizeEmailHtml(html);
    }

    return { bodyHtml: html, attachments };
  } finally {
    pstFile.close();
  }
}

export interface LiveOutlookSyncResult {
  outlookRunning: boolean;
  addedEmails: number;
  syncedFolders: number;
  categoriesCount: number;
  rulesCount: number;
  error?: string;
}

export function syncLiveOutlookViaCom(
  db: LocalMailDb,
  options?: { maxItemsPerFolder?: number }
): LiveOutlookSyncResult {
  if (process.platform !== 'win32') {
    return {
      outlookRunning: false,
      addedEmails: 0,
      syncedFolders: 0,
      categoriesCount: 0,
      rulesCount: 0,
    };
  }

  const maxItems = Math.max(10, Math.min(120, Number(options?.maxItemsPerFolder) || 60));
  const psScript = `
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
  $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
  [PSCustomObject]@{ outlookRunning = $false } | ConvertTo-Json -Depth 3
  exit 0
}
if ($null -eq $app) {
  [PSCustomObject]@{ outlookRunning = $false } | ConvertTo-Json -Depth 3
  exit 0
}

$ns = $app.GetNamespace('MAPI')
$categories = New-Object System.Collections.Generic.List[object]
try {
  foreach ($cat in $ns.Categories) {
    $categories.Add([PSCustomObject]@{ name = [string]$cat.Name; color = [string]$cat.Color })
  }
} catch {}

$rules = New-Object System.Collections.Generic.List[object]
$folders = New-Object System.Collections.Generic.List[object]
$messages = New-Object System.Collections.Generic.List[object]
$cidPropTag = 'http://schemas.microsoft.com/mapi/proptag/0x3712001F'

function Walk-Folder {
  param($folder, [string]$storeName, [string[]]$trail, [int]$depth, [ref]$fetchedCount)
  if ($null -eq $folder -or $depth -gt 3) { return }
  $name = [string]$folder.Name
  $nextTrail = if ($depth -eq 0) { @() } else { $trail + @($name) }
  $displayPath = ($nextTrail -join ' / ').Trim()

  if ($depth -gt 0 -and $displayPath) {
    $cnt = 0
    $unread = 0
    try { $cnt = [int]$folder.Items.Count; $unread = [int]$folder.UnReadItemCount } catch {}
    $folders.Add([PSCustomObject]@{
      storeName = $storeName
      folderPath = $displayPath
      itemCount = $cnt
      unreadCount = $unread
    })

    $lowerPath = $displayPath.ToLower()
    $shouldFetch = ($fetchedCount.Value -lt 240) -and ($cnt -gt 0) -and (
      $lowerPath -eq '받은 편지함' -or $lowerPath -eq 'inbox' -or
      $lowerPath -eq '보낸 편지함' -or $lowerPath -eq 'sent items' -or
      $lowerPath -eq '임시 보관함' -or $lowerPath -eq 'drafts' -or
      $unread -gt 0 -or $depth -le 2
    ) -and ($lowerPath -notmatch '지운 편지함|deleted|sync|rss|검색 폴더|search')

    if ($shouldFetch) {
      try {
        $items = $folder.Items
        $items.Sort('[ReceivedTime]', $true)
        $taken = 0
        foreach ($it in $items) {
          if ($taken -ge ${maxItems} -or $fetchedCount.Value -ge 240) { break }
          if ($it.MessageClass -and ([string]$it.MessageClass -notlike 'IPM.Note*')) { continue }
          $taken++
          $fetchedCount.Value++

          $dt = Get-Date
          try { if ($it.ReceivedTime) { $dt = [datetime]$it.ReceivedTime } } catch {}
          $atts = New-Object System.Collections.Generic.List[object]
          try {
            $acnt = [int]$it.Attachments.Count
            for ($ai = 1; $ai -le [Math]::Min($acnt, 15); $ai++) {
              $attObj = $it.Attachments.Item($ai)
              $cid = ''
              try { $cid = [string]$attObj.PropertyAccessor.GetProperty($cidPropTag) } catch {}
              $atts.Add([PSCustomObject]@{
                idx = ($ai - 1)
                filename = [string]$attObj.FileName
                sizeBytes = [int]$attObj.Size
                contentId = $cid
              })
            }
          } catch {}

          $bodyTxt = ''
          $htmlTxt = ''
          try { $bodyTxt = [string]$it.Body; if ($bodyTxt.Length -gt 8000) { $bodyTxt = $bodyTxt.Substring(0, 8000) } } catch {}
          try { $htmlTxt = [string]$it.HTMLBody; if ($htmlTxt.Length -gt 80000) { $htmlTxt = $htmlTxt.Substring(0, 80000) } } catch {}

          $messages.Add([PSCustomObject]@{
            entryId = [string]$it.EntryID
            storeName = $storeName
            folderPath = $displayPath
            subject = [string]$it.Subject
            fromName = [string]$it.SenderName
            fromEmail = [string]$it.SenderEmailAddress
            to = [string]$it.To
            receivedIso = $dt.ToUniversalTime().ToString('o')
            dateStr = $dt.ToString('yyyy-MM-dd')
            timeLabel = $dt.ToString('HH:mm')
            unread = [bool]$it.UnRead
            flagged = ([int]$it.FlagStatus -eq 2)
            importance = ([int]$it.Importance -eq 2)
            categories = [string]$it.Categories
            body = $bodyTxt
            htmlBody = $htmlTxt
            attachments = @($atts)
          })
        }
      } catch {}
    }
  }

  try {
    foreach ($sub in $folder.Folders) {
      Walk-Folder -folder $sub -storeName $storeName -trail $nextTrail -depth ($depth + 1) -fetchedCount $fetchedCount
    }
  } catch {}
}

$fetchedTotal = 0
foreach ($store in $ns.Stores) {
  $sName = [string]$store.DisplayName
  try {
    $storeRules = $store.GetRules()
    foreach ($r in $storeRules) {
      $rules.Add([PSCustomObject]@{ name = [string]$r.Name; enabled = [bool]$r.Enabled })
    }
  } catch {}
  try {
    $root = $store.GetRootFolder()
    Walk-Folder -folder $root -storeName $sName -trail @() -depth 0 -fetchedCount ([ref]$fetchedTotal)
  } catch {}
}

[PSCustomObject]@{
  outlookRunning = $true
  categories = @($categories)
  rules = @($rules)
  folders = @($folders)
  messages = @($messages)
} | ConvertTo-Json -Depth 6
`.trim();

  try {
    const rawJson = execFileSync(
      'powershell.exe',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', psScript],
      { encoding: 'utf8', timeout: 18_000, windowsHide: true, maxBuffer: 24 * 1024 * 1024 }
    );
    const parsed = JSON.parse(rawJson.trim() || '{}') as {
      outlookRunning?: boolean;
      categories?: OutlookCategoryEntry[];
      rules?: OutlookRuleEntry[];
      folders?: Array<{ storeName: string; folderPath: string; itemCount: number; unreadCount?: number }>;
      messages?: Array<{
        entryId?: string;
        storeName?: string;
        folderPath?: string;
        subject?: string;
        fromName?: string;
        fromEmail?: string;
        to?: string;
        receivedIso?: string;
        dateStr?: string;
        timeLabel?: string;
        unread?: boolean;
        flagged?: boolean;
        importance?: boolean;
        categories?: string;
        body?: string;
        htmlBody?: string;
        attachments?: Array<{ idx?: number; filename?: string; sizeBytes?: number; contentId?: string }>;
      }>;
    };

    if (!parsed.outlookRunning) {
      return {
        outlookRunning: false,
        addedEmails: 0,
        syncedFolders: 0,
        categoriesCount: 0,
        rulesCount: 0,
      };
    }

    const folderRecords: LocalFolderRecord[] = (parsed.folders || [])
      .filter((f) => f.folderPath)
      .map((f) => ({
        storeName: String(f.storeName || 'Outlook'),
        folderPath: String(f.folderPath),
        classifiedFolder: classifyOutlookFolder(String(f.folderPath)),
        itemCount: Number(f.itemCount || 0),
        unreadCount: Number(f.unreadCount || 0),
        isSystem: isSystemOutlookFolder(String(f.folderPath)),
      }));
    if (folderRecords.length > 0) {
      db.upsertFolders(folderRecords);
    }

    const categories = Array.isArray(parsed.categories) ? parsed.categories : [];
    const rules = Array.isArray(parsed.rules) ? parsed.rules : [];
    if (categories.length > 0) db.setState('outlook_categories', categories);
    if (rules.length > 0) db.setState('outlook_rules', rules);

    const records: LocalEmailRecord[] = [];
    const richQueue: Array<{ id: string; html: string; atts: MailAttachmentMeta[] }> = [];

    for (const m of parsed.messages || []) {
      const store = String(m.storeName || 'Outlook').trim();
      const entryId = String(m.entryId || '').trim();
      if (!entryId) continue;
      const hash = crypto.createHash('sha1').update(`${store}::${entryId}`).digest('hex').slice(0, 16);
      const id = `com-${hash}`;
      const folderPath = String(m.folderPath || '받은 편지함').trim();
      const subject = String(m.subject || '(제목 없음)').trim() || '(제목 없음)';
      const plainBody = String(m.body || '').trim();
      const rawHtml = String(m.htmlBody || '').trim();
      const cleanBody = plainBody || (rawHtml ? stripHtmlToText(rawHtml) : '') || '(본문 없음)';

      const atts: MailAttachmentMeta[] = (m.attachments || []).map((a, idx) => {
        const filename = String(a.filename || `attachment-${idx + 1}`).trim();
        const cid = String(a.contentId || '').trim().replace(/^<|>$/g, '') || undefined;
        const meta: MailAttachmentMeta = {
          id: `${id}-att-${idx}`,
          filename,
          mimeType: 'application/octet-stream',
          sizeBytes: Number(a.sizeBytes || 0),
          kind: inferAttachmentKind(filename, ''),
          contentId: cid,
        };
        meta.isInline = classifyAttachmentInline(meta, rawHtml);
        return meta;
      });

      if (rawHtml) {
        richQueue.push({ id, html: sanitizeEmailHtml(rawHtml), atts });
      }

      records.push({
        id,
        storeName: store,
        folder: classifyOutlookFolder(folderPath),
        folderName: folderPath,
        subject,
        fromName: String(m.fromName || m.fromEmail || '알 수 없음').trim(),
        fromEmail: String(m.fromEmail || 'unknown@outlook.local').trim(),
        to: extractRecipients(m.to),
        date: String(m.dateStr || new Date().toISOString().slice(0, 10)),
        timestamp: String(m.receivedIso || new Date().toISOString()),
        timeLabel: String(m.timeLabel || '00:00'),
        snippet: cleanBody.replace(/\s+/g, ' ').slice(0, 140),
        body: cleanBody.slice(0, MAX_BODY_CHARS),
        hasAttachment: atts.some((a) => !a.isInline),
        attachments: atts.length > 0 ? atts : undefined,
        conversationKey: normalizeConversationSubject(subject),
        importance: m.importance ? 'high' : 'normal',
        category: m.categories ? String(m.categories) : folderPath,
        isRead: !m.unread,
        isFlagged: Boolean(m.flagged),
        isDeleted: false,
        source: 'outlook-live',
      });
    }

    const inserted = db.upsertEmailsBatch(records);
    for (const r of richQueue) {
      db.setEmailRichCache(r.id, r.html, r.atts);
    }

    return {
      outlookRunning: true,
      addedEmails: inserted,
      syncedFolders: folderRecords.length,
      categoriesCount: categories.length,
      rulesCount: rules.length,
    };
  } catch (err) {
    return {
      outlookRunning: false,
      addedEmails: 0,
      syncedFolders: 0,
      categoriesCount: 0,
      rulesCount: 0,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

export interface IncrementalSyncSummary {
  ok: true;
  outlookRunning: boolean;
  addedPstEmails: number;
  addedComEmails: number;
  skippedUnchangedPsts: number;
  syncedFolders: number;
  favoritesCount: number;
  categoriesCount: number;
  rulesCount: number;
  totalEmails: number;
  windowSettings: OutlookWindowSettings;
  outlookFavorites: OutlookFavoriteLink[];
  newEmails: MailItem[];
  syncedAt: string;
}

export function syncIncrementalLocalOutlook(db: LocalMailDb): IncrementalSyncSummary {
  const inspected = syncOutlookSettingsOnly(db);
  const watermarks =
    db.getState<Record<string, { size: number; mtimeMs: number; lastSyncedAt: string }>>(
      'pst_sync_watermarks'
    ) || {};

  const existingIds = db.getExistingEmailIds();
  const knownIdsBeforeSync = new Set(existingIds);
  let addedPstEmails = 0;
  let skippedUnchangedPsts = 0;
  let pstFoldersCount = 0;

  for (const pstPath of inspected.pstPaths) {
    const resolved = path.resolve(pstPath);
    try {
      const st = fs.statSync(resolved);
      const wm = watermarks[resolved];
      const mtimeMs = Math.round(st.mtimeMs);
      if (wm && wm.size === st.size && wm.mtimeMs === mtimeMs && existingIds.size > 0) {
        skippedUnchangedPsts += 1;
        continue;
      }
      const res = indexSinglePstFile(resolved, db, undefined, existingIds);
      addedPstEmails += res.emailsIndexed;
      pstFoldersCount += res.folders.length;
      watermarks[resolved] = {
        size: st.size,
        mtimeMs,
        lastSyncedAt: new Date().toISOString(),
      };
    } catch {
      // Ignore locked .ost or transient read lock cleanly
    }
  }

  db.setState('pst_sync_watermarks', watermarks);
  db.setState('pst_file_states', inspectPstFileStates(inspected.pstPaths));

  const comRes = syncLiveOutlookViaCom(db, { maxItemsPerFolder: 60 });
  if (addedPstEmails > 0 || comRes.addedEmails > 0) {
    recomputeProjectsFromLocalDb(db);
  }

  const syncedAt = new Date().toISOString();
  const prevMeta = db.getState<Record<string, unknown>>('migration_meta') || {};
  const totalEmails = db.queryEmails({ limit: 1 }).total;
  const newEmails =
    addedPstEmails > 0 || comRes.addedEmails > 0
      ? db
          .queryEmails({ folder: 'inbox', limit: 30 })
          .items.filter((m) => !knownIdsBeforeSync.has(m.id))
          .slice(0, 10)
      : [];
  db.setState('migration_meta', {
    ...prevMeta,
    migratedAt: syncedAt,
    defaultProfileName: inspected.defaultProfileName || prevMeta.defaultProfileName,
    outlookVersion: inspected.outlookVersion || prevMeta.outlookVersion,
    totalEmailsIndexed: totalEmails,
  });

  return {
    ok: true,
    outlookRunning: comRes.outlookRunning,
    addedPstEmails,
    addedComEmails: comRes.addedEmails,
    skippedUnchangedPsts,
    syncedFolders: Math.max(pstFoldersCount, comRes.syncedFolders),
    favoritesCount: inspected.outlookFavorites.length,
    categoriesCount: comRes.categoriesCount,
    rulesCount: comRes.rulesCount,
    totalEmails,
    windowSettings: inspected.windowSettings,
    outlookFavorites: inspected.outlookFavorites,
    newEmails,
    syncedAt,
  };
}

interface RawAsyncInboxMessage {
  entryId?: string;
  storeName?: string;
  folderPath?: string;
  subject?: string;
  fromName?: string;
  fromEmail?: string;
  to?: string;
  receivedIso?: string;
  dateStr?: string;
  timeLabel?: string;
  unread?: boolean;
  flagged?: boolean;
  importance?: boolean;
  categories?: string;
  body?: string;
  hasAttachment?: boolean;
}

/**
 * Non-blocking async Outlook Inbox COM poller for the background monitor.
 * Never triggers synchronous PST re-parsing or `execFileSync`.
 */
export async function pollLiveOutlookInboxAsync(
  db: LocalMailDb,
  options?: { maxItems?: number; timeoutMs?: number; triggerSendReceive?: boolean }
): Promise<{ outlookRunning: boolean; newEmails: MailItem[]; sendReceiveTriggered?: boolean; error?: string }> {
  if (process.platform !== 'win32') {
    return { outlookRunning: false, newEmails: [] };
  }

  const maxItems = Math.max(5, Math.min(50, Math.round(options?.maxItems ?? 25)));
  const timeoutMs = Math.max(3000, Math.min(25_000, Math.round(options?.timeoutMs ?? 12_000)));
  const sendReceiveBlock = options?.triggerSendReceive
    ? `
try {
  $ns.SendAndReceive($false)
  Start-Sleep -Milliseconds 1200
} catch {}
`
    : '';

  const psScript = `
$ErrorActionPreference = 'SilentlyContinue'
try {
  $OutputEncoding = [System.Text.UTF8Encoding]::new($false)
  [Console]::OutputEncoding = $OutputEncoding
} catch {}

$app = $null
try {
  $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch {
  $app = $null
}

if ($null -eq $app) {
  @{ outlookRunning = $false; messages = @() } | ConvertTo-Json -Compress
  exit 0
}

$ns = $app.GetNamespace("MAPI")
${sendReceiveBlock}
$messages = New-Object System.Collections.Generic.List[object]

foreach ($store in $ns.Stores) {
  $sName = [string]$store.DisplayName
  $inbox = $null
  try { $inbox = $store.GetDefaultFolder(6) } catch { $inbox = $null }
  if ($null -eq $inbox) { continue }
  $folderName = [string]$inbox.Name
  if (-not $folderName) { $folderName = '받은 편지함' }
  try {
    $items = $inbox.Items
    $items.Sort('[ReceivedTime]', $true)
    $taken = 0
    foreach ($it in $items) {
      if ($taken -ge ${maxItems}) { break }
      try {
        if ([int]$it.Class -ne 43) { continue }
      } catch { continue }
      $entryId = ''
      try { $entryId = [string]$it.EntryID } catch {}
      if (-not $entryId) { continue }
      $taken++

      $dt = Get-Date
      try { if ($it.ReceivedTime) { $dt = [datetime]$it.ReceivedTime } } catch {}
      $bodyTxt = ''
      try {
        $bodyTxt = [string]$it.Body
        if ($bodyTxt.Length -gt 4000) { $bodyTxt = $bodyTxt.Substring(0, 4000) }
      } catch {}
      $attCount = 0
      try { $attCount = [int]$it.Attachments.Count } catch {}

      $messages.Add([PSCustomObject]@{
        entryId = $entryId
        storeName = $sName
        folderPath = $folderName
        subject = [string]$it.Subject
        fromName = [string]$it.SenderName
        fromEmail = [string]$it.SenderEmailAddress
        to = [string]$it.To
        receivedIso = $dt.ToUniversalTime().ToString('o')
        dateStr = $dt.ToString('yyyy-MM-dd')
        timeLabel = $dt.ToString('HH:mm')
        unread = [bool]$it.UnRead
        flagged = ([int]$it.FlagStatus -eq 2)
        importance = ([int]$it.Importance -eq 2)
        categories = [string]$it.Categories
        body = $bodyTxt
        hasAttachment = ($attCount -gt 0)
      })
    }
  } catch {}
}

[PSCustomObject]@{
  outlookRunning = $true
  messages = @($messages)
} | ConvertTo-Json -Depth 4 -Compress
`.trim();

  try {
    const rawJson = await new Promise<string>((resolve, reject) => {
      execFile(
        'powershell.exe',
        ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', psScript],
        {
          encoding: 'utf8',
          timeout: timeoutMs,
          windowsHide: true,
          maxBuffer: 8 * 1024 * 1024,
        },
        (err, stdout) => {
          if (err) {
            reject(err);
            return;
          }
          resolve(stdout || '');
        }
      );
    });

    const trimmed = rawJson.trim();
    if (!trimmed) {
      return { outlookRunning: false, newEmails: [] };
    }
    const parsed = JSON.parse(trimmed) as {
      outlookRunning?: boolean;
      messages?: RawAsyncInboxMessage[];
    };
    if (!parsed.outlookRunning) {
      return { outlookRunning: false, newEmails: [] };
    }

    const existingIds = db.getExistingEmailIds();
    const newRecords: LocalEmailRecord[] = [];
    const seenBatchIds = new Set<string>();

    for (const m of parsed.messages || []) {
      const store = String(m.storeName || 'Outlook').trim();
      const entryId = String(m.entryId || '').trim();
      if (!entryId) continue;
      const hash = crypto.createHash('sha1').update(`${store}::${entryId}`).digest('hex').slice(0, 16);
      const id = `com-${hash}`;
      if (existingIds.has(id) || seenBatchIds.has(id)) continue;
      seenBatchIds.add(id);

      const folderPath = String(m.folderPath || '받은 편지함').trim() || '받은 편지함';
      const subject = String(m.subject || '(제목 없음)').trim() || '(제목 없음)';
      const cleanBody = String(m.body || '').trim() || '(본문 없음)';

      newRecords.push({
        id,
        storeName: store,
        folder: 'inbox',
        folderName: folderPath,
        subject,
        fromName: String(m.fromName || m.fromEmail || '알 수 없음').trim(),
        fromEmail: String(m.fromEmail || 'unknown@outlook.local').trim(),
        to: extractRecipients(m.to),
        date: String(m.dateStr || new Date().toISOString().slice(0, 10)),
        timestamp: String(m.receivedIso || new Date().toISOString()),
        timeLabel: String(m.timeLabel || '00:00'),
        snippet: cleanBody.replace(/\s+/g, ' ').slice(0, 140),
        body: cleanBody.slice(0, MAX_BODY_CHARS),
        hasAttachment: Boolean(m.hasAttachment),
        conversationKey: normalizeConversationSubject(subject),
        importance: m.importance ? 'high' : 'normal',
        category: m.categories ? String(m.categories) : folderPath,
        isRead: !m.unread,
        isFlagged: Boolean(m.flagged),
        isDeleted: false,
        source: 'outlook-live',
      });
    }

    if (newRecords.length > 0) {
      db.upsertEmailsBatch(newRecords);
    }

    const newEmails: MailItem[] = [];
    for (const rec of newRecords) {
      const stored = db.getEmailById(rec.id);
      if (stored) {
        newEmails.push(stored);
      }
    }

    return {
      outlookRunning: true,
      newEmails,
    };
  } catch (err) {
    return {
      outlookRunning: false,
      newEmails: [],
      error: err instanceof Error ? err.message : String(err),
    };
  }
}
