import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import {
  verifyOutlookItemSafety,
  type OutlookActiveItemSnapshot,
  type OutlookCompanionMode,
  type OutlookCompanionPollResult,
} from '../utils/outlookCompanionVerifier';

export * from '../utils/outlookCompanionVerifier';

const execFileAsync = promisify(execFile);

function parseAddressArray(raw: unknown): string[] {
  if (Array.isArray(raw)) {
    return raw.map((x) => String(x || '').trim()).filter(Boolean);
  }
  if (typeof raw === 'string') {
    return raw
      .split(/[;,]/)
      .map((s) => s.trim())
      .filter(Boolean);
  }
  return [];
}

/**
 * Reads the currently active email or draft window from running Classic Outlook via COM
 * (ActiveInspector first, then ActiveExplorer selection) without mutating Outlook state.
 */
export async function readActiveOutlookItemAsync(options?: {
  timeoutMs?: number;
}): Promise<OutlookCompanionPollResult> {
  const polledAt = new Date().toISOString();
  if (process.platform !== 'win32') {
    return {
      outlookRunning: false,
      mode: 'idle',
      item: null,
      verification: null,
      polledAt,
    };
  }

  const timeoutMs = Math.max(1500, Math.min(15_000, Math.round(options?.timeoutMs ?? 4500)));

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
  @{ outlookRunning = $false; mode = 'idle'; item = $null } | ConvertTo-Json -Compress
  exit 0
}

function Parse-MailItem($it, [string]$srcType) {
  $subj = ''
  try { $subj = [string]$it.Subject } catch {}
  $fromName = ''
  try { $fromName = [string]$it.SenderName } catch {}
  $fromEmail = ''
  try { $fromEmail = [string]$it.SenderEmailAddress } catch {}
  if ((-not $fromEmail -or $fromEmail -like '/O=*') -and $it.SendUsingAccount) {
    try { $fromEmail = [string]$it.SendUsingAccount.SmtpAddress } catch {}
  }

  $toList = New-Object System.Collections.Generic.List[string]
  $ccList = New-Object System.Collections.Generic.List[string]
  $bccList = New-Object System.Collections.Generic.List[string]
  try {
    foreach ($rec in $it.Recipients) {
      $addr = ''
      try { $addr = [string]$rec.Address } catch {}
      if (-not $addr -or $addr -like '/O=*') {
        try {
          if ($rec.AddressEntry -and $rec.AddressEntry.Type -eq 'EX') {
            $exu = $rec.AddressEntry.GetExchangeUser()
            if ($exu -and $exu.PrimarySmtpAddress) { $addr = [string]$exu.PrimarySmtpAddress }
          }
        } catch {}
      }
      if (-not $addr) { try { $addr = [string]$rec.Name } catch {} }
      if (-not $addr) { continue }
      $rtype = 1
      try { $rtype = [int]$rec.Type } catch {}
      if ($rtype -eq 2) { $ccList.Add($addr) }
      elseif ($rtype -eq 3) { $bccList.Add($addr) }
      else { $toList.Add($addr) }
    }
  } catch {}
  if ($toList.Count -eq 0 -and $it.To) {
    foreach ($t in ([string]$it.To -split ';')) { if ($t.Trim()) { $toList.Add($t.Trim()) } }
  }
  if ($ccList.Count -eq 0 -and $it.CC) {
    foreach ($t in ([string]$it.CC -split ';')) { if ($t.Trim()) { $ccList.Add($t.Trim()) } }
  }

  $bodyTxt = ''
  try {
    $bodyTxt = [string]$it.Body
    if ($bodyTxt.Length -gt 8000) { $bodyTxt = $bodyTxt.Substring(0, 8000) }
  } catch {}

  $atts = New-Object System.Collections.Generic.List[string]
  $attCnt = 0
  try {
    $attCnt = [int]$it.Attachments.Count
    foreach ($a in $it.Attachments) {
      try { $atts.Add([string]$a.FileName) } catch {}
    }
  } catch {}

  $sentFlag = $false
  try { $sentFlag = [bool]$it.Sent } catch {}
  $entryId = ''
  try { $entryId = [string]$it.EntryID } catch {}
  $dt = Get-Date
  try { if ($it.ReceivedTime) { $dt = [datetime]$it.ReceivedTime } } catch {}
  $impNum = 1
  try { $impNum = [int]$it.Importance } catch {}
  $imp = if ($impNum -eq 2) { 'high' } elseif ($impNum -eq 0) { 'low' } else { 'normal' }

  return [PSCustomObject]@{
    entryId = $entryId
    sourceType = $srcType
    subject = $subj
    fromName = $fromName
    fromEmail = $fromEmail
    to = $toList
    cc = $ccList
    bcc = $bccList
    body = $bodyTxt
    attachments = $atts
    attachmentsCount = $attCnt
    receivedAt = $dt.ToString('o')
    sent = $sentFlag
    importance = $imp
  }
}

$snap = $null
$mode = 'idle'

try {
  $insp = $app.ActiveInspector()
  if ($null -ne $insp -and $null -ne $insp.CurrentItem -and [int]$insp.CurrentItem.Class -eq 43) {
    $snap = Parse-MailItem $insp.CurrentItem 'inspector'
    $mode = if ($snap.sent) { 'reading' } else { 'composing' }
  }
} catch {}

if ($null -eq $snap) {
  try {
    $exp = $app.ActiveExplorer()
    if ($null -ne $exp -and $null -ne $exp.Selection -and $exp.Selection.Count -ge 1) {
      $sel = $exp.Selection.Item(1)
      if ($null -ne $sel -and [int]$sel.Class -eq 43) {
        $snap = Parse-MailItem $sel 'explorer'
        $mode = if ($snap.sent) { 'reading' } else { 'composing' }
      }
    }
  } catch {}
}

[PSCustomObject]@{
  outlookRunning = $true
  mode = $mode
  item = $snap
} | ConvertTo-Json -Depth 5 -Compress
`;

  try {
    const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
    const { stdout } = await execFileAsync(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
      { timeout: timeoutMs, windowsHide: true, maxBuffer: 2 * 1024 * 1024 }
    );
    const trimmed = String(stdout || '').trim();
    if (!trimmed) {
      return { outlookRunning: false, mode: 'idle', item: null, verification: null, polledAt };
    }
    const parsed = JSON.parse(trimmed) as {
      outlookRunning?: boolean;
      mode?: OutlookCompanionMode;
      item?: Record<string, unknown> | null;
    };
    const outlookRunning = Boolean(parsed.outlookRunning);
    if (!outlookRunning || !parsed.item) {
      return {
        outlookRunning,
        mode: 'idle',
        item: null,
        verification: null,
        polledAt,
      };
    }

    const raw = parsed.item;
    const item: OutlookActiveItemSnapshot = {
      entryId: String(raw.entryId || ''),
      sourceType: raw.sourceType === 'inspector' ? 'inspector' : 'explorer',
      subject: String(raw.subject || ''),
      fromName: String(raw.fromName || ''),
      fromEmail: String(raw.fromEmail || ''),
      to: parseAddressArray(raw.to),
      cc: parseAddressArray(raw.cc),
      bcc: parseAddressArray(raw.bcc),
      body: String(raw.body || ''),
      attachments: parseAddressArray(raw.attachments),
      attachmentsCount: Number(raw.attachmentsCount || 0),
      receivedAt: String(raw.receivedAt || polledAt),
      sent: Boolean(raw.sent),
      importance: raw.importance === 'high' ? 'high' : raw.importance === 'low' ? 'low' : 'normal',
    };
    const mode: OutlookCompanionMode =
      parsed.mode === 'composing' ? 'composing' : item.sent ? 'reading' : 'composing';

    return {
      outlookRunning: true,
      mode,
      item,
      verification: verifyOutlookItemSafety(item, mode),
      polledAt,
    };
  } catch (err) {
    return {
      outlookRunning: false,
      mode: 'idle',
      item: null,
      verification: null,
      polledAt,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}
