import type { MailItem } from '../types/mail';

export interface OutlookContactRecord {
  id: string;
  fullName: string;
  email: string;
  company?: string;
  department?: string;
  jobTitle?: string;
  phone?: string;
  source: 'outlook-pc' | 'pst-derived' | 'manual';
}

export interface OutlookPcAccountRecord {
  displayName: string;
  smtpAddress: string;
  accountType: 'POP3' | 'IMAP' | 'Exchange' | 'Other';
  pop3Host?: string;
  pop3Port?: number;
  smtpHost?: string;
  smtpPort?: number;
  deliveryStorePath?: string;
}

export interface OutlookPcProfileSnapshot {
  capturedAt: string;
  machineName?: string;
  outlookVersion?: string;
  defaultProfileName?: string;
  accounts: OutlookPcAccountRecord[];
  pstPaths: string[];
  contacts: OutlookContactRecord[];
  folders: string[];
  signatureNames: string[];
}

export type AccessPolicyScope = 'org-wide' | 'per-user-override';

export interface EnterpriseApiPolicy {
  scope: AccessPolicyScope;
  allowEmployeeProviderSelection: boolean;
  allowEmployeePop3Connect: boolean;
  allowLocalPstImport: boolean;
  cloudflareWorkerUrl: string;
  centralCreditPriority: 'hybrid-gcp-first' | 'gemini-only' | 'fallback-balanced';
}

export const DEFAULT_ENTERPRISE_API_POLICY: EnterpriseApiPolicy = {
  scope: 'org-wide',
  allowEmployeeProviderSelection: true,
  allowEmployeePop3Connect: true,
  allowLocalPstImport: true,
  cloudflareWorkerUrl: '',
  centralCreditPriority: 'hybrid-gcp-first',
};

/**
 * Produces an isolated storage key per employee email/ID so 10~20+ staff sharing
 * a deployment never collide or leak personal settings/projects/drafts.
 */
export function getEmployeeStorageKey(baseKey: string, employeeEmail?: string | null): string {
  const normalized = String(employeeEmail || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9@._-]/g, '');
  if (!normalized) return baseKey;
  return `${baseKey}__u_${normalized}`;
}

/**
 * Generates a safe, read-only Windows PowerShell script that inspects the currently
 * installed Microsoft Outlook on the employee's PC (Accounts, POP3/IMAP registry settings,
 * PST/OST file paths, Address Book Contacts, Folders, Signatures) and outputs a local
 * JSON file (`bmail-outlook-pc-profile.json`) without ever exporting passwords.
 */
export function buildWindowsOutlookInspectorScript(): string {
  return [
    '# ==============================================================================',
    '# Bmail Local Windows Outlook Profile · PST · 주소록 자동 파악 스크립트',
    '# 주의: 비밀번호는 절대 읽거나 내보내지 않습니다 (계정·PST 경로·주소록·폴더만 추출).',
    '# 실행 방법: PowerShell에서 .\\bmail-outlook-pc-inspector.ps1 실행 후 생성된 JSON을 Bmail에 업로드',
    '# ==============================================================================',
    '$ErrorActionPreference = "SilentlyContinue"',
    '$outPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "bmail-outlook-pc-profile.json"',
    '',
    '$accounts = @()',
    '$contacts = @()',
    '$folders = @()',
    '$pstPaths = @()',
    '$signatureNames = @()',
    '$outlookVersion = "Unknown"',
    '$defaultProfile = ""',
    '',
    '# 1. Detect installed Outlook default profile & PST/OST paths from Windows Registry / Disk',
    '$regBase = "HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook"',
    'if (Test-Path $regBase) {',
    '  $defaultProfile = (Get-ItemProperty -Path $regBase -Name "DefaultProfile").DefaultProfile',
    '}',
    '$docPstDir = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "Outlook Files"',
    'if (Test-Path $docPstDir) {',
    '  Get-ChildItem -Path $docPstDir -Filter "*.pst" | ForEach-Object { $pstPaths += $_.FullName }',
    '}',
    '$localOstDir = Join-Path $env:LOCALAPPDATA "Microsoft\\Outlook"',
    'if (Test-Path $localOstDir) {',
    '  Get-ChildItem -Path $localOstDir -Include "*.pst","*.ost" -Recurse | ForEach-Object { $pstPaths += $_.FullName }',
    '}',
    '$sigDir = Join-Path $env:APPDATA "Microsoft\\Signatures"',
    'if (Test-Path $sigDir) {',
    '  Get-ChildItem -Path $sigDir -Filter "*.htm" | ForEach-Object { $signatureNames += $_.BaseName }',
    '}',
    '',
    '# 2. Query local Outlook COM (Accounts, Stores, Contacts Address Book, Folders)',
    'try {',
    '  $ol = New-Object -ComObject Outlook.Application',
    '  $outlookVersion = $ol.Version',
    '  $ns = $ol.GetNamespace("MAPI")',
    '',
    '  foreach ($acc in $ns.Accounts) {',
    '    $typeLabel = switch ([int]$acc.AccountType) { 0 { "Exchange" } 1 { "IMAP" } 2 { "POP3" } default { "Other" } }',
    '    $accounts += [PSCustomObject]@{',
    '      displayName = [string]$acc.DisplayName',
    '      smtpAddress = [string]$acc.SmtpAddress',
    '      accountType = $typeLabel',
    '      deliveryStorePath = [string]$acc.DeliveryStore.FilePath',
    '    }',
    '    if ($acc.DeliveryStore.FilePath) { $pstPaths += [string]$acc.DeliveryStore.FilePath }',
    '  }',
    '',
    '  # Contacts (olFolderContacts = 10)',
    '  $contactsFolder = $ns.GetDefaultFolder(10)',
    '  $cCount = 0',
    '  foreach ($item in $contactsFolder.Items) {',
    '    if ($cCount -ge 500) { break }',
    '    $addr = [string]$item.Email1Address',
    '    if ($addr -and $addr.Contains("@")) {',
    '      $contacts += [PSCustomObject]@{',
    '        id = "pc-contact-$cCount"',
    '        fullName = if ($item.FullName) { [string]$item.FullName } else { $addr }',
    '        email = $addr',
    '        company = [string]$item.CompanyName',
    '        department = [string]$item.Department',
    '        jobTitle = [string]$item.JobTitle',
    '        phone = [string]$item.BusinessTelephoneNumber',
    '        source = "outlook-pc"',
    '      }',
    '      $cCount++',
    '    }',
    '  }',
    '',
    '  # Top-level mailbox folders',
    '  foreach ($store in $ns.Stores) {',
    '    if ($store.FilePath) { $pstPaths += [string]$store.FilePath }',
    '    $root = $store.GetRootFolder()',
    '    foreach ($f in $root.Folders) {',
    '      $folders += "$($store.DisplayName) / $($f.Name)"',
    '    }',
    '  }',
    '} catch {',
    '  Write-Host "[Bmail] Outlook COM 접근 생략 (레지스트리 및 디스크 PST 정보만 수집): $($_.Exception.Message)"',
    '}',
    '',
    '$snapshot = [PSCustomObject]@{',
    '  capturedAt = (Get-Date).ToString("o")',
    '  machineName = $env:COMPUTERNAME',
    '  outlookVersion = $outlookVersion',
    '  defaultProfileName = $defaultProfile',
    '  accounts = $accounts',
    '  pstPaths = ($pstPaths | Select-Object -Unique)',
    '  contacts = $contacts',
    '  folders = ($folders | Select-Object -Unique)',
    '  signatureNames = ($signatureNames | Select-Object -Unique)',
    '}',
    '',
    '$snapshot | ConvertTo-Json -Depth 6 | Set-Content -Path $outPath -Encoding UTF8',
    'Write-Host "[Bmail] 로컬 PC Outlook 설정 추출 완료: $outPath"',
    '',
  ].join('\r\n');
}

/**
 * Validates and normalizes an uploaded bmail-outlook-pc-profile.json snapshot.
 * Strips any unexpected secret/password fields defensively.
 */
export function parseWindowsOutlookProfileJson(raw: unknown): OutlookPcProfileSnapshot {
  const obj = raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : null;
  if (!obj) {
    throw new Error('유효한 Outlook PC 프로필 JSON 형식이 아닙니다.');
  }

  const rawAccounts = Array.isArray(obj.accounts) ? obj.accounts : [];
  const accounts: OutlookPcAccountRecord[] = rawAccounts
    .slice(0, 30)
    .map((item): OutlookPcAccountRecord | null => {
      const rec = item && typeof item === 'object' ? (item as Record<string, unknown>) : {};
      const smtpAddress = String(rec.smtpAddress || '').trim().toLowerCase();
      if (!smtpAddress.includes('@')) return null;
      const rawType = String(rec.accountType || 'Other');
      const accountType: OutlookPcAccountRecord['accountType'] =
        rawType === 'POP3' || rawType === 'IMAP' || rawType === 'Exchange' ? rawType : 'Other';
      return {
        displayName: String(rec.displayName || smtpAddress).trim().slice(0, 120),
        smtpAddress: smtpAddress.slice(0, 160),
        accountType,
        pop3Host: typeof rec.pop3Host === 'string' ? rec.pop3Host.trim().slice(0, 160) : undefined,
        pop3Port: Number.isFinite(Number(rec.pop3Port)) ? Number(rec.pop3Port) : undefined,
        smtpHost: typeof rec.smtpHost === 'string' ? rec.smtpHost.trim().slice(0, 160) : undefined,
        smtpPort: Number.isFinite(Number(rec.smtpPort)) ? Number(rec.smtpPort) : undefined,
        deliveryStorePath:
          typeof rec.deliveryStorePath === 'string' ? rec.deliveryStorePath.trim().slice(0, 320) : undefined,
      };
    })
    .filter((x): x is OutlookPcAccountRecord => x !== null);

  const rawContacts = Array.isArray(obj.contacts) ? obj.contacts : [];
  const contacts: OutlookContactRecord[] = rawContacts
    .slice(0, 1000)
    .map((item, idx): OutlookContactRecord | null => {
      const rec = item && typeof item === 'object' ? (item as Record<string, unknown>) : {};
      const email = String(rec.email || '').trim().toLowerCase();
      if (!email.includes('@')) return null;
      return {
        id: String(rec.id || `pc-contact-${idx}`).slice(0, 80),
        fullName: String(rec.fullName || email).trim().slice(0, 120),
        email: email.slice(0, 160),
        company: typeof rec.company === 'string' && rec.company.trim() ? rec.company.trim().slice(0, 120) : undefined,
        department:
          typeof rec.department === 'string' && rec.department.trim() ? rec.department.trim().slice(0, 120) : undefined,
        jobTitle:
          typeof rec.jobTitle === 'string' && rec.jobTitle.trim() ? rec.jobTitle.trim().slice(0, 120) : undefined,
        phone: typeof rec.phone === 'string' && rec.phone.trim() ? rec.phone.trim().slice(0, 60) : undefined,
        source: 'outlook-pc',
      };
    })
    .filter((x): x is OutlookContactRecord => x !== null);

  const strList = (val: unknown, max: number): string[] =>
    Array.isArray(val)
      ? Array.from(
          new Set(
            val
              .map((v) => String(v || '').trim())
              .filter(Boolean)
              .slice(0, max)
          )
        )
      : [];

  return {
    capturedAt:
      typeof obj.capturedAt === 'string' && obj.capturedAt.trim()
        ? obj.capturedAt.trim().slice(0, 40)
        : new Date().toISOString(),
    machineName:
      typeof obj.machineName === 'string' && obj.machineName.trim()
        ? obj.machineName.trim().slice(0, 80)
        : undefined,
    outlookVersion:
      typeof obj.outlookVersion === 'string' && obj.outlookVersion.trim()
        ? obj.outlookVersion.trim().slice(0, 40)
        : undefined,
    defaultProfileName:
      typeof obj.defaultProfileName === 'string' && obj.defaultProfileName.trim()
        ? obj.defaultProfileName.trim().slice(0, 80)
        : undefined,
    accounts,
    pstPaths: strList(obj.pstPaths, 50),
    contacts,
    folders: strList(obj.folders, 200),
    signatureNames: strList(obj.signatureNames, 30),
  };
}

/**
 * Automatically derives a deduplicated Address Book from imported PST/POP3 emails
 * merged with any imported Windows PC Outlook Contacts.
 */
export function buildUnifiedAddressBook(
  emails: MailItem[],
  pcContacts: OutlookContactRecord[] = []
): OutlookContactRecord[] {
  const map = new Map<string, OutlookContactRecord>();

  for (const c of pcContacts) {
    const key = c.email.trim().toLowerCase();
    if (key.includes('@')) map.set(key, c);
  }

  for (const mail of emails) {
    const fromEmail = (mail.fromEmail || '').trim().toLowerCase();
    if (fromEmail.includes('@') && !map.has(fromEmail)) {
      const domain = fromEmail.split('@')[1] || '';
      map.set(fromEmail, {
        id: `mail-contact-${fromEmail}`,
        fullName: (mail.fromName || fromEmail).trim(),
        email: fromEmail,
        company: domain,
        source: 'pst-derived',
      });
    }
    for (const rawTo of mail.to || []) {
      const match = rawTo.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
      if (!match) continue;
      const toEmail = match[0].toLowerCase();
      if (!map.has(toEmail)) {
        map.set(toEmail, {
          id: `mail-contact-${toEmail}`,
          fullName: rawTo.replace(/<[^>]+>/g, '').trim() || toEmail,
          email: toEmail,
          company: toEmail.split('@')[1] || '',
          source: 'pst-derived',
        });
      }
    }
  }

  return Array.from(map.values()).sort((a, b) => a.fullName.localeCompare(b.fullName, 'ko'));
}

/**
 * Exports Contacts to RFC 6350 vCard (.vcf) so employees can write/import contacts
 * directly back into Microsoft Outlook Desktop or mobile address books.
 */
export function exportContactsToVcf(contacts: OutlookContactRecord[]): string {
  const escapeVcf = (v: string) => v.replace(/[\r\n]/g, ' ').replace(/([,;\\])/g, '\\$1');
  return contacts
    .map((c) => {
      const lines = [
        'BEGIN:VCARD',
        'VERSION:3.0',
        `FN:${escapeVcf(c.fullName)}`,
        `EMAIL;TYPE=INTERNET:${escapeVcf(c.email)}`,
      ];
      if (c.company) lines.push(`ORG:${escapeVcf(c.company)}${c.department ? `;${escapeVcf(c.department)}` : ''}`);
      if (c.jobTitle) lines.push(`TITLE:${escapeVcf(c.jobTitle)}`);
      if (c.phone) lines.push(`TEL;TYPE=WORK,VOICE:${escapeVcf(c.phone)}`);
      lines.push('END:VCARD');
      return lines.join('\r\n');
    })
    .join('\r\n');
}
