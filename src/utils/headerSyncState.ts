export type HeaderSyncSource = 'microsoft365' | 'pop3' | null;

export type HeaderAccountSyncMode = 'disconnected' | 'connected' | 'syncing';

export interface HeaderAccountSyncInput {
  microsoftConnected: boolean;
  microsoftEmail: string | null;
  pop3Connected: boolean;
  pop3Email: string | null;
  pop3Protocol?: 'imap' | 'pop3';
  pop3ProviderName?: string | null;
  isSyncingGraph: boolean;
  isSyncingPop3: boolean;
  lastSyncedAt: string | null;
  lastSyncSource: HeaderSyncSource;
}

export interface HeaderAccountSyncView {
  mode: HeaderAccountSyncMode;
  isSyncing: boolean;
  isConnected: boolean;
  primaryEmail: string | null;
  activeSourcesLabel: string;
  compactStatusLabel: string;
  lastSyncedTimeLabel: string | null;
  tooltip: string;
  ariaLabel: string;
}

/**
 * Formats an ISO timestamp into HH:MM:SS (24h local time) safely without throwing.
 */
export function formatSyncClockTime(isoTimestamp: string | null | undefined): string | null {
  if (!isoTimestamp) return null;
  const parsed = new Date(isoTimestamp);
  if (Number.isNaN(parsed.getTime())) return null;
  const hh = String(parsed.getHours()).padStart(2, '0');
  const mm = String(parsed.getMinutes()).padStart(2, '0');
  const ss = String(parsed.getSeconds()).padStart(2, '0');
  return `${hh}:${mm}:${ss}`;
}

/**
 * Derives deterministic visual & accessibility state for the header's
 * connected account indicator, including real-time background sync heartbeat/spinner state.
 */
export function buildHeaderAccountSyncView(
  input: HeaderAccountSyncInput
): HeaderAccountSyncView {
  const isSyncing = Boolean(input.isSyncingGraph || input.isSyncingPop3);
  const isConnected = Boolean(input.microsoftConnected || input.pop3Connected);
  const primaryEmail = input.pop3Email || input.microsoftEmail || null;
  const pop3ProtoLabel = (input.pop3Protocol || 'pop3').toUpperCase();
  const lastSyncedTimeLabel = formatSyncClockTime(input.lastSyncedAt);

  const syncingParts: string[] = [];
  if (input.isSyncingGraph) syncingParts.push('Microsoft 365');
  if (input.isSyncingPop3) {
    syncingParts.push(
      input.pop3ProviderName
        ? `${input.pop3ProviderName} (${pop3ProtoLabel})`
        : pop3ProtoLabel
    );
  }

  const connectedParts: string[] = [];
  if (input.microsoftConnected) connectedParts.push('Microsoft 365');
  if (input.pop3Connected) {
    connectedParts.push(
      input.pop3ProviderName
        ? `${input.pop3ProviderName} ${pop3ProtoLabel}`
        : pop3ProtoLabel
    );
  }

  if (isSyncing) {
    const activeSourcesLabel = syncingParts.join(' · ') || '메일함';
    return {
      mode: 'syncing',
      isSyncing: true,
      isConnected,
      primaryEmail,
      activeSourcesLabel,
      compactStatusLabel: '동기화 중',
      lastSyncedTimeLabel,
      tooltip: `${activeSourcesLabel} 백그라운드 동기화 진행 중...${
        primaryEmail ? ` (${primaryEmail})` : ''
      }`,
      ariaLabel: `계정 및 연결 상태 — ${activeSourcesLabel} 백그라운드 동기화 중`,
    };
  }

  if (isConnected) {
    const activeSourcesLabel = connectedParts.join(' · ');
    const syncSuffix = lastSyncedTimeLabel
      ? ` · 최신 동기화 ${lastSyncedTimeLabel}`
      : ' · 실시간 연결 유지';
    return {
      mode: 'connected',
      isSyncing: false,
      isConnected: true,
      primaryEmail,
      activeSourcesLabel,
      compactStatusLabel: lastSyncedTimeLabel ? `최신 ${lastSyncedTimeLabel.slice(0, 5)}` : '연결됨',
      lastSyncedTimeLabel,
      tooltip: `${primaryEmail || activeSourcesLabel} 연결됨 (${activeSourcesLabel}${syncSuffix}) · 클릭하여 계정·EMV 설정`,
      ariaLabel: `계정 및 연결 상태 — ${primaryEmail || activeSourcesLabel} 연결됨${syncSuffix}`,
    };
  }

  return {
    mode: 'disconnected',
    isSyncing: false,
    isConnected: false,
    primaryEmail: null,
    activeSourcesLabel: '',
    compactStatusLabel: '',
    lastSyncedTimeLabel: null,
    tooltip: 'Microsoft 365 연결 또는 계정·EMV 설정',
    ariaLabel: '계정 및 연결 상태',
  };
}
