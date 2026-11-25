import type { MailItem } from '../types/mail';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import {
  EMPTY_PROJECT_MAIL_OVERRIDES,
  type ProjectMailOverrides,
} from '../utils/outlookWorkspaceModel';
import { buildExecutiveDashboardReport } from './executiveReportEngine';
import type { LocalMailDb } from './localMailDb';
import { pollLiveOutlookInboxAsync } from './localOutlookMigrator';
import {
  summarizeNewMailsForNotification,
  type NewMailQuickSummary,
} from './newMailNotifier';
import {
  fetchImapMessages,
  fetchPop3Messages,
  type Pop3ServerSession,
  Pop3SessionStore,
} from './pop3AutoConnect';

export interface ProactiveAlertSettings {
  enabled: boolean;
  pollIntervalSec: number;
  unviewedGraceSec: number;
  sgwPushEnabled: boolean;
  sgwUrgentOnly: boolean;
  sgwTargetRoomId: string;
}

export const DEFAULT_PROACTIVE_ALERT_SETTINGS: ProactiveAlertSettings = {
  enabled: true,
  pollIntervalSec: 60,
  unviewedGraceSec: 120,
  sgwPushEnabled: true,
  sgwUrgentOnly: false,
  sgwTargetRoomId: '',
};

export interface ProactiveDashboardDigest {
  updatedAt: string;
  totalEmails: number;
  unrepliedIncomingCount: number;
  unreadCount: number;
  topUnrepliedSubjects: string[];
  recentSummaryCount: number;
}

export interface PendingMonitorAlert {
  emailId: string;
  arrivedAt: number;
  viewedByUser: boolean;
  trayNotified: boolean;
  sgwNotified: boolean;
  sgwNotifiedAt?: string;
  summary: NewMailQuickSummary;
}

export interface SgwBridgeLike {
  getStatus(): {
    connected: boolean;
    sgwBaseUrl: string;
    user: { sub: string; name: string; email: string } | null;
  };
  sendDirectMessage(payload: {
    receiverId: string;
    receiverName?: string;
    content: string;
  }): Promise<{ ok: boolean; id: string }>;
  sendRoomMessage(
    roomId: string,
    payload: { content: string }
  ): Promise<{ ok: boolean; id: string }>;
}

export interface BackgroundMonitorDependencies {
  localDb: LocalMailDb;
  pop3Sessions: Pop3SessionStore;
  sgwChatBridge: SgwBridgeLike;
  enableAiSummary?: boolean;
  nowFn?: () => number;
  pollPop3SessionFn?: (session: Pop3ServerSession) => Promise<MailItem[]>;
  pollOutlookFn?: (db: LocalMailDb) => MailItem[] | Promise<MailItem[]>;
}

export function sanitizeProactiveAlertSettings(raw: unknown): ProactiveAlertSettings {
  if (!raw || typeof raw !== 'object') {
    return { ...DEFAULT_PROACTIVE_ALERT_SETTINGS };
  }
  const obj = raw as Record<string, unknown>;
  const pollSec = Number(obj.pollIntervalSec);
  const graceSec = Number(obj.unviewedGraceSec);
  return {
    enabled:
      typeof obj.enabled === 'boolean'
        ? obj.enabled
        : DEFAULT_PROACTIVE_ALERT_SETTINGS.enabled,
    pollIntervalSec: Number.isFinite(pollSec)
      ? Math.max(30, Math.min(600, Math.round(pollSec)))
      : DEFAULT_PROACTIVE_ALERT_SETTINGS.pollIntervalSec,
    unviewedGraceSec: Number.isFinite(graceSec)
      ? Math.max(0, Math.min(1800, Math.round(graceSec)))
      : DEFAULT_PROACTIVE_ALERT_SETTINGS.unviewedGraceSec,
    sgwPushEnabled:
      typeof obj.sgwPushEnabled === 'boolean'
        ? obj.sgwPushEnabled
        : DEFAULT_PROACTIVE_ALERT_SETTINGS.sgwPushEnabled,
    sgwUrgentOnly:
      typeof obj.sgwUrgentOnly === 'boolean'
        ? obj.sgwUrgentOnly
        : DEFAULT_PROACTIVE_ALERT_SETTINGS.sgwUrgentOnly,
    sgwTargetRoomId:
      typeof obj.sgwTargetRoomId === 'string'
        ? obj.sgwTargetRoomId.trim().slice(0, 120)
        : '',
  };
}

export function formatSgwMobileMailAlert(
  item: NewMailQuickSummary,
  unrepliedCount?: number
): string {
  const badge = item.urgency === 'high' ? '[긴급 메일]' : '[새 메일]';
  const lines = [
    `${badge} ${item.sender}`,
    `제목: ${item.subject}`,
    `요약: ${item.summary}`,
  ];
  if (item.drawings.length > 0) {
    lines.push(`도면: ${item.drawings.join(', ')}`);
  }
  if (typeof unrepliedCount === 'number' && unrepliedCount > 0) {
    lines.push(`대시보드 미회신 현황: ${unrepliedCount.toLocaleString()}건`);
  }
  lines.push('— Bmail 백그라운드 선행 정리 알림');
  return lines.join('\n');
}

export class BackgroundMailMonitor {
  private readonly localDb: LocalMailDb;
  private readonly pop3Sessions: Pop3SessionStore;
  private readonly sgwChatBridge: SgwBridgeLike;
  private readonly enableAiSummary: boolean;
  private readonly nowFn: () => number;
  private readonly pollPop3SessionFn?: (session: Pop3ServerSession) => Promise<MailItem[]>;
  private readonly pollOutlookFn?: (db: LocalMailDb) => MailItem[] | Promise<MailItem[]>;

  private readonly alerts = new Map<string, PendingMonitorAlert>();
  private readonly ingestedIds = new Set<string>();
  private readonly viewedIds = new Set<string>();

  private pollTimer: NodeJS.Timeout | null = null;
  private sgwDispatchTimer: NodeJS.Timeout | null = null;
  private isPolling = false;
  private isDispatching = false;
  private lastPollAt: string | null = null;
  private lastPollAddedCount = 0;
  private totalSgwAlertsSent = 0;
  private lastDigest: ProactiveDashboardDigest | null = null;

  constructor(deps: BackgroundMonitorDependencies) {
    this.localDb = deps.localDb;
    this.pop3Sessions = deps.pop3Sessions;
    this.sgwChatBridge = deps.sgwChatBridge;
    this.enableAiSummary = deps.enableAiSummary !== false;
    this.nowFn = deps.nowFn ?? (() => Date.now());
    this.pollPop3SessionFn = deps.pollPop3SessionFn;
    this.pollOutlookFn = deps.pollOutlookFn;
  }

  getSettings(): ProactiveAlertSettings {
    try {
      const saved = this.localDb.isReady()
        ? this.localDb.getState<unknown>('proactive_alert_settings')
        : null;
      return sanitizeProactiveAlertSettings(saved);
    } catch {
      return { ...DEFAULT_PROACTIVE_ALERT_SETTINGS };
    }
  }

  updateSettings(patch: Partial<ProactiveAlertSettings>): ProactiveAlertSettings {
    const merged = sanitizeProactiveAlertSettings({
      ...this.getSettings(),
      ...patch,
    });
    if (this.localDb.isReady()) {
      this.localDb.setUserState('proactive_alert_settings', merged);
    }
    if (this.pollTimer || this.sgwDispatchTimer) {
      this.stop();
      if (merged.enabled) {
        this.start();
      }
    }
    return merged;
  }

  start(): void {
    const settings = this.getSettings();
    if (!settings.enabled) return;
    if (!this.pollTimer) {
      this.pollTimer = setInterval(() => {
        void this.runBackgroundPollCycle();
      }, settings.pollIntervalSec * 1000);
      this.pollTimer.unref?.();
    }
    if (!this.sgwDispatchTimer) {
      this.sgwDispatchTimer = setInterval(() => {
        void this.dispatchDueSgwAlerts();
      }, 15_000);
      this.sgwDispatchTimer.unref?.();
    }
  }

  stop(): void {
    if (this.pollTimer) {
      clearInterval(this.pollTimer);
      this.pollTimer = null;
    }
    if (this.sgwDispatchTimer) {
      clearInterval(this.sgwDispatchTimer);
      this.sgwDispatchTimer = null;
    }
  }

  markViewed(emailId: string): void {
    const cleanId = String(emailId || '').trim();
    if (!cleanId) return;
    this.viewedIds.add(cleanId);
    const rec = this.alerts.get(cleanId);
    if (rec) {
      rec.viewedByUser = true;
    }
  }

  refreshDashboardDigest(force = false): ProactiveDashboardDigest | null {
    if (!this.localDb.isReady()) return null;
    if (
      !force &&
      this.lastDigest &&
      this.nowFn() - Date.parse(this.lastDigest.updatedAt) < 60_000
    ) {
      return this.lastDigest;
    }
    try {
      const projects = this.localDb.getState<ProjectDefinition[]>('projects') || [];
      const overrides =
        this.localDb.getState<ProjectMailOverrides>('projectMailOverrides') ||
        EMPTY_PROJECT_MAIL_OVERRIDES;
      const report = buildExecutiveDashboardReport(
        this.localDb,
        { periodType: 'all' },
        projects,
        overrides
      );
      const digest: ProactiveDashboardDigest = {
        updatedAt: new Date(this.nowFn()).toISOString(),
        totalEmails: report.kpis.totalEmails,
        unrepliedIncomingCount: report.kpis.unrepliedIncomingCount,
        unreadCount: report.kpis.unreadCount,
        topUnrepliedSubjects: report.unrepliedIncoming
          .slice(0, 5)
          .map((item) => `[${item.fromName}] ${item.subject}`),
        recentSummaryCount: this.alerts.size,
      };
      this.lastDigest = digest;
      this.localDb.setState('proactive_dashboard_digest', digest);
      return digest;
    } catch {
      return null;
    }
  }

  async ingestNewEmails(freshEmails: MailItem[]): Promise<NewMailQuickSummary[]> {
    const candidates = freshEmails.filter(
      (m) => m && m.id && m.folder === 'inbox' && !this.ingestedIds.has(m.id)
    );
    for (const m of freshEmails) {
      if (m?.id) this.ingestedIds.add(m.id);
    }
    if (candidates.length === 0) return [];

    const summaries = await summarizeNewMailsForNotification(candidates, {
      maxItems: 5,
      enableAi: this.enableAiSummary,
    });
    const now = this.nowFn();
    for (const s of summaries) {
      if (this.alerts.size >= 200) {
        const oldestKey = this.alerts.keys().next().value;
        if (oldestKey) this.alerts.delete(oldestKey);
      }
      const alreadyViewed =
        this.viewedIds.has(s.emailId) ||
        Boolean(this.localDb.getEmailById(s.emailId)?.isRead);
      this.alerts.set(s.emailId, {
        emailId: s.emailId,
        arrivedAt: now,
        viewedByUser: alreadyViewed,
        trayNotified: alreadyViewed,
        sgwNotified: false,
        summary: s,
      });
    }

    this.refreshDashboardDigest(true);
    return summaries;
  }

  consumeUnnotifiedTrayAlerts(): Array<{
    emailId: string;
    sender: string;
    subject: string;
    summary: string;
    urgency: 'high' | 'normal';
  }> {
    const out: Array<{
      emailId: string;
      sender: string;
      subject: string;
      summary: string;
      urgency: 'high' | 'normal';
    }> = [];
    const list = Array.from(this.alerts.values()).sort((a, b) => b.arrivedAt - a.arrivedAt);
    for (const item of list) {
      if (item.trayNotified || item.viewedByUser || this.viewedIds.has(item.emailId)) {
        continue;
      }
      item.trayNotified = true;
      out.push({
        emailId: item.summary.emailId,
        sender: item.summary.sender,
        subject: item.summary.subject,
        summary: item.summary.summary,
        urgency: item.summary.urgency,
      });
      if (out.length >= 3) break;
    }
    return out;
  }

  async dispatchDueSgwAlerts(): Promise<number> {
    if (this.isDispatching) return 0;
    const settings = this.getSettings();
    if (!settings.enabled || !settings.sgwPushEnabled) return 0;
    const sgwStatus = this.sgwChatBridge.getStatus();
    if (!sgwStatus.connected || !sgwStatus.user) return 0;

    this.isDispatching = true;
    let sent = 0;
    try {
      const now = this.nowFn();
      const graceMs = settings.unviewedGraceSec * 1000;
      for (const alert of this.alerts.values()) {
        if (alert.sgwNotified || alert.viewedByUser) continue;
        if (this.viewedIds.has(alert.emailId) || this.localDb.getEmailById(alert.emailId)?.isRead) {
          alert.viewedByUser = true;
          continue;
        }
        if (now - alert.arrivedAt < graceMs) continue;
        if (settings.sgwUrgentOnly && alert.summary.urgency !== 'high') {
          alert.sgwNotified = true;
          continue;
        }

        const content = formatSgwMobileMailAlert(
          alert.summary,
          this.lastDigest?.unrepliedIncomingCount
        );
        try {
          if (settings.sgwTargetRoomId) {
            await this.sgwChatBridge.sendRoomMessage(settings.sgwTargetRoomId, { content });
          } else {
            await this.sgwChatBridge.sendDirectMessage({
              receiverId: sgwStatus.user.sub,
              receiverName: sgwStatus.user.name,
              content,
            });
          }
          alert.sgwNotified = true;
          alert.sgwNotifiedAt = new Date(now).toISOString();
          this.totalSgwAlertsSent += 1;
          sent += 1;
        } catch {
          // Stop burst on transient SGW error and retry next tick
          break;
        }
      }
      return sent;
    } finally {
      this.isDispatching = false;
    }
  }

  async runBackgroundPollCycle(): Promise<{ addedCount: number; sgwSentCount: number }> {
    if (this.isPolling || !this.localDb.isReady()) {
      return { addedCount: 0, sgwSentCount: 0 };
    }
    const settings = this.getSettings();
    if (!settings.enabled) {
      return { addedCount: 0, sgwSentCount: 0 };
    }

    this.isPolling = true;
    let addedCount = 0;
    try {
      const existingBefore = this.localDb.getExistingEmailIds();
      const newlyDiscovered: MailItem[] = [];

      // 1. Poll persisted POP3/IMAP sessions
      for (const acc of this.pop3Sessions.listAccounts()) {
        const sess = this.pop3Sessions.get(acc.token);
        if (!sess) continue;
        try {
          let fetchedEmails: MailItem[];
          if (this.pollPop3SessionFn) {
            fetchedEmails = await this.pollPop3SessionFn(sess);
          } else {
            const res =
              sess.protocol === 'imap' && sess.imapHost
                ? await fetchImapMessages({
                    email: sess.email,
                    password: sess.password,
                    imapHost: sess.imapHost,
                    imapPort: sess.imapPort,
                    imapTls: sess.imapTls,
                    providerId: sess.providerId,
                    usernameFormat: sess.usernameFormat,
                    maxMessages: 25,
                  })
                : await fetchPop3Messages({
                    email: sess.email,
                    password: sess.password,
                    pop3Host: sess.pop3Host,
                    pop3Port: sess.pop3Port,
                    pop3Tls: sess.pop3Tls,
                    providerId: sess.providerId,
                    usernameFormat: sess.usernameFormat,
                    maxMessages: 25,
                  });
            fetchedEmails = res.emails;
          }
          const fresh = fetchedEmails.filter((m) => m?.id && !existingBefore.has(m.id));
          if (fresh.length > 0) {
            this.localDb.upsertEmailsBatch(
              fresh.map((m) => ({
                ...m,
                storeName: sess.email,
              }))
            );
            for (const m of fresh) {
              existingBefore.add(m.id);
              newlyDiscovered.push(m);
            }
          }
        } catch {
          // Ignore transient network error per account
        }
      }

      // 2. Poll running local Outlook Inbox asynchronously on Windows (non-blocking COM)
      if (this.pollOutlookFn) {
        try {
          const items = await this.pollOutlookFn(this.localDb);
          for (const m of items) {
            if (m?.id && !existingBefore.has(m.id)) {
              existingBefore.add(m.id);
              newlyDiscovered.push(m);
            }
          }
        } catch {
          // ignore
        }
      } else if (process.platform === 'win32') {
        try {
          const comSync = await pollLiveOutlookInboxAsync(this.localDb, {
            maxItems: 25,
            timeoutMs: 12_000,
          });
          for (const m of comSync.newEmails || []) {
            if (m?.id && !existingBefore.has(m.id)) {
              existingBefore.add(m.id);
              newlyDiscovered.push(m);
            }
          }
        } catch {
          // ignore
        }
      }

      if (newlyDiscovered.length > 0) {
        addedCount = newlyDiscovered.length;
        await this.ingestNewEmails(newlyDiscovered);
      } else if (!this.lastDigest) {
        this.refreshDashboardDigest();
      }

      const sgwSentCount = await this.dispatchDueSgwAlerts();
      this.lastPollAt = new Date(this.nowFn()).toISOString();
      this.lastPollAddedCount = addedCount;
      return { addedCount, sgwSentCount };
    } finally {
      this.isPolling = false;
    }
  }

  async sendTestSgwAlert(): Promise<{ ok: boolean; reason: string; target?: string }> {
    const settings = this.getSettings();
    const sgwStatus = this.sgwChatBridge.getStatus();
    if (!sgwStatus.connected || !sgwStatus.user) {
      return {
        ok: false,
        reason: 'SGW 채팅 계정이 연결되어 있지 않습니다. 먼저 SGW 채팅에 로그인해 주세요.',
      };
    }
    const digest = this.refreshDashboardDigest();
    const sampleSummary: NewMailQuickSummary = {
      emailId: 'test-alert',
      sender: 'Bmail 백그라운드 모니터',
      subject: '[테스트] 미열람 새 메일 · 대시보드 모바일 알림 연동 확인',
      urgency: 'normal',
      summary: `대시보드 선행 정리 정상 작동 (미회신 ${digest?.unrepliedIncomingCount ?? 0}건 추적 중)`,
      drawings: [],
      deadlines: [],
      summarizedBy: 'deterministic',
    };
    const content = formatSgwMobileMailAlert(sampleSummary, digest?.unrepliedIncomingCount);
    if (settings.sgwTargetRoomId) {
      await this.sgwChatBridge.sendRoomMessage(settings.sgwTargetRoomId, { content });
      return {
        ok: true,
        reason: '지정된 SGW 대화방으로 테스트 알림을 전송했습니다.',
        target: settings.sgwTargetRoomId,
      };
    }
    await this.sgwChatBridge.sendDirectMessage({
      receiverId: sgwStatus.user.sub,
      receiverName: sgwStatus.user.name,
      content,
    });
    return {
      ok: true,
      reason: `내 SGW 계정(${sgwStatus.user.email}) 쪽지로 모바일 테스트 알림을 전송했습니다.`,
      target: sgwStatus.user.email,
    };
  }

  getStatus() {
    const settings = this.getSettings();
    const sgw = this.sgwChatBridge.getStatus();
    const pendingUnviewedCount = Array.from(this.alerts.values()).filter(
      (a) => !a.viewedByUser && !this.viewedIds.has(a.emailId)
    ).length;
    const digest =
      this.lastDigest ||
      (this.localDb.isReady()
        ? this.localDb.getState<ProactiveDashboardDigest>('proactive_dashboard_digest')
        : null);
    return {
      running: Boolean(this.pollTimer),
      settings,
      lastPollAt: this.lastPollAt,
      lastPollAddedCount: this.lastPollAddedCount,
      pendingUnviewedCount,
      totalSgwAlertsSent: this.totalSgwAlertsSent,
      sgwConnected: sgw.connected,
      sgwUserEmail: sgw.user?.email || null,
      digest,
      recentAlerts: Array.from(this.alerts.values())
        .sort((a, b) => b.arrivedAt - a.arrivedAt)
        .slice(0, 10)
        .map((a) => ({
          emailId: a.emailId,
          arrivedAt: new Date(a.arrivedAt).toISOString(),
          viewedByUser: a.viewedByUser,
          sgwNotified: a.sgwNotified,
          sender: a.summary.sender,
          subject: a.summary.subject,
          summary: a.summary.summary,
          urgency: a.summary.urgency,
        })),
    };
  }
}
