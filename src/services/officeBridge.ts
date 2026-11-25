import type { MailItem, OutlookHostInfo } from '../types/mail';

export const TASKPANE_PIN_STORAGE_KEY = 'maillog_outlook_taskpane_pinned_v1';
const ROAMING_PIN_KEY = 'maillog_taskpane_pinned';

interface OfficeAsyncResult<T> {
  status: string;
  value: T;
  error?: { message?: string };
}

interface OfficeEmailAddress {
  displayName?: string;
  emailAddress?: string;
}

interface OfficeMailboxItem {
  itemId?: string;
  subject?: string;
  from?: OfficeEmailAddress;
  to?: OfficeEmailAddress[];
  dateTimeCreated?: Date;
  itemType?: string;
  body?: {
    getAsync: (coercionType: string, cb: (result: OfficeAsyncResult<string>) => void) => void;
    setSelectedDataAsync?: (
      data: string,
      options: { coercionType: string },
      cb?: (result: OfficeAsyncResult<void>) => void
    ) => void;
  };
}

interface OfficeRoamingSettings {
  get: (name: string) => unknown;
  set: (name: string, value: unknown) => void;
  saveAsync?: (cb?: (result: OfficeAsyncResult<void>) => void) => void;
}

interface OfficeGlobal {
  onReady: (cb?: (info: { host?: string; platform?: string }) => void) => Promise<{ host?: string; platform?: string }>;
  CoercionType?: { Text: string; Html: string };
  EventType?: {
    ItemChanged: string;
    SelectedItemsChanged?: string;
  };
  MailboxEnums?: { RestVersion?: { v2_0: string } };
  StartupBehavior?: { load: string; none: string };
  addin?: {
    setStartupBehavior?: (behavior: string) => Promise<void>;
    showAsTaskpane?: () => Promise<void>;
  };
  context?: {
    requirements?: {
      isSetSupported?: (name: string, minVersion?: string) => boolean;
    };
    roamingSettings?: OfficeRoamingSettings;
    mailbox?: {
      item?: OfficeMailboxItem;
      diagnostics?: { hostName?: string };
      userProfile?: { emailAddress?: string; displayName?: string };
      addHandlerAsync?: (
        eventType: string,
        handler: () => void,
        cb?: (result: OfficeAsyncResult<void>) => void
      ) => void;
      removeHandlerAsync?: (
        eventType: string,
        options?: { handler?: () => void },
        cb?: (result: OfficeAsyncResult<void>) => void
      ) => void;
      getSelectedItemsAsync?: (
        cb: (result: OfficeAsyncResult<Array<{ itemId: string; subject?: string }>>) => void
      ) => void;
      convertToRestId?: (itemId: string, restVersion: string) => string;
      displayNewMessageForm?: (params: {
        toRecipients?: string[];
        subject?: string;
        htmlBody?: string;
      }) => void;
    };
  };
}

declare global {
  interface Window {
    Office?: OfficeGlobal;
  }
}

export interface OfficeBridgeInitOptions {
  /** Reflects manifest.xml <SupportsPinning> & config.json plugin.supportsPinning */
  supportsPinning?: boolean;
  /** User default pin state (defaults to supportsPinning ?? true) */
  defaultPinned?: boolean;
  /** Fired whenever the user selects a different email while the taskpane stays pinned open */
  onItemChanged?: (mail: MailItem) => void;
  /** Fired on Mailbox 1.15 multi-select changes while taskpane stays pinned open */
  onSelectedItemsChanged?: (restIds: string[]) => void;
}

export interface OfficeBridgeInitResult {
  hostInfo: OutlookHostInfo;
  cleanup: () => void;
}

export function isTaskpaneRoute(): boolean {
  if (typeof window === 'undefined') return false;
  const pathname = window.location.pathname;
  const params = new URLSearchParams(window.location.search);
  return (
    pathname.startsWith('/taskpane') ||
    params.get('mode') === 'taskpane' ||
    params.has('_host_Info')
  );
}

/**
 * Loads the persisted Taskpane pinned preference (Office roamingSettings -> localStorage -> defaultPinned).
 * Defaults to true when manifest SupportsPinning is enabled so the taskpane stays open by default.
 */
export function loadTaskpanePinnedPreference(defaultPinned = true): boolean {
  if (typeof window === 'undefined') return defaultPinned;

  try {
    const roaming = window.Office?.context?.roamingSettings;
    if (roaming) {
      const roamingVal = roaming.get(ROAMING_PIN_KEY);
      if (typeof roamingVal === 'boolean') {
        return roamingVal;
      }
    }
  } catch {
    // Fallback to localStorage
  }

  try {
    const raw = window.localStorage.getItem(TASKPANE_PIN_STORAGE_KEY);
    if (raw === 'true') return true;
    if (raw === 'false') return false;
  } catch {
    // Storage blocked in restricted iframe
  }

  return defaultPinned;
}

/**
 * Persists the user's Taskpane pinned preference to both localStorage and
 * Outlook roamingSettings (when inside Outlook) so it roams across Desktop, New Outlook, and OWA.
 */
export async function saveTaskpanePinnedPreference(pinned: boolean): Promise<boolean> {
  if (typeof window === 'undefined') return false;

  try {
    window.localStorage.setItem(TASKPANE_PIN_STORAGE_KEY, pinned ? 'true' : 'false');
  } catch {
    // Ignore storage quota / iframe restriction
  }

  const roaming = window.Office?.context?.roamingSettings;
  if (!roaming) return true;

  try {
    roaming.set(ROAMING_PIN_KEY, pinned);
    if (roaming.saveAsync) {
      return await new Promise<boolean>((resolve) => {
        roaming.saveAsync?.((res) => resolve(res?.status === 'succeeded'));
      });
    }
    return true;
  } catch {
    return false;
  }
}

/**
 * Enables Shared-Runtime startup load & always-open behavior when Outlook supports it
 * and SupportsPinning is active in the manifest.
 */
export async function applyTaskpaneAlwaysOpenBehavior(enableAlwaysOpen: boolean): Promise<boolean> {
  if (typeof window === 'undefined') return false;
  const office = window.Office;
  if (!office?.addin) return false;

  try {
    if (office.addin.setStartupBehavior && office.StartupBehavior) {
      await office.addin.setStartupBehavior(
        enableAlwaysOpen ? office.StartupBehavior.load : office.StartupBehavior.none
      );
    }
    if (enableAlwaysOpen && office.addin.showAsTaskpane) {
      await office.addin.showAsTaskpane();
    }
    return true;
  } catch {
    // Not all Outlook hosts expose Office.addin.setStartupBehavior; SupportsPinning + ItemChanged still apply
    return false;
  }
}

function loadOfficeScript(timeoutMs = 3500): Promise<boolean> {
  if (typeof window === 'undefined') return Promise.resolve(false);
  if (window.Office) return Promise.resolve(true);

  return new Promise((resolve) => {
    const existing = document.querySelector<HTMLScriptElement>('script[data-office-js="1"]');
    if (existing) {
      const timer = window.setTimeout(() => resolve(Boolean(window.Office)), timeoutMs);
      existing.addEventListener('load', () => {
        window.clearTimeout(timer);
        resolve(Boolean(window.Office));
      });
      existing.addEventListener('error', () => {
        window.clearTimeout(timer);
        resolve(false);
      });
      return;
    }

    const script = document.createElement('script');
    script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    script.async = true;
    script.dataset.officeJs = '1';

    const timer = window.setTimeout(() => resolve(Boolean(window.Office)), timeoutMs);
    script.onload = () => {
      window.clearTimeout(timer);
      resolve(Boolean(window.Office));
    };
    script.onerror = () => {
      window.clearTimeout(timer);
      resolve(false);
    };
    document.head.appendChild(script);
  });
}

/**
 * Reads multi-selected email REST IDs in Outlook (Mailbox 1.15 SupportsMultiSelect).
 */
export async function getSelectedOutlookRestIds(): Promise<string[]> {
  const office = window.Office;
  const mailbox = office?.context?.mailbox;
  if (!office || !mailbox?.getSelectedItemsAsync) return [];

  return new Promise((resolve) => {
    mailbox.getSelectedItemsAsync?.((res) => {
      if (res.status !== 'succeeded' || !Array.isArray(res.value)) {
        resolve([]);
        return;
      }
      const isIos = mailbox.diagnostics?.hostName === 'OutlookIOS';
      const restVer = office.MailboxEnums?.RestVersion?.v2_0 ?? 'v2.0';
      const ids = res.value
        .map((item) => {
          if (isIos || !mailbox.convertToRestId) return item.itemId;
          try {
            return mailbox.convertToRestId(item.itemId, restVer);
          } catch {
            return item.itemId;
          }
        })
        .filter(Boolean);
      resolve(ids);
    });
  });
}

export async function readCurrentOutlookItem(): Promise<MailItem | null> {
  const office = window.Office;
  const item = office?.context?.mailbox?.item;
  if (!office || !item) return null;

  const userEmail = office.context?.mailbox?.userProfile?.emailAddress || '';
  const fromEmail = item.from?.emailAddress || userEmail || 'unknown@outlook.com';
  const fromName = item.from?.displayName || fromEmail;
  const isSent = Boolean(userEmail && fromEmail.toLowerCase() === userEmail.toLowerCase());

  const created = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
  const dateStr = created.toISOString().slice(0, 10);
  const timeLabel = created.toTimeString().slice(0, 5);

  const bodyText = await new Promise<string>((resolve) => {
    if (!item.body?.getAsync) {
      resolve('');
      return;
    }
    const coercion = office.CoercionType?.Text ?? 'text';
    item.body.getAsync(coercion, (res) => {
      if (res.status === 'succeeded') {
        resolve(res.value || '');
      } else {
        resolve('');
      }
    });
  });

  return {
    id: item.itemId || `office-item-${Date.now()}`,
    folder: isSent ? 'sent' : 'inbox',
    subject: item.subject || '(제목 없음)',
    fromName,
    fromEmail,
    to: (item.to ?? []).map((r) => r.emailAddress || r.displayName || '').filter(Boolean),
    date: dateStr,
    timestamp: created.toISOString(),
    timeLabel,
    snippet: bodyText.replace(/\s+/g, ' ').trim().slice(0, 140),
    body: bodyText.trim() || '(본문 텍스트를 불러오지 못했습니다)',
    isRead: true,
    source: 'outlook-live',
  };
}

/**
 * Initializes the Office.js Add-in bridge with SupportsPinning persistence,
 * Mailbox 1.5+ ItemChanged & Mailbox 1.15 SelectedItemsChanged listeners,
 * and automatic always-open startup behavior so the Taskpane stays open across email switches.
 */
export async function initializeOfficeBridge(
  optionsOrHandler?: OfficeBridgeInitOptions | ((mail: MailItem) => void)
): Promise<OfficeBridgeInitResult> {
  const opts: OfficeBridgeInitOptions =
    typeof optionsOrHandler === 'function'
      ? { onItemChanged: optionsOrHandler }
      : optionsOrHandler ?? {};

  const supportsPinning = opts.supportsPinning !== false;
  const effectivePinned = supportsPinning
    ? loadTaskpanePinnedPreference(opts.defaultPinned ?? true)
    : false;
  const noopCleanup = () => {};

  if (!isTaskpaneRoute()) {
    return {
      hostInfo: {
        isInsideOutlook: false,
        hostName: 'Web Preview',
        supportsPinning,
        pinningSupported: supportsPinning,
        isPinned: effectivePinned,
      },
      cleanup: noopCleanup,
    };
  }

  const loaded = await loadOfficeScript();
  if (!loaded || !window.Office?.onReady) {
    return {
      hostInfo: {
        isInsideOutlook: false,
        hostName: 'Taskpane Preview (Browser)',
        supportsPinning,
        pinningSupported: supportsPinning,
        isPinned: effectivePinned,
      },
      cleanup: noopCleanup,
    };
  }

  try {
    const info = await Promise.race([
      window.Office.onReady(),
      new Promise<{ host?: string }>((resolve) => window.setTimeout(() => resolve({}), 2500)),
    ]);

    const hostName = info?.host || '';
    const isOutlook = Boolean(hostName && hostName.toLowerCase().includes('outlook'));
    const mailbox = window.Office.context?.mailbox;

    if (isOutlook && mailbox) {
      const pinningSupported = Boolean(
        supportsPinning &&
          (window.Office.context?.requirements?.isSetSupported?.('Mailbox', '1.5') ?? true)
      );

      // Re-evaluate persisted preference after Office.onReady (roamingSettings is now ready)
      const isPinned = pinningSupported
        ? loadTaskpanePinnedPreference(opts.defaultPinned ?? true)
        : false;

      // Persist default pinned preference so future Outlook launches keep taskpane pinned
      if (pinningSupported) {
        void saveTaskpanePinnedPreference(isPinned);
        void applyTaskpaneAlwaysOpenBehavior(isPinned);
      }

      const current = await readCurrentOutlookItem();
      if (current && opts.onItemChanged) {
        opts.onItemChanged(current);
      }

      const itemChangedHandler = () => {
        void readCurrentOutlookItem().then((updated) => {
          if (updated && opts.onItemChanged) {
            opts.onItemChanged(updated);
          }
        });
      };

      const selectedItemsChangedHandler = () => {
        void getSelectedOutlookRestIds().then((ids) => {
          if (opts.onSelectedItemsChanged) {
            opts.onSelectedItemsChanged(ids);
          }
        });
      };

      const itemChangedEvent = window.Office.EventType?.ItemChanged;
      const selectedItemsEvent = window.Office.EventType?.SelectedItemsChanged;

      // Register ItemChanged so the pinned taskpane continuously updates as user browses emails
      if (mailbox.addHandlerAsync && itemChangedEvent && opts.onItemChanged) {
        mailbox.addHandlerAsync(itemChangedEvent, itemChangedHandler);
      }

      // Register SelectedItemsChanged (Mailbox 1.15 SupportsMultiSelect) while pinned
      if (mailbox.addHandlerAsync && selectedItemsEvent && opts.onSelectedItemsChanged) {
        mailbox.addHandlerAsync(selectedItemsEvent, selectedItemsChangedHandler);
      }

      const cleanup = () => {
        if (mailbox.removeHandlerAsync && itemChangedEvent && opts.onItemChanged) {
          mailbox.removeHandlerAsync(itemChangedEvent, { handler: itemChangedHandler });
        }
        if (mailbox.removeHandlerAsync && selectedItemsEvent && opts.onSelectedItemsChanged) {
          mailbox.removeHandlerAsync(selectedItemsEvent, { handler: selectedItemsChangedHandler });
        }
      };

      return {
        hostInfo: {
          isInsideOutlook: true,
          hostName: info.host || 'Outlook',
          userEmail: mailbox.userProfile?.emailAddress,
          currentItemSubject: current?.subject,
          supportsPinning,
          pinningSupported,
          isPinned,
        },
        cleanup,
      };
    }

    return {
      hostInfo: {
        isInsideOutlook: false,
        hostName: 'Taskpane Preview (Browser)',
        supportsPinning,
        pinningSupported: supportsPinning,
        isPinned: effectivePinned,
      },
      cleanup: noopCleanup,
    };
  } catch {
    return {
      hostInfo: {
        isInsideOutlook: false,
        hostName: 'Taskpane Preview (Browser)',
        supportsPinning,
        pinningSupported: supportsPinning,
        isPinned: effectivePinned,
      },
      cleanup: noopCleanup,
    };
  }
}

export function exportDraftToOutlook(subject: string, plainTextBody: string): {
  deliveredVia: 'outlook-compose' | 'mailto';
} {
  const mailbox = window.Office?.context?.mailbox;
  if (mailbox?.displayNewMessageForm) {
    const escapedHtml = plainTextBody
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\n/g, '<br/>');
    mailbox.displayNewMessageForm({
      subject,
      htmlBody: `<div style="font-family:'Malgun Gothic',sans-serif;line-height:1.6;">${escapedHtml}</div>`,
    });
    return { deliveredVia: 'outlook-compose' };
  }

  const mailto = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(
    plainTextBody.slice(0, 1800)
  )}`;
  window.location.href = mailto;
  return { deliveredVia: 'mailto' };
}
