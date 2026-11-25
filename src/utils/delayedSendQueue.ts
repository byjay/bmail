export type DelayedSendStatus = 'queued' | 'sending' | 'sent' | 'cancelled' | 'failed';

export interface DelayedSendItem {
  id: string;
  to: string;
  cc: string;
  subject: string;
  body: string;
  kind: 'new' | 'reply' | 'replyAll' | 'forward' | 'worklog';
  queuedAt: number;
  sendAtMs: number;
  delaySeconds: number;
  status: DelayedSendStatus;
  errorMessage?: string;
}

export function createDelayedSendItem(
  payload: {
    to: string;
    cc: string;
    subject: string;
    body: string;
    kind: DelayedSendItem['kind'];
  },
  delaySeconds: number,
  nowMs = Date.now()
): DelayedSendItem {
  const safeDelay = Math.max(0, Math.min(600, Math.round(delaySeconds)));
  return {
    id: `dsq-${nowMs}-${Math.random().toString(36).slice(2, 7)}`,
    to: payload.to.trim(),
    cc: payload.cc.trim(),
    subject: payload.subject.trim() || '(제목 없음)',
    body: payload.body,
    kind: payload.kind,
    queuedAt: nowMs,
    sendAtMs: nowMs + safeDelay * 1000,
    delaySeconds: safeDelay,
    status: 'queued',
  };
}

export function getRemainingSeconds(item: DelayedSendItem, nowMs = Date.now()): number {
  if (item.status !== 'queued') return 0;
  const diff = item.sendAtMs - nowMs;
  if (diff <= 0) return 0;
  return Math.ceil(diff / 1000);
}

export function cancelDelayedSendItem(queue: DelayedSendItem[], id: string): {
  cancelled: DelayedSendItem | null;
  remainingQueue: DelayedSendItem[];
} {
  let cancelled: DelayedSendItem | null = null;
  const remainingQueue: DelayedSendItem[] = [];
  for (const item of queue) {
    if (item.id === id && item.status === 'queued') {
      cancelled = { ...item, status: 'cancelled' };
    } else {
      remainingQueue.push(item);
    }
  }
  return { cancelled, remainingQueue };
}

export function formatCountdownLabel(remainingSec: number): string {
  const s = Math.max(0, Math.round(remainingSec));
  if (s >= 60) {
    const mins = Math.floor(s / 60);
    const secs = s % 60;
    return `${mins}분 ${String(secs).padStart(2, '0')}초`;
  }
  return `${s}초`;
}
