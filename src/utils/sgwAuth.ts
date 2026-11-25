import crypto from 'node:crypto';
import type { SgwEmployee, SgwUserIdentity } from '../types/sgwSedms';

export interface SgwSsoTokenPayload {
  sub: string;
  name: string;
  department: string;
  email: string;
  role: string;
  exp: number;
  nonce: string;
}

const usedNonces = new Map<string, number>();

export function clearSgwNonceCacheForTest(): void {
  usedNonces.clear();
}

function base64UrlEncode(input: Buffer | string): string {
  return Buffer.from(input).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function base64UrlDecode(input: string): string {
  const padded = input.replace(/-/g, '+').replace(/_/g, '/');
  const padLen = (4 - (padded.length % 4)) % 4;
  return Buffer.from(padded + '='.repeat(padLen), 'base64').toString('utf8');
}

function timingSafeHexEqual(a: string, b: string): boolean {
  const bufA = Buffer.from(a, 'utf8');
  const bufB = Buffer.from(b, 'utf8');
  if (bufA.length !== bufB.length) return false;
  return crypto.timingSafeEqual(bufA, bufB);
}

export function signSgwSsoToken(payload: SgwSsoTokenPayload, secret: string): string {
  if (!secret.trim()) {
    throw new Error('SGW_SSO_SECRET가 설정되지 않았습니다.');
  }
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const sig = crypto.createHmac('sha256', secret).update(encodedPayload).digest('hex');
  return `${encodedPayload}.${sig}`;
}

export function verifySgwSsoToken(
  token: string,
  secret: string,
  options?: { nowSec?: number; enforceNonceReplay?: boolean }
): { ok: boolean; identity?: SgwUserIdentity; error?: string } {
  if (!secret || !secret.trim()) {
    return { ok: false, error: 'SGW_SSO_SECRET 환경변수가 설정되지 않았습니다.' };
  }
  const trimmed = String(token || '').trim();
  const dotIdx = trimmed.lastIndexOf('.');
  if (dotIdx <= 0 || dotIdx === trimmed.length - 1) {
    return { ok: false, error: '유효하지 않은 SGW SSO 토큰 형식입니다.' };
  }

  const encodedPayload = trimmed.slice(0, dotIdx);
  const providedSig = trimmed.slice(dotIdx + 1);
  const expectedSig = crypto.createHmac('sha256', secret).update(encodedPayload).digest('hex');

  if (!timingSafeHexEqual(providedSig, expectedSig)) {
    return { ok: false, error: 'SGW SSO 토큰 서명이 일치하지 않습니다.' };
  }

  let parsed: Partial<SgwSsoTokenPayload>;
  try {
    parsed = JSON.parse(base64UrlDecode(encodedPayload)) as Partial<SgwSsoTokenPayload>;
  } catch {
    return { ok: false, error: 'SGW SSO 토큰 페이로드를 파싱할 수 없습니다.' };
  }

  const nowSec = options?.nowSec ?? Math.floor(Date.now() / 1000);
  if (typeof parsed.exp !== 'number' || !Number.isFinite(parsed.exp) || parsed.exp <= nowSec) {
    return { ok: false, error: 'SGW SSO 토큰이 만료되었습니다.' };
  }

  if (!parsed.sub || !parsed.nonce) {
    return { ok: false, error: 'SGW SSO 토큰에 필수 클레임(sub/nonce)이 누락되었습니다.' };
  }

  const enforceReplay = options?.enforceNonceReplay !== false;
  if (enforceReplay) {
    for (const [k, exp] of usedNonces.entries()) {
      if (exp <= nowSec) usedNonces.delete(k);
    }
    if (usedNonces.has(parsed.nonce)) {
      return { ok: false, error: 'SGW SSO 토큰 nonce 재사용 차단됨' };
    }
    usedNonces.set(parsed.nonce, parsed.exp);
  }

  return {
    ok: true,
    identity: {
      sub: String(parsed.sub),
      name: String(parsed.name || parsed.sub),
      department: String(parsed.department || ''),
      email: String(parsed.email || ''),
      role: String(parsed.role || 'user'),
      authSource: 'sgw-sso',
    },
  };
}

export function signAttachmentDownloadToken(params: {
  attachmentId: string;
  ownerSub: string;
  expiresAtSec: number;
  secret: string;
}): string {
  const payload = `${params.attachmentId}|${params.ownerSub}|${params.expiresAtSec}`;
  const sig = crypto.createHmac('sha256', params.secret).update(payload).digest('hex');
  return `${base64UrlEncode(payload)}.${sig}`;
}

export function verifyAttachmentDownloadToken(params: {
  token: string;
  attachmentId: string;
  requesterSub: string;
  secret: string;
  nowSec?: number;
}): boolean {
  if (!params.token || !params.secret) return false;
  const dotIdx = params.token.lastIndexOf('.');
  if (dotIdx <= 0) return false;

  const encoded = params.token.slice(0, dotIdx);
  const sig = params.token.slice(dotIdx + 1);
  let payload = '';
  try {
    payload = base64UrlDecode(encoded);
  } catch {
    return false;
  }

  const expectedSig = crypto.createHmac('sha256', params.secret).update(payload).digest('hex');
  if (!timingSafeHexEqual(sig, expectedSig)) return false;

  const [attId, ownerSub, expStr] = payload.split('|');
  const exp = Number.parseInt(expStr || '', 10);
  const nowSec = params.nowSec ?? Math.floor(Date.now() / 1000);
  if (!Number.isFinite(exp) || exp <= nowSec) return false;

  return attId === params.attachmentId && ownerSub === params.requesterSub;
}

export class ShortTtlEmployeeCache {
  private readonly ttlMs: number;
  private items: SgwEmployee[] | null = null;
  private cachedAt = 0;

  constructor(ttlMs = 60_000) {
    this.ttlMs = Math.max(1_000, ttlMs);
  }

  set(items: SgwEmployee[], nowMs = Date.now()): void {
    this.items = items.map((e) => ({ ...e }));
    this.cachedAt = nowMs;
  }

  get(nowMs = Date.now()): SgwEmployee[] | null {
    if (!this.items) return null;
    if (nowMs - this.cachedAt > this.ttlMs) {
      this.items = null;
      return null;
    }
    return this.items.map((e) => ({ ...e }));
  }

  clear(): void {
    this.items = null;
    this.cachedAt = 0;
  }
}

export async function fetchSgwAuthMe(params: {
  sgwBaseUrl: string;
  authHeader?: string;
  cookieHeader?: string;
}): Promise<SgwUserIdentity | null> {
  const base = params.sgwBaseUrl.replace(/\/+$/, '');
  const headers: Record<string, string> = { Accept: 'application/json' };
  if (params.authHeader) headers.Authorization = params.authHeader;
  if (params.cookieHeader) headers.Cookie = params.cookieHeader;

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 5_000);
  try {
    const res = await fetch(`${base}/auth/me`, { headers, signal: controller.signal });
    if (!res.ok) return null;
    const data = (await res.json()) as Partial<SgwUserIdentity> & { id?: string };
    const sub = String(data.sub || data.id || '').trim();
    if (!sub) return null;
    return {
      sub,
      name: String(data.name || sub),
      department: String(data.department || ''),
      email: String(data.email || ''),
      role: String(data.role || 'user'),
      authSource: 'sgw-me',
    };
  } catch {
    return null;
  } finally {
    clearTimeout(timer);
  }
}

export async function fetchSgwEmployees(params: {
  sgwBaseUrl: string;
  authHeader?: string;
  cookieHeader?: string;
}): Promise<SgwEmployee[]> {
  const base = params.sgwBaseUrl.replace(/\/+$/, '');
  const headers: Record<string, string> = { Accept: 'application/json' };
  if (params.authHeader) headers.Authorization = params.authHeader;
  if (params.cookieHeader) headers.Cookie = params.cookieHeader;

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8_000);
  try {
    const res = await fetch(`${base}/api/employees`, { headers, signal: controller.signal });
    if (!res.ok) {
      throw new Error(`SGW 직원 목록 조회 실패 (HTTP ${res.status})`);
    }
    const data = (await res.json()) as unknown;
    const rawList = Array.isArray(data)
      ? data
      : data && typeof data === 'object' && Array.isArray((data as { employees?: unknown[] }).employees)
        ? (data as { employees: unknown[] }).employees
        : [];
    return rawList
      .filter((item): item is Record<string, unknown> => Boolean(item && typeof item === 'object'))
      .map((item) => ({
        sub: String(item.sub || item.id || '').trim(),
        name: String(item.name || item.sub || item.id || '').trim(),
        department: typeof item.department === 'string' ? item.department : undefined,
        email: typeof item.email === 'string' ? item.email : undefined,
        title: typeof item.title === 'string' ? item.title : undefined,
      }))
      .filter((emp) => Boolean(emp.sub));
  } finally {
    clearTimeout(timer);
  }
}
