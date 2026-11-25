import fs from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import type { EmailFolder, MailAttachmentMeta, MailItem } from '../types/mail';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import type { OutlookContactRecord } from './outlookLocalBridge';

interface SqliteStatementLike {
  run: (...params: unknown[]) => unknown;
  get: (...params: unknown[]) => Record<string, unknown> | undefined;
  all: (...params: unknown[]) => Array<Record<string, unknown>>;
}

interface SqliteDatabaseLike {
  exec: (sql: string) => void;
  prepare: (sql: string) => SqliteStatementLike;
  close: () => void;
}

interface NodeSqliteModuleLike {
  DatabaseSync: new (filePath: string) => SqliteDatabaseLike;
}

export interface LocalEmailRecord extends MailItem {
  storeName?: string;
  pstPath?: string;
  descriptorNodeId?: number;
  isDeleted?: boolean;
}

export interface LocalFolderRecord {
  storeName: string;
  folderPath: string;
  classifiedFolder: EmailFolder;
  itemCount: number;
  unreadCount?: number;
  isSystem: boolean;
}

export interface LocalSignatureRecord {
  name: string;
  htmlContent: string;
  textContent: string;
  isDefaultNew: boolean;
  isDefaultReply: boolean;
}

export interface LocalAccountRecord {
  email: string;
  displayName: string;
  providerPreset: string;
  protocol: 'pop3' | 'imap';
  pop3Host: string;
  pop3Port: number;
  pop3Tls: boolean;
  imapHost: string;
  imapPort: number;
  imapTls: boolean;
  smtpHost: string;
  smtpPort: number;
  smtpTls: boolean;
  leaveOnServer: boolean;
  deliveryStorePath: string;
}

export interface LocalBootstrapPayload {
  available: boolean;
  migratedAt: string | null;
  totalEmails: number;
  totalContacts: number;
  accounts: LocalAccountRecord[];
  signatures: LocalSignatureRecord[];
  activeSignatureName: string | null;
  folders: LocalFolderRecord[];
  pstPaths: string[];
  contacts: OutlookContactRecord[];
  projects: ProjectDefinition[];
  emails: MailItem[];
  kvState: Record<string, unknown>;
}

const ALLOWED_KV_STATE_KEYS = new Set([
  'outlookPrefs',
  'projects',
  'drafts',
  'calendar',
  'favoriteFolders',
  'pinnedEmails',
  'projectSubfolders',
  'projectMailOverrides',
  'mail_action_states',
  'mail_chat_histories',
  'outlook_window_settings',
  'outlook_categories',
  'outlook_rules',
  'outlook_favorites',
  'saved_executive_reports',
  'pst_sync_watermarks',
  'sticky_notes',
  'outlook_tasks',
  'mail_categories_map',
  'mail_snooze_map',
  'proactive_alert_settings',
  'proactive_dashboard_digest',
]);

function loadNodeSqlite(): NodeSqliteModuleLike | null {
  try {
    const req = createRequire(import.meta.url);
    const mod = req('node:sqlite') as NodeSqliteModuleLike;
    if (mod && typeof mod.DatabaseSync === 'function') {
      return mod;
    }
  } catch {
    // Node < 22.5 or runtime without node:sqlite
  }
  return null;
}

function safeParseJson<T>(raw: unknown, fallback: T): T {
  if (typeof raw !== 'string' || !raw.trim()) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

function buildFtsQuery(rawQuery: string): string {
  const tokens = rawQuery
    .trim()
    .split(/\s+/)
    .map((t) => t.replace(/"/g, '').trim())
    .filter(Boolean);
  if (tokens.length === 0) return '';
  return tokens.map((t) => `"${t}"*`).join(' AND ');
}

function rowToMailItem(row: Record<string, unknown>): MailItem {
  const attachments = safeParseJson<MailAttachmentMeta[] | undefined>(row.attachments_json, undefined);
  const toList = safeParseJson<string[]>(row.to_json, []);
  return {
    id: String(row.id || ''),
    folder: row.folder === 'sent' ? 'sent' : 'inbox',
    subject: String(row.subject || '(제목 없음)'),
    fromName: String(row.from_name || ''),
    fromEmail: String(row.from_email || ''),
    to: Array.isArray(toList) ? toList : [],
    date: String(row.date || ''),
    timestamp: String(row.timestamp || ''),
    timeLabel: String(row.time_label || ''),
    snippet: String(row.snippet || ''),
    body: String(row.body || ''),
    hasAttachment: Boolean(row.has_attachment),
    attachments: Array.isArray(attachments) && attachments.length > 0 ? attachments : undefined,
    conversationKey: row.conversation_key ? String(row.conversation_key) : undefined,
    importance: row.importance === 'high' ? 'high' : 'normal',
    category: row.category ? String(row.category) : undefined,
    folderName: row.folder_name ? String(row.folder_name) : undefined,
    storeName: row.store_name ? String(row.store_name) : undefined,
    isRead: Boolean(row.is_read),
    isFlagged: Boolean(row.is_flagged),
    source:
      row.source === 'graph-live'
        ? 'graph-live'
        : row.source === 'outlook-live'
          ? 'outlook-live'
          : 'pst-import',
  };
}

export class LocalMailDb {
  private readonly dbPath: string;
  private db: SqliteDatabaseLike | null = null;
  private initError: string | null = null;
  private initialized = false;

  constructor(baseDir: string) {
    const resolvedDir = path.resolve(baseDir);
    this.dbPath = path.join(resolvedDir, 'bmail-local.sqlite');
  }

  public getDbPath(): string {
    return this.dbPath;
  }

  public getInitError(): string | null {
    return this.initError;
  }

  public isReady(): boolean {
    return this.ensureOpen() !== null;
  }

  private ensureOpen(): SqliteDatabaseLike | null {
    if (this.initialized) return this.db;
    this.initialized = true;

    const sqliteMod = loadNodeSqlite();
    if (!sqliteMod) {
      this.initError = 'node:sqlite 런타임이 사용 불가능합니다 (Node 24+ 필요).';
      return null;
    }

    try {
      fs.mkdirSync(path.dirname(this.dbPath), { recursive: true });
      const db = new sqliteMod.DatabaseSync(this.dbPath);
      db.exec('PRAGMA busy_timeout = 15000;');
      db.exec('PRAGMA journal_mode = WAL;');
      db.exec('PRAGMA synchronous = NORMAL;');
      db.exec('PRAGMA temp_store = MEMORY;');
      this.createSchema(db);
      this.db = db;
      return db;
    } catch (err) {
      this.initError = err instanceof Error ? err.message : String(err);
      this.db = null;
      return null;
    }
  }

  private createSchema(db: SqliteDatabaseLike): void {
    db.exec(`
      CREATE TABLE IF NOT EXISTS emails (
        id TEXT PRIMARY KEY,
        store_name TEXT,
        pst_path TEXT,
        descriptor_node_id INTEGER,
        folder TEXT NOT NULL,
        folder_name TEXT,
        subject TEXT NOT NULL,
        from_name TEXT NOT NULL,
        from_email TEXT NOT NULL,
        to_json TEXT NOT NULL,
        date TEXT NOT NULL,
        timestamp TEXT NOT NULL,
        time_label TEXT NOT NULL,
        snippet TEXT NOT NULL,
        body TEXT NOT NULL,
        has_attachment INTEGER NOT NULL DEFAULT 0,
        attachments_json TEXT,
        conversation_key TEXT,
        importance TEXT NOT NULL DEFAULT 'normal',
        category TEXT,
        is_read INTEGER NOT NULL DEFAULT 1,
        is_flagged INTEGER NOT NULL DEFAULT 0,
        is_deleted INTEGER NOT NULL DEFAULT 0,
        source TEXT NOT NULL DEFAULT 'pst-import'
      );

      CREATE INDEX IF NOT EXISTS idx_emails_timestamp ON emails(timestamp DESC);
      CREATE INDEX IF NOT EXISTS idx_emails_folder ON emails(folder);
      CREATE INDEX IF NOT EXISTS idx_emails_folder_name ON emails(folder_name);
      CREATE INDEX IF NOT EXISTS idx_emails_store_name ON emails(store_name);

      CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts5(
        id UNINDEXED,
        subject,
        from_name,
        from_email,
        body,
        folder_name
      );

      CREATE TABLE IF NOT EXISTS folders (
        store_name TEXT NOT NULL,
        folder_path TEXT NOT NULL,
        classified_folder TEXT NOT NULL,
        item_count INTEGER NOT NULL DEFAULT 0,
        is_system INTEGER NOT NULL DEFAULT 0,
        PRIMARY KEY(store_name, folder_path)
      );

      CREATE TABLE IF NOT EXISTS contacts (
        id TEXT PRIMARY KEY,
        full_name TEXT NOT NULL,
        email TEXT NOT NULL,
        company TEXT,
        department TEXT,
        job_title TEXT,
        phone TEXT,
        source TEXT NOT NULL DEFAULT 'outlook-pc'
      );

      CREATE TABLE IF NOT EXISTS signatures (
        name TEXT PRIMARY KEY,
        html_content TEXT NOT NULL,
        text_content TEXT NOT NULL,
        is_default_new INTEGER NOT NULL DEFAULT 0,
        is_default_reply INTEGER NOT NULL DEFAULT 0
      );

      CREATE TABLE IF NOT EXISTS accounts (
        email TEXT PRIMARY KEY,
        display_name TEXT NOT NULL,
        provider_preset TEXT NOT NULL,
        protocol TEXT NOT NULL,
        pop3_host TEXT,
        pop3_port INTEGER,
        pop3_tls INTEGER NOT NULL DEFAULT 1,
        imap_host TEXT,
        imap_port INTEGER,
        imap_tls INTEGER NOT NULL DEFAULT 1,
        smtp_host TEXT,
        smtp_port INTEGER,
        smtp_tls INTEGER NOT NULL DEFAULT 1,
        leave_on_server INTEGER NOT NULL DEFAULT 1,
        delivery_store_path TEXT
      );

      CREATE TABLE IF NOT EXISTS kv_state (
        key TEXT PRIMARY KEY,
        value_json TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS email_rich_cache (
        id TEXT PRIMARY KEY,
        body_html TEXT NOT NULL,
        attachments_json TEXT,
        updated_at TEXT NOT NULL
      );
    `);
  }

  public getRawDb(): SqliteDatabaseLike | null {
    return this.ensureOpen();
  }

  public close(): void {
    if (this.db) {
      try {
        this.db.close();
      } catch {
        // ignore
      }
      this.db = null;
      this.initialized = false;
    }
  }

  public upsertEmailsBatch(records: LocalEmailRecord[]): number {
    if (records.length === 0) return 0;
    const db = this.ensureOpen();
    if (!db) throw new Error(this.initError || 'SQLite DB unavailable');

    const insertEmail = db.prepare(`
      INSERT INTO emails (
        id, store_name, pst_path, descriptor_node_id, folder, folder_name,
        subject, from_name, from_email, to_json, date, timestamp, time_label,
        snippet, body, has_attachment, attachments_json, conversation_key,
        importance, category, is_read, is_flagged, is_deleted, source
      ) VALUES (
        ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?, ?
      )
      ON CONFLICT(id) DO UPDATE SET
        store_name = excluded.store_name,
        pst_path = COALESCE(excluded.pst_path, emails.pst_path),
        descriptor_node_id = COALESCE(excluded.descriptor_node_id, emails.descriptor_node_id),
        folder = excluded.folder,
        folder_name = excluded.folder_name,
        subject = excluded.subject,
        from_name = excluded.from_name,
        from_email = excluded.from_email,
        to_json = excluded.to_json,
        date = excluded.date,
        timestamp = excluded.timestamp,
        time_label = excluded.time_label,
        snippet = excluded.snippet,
        body = excluded.body,
        has_attachment = excluded.has_attachment,
        attachments_json = excluded.attachments_json,
        conversation_key = excluded.conversation_key,
        importance = excluded.importance,
        category = excluded.category,
        is_read = CASE WHEN excluded.source = 'outlook-live' THEN excluded.is_read ELSE emails.is_read END,
        is_flagged = CASE WHEN excluded.source = 'outlook-live' THEN excluded.is_flagged ELSE emails.is_flagged END,
        source = excluded.source
    `);
    const deleteFts = db.prepare('DELETE FROM emails_fts WHERE id = ?');
    const insertFts = db.prepare(
      'INSERT INTO emails_fts (id, subject, from_name, from_email, body, folder_name) VALUES (?, ?, ?, ?, ?, ?)'
    );

    db.exec('BEGIN');
    try {
      for (const item of records) {
        const toJson = JSON.stringify(Array.isArray(item.to) ? item.to : []);
        const attJson =
          Array.isArray(item.attachments) && item.attachments.length > 0
            ? JSON.stringify(item.attachments)
            : null;
        const folderName = item.folderName || item.category || '';
        insertEmail.run(
          item.id,
          item.storeName || null,
          item.pstPath || null,
          typeof item.descriptorNodeId === 'number' ? item.descriptorNodeId : null,
          item.folder === 'sent' ? 'sent' : 'inbox',
          folderName || null,
          item.subject || '(제목 없음)',
          item.fromName || '',
          item.fromEmail || '',
          toJson,
          item.date || '',
          item.timestamp || '',
          item.timeLabel || '',
          item.snippet || '',
          item.body || '',
          item.hasAttachment ? 1 : 0,
          attJson,
          item.conversationKey || null,
          item.importance === 'high' ? 'high' : 'normal',
          item.category || null,
          item.isRead ? 1 : 0,
          item.isFlagged ? 1 : 0,
          item.isDeleted ? 1 : 0,
          item.source || 'pst-import'
        );
        deleteFts.run(item.id);
        insertFts.run(
          item.id,
          item.subject || '',
          item.fromName || '',
          item.fromEmail || '',
          item.body || '',
          folderName
        );
      }
      db.exec('COMMIT');
      return records.length;
    } catch (err) {
      try {
        db.exec('ROLLBACK');
      } catch {
        // ignore
      }
      throw err;
    }
  }

  public queryEmails(params: {
    q?: string;
    folder?: string;
    store?: string;
    includeDeleted?: boolean;
    deletedOnly?: boolean;
    limit?: number;
    offset?: number;
  }): { total: number; items: MailItem[] } {
    const db = this.ensureOpen();
    if (!db) return { total: 0, items: [] };

    const limit = Math.max(1, Math.min(1000, Number(params.limit) || 200));
    const offset = Math.max(0, Number(params.offset) || 0);
    const rawQ = (params.q || '').trim();
    const folder = (params.folder || '').trim();
    const store = (params.store || '').trim();
    const isTrashFolderQuery =
      folder === '지운 편지함' ||
      folder.startsWith('지운 편지함 /') ||
      folder.toLowerCase() === 'deleted items' ||
      folder.toLowerCase().startsWith('deleted items /');
    const includeDeleted = Boolean(params.includeDeleted) || isTrashFolderQuery;

    const clauses: string[] = [];
    const args: unknown[] = [];

    if (params.deletedOnly) {
      clauses.push('e.is_deleted = 1');
    } else if (!includeDeleted) {
      clauses.push('e.is_deleted = 0');
    }
    if (folder === 'inbox' || folder === 'sent') {
      clauses.push('e.folder = ?');
      args.push(folder);
    } else if (folder && folder !== 'all') {
      clauses.push('e.folder_name = ?');
      args.push(folder);
    }
    if (store) {
      clauses.push('e.store_name = ?');
      args.push(store);
    }

    if (rawQ) {
      const ftsExpr = buildFtsQuery(rawQ);
      if (ftsExpr) {
        try {
          const ftsWhere = [...clauses, 'emails_fts MATCH ?'].join(' AND ');
          const countRow = db
            .prepare(
              `SELECT COUNT(*) AS cnt FROM emails e JOIN emails_fts ON e.id = emails_fts.id WHERE ${ftsWhere}`
            )
            .get(...args, ftsExpr);
          const rows = db
            .prepare(
              `SELECT e.* FROM emails e JOIN emails_fts ON e.id = emails_fts.id WHERE ${ftsWhere} ORDER BY e.timestamp DESC LIMIT ? OFFSET ?`
            )
            .all(...args, ftsExpr, limit, offset);
          return {
            total: Number(countRow?.cnt || 0),
            items: rows.map(rowToMailItem),
          };
        } catch {
          // Fall back to LIKE below if FTS syntax cannot parse input
        }
      }

      const likeVal = `%${rawQ}%`;
      const likeWhere = [
        ...clauses,
        '(e.subject LIKE ? OR e.from_name LIKE ? OR e.from_email LIKE ? OR e.body LIKE ? OR COALESCE(e.folder_name, \'\') LIKE ?)',
      ].join(' AND ');
      const likeArgs = [...args, likeVal, likeVal, likeVal, likeVal, likeVal];
      const countRow = db
        .prepare(`SELECT COUNT(*) AS cnt FROM emails e WHERE ${likeWhere}`)
        .get(...likeArgs);
      const rows = db
        .prepare(
          `SELECT e.* FROM emails e WHERE ${likeWhere} ORDER BY e.timestamp DESC LIMIT ? OFFSET ?`
        )
        .all(...likeArgs, limit, offset);
      return {
        total: Number(countRow?.cnt || 0),
        items: rows.map(rowToMailItem),
      };
    }

    const whereSql = clauses.length > 0 ? `WHERE ${clauses.join(' AND ')}` : '';
    const countRow = db.prepare(`SELECT COUNT(*) AS cnt FROM emails e ${whereSql}`).get(...args);
    const rows = db
      .prepare(`SELECT e.* FROM emails e ${whereSql} ORDER BY e.timestamp DESC LIMIT ? OFFSET ?`)
      .all(...args, limit, offset);

    return {
      total: Number(countRow?.cnt || 0),
      items: rows.map(rowToMailItem),
    };
  }

  /** Pure read-only lookup of unread inbox messages without mutating `is_read`. */
  public queryUnreadInboxEmails(limit = 300): MailItem[] {
    const db = this.ensureOpen();
    if (!db) return [];
    const safeLimit = Math.max(1, Math.min(1000, Number(limit) || 300));
    const rows = db
      .prepare(
        `SELECT e.* FROM emails e
         WHERE e.is_deleted = 0 AND e.folder = 'inbox' AND e.is_read = 0
         ORDER BY e.timestamp DESC
         LIMIT ?`
      )
      .all(safeLimit);
    return rows.map(rowToMailItem);
  }

  public getEmailById(
    id: string
  ): (MailItem & { pstPath: string | null; descriptorNodeId: number | null; storeName: string | null }) | null {
    const db = this.ensureOpen();
    if (!db) return null;
    const row = db.prepare('SELECT * FROM emails WHERE id = ?').get(id);
    if (!row) return null;
    return {
      ...rowToMailItem(row),
      pstPath: row.pst_path ? String(row.pst_path) : null,
      descriptorNodeId:
        typeof row.descriptor_node_id === 'number' ? row.descriptor_node_id : null,
      storeName: row.store_name ? String(row.store_name) : null,
    };
  }

  public patchEmailState(
    id: string,
    patch: { isRead?: boolean; isFlagged?: boolean; isDeleted?: boolean }
  ): boolean {
    const db = this.ensureOpen();
    if (!db) return false;
    const updates: string[] = [];
    const params: unknown[] = [];
    if (typeof patch.isRead === 'boolean') {
      updates.push('is_read = ?');
      params.push(patch.isRead ? 1 : 0);
    }
    if (typeof patch.isFlagged === 'boolean') {
      updates.push('is_flagged = ?');
      params.push(patch.isFlagged ? 1 : 0);
    }
    if (typeof patch.isDeleted === 'boolean') {
      updates.push('is_deleted = ?');
      params.push(patch.isDeleted ? 1 : 0);
    }
    if (updates.length === 0) return false;
    params.push(id);
    const res = db
      .prepare(`UPDATE emails SET ${updates.join(', ')} WHERE id = ?`)
      .run(...params) as { changes?: number };
    return Number(res?.changes || 0) > 0;
  }

  public upsertFolders(folders: LocalFolderRecord[]): void {
    if (folders.length === 0) return;
    const db = this.ensureOpen();
    if (!db) return;
    const stmt = db.prepare(`
      INSERT INTO folders (store_name, folder_path, classified_folder, item_count, is_system)
      VALUES (?, ?, ?, ?, ?)
      ON CONFLICT(store_name, folder_path) DO UPDATE SET
        classified_folder = excluded.classified_folder,
        item_count = excluded.item_count,
        is_system = excluded.is_system
    `);
    db.exec('BEGIN');
    try {
      for (const f of folders) {
        stmt.run(
          f.storeName,
          f.folderPath,
          f.classifiedFolder,
          f.itemCount,
          f.isSystem ? 1 : 0
        );
      }
      db.exec('COMMIT');
    } catch (err) {
      db.exec('ROLLBACK');
      throw err;
    }
  }

  public upsertContacts(contacts: OutlookContactRecord[]): void {
    if (contacts.length === 0) return;
    const db = this.ensureOpen();
    if (!db) return;
    const stmt = db.prepare(`
      INSERT INTO contacts (id, full_name, email, company, department, job_title, phone, source)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        full_name = excluded.full_name,
        email = excluded.email,
        company = COALESCE(excluded.company, contacts.company),
        department = COALESCE(excluded.department, contacts.department),
        job_title = COALESCE(excluded.job_title, contacts.job_title),
        phone = COALESCE(excluded.phone, contacts.phone),
        source = excluded.source
    `);
    db.exec('BEGIN');
    try {
      for (const c of contacts) {
        stmt.run(
          c.id,
          c.fullName,
          c.email,
          c.company || null,
          c.department || null,
          c.jobTitle || null,
          c.phone || null,
          c.source || 'outlook-pc'
        );
      }
      db.exec('COMMIT');
    } catch (err) {
      db.exec('ROLLBACK');
      throw err;
    }
  }

  public upsertSignatures(signatures: LocalSignatureRecord[]): void {
    if (signatures.length === 0) return;
    const db = this.ensureOpen();
    if (!db) return;
    const stmt = db.prepare(`
      INSERT INTO signatures (name, html_content, text_content, is_default_new, is_default_reply)
      VALUES (?, ?, ?, ?, ?)
      ON CONFLICT(name) DO UPDATE SET
        html_content = excluded.html_content,
        text_content = excluded.text_content,
        is_default_new = excluded.is_default_new,
        is_default_reply = excluded.is_default_reply
    `);
    db.exec('BEGIN');
    try {
      for (const s of signatures) {
        stmt.run(
          s.name,
          s.htmlContent,
          s.textContent,
          s.isDefaultNew ? 1 : 0,
          s.isDefaultReply ? 1 : 0
        );
      }
      db.exec('COMMIT');
    } catch (err) {
      db.exec('ROLLBACK');
      throw err;
    }
  }

  public upsertAccounts(accounts: LocalAccountRecord[]): void {
    if (accounts.length === 0) return;
    const db = this.ensureOpen();
    if (!db) return;
    const stmt = db.prepare(`
      INSERT INTO accounts (
        email, display_name, provider_preset, protocol,
        pop3_host, pop3_port, pop3_tls,
        imap_host, imap_port, imap_tls,
        smtp_host, smtp_port, smtp_tls,
        leave_on_server, delivery_store_path
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(email) DO UPDATE SET
        display_name = excluded.display_name,
        provider_preset = excluded.provider_preset,
        protocol = excluded.protocol,
        pop3_host = excluded.pop3_host,
        pop3_port = excluded.pop3_port,
        pop3_tls = excluded.pop3_tls,
        imap_host = excluded.imap_host,
        imap_port = excluded.imap_port,
        imap_tls = excluded.imap_tls,
        smtp_host = excluded.smtp_host,
        smtp_port = excluded.smtp_port,
        smtp_tls = excluded.smtp_tls,
        leave_on_server = excluded.leave_on_server,
        delivery_store_path = excluded.delivery_store_path
    `);
    db.exec('BEGIN');
    try {
      for (const a of accounts) {
        stmt.run(
          a.email,
          a.displayName,
          a.providerPreset,
          a.protocol,
          a.pop3Host || null,
          a.pop3Port || null,
          a.pop3Tls ? 1 : 0,
          a.imapHost || null,
          a.imapPort || null,
          a.imapTls ? 1 : 0,
          a.smtpHost || null,
          a.smtpPort || null,
          a.smtpTls ? 1 : 0,
          a.leaveOnServer ? 1 : 0,
          a.deliveryStorePath || null
        );
      }
      db.exec('COMMIT');
    } catch (err) {
      db.exec('ROLLBACK');
      throw err;
    }
  }

  public getState<T = unknown>(key: string): T | null {
    const db = this.ensureOpen();
    if (!db) return null;
    const row = db.prepare('SELECT value_json FROM kv_state WHERE key = ?').get(key);
    if (!row) return null;
    return safeParseJson<T | null>(row.value_json, null);
  }

  public setState(key: string, value: unknown): void {
    const db = this.ensureOpen();
    if (!db) throw new Error(this.initError || 'SQLite DB unavailable');
    db.prepare(`
      INSERT INTO kv_state (key, value_json, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(key) DO UPDATE SET
        value_json = excluded.value_json,
        updated_at = excluded.updated_at
    `).run(key, JSON.stringify(value), new Date().toISOString());
  }

  public isAllowedStateKey(key: string): boolean {
    return ALLOWED_KV_STATE_KEYS.has(key);
  }

  public getUserState<T = unknown>(key: string): T | null {
    if (!ALLOWED_KV_STATE_KEYS.has(key)) {
      throw new Error(`허용되지 않은 상태 키입니다: ${key}`);
    }
    return this.getState<T>(key);
  }

  public setUserState(key: string, value: unknown): void {
    if (!ALLOWED_KV_STATE_KEYS.has(key)) {
      throw new Error(`허용되지 않은 상태 키입니다: ${key}`);
    }
    this.setState(key, value);
  }

  public getEmailRichCache(
    id: string
  ): { bodyHtml: string; attachments: MailAttachmentMeta[] } | null {
    const db = this.ensureOpen();
    if (!db) return null;
    const row = db
      .prepare('SELECT body_html, attachments_json FROM email_rich_cache WHERE id = ?')
      .get(id);
    if (!row) return null;
    return {
      bodyHtml: String(row.body_html || ''),
      attachments: row.attachments_json
        ? safeParseJson<MailAttachmentMeta[]>(row.attachments_json, [])
        : [],
    };
  }

  public setEmailRichCache(
    id: string,
    bodyHtml: string,
    attachments: MailAttachmentMeta[] = []
  ): void {
    const db = this.ensureOpen();
    if (!db) return;
    db.prepare(`
      INSERT INTO email_rich_cache (id, body_html, attachments_json, updated_at)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        body_html = excluded.body_html,
        attachments_json = excluded.attachments_json,
        updated_at = excluded.updated_at
    `).run(
      id,
      bodyHtml,
      attachments.length > 0 ? JSON.stringify(attachments) : null,
      new Date().toISOString()
    );
  }

  public queryEmailsWithAttachments(filters: {
    storeName?: string;
    folderPaths?: string[];
    emailIds?: string[];
    fromDate?: string;
    toDate?: string;
    includeDeleted?: boolean;
    limit?: number;
  }): Array<MailItem & { pstPath: string | null; descriptorNodeId: number | null; storeName: string | null }> {
    const db = this.ensureOpen();
    if (!db) return [];

    const limit = Math.max(1, Math.min(5000, Number(filters.limit) || 2000));
    const clauses: string[] = ['has_attachment = 1'];
    const args: unknown[] = [];

    if (!filters.includeDeleted) {
      clauses.push('is_deleted = 0');
    }
    if (filters.storeName && filters.storeName.trim()) {
      clauses.push('store_name = ?');
      args.push(filters.storeName.trim());
    }
    if (Array.isArray(filters.emailIds) && filters.emailIds.length > 0) {
      const cleanIds = filters.emailIds.slice(0, 500);
      clauses.push(`id IN (${cleanIds.map(() => '?').join(', ')})`);
      args.push(...cleanIds);
    } else if (Array.isArray(filters.folderPaths) && filters.folderPaths.length > 0) {
      const folderSubClauses: string[] = [];
      for (const rawPath of filters.folderPaths.slice(0, 40)) {
        const clean = rawPath.trim();
        if (!clean) continue;
        folderSubClauses.push('(folder_name = ? OR folder_name LIKE ?)');
        args.push(clean, `${clean} / %`);
      }
      if (folderSubClauses.length > 0) {
        clauses.push(`(${folderSubClauses.join(' OR ')})`);
      }
    }
    if (filters.fromDate && /^\d{4}-\d{2}-\d{2}$/.test(filters.fromDate)) {
      clauses.push('date >= ?');
      args.push(filters.fromDate);
    }
    if (filters.toDate && /^\d{4}-\d{2}-\d{2}$/.test(filters.toDate)) {
      clauses.push('date <= ?');
      args.push(filters.toDate);
    }

    const rows = db
      .prepare(`SELECT * FROM emails WHERE ${clauses.join(' AND ')} ORDER BY timestamp DESC LIMIT ?`)
      .all(...args, limit);

    return rows.map((row) => ({
      ...rowToMailItem(row),
      pstPath: row.pst_path ? String(row.pst_path) : null,
      descriptorNodeId:
        typeof row.descriptor_node_id === 'number' ? row.descriptor_node_id : null,
      storeName: row.store_name ? String(row.store_name) : null,
    }));
  }

  public getExistingEmailIds(storeName?: string): Set<string> {
    const db = this.ensureOpen();
    if (!db) return new Set();
    const rows = storeName && storeName.trim()
      ? (db.prepare('SELECT id FROM emails WHERE store_name = ?').all(storeName.trim()) as Array<{ id: string }>)
      : (db.prepare('SELECT id FROM emails').all() as Array<{ id: string }>);
    const out = new Set<string>();
    for (const r of rows) {
      if (r.id) out.add(String(r.id));
    }
    return out;
  }

  public queryAllEmailsForAnalytics(filters?: {
    dateFrom?: string;
    dateTo?: string;
    direction?: 'all' | 'received' | 'sent';
    limit?: number;
  }): MailItem[] {
    const db = this.ensureOpen();
    if (!db) return [];

    const clauses: string[] = ['is_deleted = 0'];
    const args: unknown[] = [];

    if (filters?.dateFrom && /^\d{4}-\d{2}-\d{2}$/.test(filters.dateFrom)) {
      clauses.push('date >= ?');
      args.push(filters.dateFrom);
    }
    if (filters?.dateTo && /^\d{4}-\d{2}-\d{2}$/.test(filters.dateTo)) {
      clauses.push('date <= ?');
      args.push(filters.dateTo);
    }
    if (filters?.direction === 'received') {
      clauses.push("folder = 'inbox'");
    } else if (filters?.direction === 'sent') {
      clauses.push("folder = 'sent'");
    }

    const limitSql =
      typeof filters?.limit === 'number' && filters.limit > 0
        ? ` LIMIT ${Math.min(100_000, Math.floor(filters.limit))}`
        : '';

    const rows = db
      .prepare(
        `SELECT id, folder, folder_name, store_name, subject, from_name, from_email, to_json, date, timestamp, time_label, snippet, SUBSTR(body, 1, 1200) AS body, has_attachment, attachments_json, conversation_key, importance, category, is_read, is_flagged, source FROM emails WHERE ${clauses.join(' AND ')} ORDER BY timestamp DESC${limitSql}`
      )
      .all(...args);

    return rows.map(rowToMailItem);
  }

  public getAllContacts(): OutlookContactRecord[] {
    const db = this.ensureOpen();
    if (!db) return [];
    const rows = db.prepare('SELECT * FROM contacts ORDER BY full_name ASC').all();
    return rows.map((r) => ({
      id: String(r.id),
      fullName: String(r.full_name),
      email: String(r.email),
      company: r.company ? String(r.company) : undefined,
      department: r.department ? String(r.department) : undefined,
      jobTitle: r.job_title ? String(r.job_title) : undefined,
      phone: r.phone ? String(r.phone) : undefined,
      source:
        r.source === 'pst-derived'
          ? 'pst-derived'
          : r.source === 'manual'
            ? 'manual'
            : 'outlook-pc',
    }));
  }

  public getFolders(options?: {
    nonSystemOnly?: boolean;
    minItemCount?: number;
    limit?: number;
  }): LocalFolderRecord[] {
    const db = this.ensureOpen();
    if (!db) return [];
    const clauses: string[] = [];
    const args: unknown[] = [];
    if (options?.nonSystemOnly) {
      clauses.push('is_system = 0');
    }
    if (typeof options?.minItemCount === 'number' && options.minItemCount > 0) {
      clauses.push('item_count >= ?');
      args.push(options.minItemCount);
    }
    const whereSql = clauses.length > 0 ? `WHERE ${clauses.join(' AND ')}` : '';
    const limitSql =
      typeof options?.limit === 'number' && options.limit > 0
        ? ` LIMIT ${Math.min(1000, Math.floor(options.limit))}`
        : '';
    const rows = db
      .prepare(`SELECT * FROM folders ${whereSql} ORDER BY item_count DESC, folder_path ASC${limitSql}`)
      .all(...args);

    const unreadRows = db
      .prepare(
        `SELECT COALESCE(store_name, '') AS store_name, COALESCE(folder_name, '') AS folder_name, COUNT(*) AS unread_cnt
         FROM emails WHERE is_deleted = 0 AND is_read = 0
         GROUP BY store_name, folder_name`
      )
      .all() as Array<{ store_name: string; folder_name: string; unread_cnt: number }>;
    const unreadMap = new Map<string, number>();
    for (const u of unreadRows) {
      unreadMap.set(`${u.store_name}::${u.folder_name}`, Number(u.unread_cnt || 0));
    }

    return rows.map((f) => {
      const storeName = String(f.store_name);
      const folderPath = String(f.folder_path);
      return {
        storeName,
        folderPath,
        classifiedFolder: f.classified_folder === 'sent' ? 'sent' : 'inbox',
        itemCount: Number(f.item_count || 0),
        unreadCount: unreadMap.get(`${storeName}::${folderPath}`) ?? 0,
        isSystem: Boolean(f.is_system),
      };
    });
  }

  public isAllowedPstPath(candidatePath: string): boolean {
    if (!candidatePath) return false;
    const resolvedCandidate = path.resolve(candidatePath).toLowerCase();
    const pstPaths = this.getState<string[]>('pstPaths') || [];
    return pstPaths.some((p) => path.resolve(p).toLowerCase() === resolvedCandidate);
  }

  public getBootstrapPayload(initialLimit = 300): LocalBootstrapPayload {
    const db = this.ensureOpen();
    if (!db) {
      return {
        available: false,
        migratedAt: null,
        totalEmails: 0,
        totalContacts: 0,
        accounts: [],
        signatures: [],
        activeSignatureName: null,
        folders: [],
        pstPaths: [],
        contacts: [],
        projects: [],
        emails: [],
        kvState: {},
      };
    }

    const totalEmailsRow = db
      .prepare('SELECT COUNT(*) AS cnt FROM emails WHERE is_deleted = 0')
      .get();
    const totalContactsRow = db.prepare('SELECT COUNT(*) AS cnt FROM contacts').get();

    const accountRows = db.prepare('SELECT * FROM accounts').all();
    const accounts: LocalAccountRecord[] = accountRows.map((a) => ({
      email: String(a.email),
      displayName: String(a.display_name || a.email),
      providerPreset: String(a.provider_preset || 'mailplug'),
      protocol: a.protocol === 'imap' ? 'imap' : 'pop3',
      pop3Host: String(a.pop3_host || ''),
      pop3Port: Number(a.pop3_port || 995),
      pop3Tls: Boolean(a.pop3_tls),
      imapHost: String(a.imap_host || ''),
      imapPort: Number(a.imap_port || 993),
      imapTls: Boolean(a.imap_tls),
      smtpHost: String(a.smtp_host || ''),
      smtpPort: Number(a.smtp_port || 465),
      smtpTls: Boolean(a.smtp_tls),
      leaveOnServer: Boolean(a.leave_on_server),
      deliveryStorePath: String(a.delivery_store_path || ''),
    }));

    const sigRows = db.prepare('SELECT * FROM signatures ORDER BY is_default_new DESC, name ASC').all();
    const signatures: LocalSignatureRecord[] = sigRows.map((s) => ({
      name: String(s.name),
      htmlContent: String(s.html_content || ''),
      textContent: String(s.text_content || ''),
      isDefaultNew: Boolean(s.is_default_new),
      isDefaultReply: Boolean(s.is_default_reply),
    }));
    const activeSig =
      signatures.find((s) => s.isDefaultNew) || signatures.find((s) => s.isDefaultReply) || signatures[0];

    const folders = this.getFolders();

    const emailQuery = this.queryEmails({ limit: initialLimit, offset: 0 });
    const contacts = this.getAllContacts().slice(0, 1000);
    const pstPaths = this.getState<string[]>('pstPaths') || [];
    const projects = this.getState<ProjectDefinition[]>('projects') || [];
    const meta = this.getState<{ migratedAt?: string }>('migration_meta');

    const kvState: Record<string, unknown> = {};
    for (const k of ALLOWED_KV_STATE_KEYS) {
      const val = this.getState(k);
      if (val !== null) kvState[k] = val;
    }

    return {
      available: true,
      migratedAt: meta?.migratedAt || null,
      totalEmails: Number(totalEmailsRow?.cnt || 0),
      totalContacts: Number(totalContactsRow?.cnt || 0),
      accounts,
      signatures,
      activeSignatureName: activeSig ? activeSig.name : null,
      folders,
      pstPaths,
      contacts,
      projects,
      emails: emailQuery.items,
      kvState,
    };
  }
}
