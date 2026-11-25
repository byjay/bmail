import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

export interface BackupSnapshotRecord {
  filename: string;
  fullPath: string;
  createdAt: string;
  sizeBytes: number;
  sha256: string;
  reason: 'daily-auto' | 'manual';
}

export interface BackupSchedulerStatus {
  backupDir: string;
  dbExists: boolean;
  dbSizeBytes: number;
  maxRetention: number;
  intervalHours: number;
  snapshots: BackupSnapshotRecord[];
  lastBackupAt: string | null;
  lastError: string | null;
}

export interface LocalDbCheckpointLike {
  isReady: () => boolean;
  getRawDb?: () => { exec: (sql: string) => void } | null;
}

export class LocalBackupScheduler {
  private readonly dataDir: string;
  private readonly dbPath: string;
  private readonly backupDir: string;
  private readonly maxRetention: number;
  private readonly intervalMs: number;
  private readonly dbLike?: LocalDbCheckpointLike;
  private timer: NodeJS.Timeout | null = null;
  private lastError: string | null = null;

  constructor(
    dataDir: string,
    dbLike?: LocalDbCheckpointLike,
    options?: { maxRetention?: number; intervalHours?: number; dbFilename?: string }
  ) {
    this.dataDir = path.resolve(dataDir);
    this.dbPath = path.join(this.dataDir, options?.dbFilename || 'bmail-local.sqlite');
    this.backupDir = path.join(this.dataDir, 'backups');
    this.maxRetention = Math.max(1, options?.maxRetention ?? 7);
    const hours = Math.max(1, options?.intervalHours ?? 24);
    this.intervalMs = hours * 3600 * 1000;
    this.dbLike = dbLike;
  }

  public listSnapshots(): BackupSnapshotRecord[] {
    if (!fs.existsSync(this.backupDir)) return [];
    const entries = fs.readdirSync(this.backupDir, { withFileTypes: true });
    const list: BackupSnapshotRecord[] = [];
    for (const ent of entries) {
      if (!ent.isFile() || !ent.name.endsWith('.sqlite.gz')) continue;
      const fullPath = path.join(this.backupDir, ent.name);
      try {
        const st = fs.statSync(fullPath);
        const shaPath = `${fullPath}.sha256`;
        const sha256 = fs.existsSync(shaPath)
          ? fs.readFileSync(shaPath, 'utf8').trim().slice(0, 64)
          : '';
        list.push({
          filename: ent.name,
          fullPath,
          createdAt: st.mtime.toISOString(),
          sizeBytes: st.size,
          sha256,
          reason: ent.name.includes('-manual-') ? 'manual' : 'daily-auto',
        });
      } catch {
        // ignore stat errors
      }
    }
    return list.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  public getStatus(): BackupSchedulerStatus {
    const dbExists = fs.existsSync(this.dbPath);
    const dbSizeBytes = dbExists ? fs.statSync(this.dbPath).size : 0;
    const snapshots = this.listSnapshots();
    return {
      backupDir: this.backupDir,
      dbExists,
      dbSizeBytes,
      maxRetention: this.maxRetention,
      intervalHours: Math.round(this.intervalMs / (3600 * 1000)),
      snapshots,
      lastBackupAt: snapshots[0]?.createdAt || null,
      lastError: this.lastError,
    };
  }

  public runSnapshotNow(reason: 'daily-auto' | 'manual' = 'manual'): BackupSnapshotRecord | null {
    try {
      if (!fs.existsSync(this.dbPath)) {
        this.lastError = '로컬 SQLite 데이터베이스 파일이 아직 생성되지 않았습니다.';
        return null;
      }

      // Checkpoint WAL cleanly before reading DB file
      try {
        const raw = this.dbLike?.getRawDb?.();
        raw?.exec('PRAGMA wal_checkpoint(PASSIVE);');
      } catch {
        // non-fatal if read-only or locked
      }

      fs.mkdirSync(this.backupDir, { recursive: true });
      const now = new Date();
      const stamp = `${now.toISOString().slice(0, 10)}_${now.toISOString().slice(11, 19).replace(/:/g, '')}`;
      const filename = `bmail-backup-${reason}-${stamp}.sqlite.gz`;
      const fullPath = path.join(this.backupDir, filename);

      const rawDbBytes = fs.readFileSync(this.dbPath);
      const gzBytes = zlib.gzipSync(rawDbBytes, { level: 6 });
      const sha256 = crypto.createHash('sha256').update(gzBytes).digest('hex');

      fs.writeFileSync(fullPath, gzBytes);
      fs.writeFileSync(`${fullPath}.sha256`, `${sha256}  ${filename}\n`, 'utf8');

      this.pruneOldSnapshots();
      this.lastError = null;
      return {
        filename,
        fullPath,
        createdAt: now.toISOString(),
        sizeBytes: gzBytes.length,
        sha256,
        reason,
      };
    } catch (err) {
      this.lastError = err instanceof Error ? err.message : String(err);
      return null;
    }
  }

  private pruneOldSnapshots(): void {
    const snapshots = this.listSnapshots();
    if (snapshots.length <= this.maxRetention) return;
    const toRemove = snapshots.slice(this.maxRetention);
    for (const item of toRemove) {
      try {
        fs.unlinkSync(item.fullPath);
        if (fs.existsSync(`${item.fullPath}.sha256`)) {
          fs.unlinkSync(`${item.fullPath}.sha256`);
        }
      } catch {
        // ignore unlink errors
      }
    }
  }

  public runDailyIfNeeded(): BackupSnapshotRecord | null {
    if (!fs.existsSync(this.dbPath)) return null;
    const todayPrefix = new Date().toISOString().slice(0, 10);
    const existing = this.listSnapshots();
    const hasToday = existing.some((s) => s.createdAt.slice(0, 10) === todayPrefix);
    if (hasToday) return null;
    return this.runSnapshotNow('daily-auto');
  }

  public startDailySchedule(): void {
    if (this.timer) return;
    // Defer initial check by 15 seconds so server startup remains instantaneous
    const warmup = setTimeout(() => {
      this.runDailyIfNeeded();
    }, 15_000);
    warmup.unref();

    this.timer = setInterval(() => {
      this.runDailyIfNeeded();
    }, this.intervalMs);
    this.timer.unref();
  }

  public stopSchedule(): void {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
    }
  }
}
