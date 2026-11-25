/**
 * Dedicated child-process PST/OST & Outlook migration worker.
 * Runs outside server.ts's Node event loop so HTTP/UI responsiveness stays <10ms
 * while indexing tens of gigabytes of local Outlook PST archives into SQLite WAL.
 */
import { LocalMailDb } from './localMailDb';
import { migrateLocalOutlookToDb, type MigrationProgressSummary } from './localOutlookMigrator';

function emitProgress(progress: MigrationProgressSummary): void {
  process.stdout.write(`__BMAIL_MIGRATION_PROGRESS__:${JSON.stringify(progress)}\n`);
}

function main(): void {
  const dataDir = process.env.BMAIL_DATA_DIR?.trim();
  if (!dataDir) {
    emitProgress({
      status: 'failed',
      startedAt: new Date().toISOString(),
      completedAt: new Date().toISOString(),
      storesProcessed: 0,
      totalEmailsIndexed: 0,
      totalContactsIndexed: 0,
      totalFoldersIndexed: 0,
      totalSignaturesIndexed: 0,
      discoveredProjectsCount: 0,
      error: 'BMAIL_DATA_DIR 환경변수가 지정되지 않았습니다.',
    });
    process.exit(1);
  }

  let extraPstPaths: string[] | undefined;
  const rawExtra = process.env.BMAIL_EXTRA_PST_PATHS_JSON?.trim();
  if (rawExtra) {
    try {
      const parsed = JSON.parse(rawExtra) as unknown;
      if (Array.isArray(parsed)) {
        extraPstPaths = parsed.filter((p): p is string => typeof p === 'string' && p.trim().length > 0);
      }
    } catch {
      // ignore invalid extra paths JSON
    }
  }

  const db = new LocalMailDb(dataDir);
  if (!db.isReady()) {
    emitProgress({
      status: 'failed',
      startedAt: new Date().toISOString(),
      completedAt: new Date().toISOString(),
      storesProcessed: 0,
      totalEmailsIndexed: 0,
      totalContactsIndexed: 0,
      totalFoldersIndexed: 0,
      totalSignaturesIndexed: 0,
      discoveredProjectsCount: 0,
      error: db.getInitError() || '로컬 SQLite DB 초기화 실패',
    });
    process.exit(1);
  }

  try {
    const finalSummary = migrateLocalOutlookToDb(db, {
      extraPstPaths,
      onProgress: (next) => {
        emitProgress(next);
      },
    });
    emitProgress(finalSummary);
    process.exit(finalSummary.status === 'failed' ? 1 : 0);
  } catch (err) {
    emitProgress({
      status: 'failed',
      startedAt: new Date().toISOString(),
      completedAt: new Date().toISOString(),
      storesProcessed: 0,
      totalEmailsIndexed: 0,
      totalContactsIndexed: 0,
      totalFoldersIndexed: 0,
      totalSignaturesIndexed: 0,
      discoveredProjectsCount: 0,
      error: err instanceof Error ? err.message : 'PST 워커 처리 중 알 수 없는 오류',
    });
    process.exit(1);
  }
}

main();
