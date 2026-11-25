import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { LocalMailDb } from '../src/services/localMailDb';
import { migrateLocalOutlookToDb } from '../src/services/localOutlookMigrator';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const APP_ROOT = path.resolve(__dirname, '..');
const DATA_DIR = path.resolve(APP_ROOT, 'data');

const startedMs = Date.now();
const db = new LocalMailDb(DATA_DIR);

if (!db.isReady()) {
  console.error('[Bmail Migration] SQLite DB unavailable:', db.getInitError());
  process.exit(1);
}

try {
  let lastLogged = 0;
  const summary = migrateLocalOutlookToDb(db, {
    onProgress: (p) => {
      if (p.totalEmailsIndexed - lastLogged >= 2000 || p.status === 'completed') {
        lastLogged = p.totalEmailsIndexed;
        console.log(
          `[Bmail Migration] status=${p.status} emails=${p.totalEmailsIndexed} contacts=${p.totalContactsIndexed} folders=${p.totalFoldersIndexed}`
        );
      }
    },
  });

  const elapsedSec = ((Date.now() - startedMs) / 1000).toFixed(2);
  const dbSizeMb = fs.existsSync(db.getDbPath())
    ? (fs.statSync(db.getDbPath()).size / (1024 * 1024)).toFixed(2)
    : '0.00';

  console.log(
    JSON.stringify(
      {
        status: summary.status,
        storesProcessed: summary.storesProcessed,
        totalEmailsIndexed: summary.totalEmailsIndexed,
        totalContactsIndexed: summary.totalContactsIndexed,
        totalFoldersIndexed: summary.totalFoldersIndexed,
        totalSignaturesIndexed: summary.totalSignaturesIndexed,
        discoveredProjectsCount: summary.discoveredProjectsCount,
        elapsedSeconds: Number(elapsedSec),
        sqliteSizeMb: Number(dbSizeMb),
        dbPath: db.getDbPath(),
      },
      null,
      2
    )
  );
} finally {
  db.close();
}
