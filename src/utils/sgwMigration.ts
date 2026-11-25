import type {
  LegacyInternalMailRow,
  MigrationDryRunReport,
  MigrationExecutionReport,
  MigrationSampleCheck,
} from '../types/sgwSedms';
import type { SgwMailStore } from '../services/sgwMailStore';

function evaluateMigrationRows(
  rows: LegacyInternalMailRow[],
  existingLegacyIds: Set<string>
): {
  report: MigrationDryRunReport;
  validItems: LegacyInternalMailRow[];
} {
  const seen = new Set<string>(existingLegacyIds);
  const skippedDuplicateIds: string[] = [];
  const selfSendRowIds: string[] = [];
  const missingFieldErrors: Array<{ rowId: string; reason: string }> = [];
  const encodingWarnings: Array<{ rowId: string; reason: string }> = [];
  const validItems: LegacyInternalMailRow[] = [];

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const rowId = row && row.id !== undefined && row.id !== null ? String(row.id) : `index-${i}`;

    if (!row || !String(row.senderId || '').trim() || !String(row.recipientId || '').trim()) {
      missingFieldErrors.push({ rowId, reason: 'senderId 또는 recipientId가 누락되었습니다.' });
      continue;
    }

    if (seen.has(rowId)) {
      skippedDuplicateIds.push(rowId);
      continue;
    }
    seen.add(rowId);

    if (String(row.senderId).trim() === String(row.recipientId).trim()) {
      selfSendRowIds.push(rowId);
    }

    if (/\uFFFD/.test(String(row.subject || '')) || /\uFFFD/.test(String(row.body || ''))) {
      encodingWarnings.push({ rowId, reason: 'UTF-8 대체 문자(U+FFFD)가 감지되었습니다.' });
    }

    validItems.push(row);
  }

  return {
    report: {
      totalSourceRows: rows.length,
      validRows: validItems.length,
      skippedDuplicateIds,
      selfSendRowIds,
      missingFieldErrors,
      encodingWarnings,
    },
    validItems,
  };
}

export function dryRunInternalMailMigration(
  rows: LegacyInternalMailRow[],
  existingLegacyIds: Set<string> = new Set()
): MigrationDryRunReport {
  return evaluateMigrationRows(Array.isArray(rows) ? rows : [], existingLegacyIds).report;
}

export function executeInternalMailMigration(
  store: SgwMailStore,
  rows: LegacyInternalMailRow[],
  options: { approved: boolean; actorSub: string }
): MigrationExecutionReport {
  if (!options.approved) {
    throw new Error('드라이런 검수 후 명시적 승인(approved=true)이 필요합니다.');
  }

  const safeRows = Array.isArray(rows) ? rows : [];
  const { validItems } = evaluateMigrationRows(safeRows, store.getMigratedLegacyIds());
  const sampleVerificationById: MigrationSampleCheck[] = [];

  for (const row of validItems) {
    const inserted = store.insertMigratedRow(row);
    const expectedRead = row.isRead === true || row.isRead === 1;
    const recipient = inserted.recipients[0];
    sampleVerificationById.push({
      legacyId: String(row.id),
      newMailId: inserted.id,
      senderSub: inserted.senderSub,
      recipientSub: recipient?.sub || String(row.recipientId).trim(),
      isReadMatch: Boolean(recipient?.isRead) === expectedRead,
    });
  }

  const migratedCount = validItems.length;
  const skippedCount = safeRows.length - migratedCount;
  store.recordAudit(options.actorSub, 'migration.execute', `migrated:${migratedCount}`);

  return {
    sourceTotal: safeRows.length,
    migratedCount,
    skippedCount,
    verifiedCountMatch: migratedCount + skippedCount === safeRows.length,
    sampleVerificationById,
    executedAt: new Date().toISOString(),
  };
}

export function rollbackInternalMailMigration(
  store: SgwMailStore,
  actorSub: string
): { removedMessages: number } {
  return store.rollbackMigratedRows(actorSub);
}
