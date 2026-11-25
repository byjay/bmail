import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  EmvVaultService,
  maskEmvValue,
  normalizeEmployeeId,
} from './emvVaultService';

test('normalizeEmployeeId accepts valid corporate emails/staff IDs and blocks traversal or prototype keys', () => {
  assert.equal(normalizeEmployeeId('Hong.Gildong@Company.co.kr '), 'hong.gildong@company.co.kr');
  assert.equal(normalizeEmployeeId('emp_2026-01'), 'emp_2026-01');
  assert.equal(normalizeEmployeeId('__proto__'), '');
  assert.equal(normalizeEmployeeId('constructor'), '');
  assert.equal(normalizeEmployeeId('../etc/passwd'), '');
  assert.equal(normalizeEmployeeId('user\r\n@evil.com'), '');
  assert.equal(maskEmvValue('GEMINI_API_KEY', 'TEST_KEY_1234567890ABCD'), 'TEST••••ABCD');
  assert.equal(maskEmvValue('GEMINI_MODEL', 'gemini-4-hydrogen'), 'gemini-4-hydrogen');
});

test('EmvVaultService encrypts at rest, isolates per-employee keys, and enforces strict resolution precedence', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-emv-test-'));
  try {
    const env: NodeJS.ProcessEnv = {
      EMV_MASTER_KEY: 'super-strong-test-master-secret-999',
      CONFIG_ADMIN_PASSWORD: 'admin-pass-test-2026',
      GEMINI_API_KEY: 'ENV_GEMINI_FALLBACK_KEY_0001',
      OPENAI_API_KEY: 'ENV_OPENAI_KEY_0002',
    };

    const vault = new EmvVaultService({ baseDir: tmpDir, env });

    // Admin auth requires matching CONFIG_ADMIN_PASSWORD
    assert.throws(() => vault.authenticateAdmin('wrong-pw'), /일치하지 않습니다/);
    const session = vault.authenticateAdmin('admin-pass-test-2026');
    assert.ok(vault.verifyAdminSession(session.token));

    // Save org-wide global Gemini key and employee-specific Gemini + Groq keys
    const globalSaved = vault.setEntry({
      scope: 'global',
      key: 'GEMINI_API_KEY',
      value: 'GLOBAL_ORG_GEMINI_KEY_7777',
    });
    assert.equal(globalSaved.scope, 'global');
    assert.ok(!globalSaved.maskedPreview.includes('ORG_GEMINI_KEY'));

    vault.setEntry({
      scope: 'employee',
      employeeId: 'alice@shipyard.co.kr',
      key: 'GEMINI_API_KEY',
      value: 'ALICE_PERSONAL_GEMINI_KEY_8888',
    });
    vault.setEntry({
      scope: 'employee',
      employeeId: 'alice@shipyard.co.kr',
      key: 'GROQ_API_KEY',
      value: 'ALICE_GROQ_KEY_9999',
    });

    // Verify encrypted vault file NEVER contains plaintext secrets
    const rawFile = fs.readFileSync(path.join(tmpDir, 'emv-vault.enc.json'), 'utf8');
    assert.ok(!rawFile.includes('GLOBAL_ORG_GEMINI_KEY_7777'));
    assert.ok(!rawFile.includes('ALICE_PERSONAL_GEMINI_KEY_8888'));
    assert.ok(!rawFile.includes('ALICE_GROQ_KEY_9999'));

    // Reload from disk in a fresh instance to verify AES-256-GCM decryption works across restarts
    const reloaded = new EmvVaultService({ baseDir: tmpDir, env });

    // 1) Verified employee Alice gets her personal Gemini + Groq keys and falls back to env OpenAI
    const aliceEnv = reloaded.resolveRuntimeEnv('alice@shipyard.co.kr');
    assert.equal(aliceEnv.GEMINI_API_KEY, 'ALICE_PERSONAL_GEMINI_KEY_8888');
    assert.equal(aliceEnv.GROQ_API_KEY, 'ALICE_GROQ_KEY_9999');
    assert.equal(aliceEnv.OPENAI_API_KEY, 'ENV_OPENAI_KEY_0002');

    // 2) Verified employee Bob (no personal override) gets Global EMV Gemini key, NOT Alice's key
    const bobEnv = reloaded.resolveRuntimeEnv('bob@shipyard.co.kr');
    assert.equal(bobEnv.GEMINI_API_KEY, 'GLOBAL_ORG_GEMINI_KEY_7777');
    assert.equal(bobEnv.GROQ_API_KEY, undefined);
    assert.equal(bobEnv.OPENAI_API_KEY, 'ENV_OPENAI_KEY_0002');

    // 3) Unauthenticated/null caller NEVER receives any employee key
    const anonEnv = reloaded.resolveRuntimeEnv(null);
    assert.equal(anonEnv.GEMINI_API_KEY, 'GLOBAL_ORG_GEMINI_KEY_7777');
    assert.equal(anonEnv.GROQ_API_KEY, undefined);

    // 4) Resolution preview exposes only source labels & masked previews
    const preview = reloaded.getResolutionPreview('alice@shipyard.co.kr');
    const geminiItem = preview.items.find((i) => i.key === 'GEMINI_API_KEY');
    assert.equal(geminiItem?.source, 'employee-emv');
    assert.ok(!geminiItem?.maskedPreview.includes('ALICE_PERSONAL'));

    // 5) Deleting Alice's override causes immediate fallback to global EMV
    assert.equal(
      reloaded.deleteEntry({ scope: 'employee', employeeId: 'alice@shipyard.co.kr', key: 'GEMINI_API_KEY' }),
      true
    );
    assert.equal(reloaded.resolveRuntimeEnv('alice@shipyard.co.kr').GEMINI_API_KEY, 'GLOBAL_ORG_GEMINI_KEY_7777');
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('EmvVaultService supports atomic first-run admin claim and rejects unsafe inputs', () => {
  const vault = new EmvVaultService({ env: {} });
  const initialStatus = vault.getAuthStatus();
  assert.equal(initialStatus.requiresInitialClaim, true);
  assert.equal(initialStatus.persistenceMode, 'encrypted-memory');

  // Weak password rejected
  assert.throws(() => vault.claimInitialAdminPassword('1234'), /8자 이상/);

  // Atomic first-run claim succeeds once and locks afterwards
  const session = vault.claimInitialAdminPassword('AdminClaim#2026');
  assert.ok(vault.verifyAdminSession(session.token));
  assert.throws(() => vault.claimInitialAdminPassword('AnotherPass#2026'), /이미 등록되어 있습니다/);

  // Login works with claimed password, fails with wrong password
  assert.throws(() => vault.authenticateAdmin('WrongPass#2026'), /일치하지 않습니다/);
  const login2 = vault.authenticateAdmin('AdminClaim#2026');
  assert.ok(vault.verifyAdminSession(login2.token));

  // Rejects disallowed key names & non-HTTPS Worker URLs
  assert.throws(
    () => vault.setEntry({ scope: 'global', key: 'PATH', value: '/usr/bin' }),
    /허용되지 않은 EMV 키/
  );
  assert.throws(
    () =>
      vault.setEntry({
        scope: 'global',
        key: 'AGENT_BUILDER_WORKER_URL',
        value: 'http://127.0.0.1:8080/internal',
      }),
    /https:\/\//
  );
});
