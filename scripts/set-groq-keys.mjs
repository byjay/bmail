#!/usr/bin/env node
/**
 * Configure up to 5 Groq API keys in a single terminal command (comma/whitespace separated)
 * or interactive stdin prompt. Syncs:
 *   1) <root>/groq-keys.txt
 *   2) <root>/data/groq-keys.txt (picked up automatically by packaged Windows .exe)
 *   3) <root>/.env (GROQ_API_KEY, GROQ_API_KEY_1..5)
 */
import fs from 'node:fs';
import path from 'node:path';
import readline from 'node:readline/promises';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DEFAULT_ROOT_DIR = path.resolve(__dirname, '..');
const MAX_SLOTS = 5;
const ENV_KEYS = [
  'GROQ_API_KEY',
  'GROQ_API_KEY_1',
  'GROQ_API_KEY_2',
  'GROQ_API_KEY_3',
  'GROQ_API_KEY_4',
  'GROQ_API_KEY_5',
];

export function maskGroqKey(key) {
  const clean = String(key || '').trim();
  if (clean.length <= 10) return '****';
  return `${clean.slice(0, 6)}...${clean.slice(-4)}`;
}

export function parseGroqKeysInput(raw) {
  const clean = String(raw || '').replace(/^\uFEFF/, '');
  const seen = new Set();
  const keys = [];
  for (const rawLine of clean.split(/\r?\n/)) {
    const trimmed = rawLine.trim();
    if (!trimmed || trimmed.startsWith('#') || trimmed.startsWith('//')) continue;
    const eqMatch = trimmed.match(/^[A-Za-z0-9_]+\s*=\s*(.+)$/);
    const valuePart = (eqMatch ? eqMatch[1] : trimmed).trim().replace(/^['"]|['"]$/g, '');
    for (const part of valuePart.split(/[,;\s]+/)) {
      const token = part.trim().replace(/^['"]|['"]$/g, '');
      if (token.length < 10) continue;
      if (/^(placeholder|changeme|example|none|null|undefined|MY_GROQ.*)$/i.test(token)) continue;
      if (seen.has(token)) continue;
      seen.add(token);
      keys.push(token);
      if (keys.length >= MAX_SLOTS) return keys;
    }
  }
  return keys;
}

function buildGroqKeysFileBody(keys) {
  return [
    '# ==============================================================================',
    '# Bmail — Groq API 5-키 자동 순환 설정 파일 (Git 커밋 자동 제외)',
    '# ==============================================================================',
    '# 한 줄에 1개씩(gsk_...) 넣거나 쉼표(,)로 구분해 5개를 한 줄에 붙여넣을 수 있습니다.',
    '# 터미널 일괄 적용: npm run groq:keys -- "gsk_키1,gsk_키2,gsk_키3,gsk_키4,gsk_키5"',
    '# ==============================================================================',
    '',
    ...keys,
    '',
  ].join('\n');
}

function upsertEnvKeys(envPath, keys) {
  const existing = fs.existsSync(envPath) ? fs.readFileSync(envPath, 'utf8') : '';
  const lines = existing ? existing.split(/\r?\n/) : [];
  const managedSet = new Set(ENV_KEYS);
  const filtered = lines.filter((line) => {
    const m = line.match(/^\s*([A-Za-z0-9_]+)\s*=/);
    return !m || !managedSet.has(m[1]);
  });
  while (filtered.length > 0 && filtered[filtered.length - 1] === '') {
    filtered.pop();
  }

  const slotEntries = [
    `GROQ_API_KEY=${keys[0] || ''}`,
    `GROQ_API_KEY_1=${keys[0] || ''}`,
    `GROQ_API_KEY_2=${keys[1] || ''}`,
    `GROQ_API_KEY_3=${keys[2] || ''}`,
    `GROQ_API_KEY_4=${keys[3] || ''}`,
    `GROQ_API_KEY_5=${keys[4] || ''}`,
  ];
  const nextText = [...filtered, '', ...slotEntries, ''].join('\n');
  fs.writeFileSync(envPath, nextText, 'utf8');
}

export function applyGroqKeys(rawInput, opts = {}) {
  const keys = parseGroqKeysInput(rawInput);
  if (keys.length === 0) {
    throw new Error('유효한 Groq API 키(10자 이상)를 1개 이상 입력해 주세요.');
  }
  const rootDir = path.resolve(opts.rootDir || DEFAULT_ROOT_DIR);
  const dataDir = path.resolve(opts.dataDir || path.join(rootDir, 'data'));

  fs.mkdirSync(rootDir, { recursive: true });
  fs.mkdirSync(dataDir, { recursive: true });

  const body = buildGroqKeysFileBody(keys);
  const rootKeysFile = path.join(rootDir, 'groq-keys.txt');
  const dataKeysFile = path.join(dataDir, 'groq-keys.txt');
  const rootEnvFile = path.join(rootDir, '.env');

  fs.writeFileSync(rootKeysFile, body, 'utf8');
  fs.writeFileSync(dataKeysFile, body, 'utf8');
  upsertEnvKeys(rootEnvFile, keys);

  return {
    count: keys.length,
    maskedKeys: keys.map(maskGroqKey),
    writtenFiles: [rootKeysFile, dataKeysFile, rootEnvFile],
  };
}

async function runCli() {
  let raw = process.argv.slice(2).join(' ').trim();
  if (!raw) {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    try {
      raw = await rl.question(
        'Groq API 키를 쉼표(,)로 구분해 최대 5개 붙여넣고 Enter를 누르세요:\n> '
      );
    } finally {
      rl.close();
    }
  }
  const res = applyGroqKeys(raw);
  console.log(`✅ Groq API ${res.count}/5개 키 적용 완료 (자동 순환 + 5-키 병렬 처리 활성화):`);
  res.maskedKeys.forEach((masked, idx) => {
    console.log(`   Slot #${idx + 1}: ${masked}`);
  });
  console.log('   적용 파일:');
  for (const p of res.writtenFiles) {
    console.log(`   - ${p}`);
  }
}

const isDirectMain =
  process.argv[1] && path.resolve(process.argv[1]) === path.resolve(__filename);

if (isDirectMain) {
  runCli().catch((err) => {
    console.error('❌ 오류:', err instanceof Error ? err.message : String(err));
    process.exit(1);
  });
}
