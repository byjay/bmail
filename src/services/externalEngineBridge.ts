import { execFile } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { promisify } from 'node:util';
import type { ExternalEngineProbeResult, ExternalEnginesConfig } from '../types/mail';

const execFileAsync = promisify(execFile);

const NODE_ENTRY_CANDIDATES = [
  'maillog-adapter.mjs',
  'maillog-adapter.js',
  'maillog-adapter.ts',
  'cli.mjs',
  'cli.js',
  'index.mjs',
  'index.js',
  'src/index.ts',
];

const PYTHON_ENTRY_CANDIDATES = [
  'maillog_adapter.py',
  'analyze.py',
  'extract.py',
  'main.py',
  'app.py',
];

export interface ExternalEngineRunOutput {
  text: string;
  engineLabel: string;
  pageCount?: number;
  sheetNames?: string[];
  rowCount?: number;
  numericHighlights?: string[];
  warnings?: string[];
}

export function resolveExternalEngineDir(
  kind: 'pdf' | 'excel',
  config?: ExternalEnginesConfig
): { dir: string; command: string } {
  const envDir =
    kind === 'pdf'
      ? process.env.EXTERNAL_PDF_ENGINE_DIR?.trim() || ''
      : process.env.EXTERNAL_EXCEL_ENGINE_DIR?.trim() || '';
  const envCmd =
    kind === 'pdf'
      ? process.env.EXTERNAL_PDF_ENGINE_CMD?.trim() || ''
      : process.env.EXTERNAL_EXCEL_ENGINE_CMD?.trim() || '';

  const cfgDir = kind === 'pdf' ? config?.pdfEngineDir?.trim() || '' : config?.excelEngineDir?.trim() || '';
  const cfgCmd =
    kind === 'pdf' ? config?.pdfEngineCommand?.trim() || '' : config?.excelEngineCommand?.trim() || '';

  return {
    dir: envDir || cfgDir,
    command: envCmd || cfgCmd,
  };
}

/**
 * Inspects a local project folder to detect whether it exposes a runnable
 * Node.js, Python, or custom CLI adapter for PDF / Excel parsing.
 */
export function probeExternalEngine(
  kind: 'pdf' | 'excel',
  rawDir: string,
  rawCommand = ''
): ExternalEngineProbeResult {
  const configuredDir = rawDir.trim();
  if (!configuredDir && !rawCommand.trim()) {
    return {
      kind,
      configuredDir: '',
      exists: false,
      detectedRuntime: 'none',
      entryFile: null,
      message: `외부 ${kind.toUpperCase()} 엔진 폴더가 미설정되어 내장 엔진을 사용합니다.`,
    };
  }

  if (configuredDir && !fs.existsSync(configuredDir)) {
    return {
      kind,
      configuredDir,
      exists: false,
      detectedRuntime: 'none',
      entryFile: null,
      message: `지정한 폴더를 찾을 수 없습니다: ${configuredDir}`,
    };
  }

  if (rawCommand.trim()) {
    return {
      kind,
      configuredDir,
      exists: configuredDir ? fs.existsSync(configuredDir) : true,
      detectedRuntime: 'cli',
      entryFile: rawCommand.trim(),
      message: `사용자 지정 CLI 명령어(${rawCommand.trim().slice(0, 48)})가 설정되었습니다.`,
    };
  }

  // Check Python candidates first for PDF/Excel science projects, then Node
  for (const pyFile of PYTHON_ENTRY_CANDIDATES) {
    const full = path.join(configuredDir, pyFile);
    if (fs.existsSync(full)) {
      return {
        kind,
        configuredDir,
        exists: true,
        detectedRuntime: 'python',
        entryFile: pyFile,
        message: `Python 외부 ${kind.toUpperCase()} 엔진 감지됨 (${pyFile}). 파일 경로 인자로 호출합니다.`,
      };
    }
  }

  for (const nodeFile of NODE_ENTRY_CANDIDATES) {
    const full = path.join(configuredDir, nodeFile);
    if (fs.existsSync(full)) {
      return {
        kind,
        configuredDir,
        exists: true,
        detectedRuntime: 'node',
        entryFile: nodeFile,
        message: `Node.js 외부 ${kind.toUpperCase()} 엔진 감지됨 (${nodeFile}). 파일 경로 인자로 호출합니다.`,
      };
    }
  }

  const hasPackageJson = fs.existsSync(path.join(configuredDir, 'package.json'));
  const hasPyproject =
    fs.existsSync(path.join(configuredDir, 'pyproject.toml')) ||
    fs.existsSync(path.join(configuredDir, 'requirements.txt'));

  return {
    kind,
    configuredDir,
    exists: true,
    detectedRuntime: hasPyproject ? 'python' : hasPackageJson ? 'node' : 'none',
    entryFile: null,
    message:
      '폴더는 확인되었으나 표준 어댑터 엔트리(maillog_adapter.py 또는 maillog-adapter.mjs)가 없습니다. 외부 어댑터 설정 가이드를 참고해 엔트리를 추가하거나 실행 명령어를 지정하세요.',
  };
}

function splitSimpleCommand(cmd: string, filePath: string): { bin: string; args: string[] } | null {
  const tokens = cmd.match(/(?:[^\s"]+|"[^"]*")+/g);
  if (!tokens || tokens.length === 0) return null;
  const cleaned = tokens.map((t) => t.replace(/^"|"$/g, '').replace(/\{file\}/g, filePath));
  if (!cmd.includes('{file}')) {
    cleaned.push(filePath);
  }
  return { bin: cleaned[0], args: cleaned.slice(1) };
}

/**
 * Invokes an external local PDF/Excel project folder if configured and runnable.
 * Expects either JSON stdout (`{ text, pageCount, sheetNames, rowCount, numericHighlights }`)
 * or plain UTF-8 extracted text stdout.
 */
export async function tryRunExternalDocumentEngine(
  kind: 'pdf' | 'excel',
  filePath: string,
  config?: ExternalEnginesConfig
): Promise<ExternalEngineRunOutput | null> {
  const { dir, command } = resolveExternalEngineDir(kind, config);
  if (!dir && !command) return null;

  const probe = probeExternalEngine(kind, dir, command);
  if (!probe.exists || probe.detectedRuntime === 'none') {
    return null;
  }

  let bin = '';
  let args: string[] = [];
  const cwd = dir && fs.existsSync(dir) ? dir : process.cwd();

  if (probe.detectedRuntime === 'cli' && command) {
    const parsed = splitSimpleCommand(command, filePath);
    if (!parsed) return null;
    bin = parsed.bin;
    args = parsed.args;
  } else if (probe.detectedRuntime === 'python' && probe.entryFile) {
    bin = process.platform === 'win32' ? 'python' : 'python3';
    args = [path.join(dir, probe.entryFile), filePath, '--json'];
  } else if (probe.detectedRuntime === 'node' && probe.entryFile) {
    bin = probe.entryFile.endsWith('.ts') ? 'npx' : process.execPath;
    args = probe.entryFile.endsWith('.ts')
      ? ['tsx', path.join(dir, probe.entryFile), filePath, '--json']
      : [path.join(dir, probe.entryFile), filePath, '--json'];
  } else {
    return null;
  }

  const { stdout } = await execFileAsync(bin, args, {
    cwd,
    timeout: 20_000,
    maxBuffer: 12 * 1024 * 1024,
  });

  const rawOut = stdout.trim();
  if (!rawOut) {
    throw new Error(`외부 ${kind.toUpperCase()} 엔진이 빈 출력을 반환했습니다.`);
  }

  const folderLabel = dir ? path.basename(dir) : bin;
  try {
    const parsed = JSON.parse(rawOut) as {
      text?: string;
      pageCount?: number;
      sheetNames?: string[];
      rowCount?: number;
      numericHighlights?: string[];
      warnings?: string[];
    };
    if (typeof parsed.text === 'string' && parsed.text.trim()) {
      return {
        text: parsed.text.trim(),
        engineLabel: `외부 엔진 (${folderLabel})`,
        pageCount: parsed.pageCount,
        sheetNames: parsed.sheetNames,
        rowCount: parsed.rowCount,
        numericHighlights: parsed.numericHighlights,
        warnings: parsed.warnings,
      };
    }
  } catch {
    // Plain text stdout fallback
  }

  return {
    text: rawOut,
    engineLabel: `외부 엔진 (${folderLabel})`,
  };
}
