/**
 * Bmail Windows Desktop (.exe) Main Entry (Electron)
 * - Enforces single instance & captures Windows mailto: / bmail: protocol launches.
 * - Dynamically resolves a conflict-free loopback port (avoiding port 3000 MCP/dev collisions).
 * - Standalone packaged mode: spawns embedded Electron Node 24.15.0 (ELECTRON_RUN_AS_NODE=1)
 *   running dist-server/server.mjs with built-in node:sqlite + FTS5 (zero external Node dependency).
 * - Polls /api/config readiness with strict Bmail JSON check, supports Tray icon, auto-launch, & clean exit.
 */
const {
  app,
  BrowserWindow,
  Menu,
  Notification,
  Tray,
  dialog,
  ipcMain,
  nativeImage,
  screen,
  session,
  shell,
} = require('electron');
const crypto = require('node:crypto');
const fs = require('node:fs');
const http = require('node:http');
const https = require('node:https');
const net = require('node:net');
const os = require('node:os');
const path = require('node:path');
const zlib = require('node:zlib');
const { spawn } = require('node:child_process');

const DEFAULT_UPDATE_MANIFEST_URL =
  process.env.BMAIL_UPDATE_URL || 'https://mail.seastar.work/releases/latest.json';

let activePort = 3456;
function getBaseUrl(port = activePort) {
  return `http://127.0.0.1:${port}`;
}

let mainWindow = null;
let trayToastWindow = null;
let trayToastTimer = null;
let trayMonitorTimer = null;
let serverProcess = null;
let backendLogFd = null;
let tray = null;
let isQuitting = false;
let pendingMailtoUri = null;
let pendingDeepLinkUri = null;
let lastCheckedManifest = null;
let updateInFlight = false;

{
  const initialUri = extractMailtoArg(process.argv);
  if (initialUri) {
    if (/^(bmail|sgw-outlook):/i.test(initialUri)) {
      pendingDeepLinkUri = initialUri;
    } else {
      pendingMailtoUri = initialUri;
    }
  }
}

function parseSemverParts(raw) {
  const clean = String(raw || '').trim().replace(/^v/i, '').split('-')[0];
  const parts = clean.split('.').map((n) => Number.parseInt(n, 10));
  return [parts[0] || 0, parts[1] || 0, parts[2] || 0];
}

function compareSemver(a, b) {
  const pa = parseSemverParts(a);
  const pb = parseSemverParts(b);
  for (let i = 0; i < 3; i += 1) {
    if (pa[i] !== pb[i]) return pa[i] > pb[i] ? 1 : -1;
  }
  return 0;
}

function isSafeUpdateUrl(rawUrl) {
  try {
    const u = new URL(String(rawUrl || ''));
    if (u.protocol === 'https:') return true;
    if (u.protocol === 'http:' && (u.hostname === '127.0.0.1' || u.hostname === 'localhost')) {
      return true;
    }
    return false;
  } catch {
    return false;
  }
}

function fetchManifestJson(rawUrl, redirectsLeft = 3) {
  return new Promise((resolve, reject) => {
    if (!isSafeUpdateUrl(rawUrl)) {
      reject(new Error('허용되지 않은 업데이트 URL 스키마입니다.'));
      return;
    }
    const u = new URL(rawUrl);
    const client = u.protocol === 'https:' ? https : http;
    const req = client.get(u, { timeout: 10000, headers: { Accept: 'application/json' } }, (res) => {
      if (res.statusCode && res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirectsLeft > 0) {
        res.resume();
        const nextUrl = new URL(res.headers.location, u).toString();
        resolve(fetchManifestJson(nextUrl, redirectsLeft - 1));
        return;
      }
      if (res.statusCode !== 200) {
        res.resume();
        reject(new Error(`업데이트 서버 응답 오류 (HTTP ${res.statusCode})`));
        return;
      }
      let raw = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => {
        if (raw.length < 131072) raw += chunk;
      });
      res.on('end', () => {
        try {
          resolve(JSON.parse(raw));
        } catch {
          reject(new Error('업데이트 매니페스트 JSON 파싱 실패'));
        }
      });
    });
    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy(new Error('업데이트 확인 시간 초과'));
    });
  });
}

function sendUpdateProgress(payload) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('bmail:updateProgress', payload);
  }
}

async function checkForUpdates() {
  const currentVersion = app.getVersion() || '1.0.0';
  try {
    const manifest = await fetchManifestJson(DEFAULT_UPDATE_MANIFEST_URL);
    const latestVersion = typeof manifest?.version === 'string' ? manifest.version.trim() : '';
    const downloadUrl =
      typeof manifest?.downloadUrl === 'string'
        ? manifest.downloadUrl.trim()
        : typeof manifest?.latestDownloadUrl === 'string'
          ? manifest.latestDownloadUrl.trim()
          : '';
    const sha256 = typeof manifest?.sha256 === 'string' ? manifest.sha256.trim().toLowerCase() : '';
    const updateAvailable =
      Boolean(latestVersion && downloadUrl && /^[a-f0-9]{64}$/.test(sha256)) &&
      compareSemver(latestVersion, currentVersion) > 0 &&
      isSafeUpdateUrl(downloadUrl);

    if (updateAvailable) {
      lastCheckedManifest = {
        version: latestVersion,
        downloadUrl,
        sha256,
        sizeBytes: Number(manifest.sizeBytes || 0),
        releaseNotes: typeof manifest.releaseNotes === 'string' ? manifest.releaseNotes : '',
        publishedAt: typeof manifest.publishedAt === 'string' ? manifest.publishedAt : '',
      };
    }

    return {
      currentVersion,
      latestVersion: latestVersion || currentVersion,
      updateAvailable,
      releaseNotes: typeof manifest?.releaseNotes === 'string' ? manifest.releaseNotes : '',
      downloadUrl: updateAvailable ? downloadUrl : '',
      sizeBytes: Number(manifest?.sizeBytes || 0),
      publishedAt: typeof manifest?.publishedAt === 'string' ? manifest.publishedAt : '',
    };
  } catch (err) {
    return {
      currentVersion,
      latestVersion: currentVersion,
      updateAvailable: false,
      releaseNotes: '',
      downloadUrl: '',
      sizeBytes: 0,
      publishedAt: '',
      error: err instanceof Error ? err.message : '업데이트 확인 실패',
    };
  }
}

function downloadFileWithSha256(rawUrl, destPath, version, expectedSizeBytes, redirectsLeft = 3) {
  return new Promise((resolve, reject) => {
    if (!isSafeUpdateUrl(rawUrl)) {
      reject(new Error('안전하지 않은 다운로드 URL입니다.'));
      return;
    }
    const u = new URL(rawUrl);
    const client = u.protocol === 'https:' ? https : http;
    const req = client.get(u, { timeout: 60000 }, (res) => {
      if (res.statusCode && res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirectsLeft > 0) {
        res.resume();
        const nextUrl = new URL(res.headers.location, u).toString();
        resolve(downloadFileWithSha256(nextUrl, destPath, version, expectedSizeBytes, redirectsLeft - 1));
        return;
      }
      if (res.statusCode !== 200) {
        res.resume();
        reject(new Error(`설치 파일 다운로드 실패 (HTTP ${res.statusCode})`));
        return;
      }
      const totalBytes = Number(res.headers['content-length'] || expectedSizeBytes || 0);
      let downloadedBytes = 0;
      let lastEmitAt = 0;
      const hash = crypto.createHash('sha256');
      const out = fs.createWriteStream(destPath);

      res.on('data', (chunk) => {
        downloadedBytes += chunk.length;
        hash.update(chunk);
        const now = Date.now();
        if (now - lastEmitAt >= 150) {
          lastEmitAt = now;
          const percent = totalBytes > 0 ? Math.min(99, Math.round((downloadedBytes / totalBytes) * 100)) : 0;
          sendUpdateProgress({ phase: 'downloading', percent, downloadedBytes, totalBytes, version });
        }
      });

      res.on('error', (err) => {
        out.destroy();
        reject(err);
      });
      out.on('error', reject);
      out.on('finish', () => {
        resolve({ sha256: hash.digest('hex'), downloadedBytes, totalBytes });
      });
      res.pipe(out);
    });
    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy(new Error('다운로드 시간 초과'));
    });
  });
}

async function downloadAndInstallUpdate() {
  if (updateInFlight) {
    return { started: false, error: '이미 업데이트가 진행 중입니다.' };
  }
  updateInFlight = true;
  let destPath = '';
  try {
    const check = await checkForUpdates();
    if (!check.updateAvailable || !lastCheckedManifest) {
      updateInFlight = false;
      return {
        started: false,
        error: check.error || `현재 최신 버전(v${check.currentVersion})을 사용 중입니다.`,
      };
    }

    const safeVer = lastCheckedManifest.version.replace(/[^0-9A-Za-z._-]/g, '_');
    destPath = path.join(os.tmpdir(), `Bmail-Outlook-AI-Update-${safeVer}.exe`);
    sendUpdateProgress({
      phase: 'downloading',
      percent: 0,
      downloadedBytes: 0,
      totalBytes: lastCheckedManifest.sizeBytes,
      version: lastCheckedManifest.version,
    });

    const dl = await downloadFileWithSha256(
      lastCheckedManifest.downloadUrl,
      destPath,
      lastCheckedManifest.version,
      lastCheckedManifest.sizeBytes
    );

    if (dl.sha256.toLowerCase() !== lastCheckedManifest.sha256.toLowerCase()) {
      try {
        fs.unlinkSync(destPath);
      } catch {
        // ignore
      }
      const msg = '다운로드 파일의 SHA-256 무결성 검증에 실패했습니다.';
      sendUpdateProgress({ phase: 'error', percent: 0, downloadedBytes: 0, totalBytes: 0, version: lastCheckedManifest.version, message: msg });
      updateInFlight = false;
      return { started: false, error: msg };
    }

    sendUpdateProgress({
      phase: 'installing',
      percent: 100,
      downloadedBytes: dl.downloadedBytes,
      totalBytes: dl.downloadedBytes,
      version: lastCheckedManifest.version,
    });

    const installerChild = spawn(destPath, ['/S'], {
      detached: true,
      stdio: 'ignore',
      windowsHide: false,
    });
    installerChild.unref();

    isQuitting = true;
    if (serverProcess) {
      try {
        serverProcess.kill();
      } catch {
        // ignore
      }
      serverProcess = null;
    }
    setTimeout(() => {
      app.quit();
    }, 300);

    return { started: true, version: lastCheckedManifest.version };
  } catch (err) {
    if (destPath) {
      try {
        fs.unlinkSync(destPath);
      } catch {
        // ignore
      }
    }
    const msg = err instanceof Error ? err.message : '업데이트 설치 중 오류가 발생했습니다.';
    sendUpdateProgress({ phase: 'error', percent: 0, downloadedBytes: 0, totalBytes: 0, version: '', message: msg });
    updateInFlight = false;
    return { started: false, error: msg };
  }
}

function extractMailtoArg(argv) {
  if (!Array.isArray(argv)) return null;
  for (const arg of argv) {
    if (typeof arg === 'string' && /^(mailto|bmail|sgw-outlook):/i.test(arg.trim())) {
      return arg.trim();
    }
  }
  return null;
}

function resolveAppRoot() {
  return path.resolve(__dirname, '..');
}

function resolveDataDir(appRoot) {
  const localWorkspaceDb = 'E:\\code-project\\bmail\\data\\bmail-local.sqlite';
  if (fs.existsSync(localWorkspaceDb)) {
    return path.dirname(localWorkspaceDb);
  }
  if (app.isPackaged) {
    return path.join(app.getPath('userData'), 'data');
  }
  return path.join(appRoot, 'data');
}

function resolveEnvFile(appRoot, dataDir) {
  const candidates = [
    path.join(dataDir, '.env'),
    'E:\\code-project\\bmail\\.env',
    path.join(appRoot, '.env'),
  ];
  for (const candidate of candidates) {
    try {
      if (fs.existsSync(candidate)) return candidate;
    } catch {
      // ignore
    }
  }
  return path.join(dataDir, '.env');
}

function resolveServerEntry(appRoot) {
  const compiled = path.join(appRoot, 'dist-server', 'server.mjs');
  if (fs.existsSync(compiled)) return compiled;
  return path.join(appRoot, 'server.ts');
}

function checkServerReady(port = activePort) {
  return new Promise((resolve) => {
    const req = http.get(`${getBaseUrl(port)}/api/config`, { timeout: 1000 }, (res) => {
      if (res.statusCode !== 200) {
        res.resume();
        resolve(false);
        return;
      }
      let raw = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => {
        if (raw.length < 32768) raw += chunk;
      });
      res.on('end', () => {
        try {
          const json = JSON.parse(raw);
          resolve(
            Boolean(
              json &&
                typeof json.appUrl === 'string' &&
                Array.isArray(json.providers)
            )
          );
        } catch {
          resolve(false);
        }
      });
    });
    req.on('error', () => resolve(false));
    req.on('timeout', () => {
      req.destroy();
      resolve(false);
    });
  });
}

function isTcpConnectRejected(port) {
  return new Promise((resolve) => {
    const socket = net.connect({ host: '127.0.0.1', port });
    const done = (rejected) => {
      socket.removeAllListeners();
      socket.destroy();
      resolve(rejected);
    };
    socket.setTimeout(350);
    socket.on('connect', () => done(false));
    socket.on('timeout', () => done(false));
    socket.on('error', () => done(true));
  });
}

function canBindLoopbackPort(port) {
  return new Promise((resolve) => {
    const srv = net.createServer();
    srv.unref();
    srv.once('error', () => resolve(false));
    srv.listen(port, '127.0.0.1', () => {
      const addr = srv.address();
      const boundPort = addr && typeof addr === 'object' ? addr.port : port;
      srv.close(() => resolve(boundPort));
    });
  });
}

async function isLoopbackPortFree(port) {
  const rejected = await isTcpConnectRejected(port);
  if (!rejected) return false;
  const bound = await canBindLoopbackPort(port);
  return Boolean(bound);
}

async function resolveDesktopPort() {
  const envPort = Number(process.env.BMAIL_DESKTOP_PORT || '');
  if (Number.isInteger(envPort) && envPort > 0) {
    return { port: envPort, alreadyRunning: await checkServerReady(envPort) };
  }
  for (const p of [3456, 3457, 3458, 3000]) {
    if (await checkServerReady(p)) {
      return { port: p, alreadyRunning: true };
    }
  }
  for (const p of [3456, 3457, 3458, 3459, 3460, 3461, 3462]) {
    if (await isLoopbackPortFree(p)) {
      return { port: p, alreadyRunning: false };
    }
  }
  const ephemeral = await canBindLoopbackPort(0);
  return { port: ephemeral || 3456, alreadyRunning: false };
}

async function waitForServerReady(port = activePort, timeoutMs = 15000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (await checkServerReady(port)) return true;
    await new Promise((r) => setTimeout(r, 200));
  }
  return false;
}

function writeRuntimeManifest(dataDir, port) {
  try {
    fs.mkdirSync(dataDir, { recursive: true });
    fs.writeFileSync(
      path.join(dataDir, 'desktop-runtime.json'),
      JSON.stringify(
        {
          port,
          url: getBaseUrl(port),
          pid: process.pid,
          startedAt: new Date().toISOString(),
        },
        null,
        2
      ) + '\n',
      'utf8'
    );
  } catch {
    // ignore non-critical runtime file errors
  }
}

async function startLocalBackendIfNeeded() {
  const appRoot = resolveAppRoot();
  const dataDir = resolveDataDir(appRoot);
  const resolved = await resolveDesktopPort();
  activePort = resolved.port;

  if (resolved.alreadyRunning) {
    writeRuntimeManifest(dataDir, activePort);
    return;
  }

  const entry = resolveServerEntry(appRoot);
  const envFile = resolveEnvFile(appRoot, dataDir);

  fs.mkdirSync(dataDir, { recursive: true });
  const logPath = path.join(dataDir, 'desktop-backend.log');
  try {
    backendLogFd = fs.openSync(logPath, 'a');
  } catch {
    backendLogFd = null;
  }

  const preferredDevNode24 = 'E:\\tools\\node-v24\\node.exe';
  const useEmbeddedNode =
    app.isPackaged ||
    !fs.existsSync(preferredDevNode24) ||
    !entry.endsWith('.mjs');

  const execBin = useEmbeddedNode ? process.execPath : preferredDevNode24;
  const env = {
    ...process.env,
    NODE_ENV: 'production',
    PORT: String(activePort),
    BMAIL_DESKTOP_PORT: String(activePort),
    BMAIL_HOST: '127.0.0.1',
    BMAIL_APP_ROOT: appRoot,
    BMAIL_DATA_DIR: dataDir,
    BMAIL_ENV_FILE: envFile,
  };
  if (useEmbeddedNode) {
    env.ELECTRON_RUN_AS_NODE = '1';
  }

  serverProcess = spawn(execBin, [entry], {
    cwd: appRoot,
    env,
    stdio: backendLogFd !== null ? ['ignore', backendLogFd, backendLogFd] : 'ignore',
    windowsHide: true,
  });

  await waitForServerReady(activePort, 15000);
  writeRuntimeManifest(dataDir, activePort);
}

function createFallbackTrayIcon() {
  const size = 16;
  const stride = size * 4 + 1;
  const raw = Buffer.alloc(stride * size);
  for (let y = 0; y < size; y += 1) {
    raw[y * stride] = 0;
    for (let x = 0; x < size; x += 1) {
      const idx = y * stride + 1 + x * 4;
      const border = x === 0 || y === 0 || x === size - 1 || y === size - 1;
      const mailStripe = y >= 4 && y <= 11 && x >= 3 && x <= 12 && (y === 4 || y === 11 || x === 3 || x === 12);
      if (border) {
        raw[idx] = 11;
        raw[idx + 1] = 79;
        raw[idx + 2] = 138;
        raw[idx + 3] = 255;
      } else if (mailStripe) {
        raw[idx] = 255;
        raw[idx + 1] = 255;
        raw[idx + 2] = 255;
        raw[idx + 3] = 255;
      } else {
        raw[idx] = 15;
        raw[idx + 1] = 108;
        raw[idx + 2] = 189;
        raw[idx + 3] = 255;
      }
    }
  }
  const crcTable = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  const crc32 = (buf) => {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i += 1) c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const t = Buffer.from(type, 'ascii');
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length, 0);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(Buffer.concat([t, data])), 0);
    return Buffer.concat([len, t, data, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const png = Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
  return nativeImage.createFromBuffer(png);
}

function loadAppIcon() {
  const pngPath = path.join(__dirname, 'icon.png');
  if (fs.existsSync(pngPath)) {
    const img = nativeImage.createFromPath(pngPath);
    if (!img.isEmpty()) return img;
  }
  return createFallbackTrayIcon();
}

function focusMainWindow() {
  if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.show();
    mainWindow.focus();
  } else {
    void createWindow();
  }
}

function setupTray() {
  if (tray) return;
  try {
    const appIcon = loadAppIcon();
    const trayImg = appIcon.resize({ width: 16, height: 16 });
    tray = new Tray(trayImg);
    tray.setToolTip('Bmail Outlook AI');

    const loginSettings = app.getLoginItemSettings();
    tray.setContextMenu(
      Menu.buildFromTemplate([
        {
          label: 'Bmail 열기',
          click: () => focusMainWindow(),
        },
        {
          label: '업데이트 확인',
          click: async () => {
            const res = await checkForUpdates();
            focusMainWindow();
            if (res.updateAvailable && mainWindow && !mainWindow.isDestroyed()) {
              mainWindow.webContents.send('bmail:updateAvailable', res);
            } else {
              void dialog.showMessageBox(mainWindow || undefined, {
                type: 'info',
                title: 'Bmail 업데이트 확인',
                message: res.error
                  ? `업데이트 확인 실패: ${res.error}`
                  : `현재 최신 버전(v${res.currentVersion})을 사용 중입니다.`,
              });
            }
          },
        },
        {
          label: 'Windows 시작 시 자동 실행',
          type: 'checkbox',
          checked: Boolean(loginSettings.openAtLogin),
          click: (menuItem) => {
            app.setLoginItemSettings({ openAtLogin: menuItem.checked });
          },
        },
        {
          label: '데이터 폴더 열기',
          click: () => {
            const dDir = resolveDataDir(resolveAppRoot());
            fs.mkdirSync(dDir, { recursive: true });
            void shell.openPath(dDir);
          },
        },
        { type: 'separator' },
        {
          label: '종료',
          click: () => {
            isQuitting = true;
            app.quit();
          },
        },
      ])
    );
    tray.on('double-click', () => focusMainWindow());
  } catch {
    tray = null;
  }
}

async function createWindow() {
  const ready = await waitForServerReady(activePort, 15000);
  const baseUrl = getBaseUrl(activePort);
  const targetUrl = pendingMailtoUri
    ? `${baseUrl}/?mailto=${encodeURIComponent(pendingMailtoUri)}`
    : baseUrl;
  pendingMailtoUri = null;

  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 380,
    minHeight: 620,
    title: 'Bmail Outlook AI',
    icon: loadAppIcon(),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith(`${baseUrl}/?compose=1`)) {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          width: 1080,
          height: 780,
          autoHideMenuBar: true,
          icon: loadAppIcon(),
          webPreferences: {
            preload: path.join(__dirname, 'preload.cjs'),
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: true,
          },
        },
      };
    }
    if (url.startsWith('https://')) {
      void shell.openExternal(url);
    }
    return { action: 'deny' };
  });

  mainWindow.on('close', (e) => {
    if (!isQuitting && tray) {
      e.preventDefault();
      mainWindow.hide();
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  const deepLinkToDeliver = pendingDeepLinkUri;
  pendingDeepLinkUri = null;

  mainWindow.webContents.on('did-finish-load', () => {
    if (deepLinkToDeliver && mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('bmail:deepLink', deepLinkToDeliver);
    }
  });

  if (ready) {
    await mainWindow.loadURL(targetUrl);
  } else {
    await mainWindow.loadURL(
      `data:text/html;charset=utf-8,${encodeURIComponent(
        `<html><body style="font-family:sans-serif;padding:24px;"><h3>Bmail 로컬 서버 시작 실패</h3><p>데이터 폴더의 desktop-backend.log 및 포트 ${activePort} 연결 상태를 확인해 주세요.</p></body></html>`
      )}`
    );
  }
}

const gotSingleLock = app.requestSingleInstanceLock();
if (!gotSingleLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv) => {
    const uri = extractMailtoArg(argv);
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.show();
      mainWindow.focus();
      if (uri) {
        if (/^(bmail|sgw-outlook):/i.test(uri)) {
          mainWindow.webContents.send('bmail:deepLink', uri);
        } else {
          mainWindow.webContents.send('bmail:mailto', uri);
        }
      }
    } else if (uri) {
      if (/^(bmail|sgw-outlook):/i.test(uri)) {
        pendingDeepLinkUri = uri;
      } else {
        pendingMailtoUri = uri;
      }
    }
  });

  app.whenReady().then(async () => {
    if (app.isPackaged) {
      app.setAsDefaultProtocolClient('mailto');
      app.setAsDefaultProtocolClient('bmail');
      app.setAsDefaultProtocolClient('sgw-outlook');
    }
    ipcMain.handle('bmail:getAppInfo', () => ({
      version: app.getVersion() || '1.0.0',
      isPackaged: app.isPackaged,
      platform: process.platform,
      runtimePort: activePort,
    }));
    ipcMain.handle('bmail:checkUpdate', async () => checkForUpdates());
    ipcMain.handle('bmail:installUpdate', async () => downloadAndInstallUpdate());
    ipcMain.handle('bmail:notify', (_event, payload) => {
      try {
        if (!Notification.isSupported()) return false;
        const p = payload && typeof payload === 'object' ? payload : {};
        const title = typeof p.title === 'string' && p.title.trim() ? p.title.trim() : 'Bmail 알림';
        const body = typeof p.body === 'string' ? p.body.trim().slice(0, 180) : '';
        const route = typeof p.route === 'string' ? p.route : '';
        if (mainWindow && !mainWindow.isDestroyed() && mainWindow.isFocused() && mainWindow.isVisible()) {
          return false;
        }
        const n = new Notification({ title, body, icon: loadAppIcon() });
        n.on('click', () => {
          focusMainWindow();
          if (route && mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.webContents.send('bmail:openRoute', route);
          }
        });
        n.show();
        return true;
      } catch {
        return false;
      }
    });

    const escapeToastHtml = (raw) =>
      String(raw || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');

    const closeTrayToast = () => {
      if (trayToastTimer) {
        clearTimeout(trayToastTimer);
        trayToastTimer = null;
      }
      if (trayToastWindow && !trayToastWindow.isDestroyed()) {
        trayToastWindow.close();
      }
      trayToastWindow = null;
    };

    const showTrayMiniToast = (payload) => {
      try {
        closeTrayToast();
        const p = payload && typeof payload === 'object' ? payload : {};
        const emailId = typeof p.emailId === 'string' ? p.emailId.trim() : '';
        const sender = escapeToastHtml(p.sender || '새 메일 도착');
        const subject = escapeToastHtml(p.subject || '(제목 없음)');
        const summary = escapeToastHtml(String(p.summary || p.snippet || '').slice(0, 140));
        const isUrgent = p.urgency === 'high';

        const wa = screen.getPrimaryDisplay().workArea;
        const w = 380;
        const h = 168;
        const x = Math.max(wa.x + 8, wa.x + wa.width - w - 16);
        const y = Math.max(wa.y + 8, wa.y + wa.height - h - 16);

        trayToastWindow = new BrowserWindow({
          width: w,
          height: h,
          x,
          y,
          frame: false,
          resizable: false,
          minimizable: false,
          maximizable: false,
          skipTaskbar: true,
          alwaysOnTop: true,
          show: false,
          focusable: true,
          transparent: false,
          backgroundColor: '#FFFFFF',
          webPreferences: {
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: true,
          },
        });

        const safeIdJs = JSON.stringify(emailId);
        const html = `<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<style>
  *{box-sizing:border-box;font-family:'Segoe UI','Malgun Gothic',sans-serif;}
  body{margin:0;background:#fff;border:1px solid ${isUrgent ? '#D13438' : '#0F6CBD'};height:168px;display:flex;flex-direction:column;justify-content:space-between;padding:10px 12px;color:#242424;}
  .top{display:flex;align-items:center;justify-content:space-between;gap:8px;}
  .badges{display:flex;align-items:center;gap:5px;}
  .badge{font-size:10px;font-weight:700;color:#0F6CBD;background:#EBF3FC;padding:2px 6px;border-radius:4px;}
  .urgent{font-size:10px;font-weight:700;color:#fff;background:#D13438;padding:2px 6px;border-radius:4px;}
  .close{border:none;background:transparent;color:#707070;cursor:pointer;font-size:13px;line-height:1;}
  .sender{font-size:12px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:4px;}
  .subj{font-size:11px;font-weight:600;color:#0F6CBD;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:2px;}
  .sum{font-size:11px;color:#323130;background:#FAF9F8;border-left:3px solid ${isUrgent ? '#D13438' : '#0F6CBD'};padding:4px 7px;margin-top:5px;border-radius:2px;line-height:1.35;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}
  .actions{display:flex;gap:6px;margin-top:6px;}
  button.act{flex:1;padding:5px 6px;font-size:11px;font-weight:600;border-radius:4px;border:1px solid #D1D1D1;background:#FAF9F8;color:#242424;cursor:pointer;}
  button.act.primary{background:#0F6CBD;border-color:#0F6CBD;color:#fff;}
</style>
</head>
<body>
  <div>
    <div class="top">
      <div class="badges">
        <span class="badge">새 메일 도착 · 간단 요약</span>
        ${isUrgent ? '<span class="urgent">긴급</span>' : ''}
      </div>
      <button class="close" onclick="location.href='bmail-toast://dismiss'">✕</button>
    </div>
    <div class="sender">${sender}</div>
    <div class="subj">${subject}</div>
    <div class="sum">${summary || '새 메일이 접수되었습니다.'}</div>
  </div>
  <div class="actions">
    <button class="act" onclick="location.href='bmail-toast://markRead?id='+encodeURIComponent(${safeIdJs})">✓ 즉시 읽음</button>
    <button class="act primary" onclick="location.href='bmail-toast://reply?id='+encodeURIComponent(${safeIdJs})">↩ 즉시 회신</button>
    <button class="act" onclick="location.href='bmail-toast://open?id='+encodeURIComponent(${safeIdJs})">열기</button>
  </div>
</body>
</html>`;

        trayToastWindow.webContents.on('will-navigate', (ev, url) => {
          if (!url.startsWith('bmail-toast://')) return;
          ev.preventDefault();
          try {
            const parsed = new URL(url);
            const action = parsed.hostname;
            const targetId = parsed.searchParams.get('id') || emailId;
            if (action === 'dismiss') {
              closeTrayToast();
              return;
            }
            if (action === 'markRead' && targetId) {
              const patchBody = JSON.stringify({ isRead: true });
              const req = http.request({
                hostname: '127.0.0.1',
                port: activePort,
                path: `/api/local/emails/${encodeURIComponent(targetId)}`,
                method: 'PATCH',
                headers: {
                  'Content-Type': 'application/json',
                  'Content-Length': Buffer.byteLength(patchBody),
                },
              });
              req.on('error', () => {});
              req.write(patchBody);
              req.end();
              if (mainWindow && !mainWindow.isDestroyed()) {
                mainWindow.webContents.send('bmail:toastAction', { action: 'markRead', emailId: targetId });
              }
            } else if ((action === 'reply' || action === 'open') && targetId) {
              focusMainWindow();
              if (mainWindow && !mainWindow.isDestroyed()) {
                mainWindow.webContents.send('bmail:toastAction', { action, emailId: targetId });
              }
            }
          } finally {
            closeTrayToast();
          }
        });

        trayToastWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
        trayToastWindow.once('ready-to-show', () => {
          if (trayToastWindow && !trayToastWindow.isDestroyed()) {
            trayToastWindow.showInactive();
          }
        });
        trayToastTimer = setTimeout(() => closeTrayToast(), 12_000);
        return true;
      } catch {
        return false;
      }
    };

    ipcMain.handle('bmail:showMailToast', (_event, payload) => showTrayMiniToast(payload));
    ipcMain.handle('bmail:pickLocalFiles', async () => {
      const targetWin = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
      const options = {
        title: '로컬 문서/도면 파일 선택',
        properties: ['openFile', 'multiSelections'],
        filters: [
          {
            name: '문서 및 데이터',
            extensions: ['pdf', 'xlsx', 'xls', 'xlsm', 'csv', 'docx', 'txt', 'md', 'json', 'eml', 'msg'],
          },
          { name: '모든 파일', extensions: ['*'] },
        ],
      };
      const result = targetWin
        ? await dialog.showOpenDialog(targetWin, options)
        : await dialog.showOpenDialog(options);
      return result.canceled ? [] : result.filePaths;
    });
    ipcMain.handle('bmail:pickLocalFolder', async () => {
      const targetWin = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
      const options = {
        title: '로컬 PST / 메일 보관 폴더 선택',
        properties: ['openDirectory'],
      };
      const result = targetWin
        ? await dialog.showOpenDialog(targetWin, options)
        : await dialog.showOpenDialog(options);
      return result.canceled ? '' : result.filePaths[0] || '';
    });

    ipcMain.handle('bmail:openComposeWindow', async (_event, draftKey) => {
      try {
        const key = typeof draftKey === 'string' ? draftKey.trim() : '';
        if (!key) return { ok: false, error: 'missing draftKey' };
        const baseUrl = getBaseUrl(activePort);
        const composeUrl = `${baseUrl}/?compose=1&draftKey=${encodeURIComponent(key)}`;
        const composeWin = new BrowserWindow({
          width: 1080,
          height: 780,
          minWidth: 720,
          minHeight: 540,
          title: 'Bmail Outlook — 메세지 작성 (HTML)',
          autoHideMenuBar: true,
          icon: loadAppIcon(),
          webPreferences: {
            preload: path.join(__dirname, 'preload.cjs'),
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: true,
          },
        });
        await composeWin.loadURL(composeUrl);
        return { ok: true };
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) };
      }
    });

    let savedPreDockBounds = null;
    ipcMain.handle('bmail:setCompanionDock', async (_event, opts) => {
      try {
        if (!mainWindow || mainWindow.isDestroyed()) {
          return { ok: false, error: 'mainWindow unavailable' };
        }
        const enabled = Boolean(opts && typeof opts === 'object' && opts.enabled);
        if (enabled) {
          if (!savedPreDockBounds && !mainWindow.isMaximized()) {
            savedPreDockBounds = mainWindow.getBounds();
          }
          if (mainWindow.isMaximized()) {
            mainWindow.unmaximize();
          }
          const display = screen.getDisplayMatching(mainWindow.getBounds()) || screen.getPrimaryDisplay();
          const wa = display.workArea;
          const requestedW = Number(opts?.width ?? 460);
          const dockW = Math.max(
            380,
            Math.min(Math.min(680, wa.width - 120), Number.isFinite(requestedW) ? Math.round(requestedW) : 460)
          );
          mainWindow.setBounds({
            x: wa.x + wa.width - dockW,
            y: wa.y,
            width: dockW,
            height: wa.height,
          });
          const alwaysOnTop = opts?.alwaysOnTop !== false;
          mainWindow.setAlwaysOnTop(alwaysOnTop, 'floating');
          return { ok: true, docked: true, alwaysOnTop: mainWindow.isAlwaysOnTop(), width: dockW };
        }
        mainWindow.setAlwaysOnTop(false);
        if (savedPreDockBounds) {
          mainWindow.setBounds(savedPreDockBounds);
          savedPreDockBounds = null;
        }
        return { ok: true, docked: false, alwaysOnTop: false };
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) };
      }
    });

    ipcMain.handle('bmail:clearSgwAuthSession', async (_event, rawBaseUrl) => {
      try {
        const cleanUrl =
          typeof rawBaseUrl === 'string' && rawBaseUrl.trim()
            ? rawBaseUrl.trim()
            : 'https://sgw.seastar.work';
        const origin = new URL(cleanUrl).origin;
        const sgwPartition = session.fromPartition('persist:sgw-auth');
        await sgwPartition.clearStorageData({ origin });
        return { ok: true };
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) };
      }
    });

    ipcMain.handle('bmail:openSgwLoginWindow', async (_event, rawBaseUrl) => {
      const candidate =
        typeof rawBaseUrl === 'string' && rawBaseUrl.trim()
          ? rawBaseUrl.trim()
          : 'https://sgw.seastar.work';
      let origin = 'https://sgw.seastar.work';
      try {
        const u = new URL(candidate);
        if (
          u.protocol === 'https:' ||
          (u.protocol === 'http:' && (u.hostname === '127.0.0.1' || u.hostname === 'localhost'))
        ) {
          origin = u.origin;
        }
      } catch {
        origin = 'https://sgw.seastar.work';
      }

      const sgwPartition = session.fromPartition('persist:sgw-auth');

      return new Promise((resolve) => {
        let settled = false;
        let verifyInFlight = false;
        let lastTestedKey = '';

        const authWin = new BrowserWindow({
          width: 520,
          height: 740,
          title: 'SGW 사내 그룹웨어 공식 로그인 연동',
          parent: mainWindow && !mainWindow.isDestroyed() ? mainWindow : undefined,
          autoHideMenuBar: true,
          icon: loadAppIcon(),
          webPreferences: {
            partition: 'persist:sgw-auth',
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: true,
          },
        });

        let pollTimer = null;

        const finish = (payload) => {
          if (settled) return;
          settled = true;
          if (pollTimer) {
            clearInterval(pollTimer);
            pollTimer = null;
          }
          try {
            sgwPartition.webRequest.onBeforeSendHeaders(null);
          } catch {
            // ignore
          }
          try {
            if (!authWin.isDestroyed()) authWin.close();
          } catch {
            // ignore
          }
          resolve(payload);
        };

        const postSessionToLocalBackend = (token, cookieHeader) =>
          new Promise((resPost) => {
            const bodyStr = JSON.stringify({
              sgwBaseUrl: origin,
              token: token || '',
              cookieHeader: cookieHeader || '',
            });
            const req = http.request(
              {
                hostname: '127.0.0.1',
                port: activePort,
                path: '/api/sgw-chat/session-connect',
                method: 'POST',
                timeout: 8000,
                headers: {
                  'Content-Type': 'application/json; charset=utf-8',
                  'Content-Length': Buffer.byteLength(bodyStr),
                },
              },
              (res) => {
                let raw = '';
                res.setEncoding('utf8');
                res.on('data', (c) => {
                  if (raw.length < 65536) raw += c;
                });
                res.on('end', () => {
                  try {
                    const parsed = JSON.parse(raw || '{}');
                    if (res.statusCode === 200 && parsed && parsed.connected) {
                      resPost(parsed);
                      return;
                    }
                  } catch {
                    // ignore
                  }
                  resPost(null);
                });
              }
            );
            req.on('error', () => resPost(null));
            req.on('timeout', () => {
              req.destroy();
              resPost(null);
            });
            req.write(bodyStr);
            req.end();
          });

        const tryCandidateAuth = async (candidateToken) => {
          if (settled || verifyInFlight) return;
          verifyInFlight = true;
          try {
            const cookies = await sgwPartition.cookies.get({ url: origin }).catch(() => []);
            const cookieHeader = Array.isArray(cookies) && cookies.length > 0
              ? cookies.map((c) => `${c.name}=${c.value}`).join('; ')
              : '';
            const token = String(candidateToken || '').trim().replace(/^Bearer\s+/i, '');
            if (!token && !cookieHeader) {
              return;
            }
            const dedupKey = `${token}::${cookieHeader}`;
            if (dedupKey === lastTestedKey) {
              return;
            }
            lastTestedKey = dedupKey;
            const connectedStatus = await postSessionToLocalBackend(token, cookieHeader);
            if (connectedStatus && connectedStatus.connected) {
              finish(connectedStatus);
            }
          } finally {
            verifyInFlight = false;
          }
        };

        try {
          sgwPartition.webRequest.onBeforeSendHeaders(
            { urls: [`${origin}/api/*`, `${origin}/*`] },
            (details, callback) => {
              const hdrs = details.requestHeaders || {};
              const authVal = hdrs.Authorization || hdrs.authorization || '';
              if (typeof authVal === 'string' && /^Bearer\s+\S+/i.test(authVal)) {
                void tryCandidateAuth(authVal);
              }
              callback({ requestHeaders: details.requestHeaders });
            }
          );
        } catch {
          // ignore if webRequest listener attachment fails
        }

        const probeWindowStorage = async () => {
          if (settled || authWin.isDestroyed()) return;
          try {
            const foundToken = await authWin.webContents.executeJavaScript(`
              (() => {
                const stripQuotes = (v) => String(v || '').trim().replace(/^"+|"+$/g, '');
                const stores = [window.localStorage, window.sessionStorage];
                for (const st of stores) {
                  if (!st) continue;
                  for (let i = 0; i < st.length; i++) {
                    const k = st.key(i) || '';
                    const raw = stripQuotes(st.getItem(k));
                    if (!raw) continue;
                    if (/^eyJ[A-Za-z0-9_-]+\\.[A-Za-z0-9_-]+\\.[A-Za-z0-9_-]+$/.test(raw)) {
                      return raw;
                    }
                    if (/token|jwt|auth|session/i.test(k) && raw.length >= 16 && raw.length <= 4096) {
                      if (raw.startsWith('{')) {
                        try {
                          const obj = JSON.parse(raw);
                          if (typeof obj.token === 'string' && obj.token.length >= 16) return obj.token;
                          if (typeof obj.access_token === 'string') return obj.access_token;
                        } catch {}
                      } else {
                        return raw;
                      }
                    }
                  }
                }
                return '';
              })()
            `);
            await tryCandidateAuth(typeof foundToken === 'string' ? foundToken : '');
          } catch {
            // ignore cross-navigation script error
          }
        };

        authWin.webContents.on('did-finish-load', () => {
          void probeWindowStorage();
        });
        authWin.webContents.on('did-navigate-in-page', () => {
          void probeWindowStorage();
        });
        pollTimer = setInterval(() => {
          void probeWindowStorage();
        }, 1200);

        authWin.on('closed', () => {
          finish({ connected: false, canceled: true });
        });

        authWin.loadURL(origin).catch((err) => {
          finish({
            connected: false,
            error: err instanceof Error ? err.message : 'SGW 로그인 창을 열지 못했습니다.',
          });
        });
      });
    });
    await startLocalBackendIfNeeded();
    setupTray();
    await createWindow();
    setTimeout(() => {
      void checkForUpdates().then((res) => {
        if (res.updateAvailable && mainWindow && !mainWindow.isDestroyed()) {
          mainWindow.webContents.send('bmail:updateAvailable', res);
        }
      });
    }, 6000);

    // Poll backend proactive mail alerts every 20s when main window is minimized/hidden/unfocused
    const pollTrayAlerts = () => {
      const hiddenOrUnfocused =
        !mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible() || !mainWindow.isFocused();
      if (!hiddenOrUnfocused) return;
      const req = http.request(
        {
          hostname: '127.0.0.1',
          port: activePort,
          path: '/api/local/monitor/consume-tray-alerts',
          method: 'POST',
          timeout: 3000,
          headers: { 'Content-Type': 'application/json', 'Content-Length': 2 },
        },
        (res) => {
          let raw = '';
          res.setEncoding('utf8');
          res.on('data', (c) => {
            if (raw.length < 32768) raw += c;
          });
          res.on('end', () => {
            try {
              const parsed = JSON.parse(raw || '{}');
              const alerts = Array.isArray(parsed.alerts) ? parsed.alerts : [];
              if (alerts.length > 0) {
                showTrayMiniToast(alerts[0]);
              }
            } catch {
              // ignore transient startup JSON error
            }
          });
        }
      );
      req.on('error', () => {});
      req.on('timeout', () => req.destroy());
      req.write('{}');
      req.end();
    };
    trayMonitorTimer = setInterval(pollTrayAlerts, 20_000);
  });

  app.on('before-quit', () => {
    isQuitting = true;
    if (trayMonitorTimer) {
      clearInterval(trayMonitorTimer);
      trayMonitorTimer = null;
    }
    if (serverProcess) {
      try {
        serverProcess.kill();
      } catch {
        // ignore
      }
      serverProcess = null;
    }
    if (backendLogFd !== null) {
      try {
        fs.closeSync(backendLogFd);
      } catch {
        // ignore
      }
      backendLogFd = null;
    }
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin' && (!tray || isQuitting)) {
      app.quit();
    }
  });
}
