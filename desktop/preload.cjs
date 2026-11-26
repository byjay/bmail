const { contextBridge, ipcRenderer } = require('electron');

function subscribeChannel(channel, callback) {
  if (typeof callback !== 'function') return () => {};
  const listener = (_event, payload) => {
    callback(payload);
  };
  ipcRenderer.on(channel, listener);
  return () => {
    ipcRenderer.removeListener(channel, listener);
  };
}

contextBridge.exposeInMainWorld('bmailDesktop', {
  isDesktopExe: true,
  platform: process.platform,
  onMailto: (callback) =>
    subscribeChannel('bmail:mailto', (uri) => {
      if (typeof uri === 'string') callback(uri);
    }),
  onDeepLink: (callback) =>
    subscribeChannel('bmail:deepLink', (uri) => {
      if (typeof uri === 'string') callback(uri);
    }),
  onOpenRoute: (callback) =>
    subscribeChannel('bmail:openRoute', (route) => {
      if (typeof route === 'string') callback(route);
    }),
  onUpdateAvailable: (callback) => subscribeChannel('bmail:updateAvailable', callback),
  onUpdateProgress: (callback) => subscribeChannel('bmail:updateProgress', callback),
  onToastAction: (callback) => subscribeChannel('bmail:toastAction', callback),
  pickLocalFiles: () => ipcRenderer.invoke('bmail:pickLocalFiles'),
  pickLocalFolder: () => ipcRenderer.invoke('bmail:pickLocalFolder'),
  getAppInfo: () => ipcRenderer.invoke('bmail:getAppInfo'),
  checkUpdate: () => ipcRenderer.invoke('bmail:checkUpdate'),
  installUpdate: () => ipcRenderer.invoke('bmail:installUpdate'),
  notify: (payload) => ipcRenderer.invoke('bmail:notify', payload),
  showMailToast: (payload) => ipcRenderer.invoke('bmail:showMailToast', payload),
  openSgwLoginWindow: (sgwBaseUrl) => ipcRenderer.invoke('bmail:openSgwLoginWindow', sgwBaseUrl),
  clearSgwAuthSession: (sgwBaseUrl) => ipcRenderer.invoke('bmail:clearSgwAuthSession', sgwBaseUrl),
  openComposeWindow: (draftKey) => ipcRenderer.invoke('bmail:openComposeWindow', draftKey),
  setCompanionDock: (opts) => ipcRenderer.invoke('bmail:setCompanionDock', opts),
});
