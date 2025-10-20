const { contextBridge, ipcRenderer } = require('electron');

function createSubscription(channel) {
  return (callback) => {
    const listener = (_event, data) => callback(data);
    ipcRenderer.on(channel, listener);
    return () => ipcRenderer.removeListener(channel, listener);
  };
}

contextBridge.exposeInMainWorld('carteraApi', {
  selectFile: () => ipcRenderer.invoke('dialog:select-file'),
  processFile: (filePath) => ipcRenderer.invoke('processor:run', filePath),
  sendLog: (message) => ipcRenderer.send('log:forward', { message }),
  revealInFolder: (filePath) => ipcRenderer.invoke('system:reveal-in-folder', filePath),
  onLogMessage: createSubscription('log-message'),
  onProcessingState: createSubscription('processing-state'),
  onUpdateEvent: createSubscription('update-event')
});
