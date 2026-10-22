const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld("documentTabs", {
  request: (request) => ipcRenderer.invoke("office-tabs", request),
  onState: (callback) => ipcRenderer.on("office-tabs-state", (_event, state) => callback(state)),
});
