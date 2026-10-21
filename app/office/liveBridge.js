const { classifyOfficeUrl } = require("./inspector");
const { messages } = require("./addin/operations");

class LiveOfficeBridge {
  constructor(getWebContents, origin, { timeoutMs = 15000 } = {}) {
    this.getWebContents = getWebContents;
    this.origin = origin;
    this.timeoutMs = timeoutMs;
    this.documents = new Map();
    this.busy = new Set();
  }

  isAddin(frame) {
    try {
      const url = new URL(frame.url);
      if (url.origin !== this.origin || url.pathname !== "/taskpane.html" || frame.detached) return false;
      for (let parent = frame.parent, depth = 0; parent && depth < 16; parent = parent.parent, depth++) {
        if (classifyOfficeUrl(parent.url)?.application === "excel") return true;
      }
    } catch { /* Closed or navigated frame. */ }
    return false;
  }

  async evaluate(frame, source) {
    let timer;
    try {
      return await Promise.race([
        frame.executeJavaScript(source, false),
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(messages.TIMEOUT)), this.timeoutMs); }),
      ]);
    } finally { clearTimeout(timer); }
  }

  async listDocuments() {
    const found = new Map();
    for (const contents of this.getWebContents()) {
      try {
        if (contents.isDestroyed()) continue;
        for (const frame of contents.mainFrame.framesInSubtree) {
          if (found.size >= 32 || !this.isAddin(frame)) continue;
          try {
            const info = await this.evaluate(frame, "globalThis.TeamsOfficeBridge?.describe()");
            if (info?.ready === true && info.application === "excel" && /^[a-f0-9-]{36}$/i.test(info.documentId) && this.isAddin(frame)) {
              found.set(info.documentId, { frame, id: info.documentId, name: typeof info.name === "string" ? info.name.slice(0, 256) : "Excel workbook", application: "excel", webContentsId: contents.id });
            }
          } catch { /* An add-in may still be initializing. */ }
        }
      } catch { /* Window closed during discovery. */ }
    }
    this.documents = found;
    return { documents: [...found.values()].map(({ frame: _frame, ...metadata }) => metadata) };
  }

  async execute(documentId, operation, args) {
    const entry = this.documents.get(documentId);
    if (!entry || !this.isAddin(entry.frame)) throw new Error(messages.WRONG_DOCUMENT);
    if (this.busy.has(documentId)) throw new Error(messages.BUSY);
    this.busy.add(documentId);
    try {
      const request = { documentId, operation, args, deadline: Date.now() + this.timeoutMs - 1000 };
      // Only a structured command reaches our own, locally served task pane.
      // JSON encoding keeps cell data out of executable JavaScript syntax.
      const source = `globalThis.TeamsOfficeBridge?.execute(${JSON.stringify(request)})`;
      const result = await this.evaluate(entry.frame, source);
      if (!result?.ok) throw new Error(messages[result?.code] || messages.OFFICE_ERROR);
      return result.data;
    } catch (error) {
      if (Object.values(messages).includes(error.message)) throw error;
      throw new Error(messages.OFFICE_ERROR, { cause: error });
    } finally { this.busy.delete(documentId); }
  }
}

module.exports = { LiveOfficeBridge };
