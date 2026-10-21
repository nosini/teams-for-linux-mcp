const { classifyOfficeUrl } = require("./inspector");
const { editing } = require("./addin/operations");
const { messages } = require("./addin/operations");
const { evaluateInFrame } = require("../mcp/evaluateInFrame");

class LiveOfficeBridge {
  #getWebContents;
  #origin;
  #timeoutMs;
  #probeTimeoutMs;
  #documents = new Map();
  #busy = new Set();

  constructor(getWebContents, origin, { timeoutMs = 15000, probeTimeoutMs = 2000 } = {}) {
    this.#getWebContents = getWebContents;
    this.#origin = origin;
    this.#timeoutMs = timeoutMs;
    this.#probeTimeoutMs = probeTimeoutMs;
  }

  applicationOf(frame) {
    try {
      const url = new URL(frame.url);
      if (url.origin !== this.#origin || url.pathname !== "/taskpane.html" || frame.detached) return null;
      for (let parent = frame.parent, depth = 0; parent && depth < 16; parent = parent.parent, depth++) {
        const application = classifyOfficeUrl(parent.url)?.application;
        if (["excel", "word"].includes(application)) return application;
      }
    } catch { /* Closed or navigated frame. */ }
    return null;
  }

  isAddin(frame) { return !!this.applicationOf(frame); }

  evaluate(frame, source, timeoutMs = this.#timeoutMs) {
    return evaluateInFrame(frame, source, timeoutMs, messages.TIMEOUT);
  }

  async listDocuments() {
    const candidates = [];
    for (const contents of this.#getWebContents()) {
      try {
        if (contents.isDestroyed()) continue;
        for (const frame of contents.mainFrame.framesInSubtree) {
          if (candidates.length < 64 && this.isAddin(frame)) candidates.push({ frame, contents });
        }
      } catch { /* Window closed during discovery. */ }
    }
    // describe() only reads task-pane state. Probe panes in parallel with a
    // short deadline so a stalled or still-loading pane cannot delay the rest.
    const results = await Promise.allSettled(candidates.map(({ frame }) =>
      this.evaluate(frame, "globalThis.TeamsOfficeBridge?.describe()", this.#probeTimeoutMs)));
    const found = new Map();
    for (let index = 0; index < results.length && found.size < 32; index++) {
      const { frame, contents } = candidates[index], info = results[index].value;
      try {
        if (info?.ready === true && ["excel", "word"].includes(info.application) && info.application === this.applicationOf(frame) && /^[a-f0-9-]{36}$/i.test(info.documentId)) {
          found.set(info.documentId, { frame, id: info.documentId, name: typeof info.name === "string" ? info.name.slice(0, 256) : "Office document", application: info.application, webContentsId: contents.id });
        }
      } catch { /* Window closed during discovery. */ }
    }
    this.#documents = found;
    return { documents: [...found.values()].map(({ frame: _frame, ...metadata }) => metadata) };
  }

  async execute(documentId, operation, args) {
    const entry = this.#documents.get(documentId);
    if (!entry || this.applicationOf(entry.frame) !== entry.application) throw new Error(messages.WRONG_DOCUMENT);
    const operations = entry.application === "word" ? ["read_paragraphs", "read_paragraph", "replace_text", ...editing.wordOperations] :
      ["list_worksheets", "read_range", "write_range", "set_formulas", ...editing.excelOperations];
    if (!operations.includes(operation)) throw new Error(messages.WRONG_APPLICATION);
    if (this.#busy.has(documentId)) throw new Error(messages.BUSY);
    this.#busy.add(documentId);
    try {
      const request = { documentId, operation, args, deadline: Date.now() + this.#timeoutMs - 1000 };
      // Only a structured command reaches our own, locally served task pane.
      // JSON encoding keeps cell data out of executable JavaScript syntax.
      const source = `globalThis.TeamsOfficeBridge?.execute(${JSON.stringify(request)})`;
      const result = await this.evaluate(entry.frame, source);
      if (!result?.ok) throw new Error(messages[result?.code] || messages[entry.application === "word" ? "WORD_ERROR" : "OFFICE_ERROR"]);
      return result.data;
    } catch (error) {
      if (Object.values(messages).includes(error.message)) throw error;
      throw new Error(messages[entry.application === "word" ? "WORD_ERROR" : "OFFICE_ERROR"], { cause: error });
    } finally { this.#busy.delete(documentId); }
  }
}

module.exports = { LiveOfficeBridge };
