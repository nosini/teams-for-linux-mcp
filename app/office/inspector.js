// Discovery only: never load Office.js into the editor, access document content,
// or call private editor APIs. An Office global alone is not an editing bridge.
const { evaluateInFrame } = require("../mcp/evaluateInFrame");

function classifyOfficeUrl(value) {
  let url;
  try {
    url = new URL(value);
  } catch {
    return null;
  }
  if (url.protocol !== "https:" || url.username || url.password) return null;
  const host = url.hostname;
  const isOffice = host.endsWith(".officeapps.live.com") ||
    ["excel.cloud.microsoft", "word.cloud.microsoft"].includes(host);
  const isWrapper = host.endsWith(".sharepoint.com") &&
    /\/_layouts\/15\/(doc|wopiframe)\.aspx$/i.test(url.pathname);
  if (!isOffice && !isWrapper) return null;
  const application = /(^|[.-])excel([.-]|$)/.test(host) ? "excel" :
    /(^|[.-])word([.-]|$)/.test(host) ? "word" : "unknown";
  return {
    application,
    surface: isWrapper ? "document-wrapper" : "office-frame",
    // Query strings and paths may contain document IDs, sharing tokens or names.
    origin: url.origin,
  };
}

function probeOfficeContext() {
  const office = globalThis.Office;
  const context = office?.context;
  const supports = (name) => {
    try {
      return context?.requirements?.isSetSupported(name, "1.1") === true;
    } catch {
      return false;
    }
  };
  return {
    officeGlobal: !!office,
    documentContext: !!context?.document,
    excelRun: typeof globalThis.Excel?.run === "function",
    wordRun: typeof globalThis.Word?.run === "function",
    excelApi11: supports("ExcelApi"),
    wordApi11: supports("WordApi"),
  };
}

const PROBE_SOURCE = `(${probeOfficeContext.toString()})()`;

class OfficeInspector {
  #getWebContents;
  #timeoutMs;
  #maxFrames;

  constructor(getWebContents, { timeoutMs = 2000, maxFrames = 64 } = {}) {
    this.#getWebContents = getWebContents;
    this.#timeoutMs = timeoutMs;
    this.#maxFrames = maxFrames;
  }

  collect() {
    const entries = [];
    let truncated = false;
    for (const contents of this.#getWebContents()) {
      try {
        if (contents.isDestroyed()) continue;
        for (const frame of contents.mainFrame.framesInSubtree) {
          // A frame can disappear during navigation, without invalidating siblings.
          try {
            const metadata = classifyOfficeUrl(frame.url);
            if (!metadata) continue;
            if (entries.length >= this.#maxFrames) {
              truncated = true;
              break;
            }
            entries.push({
              frame,
              metadata: {
                id: `${contents.id}:${frame.processId}:${frame.routingId}`,
                webContentsId: contents.id,
                ...metadata,
              },
            });
          } catch { /* Detached frame. */ }
        }
      } catch { /* Closed window or renderer without a main frame. */ }
    }
    return { entries, truncated };
  }

  listEditors() {
    const { entries, truncated } = this.collect();
    return { editors: entries.map(({ metadata }) => metadata), truncated };
  }

  async inspectEditor(id) {
    // Resolve afresh on every request; never retain a frame across navigation.
    const entry = this.collect().entries.find(({ metadata }) => metadata.id === id);
    if (!entry) throw new Error("Editor is no longer available. Call office_list_editors again.");
    try {
      const capabilities = await evaluateInFrame(entry.frame, PROBE_SOURCE, this.#timeoutMs, "Probe timed out");
      // Recheck after execution, since navigation can happen during a probe.
      if (classifyOfficeUrl(entry.frame.url)?.origin !== entry.metadata.origin) {
        throw new Error("Frame navigated");
      }
      const fields = ["officeGlobal", "documentContext", "excelRun", "wordRun", "excelApi11", "wordApi11"];
      return {
        ...entry.metadata,
        probeStatus: "completed",
        apiPresence: Object.fromEntries(fields.map((key) => [key, capabilities?.[key] === true])),
        bridgeStatus: "not-connected",
        note: "API presence does not prove access to the live document. No document content was read or changed.",
      };
    } catch {
      // Renderer exceptions can contain URLs or document data. Do not forward them.
      return { ...entry.metadata, probeStatus: "unavailable", bridgeStatus: "not-connected" };
    }
  }
}

module.exports = { OfficeInspector, classifyOfficeUrl, PROBE_SOURCE };
