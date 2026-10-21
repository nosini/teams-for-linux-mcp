const { randomUUID } = require("node:crypto");
const { codeMirrorCommand } = require("./codeMirrorCommand");
const { evaluateInFrame } = require("../mcp/evaluateInFrame");

const messages = {
  TIMEOUT: "The editor command timed out. Read the current text before deciding whether to write again.",
  UNAVAILABLE: "The CodeMirror editor changed or closed. List editors again.",
  STALE_READ: "The read token expired, was used, or the editor document changed. Read the text again.",
  INVALID_EDIT: "Use UTF-16 offsets inside the read window and at most 8,192 characters of LF text without NUL or Unicode paragraph separators.",
  READ_ONLY: "This editor is read-only.",
  COMPOSING: "The user is composing text. Wait until composition finishes before editing.",
  EDIT_NOT_CONFIRMED: "The editor did not confirm the exact requested text. It may have filtered or deferred the edit. Read the current text before deciding whether to write again.",
  BUSY: "Another command is running in this editor. Wait for it to finish before reading or writing again.",
  INTERNAL: "The CodeMirror operation failed. Read the current text before deciding whether to write again.",
};
function allowedFrame(frame) {
  try {
    const url = new URL(frame.url);
    return !frame.detached && url.protocol === "https:" && !url.username && !url.password && !url.port && url.hostname.endsWith(".sharepoint.com");
  } catch { return false; }
}

class CodeMirrorBridge {
  #inspector;
  #timeoutMs;
  #now;
  #registryKey = `__teamsMcpCodeMirror_${randomUUID()}`;
  #editors = new Map();
  #busy = new Set();

  constructor(inspector, { timeoutMs = 5000, now = Date.now } = {}) {
    this.#inspector = inspector;
    this.#timeoutMs = timeoutMs;
    this.#now = now;
  }
  async command(frame, url, operation, args = {}) {
    if (!allowedFrame(frame) || frame.url !== url) throw new Error(messages.UNAVAILABLE);
    // The page command runs synchronously, so discovery cannot interleave with
    // a read or replace. Only those wait for an earlier outcome in this frame.
    const exclusive = operation !== "list";
    if (exclusive && this.#busy.has(frame)) throw new Error(messages.BUSY);
    if (exclusive) this.#busy.add(frame);
    try {
      const request = { registryKey: this.#registryKey, operation, args, deadline: this.#now() + this.#timeoutMs - Math.min(500, this.#timeoutMs / 5) };
      const value = await evaluateInFrame(frame, `(${codeMirrorCommand.toString()})(${JSON.stringify(request)})`, this.#timeoutMs, messages.TIMEOUT);
      if (!allowedFrame(frame) || frame.url !== url) throw new Error(messages.UNAVAILABLE);
      if (!value?.ok) throw new Error(messages[value?.code] || messages.INTERNAL);
      if (!value.data || Buffer.byteLength(JSON.stringify(value.data)) > 32768) throw new Error(messages.INTERNAL);
      return value.data;
    } catch (error) {
      if (Object.values(messages).includes(error.message)) throw error;
      // Fetch/evaluation causes can contain document text or signed URLs.
      // eslint-disable-next-line preserve-caught-error
      throw new Error(messages.INTERNAL);
    } finally {
      if (exclusive) this.#busy.delete(frame);
    }
  }

  async listEditors() {
    const found = new Map(), editors = [], unavailableFrames = [];
    let unavailable = 0;
    const collected = this.#inspector.collect();
    const frames = collected.entries.filter((entry) => allowedFrame(entry.frame));
    const selected = frames.slice(0, 16);
    let truncated = frames.length > selected.length;
    // Each frame is independent; a stalled wrapper must not prevent inspecting
    // the actual editor nested beside it. Limit fanout and per-frame deadlines.
    const results = await Promise.allSettled(selected.map((entry) => this.command(entry.frame, entry.url, "list")));
    for (let index = 0; index < results.length; index++) {
      const entry = selected[index], response = results[index];
      if (response.status !== "fulfilled") { unavailableFrames.push(entry.metadata.id); continue; }
      const result = response.value;
      unavailable += Number.isInteger(result.unavailable) && result.unavailable > 0 ? Math.min(result.unavailable, 32) : 0;
      truncated ||= result.truncated === true;
      for (const candidate of Array.isArray(result.editors) ? result.editors.slice(0, 32) : []) {
        if (editors.length >= 32) { truncated = true; break; }
        if (!/^[a-f\d-]{36}$/i.test(candidate.id) || candidate.engine !== "codemirror6" || !Number.isInteger(candidate.length) || candidate.length < 0) continue;
        found.set(candidate.id, { frame: entry.frame, url: entry.url });
        editors.push({ id: candidate.id, frameId: entry.metadata.id, origin: entry.metadata.origin, engine: "codemirror6", length: candidate.length, readOnly: candidate.readOnly === true });
      }
    }
    this.#editors = found;
    return { editors, unavailable, unavailableFrames, truncated: truncated || collected.truncated };
  }

  async execute(operation, args) {
    const entry = this.#editors.get(args.editorId);
    if (!entry || !["read", "replace"].includes(operation)) throw new Error(messages.UNAVAILABLE);
    return this.command(entry.frame, entry.url, operation, args);
  }
}
module.exports = { CodeMirrorBridge, allowedFrame, messages };
