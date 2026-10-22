// Serialized as one fixed function into a selected SharePoint frame. No imports,
// caller JavaScript, DOM content writes, Graph uploads or editor reconfiguration.
function codeMirrorCommand(request) {
  const messages = ["TIMEOUT", "UNAVAILABLE", "STALE_READ", "INVALID_EDIT", "READ_ONLY", "COMPOSING", "EDIT_NOT_CONFIRMED", "INTERNAL"];
  const fail = (code) => { throw Object.assign(new Error(code), { code }); };
  const now = () => Date.now();
  const check = () => { if (!Number.isFinite(request.deadline) || now() >= request.deadline) fail("TIMEOUT"); };
  const own = (object, key) => { try { return Object.getOwnPropertyDescriptor(object, key)?.value; } catch { return undefined; } };
  const supportedText = (text) => typeof text === "string" && text.isWellFormed() && text.length <= 8192 && !text.includes("\r") && !text.includes("\0") && !text.includes("\u2028") && !text.includes("\u2029");
  const splitsSurrogate = (doc, offset) => {
    if (offset <= 0 || offset >= doc.length) return false;
    const pair = doc.sliceString(offset - 1, offset + 1, "\n");
    return /[\uD800-\uDBFF][\uDC00-\uDFFF]/.test(pair);
  };
  const bounded = (value) => new TextEncoder().encode(JSON.stringify(value)).length <= 32768;
  const supportedDocument = (doc) => doc && Number.isInteger(doc.length) && doc.length >= 0 && typeof doc.sliceString === "function";
  const roots = () => [...document.querySelectorAll(".cm-editor")].slice(0, 33);
  const findView = (root) => {
    const content = root.querySelector(".cm-content");
    for (const node of [content, root]) for (const key of ["cmView", "cmTile"]) {
      let linked = own(node, key);
      const seen = new Set();
      for (let depth = 0; linked && depth < 32 && !seen.has(linked); depth++) {
        seen.add(linked);
        const view = own(linked, "view");
        if (view && typeof view.dispatch === "function" && view.dom === root && view.contentDOM === content && root.contains(content) && supportedDocument(view.state?.doc) && view.state.config && typeof view.state.config === "object") return view;
        linked = own(linked, "parent");
      }
    }
    return null;
  };
  try {
    check();
    if (typeof request.registryKey !== "string" || !/^__teamsMcpCodeMirror_[a-f\d-]{36}$/i.test(request.registryKey)) fail("INVALID_EDIT");
    const url = location.href;
    let registry = own(globalThis, request.registryKey);
    if (!registry || registry.url !== url) {
      if (request.operation !== "list") fail("UNAVAILABLE");
      registry = { url, entries: new Map(), byView: new WeakMap(), snapshots: new Map() };
      Object.defineProperty(globalThis, request.registryKey, { value: registry, configurable: true });
    }
    for (const [token, snapshot] of registry.snapshots) if (snapshot.expires <= now()) registry.snapshots.delete(token);
    if (request.operation === "list") {
      const found = new Map(), editors = [];
      const candidates = roots();
      let unavailable = 0;
      for (const root of candidates.slice(0, 32)) {
        const view = findView(root);
        if (!view) { unavailable++; continue; }
        const existing = registry.byView.get(view);
        const id = existing && registry.entries.has(existing) ? existing : crypto.randomUUID();
        registry.byView.set(view, id);
        found.set(id, { root, view });
        editors.push({ id, engine: "codemirror6", length: view.state.doc.length, readOnly: view.state.readOnly === true || view.contentDOM.getAttribute("contenteditable") !== "true" });
      }
      registry.entries = found;
      for (const [token, snapshot] of registry.snapshots) if (!found.has(snapshot.editorId)) registry.snapshots.delete(token);
      return { ok: true, data: { editors, unavailable, truncated: candidates.length > 32 } };
    }
    const args = request.args || {};
    const entry = registry.entries.get(args.editorId);
    if (!entry || !entry.root.isConnected || findView(entry.root) !== entry.view || location.href !== registry.url) fail("UNAVAILABLE");
    const { view } = entry;
    const state = view.state, doc = state.doc;
    if (!supportedDocument(doc)) fail("UNAVAILABLE");
    if (request.operation === "read") {
      const start = args.start ?? 0, count = args.count ?? 4096;
      if (!Number.isInteger(start) || start < 0 || start > doc.length || !Number.isInteger(count) || count < 1 || count > 8192) fail("INVALID_EDIT");
      if (splitsSurrogate(doc, start)) fail("INVALID_EDIT");
      let end = Math.min(start + count, doc.length);
      if (splitsSurrogate(doc, end)) end--;
      if (end === start && start < doc.length) fail("INVALID_EDIT");
      const text = doc.sliceString(start, end, "\n");
      if (typeof text !== "string" || text.length !== end - start) fail("UNAVAILABLE");
      const readToken = crypto.randomUUID();
      const data = { editorId: args.editorId, start, end, text, length: doc.length, nextStart: end < doc.length ? end : null, readToken };
      if (!bounded(data)) fail("INVALID_EDIT");
      while (registry.snapshots.size >= 64) registry.snapshots.delete(registry.snapshots.keys().next().value);
      registry.snapshots.set(readToken, { editorId: args.editorId, root: entry.root, view, doc, config: state.config, start, end, text, expires: now() + 120000 });
      return { ok: true, data };
    }
    if (request.operation !== "replace") fail("INVALID_EDIT");
    // Consume even unsuccessful attempts. An uncertain dispatch is never retried.
    const snapshot = registry.snapshots.get(args.readToken);
    registry.snapshots.delete(args.readToken);
    if (!snapshot || snapshot.expires <= now() || snapshot.editorId !== args.editorId || snapshot.root !== entry.root || snapshot.view !== view || snapshot.doc !== doc || snapshot.config !== state.config) fail("STALE_READ");
    if (!Number.isInteger(args.start) || !Number.isInteger(args.end) || args.start < snapshot.start || args.end > snapshot.end || args.end < args.start || !supportedText(args.text) || !bounded(args)) fail("INVALID_EDIT");
    if (splitsSurrogate(doc, args.start) || splitsSurrogate(doc, args.end)) fail("INVALID_EDIT");
    if (doc.sliceString(snapshot.start, snapshot.end, "\n") !== snapshot.text) fail("STALE_READ");
    if (state.readOnly === true || view.contentDOM.getAttribute("contenteditable") !== "true") fail("READ_ONLY");
    if (view.composing === true || view.compositionStarted === true) fail("COMPOSING");
    check();
    if (view.state.doc !== doc || view.state.config !== snapshot.config || location.href !== registry.url) fail("STALE_READ");
    // Synchronous compare + dispatch in the renderer's event loop. The host's
    // own dispatch override, transaction filters and update listeners still run.
    view.dispatch({ changes: { from: args.start, to: args.end, insert: args.text }, userEvent: "input" });
    const after = view.state.doc;
    const expectedLength = doc.length - (args.end - args.start) + args.text.length;
    if (!supportedDocument(after) || after.length !== expectedLength || after.sliceString(args.start, args.start + args.text.length, "\n") !== args.text) fail("EDIT_NOT_CONFIRMED");
    return { ok: true, data: { applied: true, editorId: args.editorId, start: args.start, end: args.end, insertedLength: args.text.length, length: after.length, saveStatus: "unverified" } };
  } catch (error) {
    return { ok: false, code: messages.includes(error?.code) ? error.code : "INTERNAL" };
  }
}
module.exports = { codeMirrorCommand };
