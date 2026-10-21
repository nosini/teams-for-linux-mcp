// Fixed, read-only probes for identifying non-Office editors. Never return
// document text, model data, storage, tokens, full URLs or executable commands.
const { evaluateInFrame } = require("../mcp/evaluateInFrame");

function frameMetadata(value) {
  try {
    const url = new URL(value);
    if (url.protocol !== "https:" || url.username || url.password) return null;
    const host = url.hostname;
    const inspectable = ["sharepoint.com", "officeapps.live.com", "cloud.microsoft", "office.com", "office.net", "microsoft.com"].some((domain) => host === domain || host.endsWith(`.${domain}`));
    return { origin: url.origin, inspectable };
  } catch { return null; }
}

function probeEditorSurface() {
  const method = (value, key) => {
    // Inspect own data descriptors without invoking page-controlled getters.
    try {
      for (let object = value, depth = 0; object && depth < 4; object = Object.getPrototypeOf(object), depth++) {
        const descriptor = Object.getOwnPropertyDescriptor(object, key);
        if (descriptor) return typeof descriptor.value === "function";
      }
    } catch { /* Unavailable handle. */ }
    return false;
  };
  const ownValue = (value, key) => {
    try { return Object.getOwnPropertyDescriptor(value, key)?.value; } catch { return undefined; }
  };
  const codeMirrorView = (node) => {
    // CodeMirror 6 has used both cmView and cmTile DOM links. Their public
    // findFromDOM implementations walk to the linked tree root and its view.
    // Follow only own data properties: no getters or editor methods are run.
    for (const key of ["cmView", "cmTile"]) {
      let linked = ownValue(node, key);
      const seen = new Set();
      for (let depth = 0; linked && depth < 32 && !seen.has(linked); depth++) {
        seen.add(linked);
        const view = ownValue(linked, "view");
        if (view && method(view, "dispatch")) return view;
        linked = ownValue(linked, "parent");
      }
    }
    return null;
  };
  const count = (selector) => Math.min(document.querySelectorAll(selector).length, 100);
  const surfaces = {
    monaco: count(".monaco-editor"),
    codeMirror: count(".CodeMirror, .cm-editor"),
    prosemirror: count(".ProseMirror, [data-tiptap-editor]"),
    lexical: count("[data-lexical-editor]"),
    slate: count("[data-slate-editor]"),
    quill: count(".ql-editor"),
    contentEditable: count('[contenteditable="true"], [contenteditable="plaintext-only"]'),
    textareas: count("textarea"),
    textboxes: count('[role="textbox"]'),
  };
  const roots = [...document.querySelectorAll('.monaco-editor, .CodeMirror, .cm-editor, .ProseMirror, [data-lexical-editor], [data-slate-editor], .ql-editor, [contenteditable="true"], [contenteditable="plaintext-only"]')].slice(0, 20);
  const adapters = roots.map((node) => {
    const lexical = ownValue(node, "__lexicalEditor");
    const pm = ownValue(node, "pmViewDesc");
    const cm = ownValue(node, "CodeMirror");
    const quill = ownValue(node, "__quill");
    const cm6 = codeMirrorView(node);
    return {
      tag: node.tagName,
      lexicalEditorHandle: !!lexical,
      lexicalGetEditorState: method(lexical, "getEditorState"),
      lexicalUpdate: method(lexical, "update"),
      prosemirrorViewHandle: !!ownValue(pm, "view"),
      codeMirrorHandle: !!cm,
      codeMirrorViewLink: !!ownValue(node, "cmView"),
      codeMirrorTileLink: !!ownValue(node, "cmTile"),
      codeMirror6ViewHandle: !!cm6,
      codeMirrorDispatch: method(cm6, "dispatch"),
      codeMirrorStateHandle: !!ownValue(ownValue(cm6, "viewState"), "state"),
      codeMirrorContentDOMMatches: !!cm6 && ownValue(cm6, "contentDOM") === node,
      quillHandle: !!quill,
      reactFiberHandle: Object.keys(node).some((key) => key.startsWith("__reactFiber$") || key.startsWith("__reactInternalInstance$")),
    };
  });
  const monaco = ownValue(globalThis, "monaco");
  const editor = ownValue(monaco, "editor");
  return {
    surfaces, adapters,
    globals: {
      monaco: !!monaco,
      monacoGetEditors: method(editor, "getEditors"),
      monacoGetModels: method(editor, "getModels"),
      lexical: !!ownValue(globalThis, "Lexical"),
      fluid: !!ownValue(globalThis, "Fluid") || !!ownValue(globalThis, "fluid"),
      webpack: Object.keys(globalThis).some((key) => key.startsWith("webpackChunk")),
    },
  };
}
const PROBE_SOURCE = `(${probeEditorSurface.toString()})()`;
const surfaceNames = ["monaco", "codeMirror", "prosemirror", "lexical", "slate", "quill", "contentEditable", "textareas", "textboxes"];
const globalNames = ["monaco", "monacoGetEditors", "monacoGetModels", "lexical", "fluid", "webpack"];
const adapterNames = ["lexicalEditorHandle", "lexicalGetEditorState", "lexicalUpdate", "prosemirrorViewHandle", "codeMirrorHandle", "codeMirrorViewLink", "codeMirrorTileLink", "codeMirror6ViewHandle", "codeMirrorDispatch", "codeMirrorStateHandle", "codeMirrorContentDOMMatches", "quillHandle", "reactFiberHandle"];

class SurfaceInspector {
  constructor(getWebContents, { maxFrames = 64, timeoutMs = 2000 } = {}) {
    this.getWebContents = getWebContents;
    this.maxFrames = maxFrames;
    this.timeoutMs = timeoutMs;
  }

  collect() {
    const entries = [];
    for (const contents of this.getWebContents()) {
      try {
        if (contents.isDestroyed()) continue;
        for (const frame of contents.mainFrame.framesInSubtree) {
          try {
            const metadata = frameMetadata(frame.url);
            if (!metadata || frame.detached) continue;
            if (entries.length >= this.maxFrames) return { entries, truncated: true };
            entries.push({ frame, url: frame.url, metadata: { id: `${contents.id}:${frame.processId}:${frame.routingId}`, webContentsId: contents.id, ...metadata } });
          } catch { /* Detached frame. */ }
        }
      } catch { /* Closed window. */ }
    }
    return { entries, truncated: false };
  }

  listFrames() {
    const { entries, truncated } = this.collect();
    return { frames: entries.map(({ metadata }) => metadata), truncated };
  }

  async inspectFrame(id) {
    const entry = this.collect().entries.find(({ metadata }) => metadata.id === id);
    if (!entry || !entry.metadata.inspectable) throw new Error("Select an inspectable frame from editor_list_frames.");
    try {
      const value = await evaluateInFrame(entry.frame, PROBE_SOURCE, this.timeoutMs, "Probe timeout");
      if (entry.frame.detached || entry.frame.url !== entry.url) throw new Error("Frame navigated");
      const booleans = (data, keys) => Object.fromEntries(keys.map((key) => [key, data?.[key] === true]));
      return {
        ...entry.metadata, probeStatus: "completed",
        surfaces: Object.fromEntries(surfaceNames.map((key) => [key, Number.isInteger(value?.surfaces?.[key]) && value.surfaces[key] >= 0 ? Math.min(value.surfaces[key], 100) : 0])),
        globals: booleans(value?.globals, globalNames),
        adapters: Array.isArray(value?.adapters) ? value.adapters.slice(0, 20).map((adapter) => ({ tag: ["DIV", "SPAN", "TEXTAREA", "SECTION", "MAIN", "ARTICLE", "P", "PRE", "CODE"].includes(adapter?.tag) ? adapter.tag : "OTHER", ...booleans(adapter, adapterNames) })) : [],
        note: "Editor presence is diagnostic, not proof of live collaboration access. No document content was read or changed.",
      };
    } catch {
      return { ...entry.metadata, probeStatus: "unavailable" };
    }
  }
}
module.exports = { SurfaceInspector, frameMetadata, PROBE_SOURCE };
