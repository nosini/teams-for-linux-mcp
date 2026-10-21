// Optional integration check against an actual CodeMirror 6 package + jsdom.
// Run in a test environment with @codemirror/view, @codemirror/state and jsdom
// available to Node. These are test dependencies, not runtime bridge dependencies.
const assert = require("node:assert/strict");
const { randomUUID } = require("node:crypto");
const { JSDOM } = require("jsdom");
const dom = new JSDOM("<!doctype html><html><body></body></html>", { url: "https://tenant.sharepoint.com/editor", pretendToBeVisual: true });
for (const key of ["window", "document", "location", "navigator", "MutationObserver", "Node", "HTMLElement", "DOMRect"]) Object.defineProperty(globalThis, key, { value: dom.window[key], configurable: true });
globalThis.getComputedStyle = dom.window.getComputedStyle.bind(dom.window);
const { EditorState, Compartment } = require("@codemirror/state");
const { EditorView } = require("@codemirror/view");
const { codeMirrorCommand } = require("../../app/editor/codeMirrorCommand");
const registryKey = `__teamsMcpCodeMirror_${randomUUID()}`;
const command = (operation, args = {}) => codeMirrorCommand({ registryKey, operation, args, deadline: Date.now() + 5000 });
const success = (operation, args) => { const result = command(operation, args); assert.equal(result.ok, true, JSON.stringify(result)); return result.data; };
const readonly = new Compartment();
const updates = [], dispatches = [];
const view = new EditorView({
  parent: document.body,
  state: EditorState.create({ doc: "# Title\n\nHello world\n", extensions: [readonly.of(EditorState.readOnly.of(false)), EditorView.updateListener.of((update) => { if (update.docChanged) updates.push(update.state.doc.toString()); })] }),
  dispatchTransactions(transactions, editor) { dispatches.push(...transactions); editor.update(transactions); },
});
let filtered;
try {
  const editors = success("list").editors;
  assert.equal(editors.length, 1);
  const editorId = editors[0].id;
  const before = success("read", { editorId, start: 9, count: 5 });
  assert.equal(before.text, "Hello");
  const written = success("replace", { editorId, readToken: before.readToken, start: 9, end: 14, text: "Greetings" });
  assert.equal(written.applied, true);
  assert.equal(written.saveStatus, "unverified");
  assert.equal(view.state.doc.toString(), "# Title\n\nGreetings world\n");
  assert.equal(dispatches.length, 1);
  assert.deepEqual(updates, ["# Title\n\nGreetings world\n"]);
  assert.equal(command("replace", { editorId, readToken: before.readToken, start: 9, end: 14, text: "Again" }).code, "STALE_READ");
  const stale = success("read", { editorId });
  view.dispatch({ changes: { from: view.state.doc.length, insert: "Coauthor\n" } });
  assert.equal(command("replace", { editorId, readToken: stale.readToken, start: 0, end: 1, text: "!" }).code, "STALE_READ");
  const selectionOnly = success("read", { editorId });
  view.dispatch({ selection: { anchor: 0 } });
  assert.equal(success("replace", { editorId, readToken: selectionOnly.readToken, start: 0, end: 0, text: "Intro\n" }).applied, true);
  view.dispatch({ effects: readonly.reconfigure(EditorState.readOnly.of(true)) });
  const locked = success("read", { editorId });
  assert.equal(command("replace", { editorId, readToken: locked.readToken, start: 0, end: 0, text: "blocked" }).code, "READ_ONLY");
  filtered = new EditorView({ parent: document.body, state: EditorState.create({ doc: "Filtered", extensions: [EditorState.transactionFilter.of((transaction) => transaction.docChanged ? [] : transaction)] }) });
  const filteredId = success("list").editors.find((editor) => editor.id !== editorId).id;
  const filteredRead = success("read", { editorId: filteredId });
  assert.equal(command("replace", { editorId: filteredId, readToken: filteredRead.readToken, start: 0, end: 8, text: "Changed" }).code, "EDIT_NOT_CONFIRMED");
  assert.equal(filtered.state.doc.toString(), "Filtered");
  console.log("Real CodeMirror model integration passed: discovery, model paging, host dispatch/update listeners, stale revisions, selection-only changes, read-only state and transaction filtering. SharePoint saving/coauthor transport was not tested.");
} finally {
  filtered?.destroy();
  view.destroy();
  dom.window.close();
}
