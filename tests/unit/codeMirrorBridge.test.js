const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { randomUUID } = require("node:crypto");
const fs = require("node:fs/promises");
const path = require("node:path");
const { CodeMirrorBridge, allowedFrame } = require("../../app/editor/codeMirrorBridge");
const { codeMirrorCommand } = require("../../app/editor/codeMirrorCommand");
const { SurfaceInspector } = require("../../app/editor/inspector");
const { createOfficeMcpServer } = require("../../app/mcp/server");
const { startLocalEndpoint } = require("../../app/mcp/localEndpoint");
const { Client } = require("@modelcontextprotocol/client");
const { StdioClientTransport } = require("@modelcontextprotocol/client/stdio");
const { codeMirrorFixture } = require("../fixtures/codeMirrorApi");

function fixture(text, link) {
  const api = codeMirrorFixture(text, link);
  const frame = { processId: 1, routingId: 2, get url() { return api.model.url; }, executeJavaScript: async (source, gesture) => { assert.equal(gesture, false); return api.evaluate(source); } };
  const inspector = new SurfaceInspector(() => [{ id: 3, isDestroyed: () => false, mainFrame: { framesInSubtree: [frame] } }]);
  const bridge = new CodeMirrorBridge(inspector, { now: () => api.model.now });
  const list = () => bridge.listEditors();
  const read = async (args = {}) => { const { editors } = await list(); return bridge.execute("read", { editorId: editors[0].id, ...args }); };
  return { ...api, bridge, frame, list, read };
}

describe("Native CodeMirror bridge", () => {
  it("discovers both DOM-link generations with stable editor IDs and no content/URL disclosure", async () => {
    for (const link of ["cmView", "cmTile"]) {
      const { list } = fixture("private content", link);
      const first = await list(), second = await list();
      assert.equal(first.editors[0].engine, "codemirror6");
      assert.equal(first.editors[0].id, second.editors[0].id);
      assert.doesNotMatch(JSON.stringify(first), /private|file=test/);
    }
    for (const url of ["https://evilsharepoint.com/", "https://tenant.sharepoint.com.evil.test/", "http://tenant.sharepoint.com/", "https://user:pass@tenant.sharepoint.com/", "https://teams.cloud.microsoft/"]) assert.equal(allowedFrame({ url }), false);
  });

  it("reads paged live model text and applies only the selected change through native dispatch", async () => {
    const { read, bridge, view, model } = fixture("# Title\n\nHello world\n");
    const before = await read({ start: 9, count: 5 });
    assert.equal(before.text, "Hello");
    assert.equal(before.nextStart, 14);
    const result = await bridge.execute("replace", { editorId: before.editorId, readToken: before.readToken, start: 9, end: 14, text: "Greetings" });
    assert.equal(result.applied, true);
    assert.equal(result.saveStatus, "unverified");
    assert.equal(view.state.doc.sliceString(0, view.state.doc.length), "# Title\n\nGreetings world\n");
    assert.equal(model.dispatches.length, 1);
    assert.deepEqual(JSON.parse(JSON.stringify(model.dispatches[0])), { changes: { from: 9, to: 14, insert: "Greetings" }, userEvent: "input" });
    assert.equal(model.updates.length, 1);
    await assert.rejects(bridge.execute("replace", { editorId: before.editorId, readToken: before.readToken, start: 9, end: 14, text: "Again" }), /expired, was used/);
  });

  it("inserts/deletes, allows selection changes and rejects changes anywhere in the document", async () => {
    const { read, bridge, view, doc } = fixture("Hello");
    const before = await read();
    view.state = { ...view.state, selection: { anchor: 1 } };
    await bridge.execute("replace", { editorId: before.editorId, readToken: before.readToken, start: 5, end: 5, text: "!" });
    const after = await read();
    await bridge.execute("replace", { editorId: after.editorId, readToken: after.readToken, start: 5, end: 6, text: "" });
    const stale = await read({ count: 2 });
    view.state = { ...view.state, doc: doc("HellO") };
    await assert.rejects(bridge.execute("replace", { editorId: stale.editorId, readToken: stale.readToken, start: 0, end: 2, text: "Hi" }), /editor document changed/);
  });

  it("rejects stale revisions, reused view configuration, expired tokens and wrong editor IDs", async () => {
    for (const change of ["config", "expiry", "editor", "view", "navigate", "detach"]) {
      const { read, bridge, view, model, root, content } = fixture("Hello");
      const before = await read();
      const args = { editorId: before.editorId, readToken: before.readToken, start: 0, end: 5, text: "Changed" };
      if (change === "config") view.state = { ...view.state, config: {} };
      if (change === "expiry") model.now += 120000;
      if (change === "editor") args.editorId = randomUUID();
      if (change === "view") content.cmTile.parent.view = { ...view };
      if (change === "navigate") model.url += "-other";
      if (change === "detach") root.isConnected = false;
      await assert.rejects(bridge.execute("replace", args), /changed|expired|closed/);
      assert.equal(model.dispatches.length, 0, change);
    }
  });

  it("checks read-only states and IME composition without dispatching", async () => {
    for (const field of ["readOnly", "editable", "composing", "compositionStarted"]) {
      const { read, bridge, view, model } = fixture();
      const before = await read();
      if (field === "readOnly") view.state = { ...view.state, readOnly: true };
      else if (field === "editable") model.editable = false;
      else view[field] = true;
      await assert.rejects(bridge.execute("replace", { editorId: before.editorId, readToken: before.readToken, start: 0, end: 0, text: "x" }), /read-only|composing/);
      assert.equal(model.dispatches.length, 0);
    }
  });

  it("rejects out-of-window writes, invalid offsets and unsafe newline/Unicode boundaries", async () => {
    for (const edit of [{ start: 0, end: 4, text: "x" }, { start: 2, end: 1, text: "x" }, { start: 1, end: 2, text: "a\rb" }, { start: 1, end: 2, text: "\0" }, { start: 1, end: 2, text: "\uD800" }]) {
      const { read, bridge, model } = fixture("Hello");
      const before = await read({ start: 1, count: 2 });
      await assert.rejects(bridge.execute("replace", { editorId: before.editorId, readToken: before.readToken, ...edit }), /UTF-16/);
      assert.equal(model.dispatches.length, 0);
    }
    const { read, bridge, model } = fixture("A😀B");
    const before = await read({ count: 2 });
    assert.equal(before.text, "A");
    assert.equal(before.nextStart, 1);
    const whole = await read();
    await assert.rejects(bridge.execute("replace", { editorId: whole.editorId, readToken: whole.readToken, start: 1, end: 2, text: "x" }), /UTF-16/);
    assert.equal(model.dispatches.length, 0);
    await assert.rejects(read({ start: 2 }), /UTF-16/);
  });

  it("reports filtered/deferred/throwing transactions as uncertain, never retries, and consumes tokens", async () => {
    for (const action of ["filtered", "deferred", "error"]) {
      const { read, bridge, model } = fixture("Hello");
      const before = await read();
      if (action === "error") model.error = new Error("private contents and URLs");
      else model[action] = true;
      const args = { editorId: before.editorId, readToken: before.readToken, start: 0, end: 5, text: "Changed" };
      await assert.rejects(bridge.execute("replace", args), (error) => !error.message.includes("private") && /failed|did not confirm/.test(error.message));
      await assert.rejects(bridge.execute("replace", args), /token expired/);
      assert.equal(model.dispatches.length, 1);
    }
  });

  it("keeps JavaScript-looking text as data and bounds renderer timeouts and discovery", async () => {
    const { read, bridge, scope, frame, list } = fixture("Hello");
    const before = await read();
    const text = 'x"); globalThis.compromised = true; //';
    await bridge.execute("replace", { editorId: before.editorId, readToken: before.readToken, start: 0, end: 5, text });
    assert.equal(scope.compromised, undefined);
    frame.executeJavaScript = async () => { throw new Error("private contents"); };
    assert.equal((await list()).unavailableFrames.length, 1);
    const timed = fixture();
    const target = await timed.read();
    timed.bridge.timeoutMs = 15;
    timed.frame.executeJavaScript = () => new Promise(() => {});
    await assert.rejects(timed.bridge.execute("read", { editorId: target.editorId }), /timed out/);
  });

  it("lists editors while a read runs in the same frame and keeps reads exclusive", async () => {
    const { bridge, frame, list } = fixture("Hello");
    const { editors: [editor] } = await list();
    const evaluate = frame.executeJavaScript;
    let release;
    const gate = new Promise((resolve) => { release = resolve; });
    frame.executeJavaScript = async (source, gesture) => { frame.executeJavaScript = evaluate; await gate; return evaluate(source, gesture); };
    const reading = bridge.execute("read", { editorId: editor.id });
    const listed = await list();
    assert.deepEqual(listed.editors.map(({ id }) => id), [editor.id]);
    assert.deepEqual(listed.unavailableFrames, []);
    await assert.rejects(bridge.execute("read", { editorId: editor.id }), /Another command is running/);
    release();
    assert.equal((await reading).text, "Hello");
    assert.equal((await bridge.execute("read", { editorId: editor.id })).text, "Hello");
  });

  it("supports the full flow through the official MCP client and rejects caller scripts", async (t) => {
    const { bridge, model } = fixture("Hello");
    const root = path.resolve(__dirname, "../../test-results/mcp");
    await fs.mkdir(root, { recursive: true });
    const directory = await fs.mkdtemp(path.join(root, "codemirror-"));
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer({ listEditors: () => ({ editors: [] }) }, null, null, null, bridge) });
    const client = new Client({ name: "codemirror-test", version: "1" });
    t.after(async () => { await client.close(); await endpoint.close(); await fs.rm(directory, { recursive: true, force: true }); });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [path.resolve(__dirname, "../../app/mcp/stdio.js"), endpoint.socketPath] }));
    const call = (name, args = {}) => client.callTool({ name, arguments: args });
    const editorId = (await call("codemirror_list_editors")).structuredContent.editors[0].id;
    const before = (await call("codemirror_read_text", { editorId })).structuredContent;
    assert.equal((await call("codemirror_replace_text", { editorId, readToken: before.readToken, start: 0, end: 5, text: "Updated" })).structuredContent.applied, true);
    assert.equal((await call("codemirror_read_text", { editorId })).structuredContent.text, "Updated");
    assert.equal((await call("codemirror_read_text", { editorId, script: "alert(1)" })).isError, true);
    assert.equal(model.dispatches.length, 1);
  });

  it("rejects expired commands in the renderer even when execution starts after the main-process timeout", async () => {
    const { evaluate, model } = codeMirrorFixture();
    const result = evaluate(`(${codeMirrorCommand.toString()})(${JSON.stringify({ registryKey: `__teamsMcpCodeMirror_${randomUUID()}`, operation: "list", deadline: model.now - 1 })})`);
    assert.equal(result.code, "TIMEOUT");
    assert.equal(model.dispatches.length, 0);
  });
});
