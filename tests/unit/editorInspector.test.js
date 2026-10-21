const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const { SurfaceInspector, frameMetadata, PROBE_SOURCE } = require("../../app/editor/inspector");
const { createOfficeMcpServer } = require("../../app/mcp/server");
const { Client } = require("@modelcontextprotocol/client");
const { StdioClientTransport } = require("@modelcontextprotocol/client/stdio");
const { startLocalEndpoint } = require("../../app/mcp/localEndpoint");
const fs = require("node:fs/promises");
const path = require("node:path");

const frame = (url, executeJavaScript = async () => ({})) => ({ url, processId: 1, routingId: 2, executeJavaScript });
const inspector = (frames, options) => new SurfaceInspector(() => [{ id: 3, isDestroyed: () => false, mainFrame: { framesInSubtree: frames } }], options);

describe("Non-Office editor diagnostics", () => {
  it("lists HTTPS origins without returning paths, names, queries, fragments or credentials", () => {
    const result = inspector([
      frame("https://tenant.sharepoint.com/preview?file=Naamloos.md&token=secret#private"),
      frame("https://third-party.test/file"),
      frame("file:///private"), frame("http://localhost/"), frame("https://user:password@tenant.sharepoint.com/"),
    ]).listFrames();
    assert.equal(result.frames.length, 2);
    assert.equal(result.frames[0].inspectable, true);
    assert.equal(result.frames[1].inspectable, false);
    assert.doesNotMatch(JSON.stringify(result), /Naamloos|secret|private|token|password/);
    assert.equal(frameMetadata("https://tenant.sharepoint.com.evil.test").inspectable, false);
    assert.equal(frameMetadata("https://evilsharepoint.com").inspectable, false);
  });

  it("never executes during discovery or probes an unapproved third-party frame", async () => {
    let executions = 0;
    const service = inspector([frame("https://third-party.test", () => { executions++; })]);
    service.listFrames();
    await assert.rejects(service.inspectFrame("3:1:2"), /inspectable frame/);
    assert.equal(executions, 0);
  });

  it("runs only the fixed probe, validates results and strips renderer-controlled content", async () => {
    const target = frame("https://tenant.sharepoint.com/", async (source, gesture) => {
      assert.equal(source, PROBE_SOURCE);
      assert.equal(gesture, false);
      return { surfaces: { monaco: 1000, contentEditable: -1, textareas: "private" }, globals: { monaco: true, fluid: "yes", text: "secret" }, adapters: [{ tag: "private-document-name", lexicalEditorHandle: true, text: "secret" }], text: "secret" };
    });
    const result = await inspector([target]).inspectFrame("3:1:2");
    assert.equal(result.probeStatus, "completed");
    assert.equal(result.surfaces.monaco, 100);
    assert.equal(result.surfaces.contentEditable, 0);
    assert.equal(result.globals.monaco, true);
    assert.equal(result.globals.fluid, false);
    assert.equal(result.adapters[0].tag, "OTHER");
    assert.doesNotMatch(JSON.stringify(result), /private|secret/);
  });

  it("checks same-origin navigation, timeouts, detachments and sanitized exceptions", async () => {
    for (const action of ["navigate", "detach", "throw", "hang"]) {
      const target = frame("https://tenant.sharepoint.com/first", async () => {
        if (action === "navigate") target.url = "https://tenant.sharepoint.com/second";
        if (action === "detach") target.detached = true;
        if (action === "throw") throw new Error("private file contents");
        if (action === "hang") return new Promise(() => {});
        return {};
      });
      const result = await inspector([target], { timeoutMs: 5 }).inspectFrame("3:1:2");
      assert.equal(result.probeStatus, "unavailable", action);
      assert.doesNotMatch(JSON.stringify(result), /private file contents/);
    }
  });

  it("bounds discovery and tolerates closed/detached windows", () => {
    const service = inspector([frame("https://teams.cloud.microsoft/"), { get url() { throw new Error("closed"); } }, frame("https://tenant.sharepoint.com/")], { maxFrames: 1 });
    assert.equal(service.listFrames().frames.length, 1);
    assert.equal(service.listFrames().truncated, true);
  });

  it("detects editor signatures without reading document text, invoking editor methods or getters", () => {
    let calls = 0;
    const forbidden = () => { calls++; throw new Error("Must remain read-only"); };
    const node = { tagName: "DIV", __lexicalEditor: { getEditorState: forbidden, update: forbidden }, pmViewDesc: { view: {} }, __reactFiber$test: {} };
    Object.defineProperty(node, "textContent", { get: forbidden });
    const scope = {
      document: { querySelectorAll: (selector) => selector === ".monaco-editor" ? [] : [node] },
      monaco: { editor: { getEditors: forbidden, getModels: forbidden } },
      webpackChunkTest: [],
    };
    Object.defineProperty(scope, "Fluid", { get: forbidden });
    const value = vm.runInNewContext(PROBE_SOURCE, scope);
    assert.equal(value.adapters[0].lexicalEditorHandle, true);
    assert.equal(value.adapters[0].lexicalGetEditorState, true);
    assert.equal(value.adapters[0].prosemirrorViewHandle, true);
    assert.equal(value.globals.monacoGetModels, true);
    assert.equal(calls, 0);
  });

  it("finds both CodeMirror 6 DOM-link generations through bounded parent data without calling getters", () => {
    for (const link of ["cmView", "cmTile"]) {
      let calls = 0;
      const view = { dispatch() { calls++; }, viewState: { state: {} } };
      const root = { view };
      const node = { tagName: "DIV", [link]: { parent: root } };
      view.contentDOM = node;
      Object.defineProperty(node[link], "rootView", { get() { calls++; throw new Error("No getters"); } });
      const result = vm.runInNewContext(PROBE_SOURCE, { document: { querySelectorAll: () => [node] } });
      assert.equal(result.adapters[0].codeMirror6ViewHandle, true, link);
      assert.equal(result.adapters[0].codeMirrorDispatch, true, link);
      assert.equal(result.adapters[0].codeMirrorStateHandle, true, link);
      assert.equal(result.adapters[0].codeMirrorContentDOMMatches, true, link);
      assert.equal(calls, 0);
      node[link].parent = node[link];
      const cyclic = vm.runInNewContext(PROBE_SOURCE, { document: { querySelectorAll: () => [node] } });
      assert.equal(cyclic.adapters[0].codeMirror6ViewHandle, false);
    }
  });

  it("registers strict, read-only MCP diagnostics without a JavaScript argument", async (t) => {
    const service = inspector([frame("https://tenant.sharepoint.com/", async () => ({ surfaces: { monaco: 1 } }))]);
    const root = path.resolve(__dirname, "../../test-results/mcp");
    await fs.mkdir(root, { recursive: true });
    const directory = await fs.mkdtemp(path.join(root, "surface-"));
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer({ listEditors: () => ({ editors: [] }) }, null, null, service) });
    const client = new Client({ name: "surface-test", version: "1" });
    t.after(async () => { await client.close(); await endpoint.close(); await fs.rm(directory, { recursive: true, force: true }); });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [path.resolve(__dirname, "../../app/mcp/stdio.js"), endpoint.socketPath] }));
    const tools = (await client.listTools()).tools;
    assert.equal(tools.find((tool) => tool.name === "editor_inspect_frame").annotations.readOnlyHint, true);
    const listing = await client.callTool({ name: "editor_list_frames", arguments: {} });
    assert.equal(listing.structuredContent.frames[0].origin, "https://tenant.sharepoint.com");
    const result = await client.callTool({ name: "editor_inspect_frame", arguments: { frameId: "3:1:2" } });
    assert.equal(result.structuredContent.surfaces.monaco, 1);
    const invalid = await client.callTool({ name: "editor_inspect_frame", arguments: { frameId: "3:1:2", script: "alert(1)" } });
    assert.equal(invalid.isError, true);
  });
});
