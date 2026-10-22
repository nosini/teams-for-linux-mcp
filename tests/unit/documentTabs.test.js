const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const { DocumentTabs, documentUrl, allowedNavigation, sourceDocumentUrl } = require("../../app/office/documentTabs");
const { createOfficeMcpServer } = require("../../app/mcp/server");
const fs = require("node:fs/promises");
const path = require("node:path");
const { startLocalEndpoint } = require("../../app/mcp/localEndpoint");
const { withOfficeClient } = require("../../scripts/diagnoseOfficeMcp");

function fixture(t) {
  const contents = new Map(), windows = [], views = [], handlers = new Map();
  let nextId = 1;
  class Contents extends EventEmitter {
    constructor(options = {}) { super(); this.id = nextId++; this.session = options.session || {}; this.mainFrame = { url: "", framesInSubtree: [] }; this.dead = false; contents.set(this.id, this); }
    isDestroyed() { return this.dead; }
    getUserAgent() { return "fixture-agent"; }
    setWindowOpenHandler(handler) { this.windowOpen = handler; }
    setAudioMuted(value) { this.muted = value; }
    send(...args) { this.sent = args; }
    async loadURL(url, options) { this.loaded = { url, options }; this.emit("did-start-loading"); }
    close(options) { this.closeOptions = options; if (options?.waitForBeforeUnload && this.cancelClose) return; this.dead = true; this.emit("destroyed"); }
  }
  class Window extends EventEmitter {
    constructor(options) {
      super(); this.options = options; this.visible = false; this.webContents = new Contents();
      this.children = []; this.contentView = { addChildView: (view) => this.children.push(view), removeChildView: (view) => { this.children = this.children.filter((item) => item !== view); } };
      windows.push(this);
    }
    removeMenu() {}
    async loadFile(file) { this.file = file; }
    getContentSize() { return [1200, 850]; }
    isVisible() { return this.visible; }
    isDestroyed() { return this.dead === true; }
    isMinimized() { return false; }
    show() { this.visible = true; }
    focus() { this.focused = true; }
    hide() { this.visible = false; }
    destroy() { this.dead = true; }
  }
  class View { constructor(options) { this.options = options; this.webContents = new Contents(options.webPreferences); views.push(this); } setBounds(bounds) { this.bounds = bounds; } setVisible(value) { this.visible = value; } }
  const app = new EventEmitter();
  const electron = { app, BrowserWindow: Window, WebContentsView: View, ipcMain: { handle: (name, handler) => handlers.set(name, handler), removeHandler: (name) => handlers.delete(name) }, webContents: { fromId: (id) => contents.get(id) } };
  const main = { webContents: new Contents() };
  const otherAccount = new Contents();
  const frame = { url: "https://tenant.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=private" };
  const inspector = { collect: () => ({ entries: [{ frame, metadata: { id: "7:8:9", webContentsId: otherAccount.id, surface: "document-wrapper" } }] }) };
  let documents = [];
  const bridge = { listDocuments: async () => ({ documents }) };
  const tabs = new DocumentTabs({ getMainWindow: () => main, inspector, liveBridge: bridge, manifests: { excel: "excel manifest", word: "word manifest" }, electron });
  t.after(() => tabs.close());
  return { tabs, windows, views, main, otherAccount, app, handlers, frame, inspector, setDocuments: (value) => { documents = value; } };
}

describe("persistent document tabs", () => {
  it("accepts document links and prevents arbitrary/local destinations and navigation", () => {
    assert.equal(documentUrl("https://tenant.sharepoint.com/sites/files/book.xlsx"), "https://tenant.sharepoint.com/sites/files/book.xlsx?web=1");
    assert.match(documentUrl("https://tenant.sharepoint.com/:w:/r/sites/file.docx?e=secret"), /web=1/);
    for (const url of ["file:///etc/passwd", "javascript:alert(1)", "https://evilsharepoint.com/a.xlsx", "https://tenant.sharepoint.com.evil.test/a.xlsx", "https://user:pass@tenant.sharepoint.com/a.xlsx", "https://tenant.sharepoint.com:444/a.xlsx", "https://tenant.sharepoint.com/home.aspx", "https://example.com/a.xlsx"]) assert.throws(() => documentUrl(url));
    assert.equal(allowedNavigation("https://login.microsoftonline.com/common/authorize", "https://tenant.sharepoint.com"), true);
    assert.equal(allowedNavigation("https://evil.test/", "https://tenant.sharepoint.com"), false);
  });

  it("opens a listed embedded Office editor through its own SharePoint container", async (t) => {
    const { tabs, views, otherAccount, inspector } = fixture(t);
    const frame = { url: "https://euc-excel.officeapps.live.com/x/edit?access_token=private", parent: {
      url: "https://m365.cloud.microsoft/document/", parent: { url: "https://tenant.sharepoint.com/_layouts/15/teamslogon.aspx?dest=private-document" },
    } };
    inspector.collect = () => ({ entries: [{ frame, metadata: { id: "7:8:9", webContentsId: otherAccount.id, surface: "office-frame", application: "excel" } }] });
    assert.equal((await tabs.list()).sources[0].editorId, "7:8:9");
    const opened = tabs.open({ editorId: "7:8:9" });
    assert.equal(opened.visible, false);
    assert.equal(views[0].webContents.loaded.url, frame.parent.parent.url);
    assert.equal(views[0].webContents.session, otherAccount.session);
    assert.doesNotMatch(JSON.stringify(opened), /access_token|dest=|private-document/);
    // A pasted arbitrary container URL still fails the public validation.
    assert.throws(() => tabs.open({ url: frame.parent.parent.url }), /document link/);
  });

  it("never substitutes a sibling document or opens a raw WOPI frame", () => {
    const metadata = { surface: "office-frame", application: "word" };
    const frame = { url: "https://euc-word.officeapps.live.com/w/edit?access_token=private", parent: { url: "https://teams.cloud.microsoft/" } };
    assert.throws(() => sourceDocumentUrl({ frame, metadata }), /no reusable/);
    frame.parent = { url: "https://evilsharepoint.com/signedin" };
    assert.throws(() => sourceDocumentUrl({ frame, metadata }), /no reusable/);
    frame.parent = { url: "https://tenant.sharepoint.com:444/container" };
    assert.throws(() => sourceDocumentUrl({ frame, metadata }), /no reusable/);
    frame.parent = { url: "https://tenant.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=chosen" };
    assert.match(sourceDocumentUrl({ frame, metadata }), /sourcedoc=chosen/);
  });

  it("opens a direct Microsoft 365 embedding without requiring any SharePoint frame", async (t) => {
    const { tabs, views, otherAccount, inspector } = fixture(t);
    const frame = { url: "https://euc-excel.officeapps.live.com/x/edit?access_token=private", parent: {
      url: "https://m365.cloud.microsoft/launch/excel?file=private-document", parent: { url: "https://teams.cloud.microsoft/" },
    } };
    inspector.collect = () => ({ entries: [{ frame, metadata: { id: "7:8:9", webContentsId: otherAccount.id, surface: "office-frame", application: "excel" } }] });
    assert.equal((await tabs.list()).sources[0].editorId, "7:8:9");
    const opened = tabs.open({ editorId: "7:8:9" });
    assert.equal(views[0].webContents.loaded.url, frame.parent.url);
    assert.doesNotMatch(JSON.stringify(opened), /private-document|access_token/);
    frame.parent.url = "https://m365.cloud.microsoft.evil.test/launch/excel?file=secret";
    const unavailable = await tabs.list();
    assert.equal(unavailable.sources.length, 0);
    assert.deepEqual(unavailable.unavailableSources[0].ancestorOrigins, ["https://m365.cloud.microsoft.evil.test", "https://teams.cloud.microsoft"]);
    assert.doesNotMatch(JSON.stringify(unavailable), /file=|secret/);
  });

  it("keeps separate hidden pages alive, binds the right account, and does not focus a foreground workspace for background opens", async (t) => {
    const { tabs, windows, views, main, otherAccount } = fixture(t);
    const first = tabs.open({ url: "https://tenant.sharepoint.com/first.xlsx" });
    assert.equal(windows[0].visible, false);
    assert.equal(views[0].webContents.session, main.webContents.session);
    assert.equal(views[0].options.webPreferences.backgroundThrottling, false);
    assert.equal(views[0].options.webPreferences.preload, undefined);
    tabs.show(first.tabId);
    windows[0].focused = false;
    const second = tabs.open({ editorId: "7:8:9" });
    assert.equal(views[1].webContents.session, otherAccount.session);
    assert.equal(second.selectedTabId, first.tabId);
    assert.equal(windows[0].focused, false);
    assert.equal(views[0].webContents.isDestroyed(), false);
    assert.equal(views[1].webContents.isDestroyed(), false);
    assert.equal(views[1].visible, false);
    assert.equal(tabs.open({ url: "https://tenant.sharepoint.com/first.xlsx" }).reused, true);
    assert.equal(views.length, 2);
    tabs.hide();
    assert.equal(windows[0].visible, false);
    assert.ok(views.every((view) => !view.webContents.isDestroyed()));
    assert.equal(main.webContents.loaded, undefined);
    assert.throws(() => tabs.open({ url: "https://tenant.sharepoint.com/a.xlsx", editorId: "7:8:9" }), /either/);
    assert.throws(() => tabs.open({ editorId: "missing" }), /current/);
  });

  it("distinguishes a loaded page from connected Office and invalidates discovery on navigation", async (t) => {
    const { tabs, views, setDocuments } = fixture(t);
    tabs.open({ url: "https://tenant.sharepoint.com/first.xlsx" });
    const target = views[0].webContents;
    target.emit("did-finish-load");
    assert.equal((await tabs.list()).tabs[0].connected, false);
    setDocuments([{ id: "document-id", webContentsId: target.id, application: "excel", name: "first.xlsx" }]);
    const state = await tabs.list();
    assert.equal(state.tabs[0].connected, true);
    assert.equal(state.tabs[0].documents[0].id, "document-id");
    target.emit("did-start-loading");
    setDocuments([]);
    assert.equal((await tabs.list()).tabs[0].connected, false);
    let prevented = false;
    target.emit("will-navigate", { preventDefault: () => { prevented = true; } }, "file:///etc/passwd");
    assert.equal(prevented, true);
    assert.equal((await tabs.list()).tabs[0].blockedNavigation, true);
    prevented = false;
    target.emit("will-redirect", { isMainFrame: false, preventDefault: () => { prevented = true; } }, "https://office.com/frame-auth");
    assert.equal(prevented, false);
    assert.doesNotMatch(JSON.stringify(state), /sourcedoc|private|https:/);
  });

  it("restricts control IPC to the local shell and keeps tabs alive on window close", async (t) => {
    const { tabs, views, windows, handlers } = fixture(t);
    tabs.open({ url: "https://tenant.sharepoint.com/first.xlsx" });
    const handle = handlers.get("office-tabs");
    const wc = windows[0].webContents;
    await assert.rejects(handle({ sender: views[0].webContents }, { action: "hide" }), /window only/);
    await assert.rejects(handle({ sender: wc, senderFrame: {} }, { action: "hide" }), /window only/);
    let prevented = false;
    windows[0].emit("close", { preventDefault: () => { prevented = true; } });
    assert.equal(prevented, true);
    assert.equal(views[0].webContents.isDestroyed(), false);
    assert.equal(windows[0].visible, false);
  });

  it("respects a cancelled tab close, bounds tabs, and disposes renderers and IPC on quit", async (t) => {
    const { tabs, views, app, handlers } = fixture(t);
    const first = tabs.open({ url: "https://tenant.sharepoint.com/first.xlsx" });
    views[0].webContents.cancelClose = true;
    assert.equal(tabs.closeTab(first.tabId).requested, true);
    assert.equal((await tabs.list()).tabs.length, 1);
    assert.deepEqual(views[0].webContents.closeOptions, { waitForBeforeUnload: true });
    for (let i = 1; i < 16; i++) tabs.open({ url: `https://tenant.sharepoint.com/file${i}.docx` });
    assert.throws(() => tabs.open({ url: "https://tenant.sharepoint.com/extra.xlsx" }), /Sixteen/);
    app.emit("before-quit");
    assert.ok(views.every((view) => view.webContents.isDestroyed()));
    assert.equal(handlers.has("office-tabs"), false);
    assert.equal(app.listenerCount("office-documents-show"), 0);
    assert.throws(() => tabs.show(), /closed/);
  });

  it("requires an explicit retry after a single-step upload attempt", async (t) => {
    const { tabs, views, windows } = fixture(t);
    const opened = tabs.open({ url: "https://tenant.sharepoint.com/fresh.xlsx" });
    const editor = { url: "https://euc-excel.officeapps.live.com/x/edit", executeJavaScript: async (source) =>
      source.includes('"type":"probe"') ? { actions: ["submit"], uploadInput: true, expectedFile: true } : { performed: true } };
    views[0].webContents.mainFrame.framesInSubtree = [editor];
    assert.equal((await tabs.bridgeSetupStep(opened.tabId, "submit")).performed, true);
    await assert.rejects(tabs.bridgeSetupStep(opened.tabId, "submit"), /already attempted/);
    assert.equal((await tabs.bridgeSetupStep(opened.tabId, "submit", "pointer", true)).performed, true);
    assert.equal(windows[0].visible, false);
  });

  it("exposes background opening and tab/document state through the real MCP SDK", async (t) => {
    const { tabs } = fixture(t);
    const root = path.resolve(__dirname, "../../test-results/mcp");
    await fs.mkdir(root, { recursive: true });
    const directory = await fs.mkdtemp(path.join(root, "tabs-"));
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer({ listEditors: () => ({ editors: [] }) }, null, null, null, null, tabs) });
    t.after(async () => { await endpoint.close(); await fs.rm(directory, { recursive: true, force: true }); });
    await withOfficeClient(endpoint.socketPath, async (client) => {
      const call = (name, args = {}) => client.callTool({ name, arguments: args });
      const opened = await call("office_open_document", { url: "https://tenant.sharepoint.com/file.xlsx" });
      assert.equal(opened.structuredContent.visible, false);
      assert.equal((await call("office_list_tabs")).structuredContent.tabs.length, 1);
      assert.equal((await call("office_open_document", { url: "file:///etc/passwd" })).isError, true);
      assert.equal((await call("office_open_document", { url: "https://tenant.sharepoint.com/file.xlsx", script: "evil()" })).isError, true);
      const { tabId } = opened.structuredContent;
      const started = await call("office_setup_bridge", { tabId });
      assert.equal(started.structuredContent.bridgeSetup.status, "running");
      assert.equal((await call("office_setup_bridge", { tabId, manifestPath: "/etc/passwd" })).isError, true);
      assert.equal((await call("office_setup_bridge", { tabId, script: "evil()" })).isError, true);
      assert.equal((await call("office_setup_bridge", { tabId, selector: "input" })).isError, true);
      const repeated = await call("office_setup_bridge", { tabId });
      assert.equal(repeated.structuredContent.bridgeSetup.status, "running");
      assert.equal((await call("office_list_tabs")).structuredContent.visible, false);
    });
  });
});
