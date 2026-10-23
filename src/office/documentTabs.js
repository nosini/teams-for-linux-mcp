const { officeUrl } = require("./urlPolicy");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const { randomUUID } = require("node:crypto");

const SHELL = path.join(__dirname, "tabs", "index.html");
const HEADER_HEIGHT = 148;
const LIMIT = 16;

// Only document links, never caller-supplied renderer URLs or local files.
function documentUrl(value) {
  let url;
  try { url = new URL(value); } catch { throw new Error("Use a SharePoint Excel or Word document link."); }
  if (!officeUrl(value) || value.length > 16384) throw new Error("Use a SharePoint Excel or Word document link.");
  const sharepoint = url.hostname.endsWith(".sharepoint.com") &&
    (/\.(xlsx|xlsm|xlsb|xls|docx|docm|doc)$/i.test(url.pathname) ||
      /\/_layouts\/15\/(doc|wopiframe)\.aspx$/i.test(url.pathname) || /^\/:[xw]:\//i.test(url.pathname));
  const cloud = ["excel.cloud.microsoft", "word.cloud.microsoft"].includes(url.hostname) && url.pathname.startsWith("/open/");
  if (!sharepoint && !cloud) throw new Error("Use a SharePoint Excel or Word document link.");
  if (sharepoint) url.searchParams.set("web", "1");
  if (/\/_layouts\/15\/(doc|wopiframe)\.aspx$/i.test(url.pathname)) url.searchParams.set("action", "edit");
  return url.href;
}

function allowedNavigation(value, initialOrigin) {
  try {
    const url = officeUrl(value);
    return !!url && (url.origin === initialOrigin || !!officeUrl(value, "navigation"));
  } catch { return false; }
}

// Teams can host Office below a SharePoint container or directly below the
// Microsoft 365 file viewer. Resolve the listed frame's own ancestry, never a sibling
// frame or an agent-provided container URL. Raw Office frame URLs may depend on
// POSTed WOPI credentials, so do not open those directly as standalone pages.
function sourceDocumentUrl({ frame, metadata }) {
  let container = null;
  const office = metadata.surface === "office-frame" && ["excel", "word"].includes(metadata.application);
  if (!office && metadata.surface !== "document-wrapper") throw new Error("Select an Excel/Word editor or document wrapper.");
  for (let current = frame, depth = 0; current && depth < 24; current = current.parent, depth++) {
    try { return documentUrl(current.url); } catch { /* Look for the document's wrapper. */ }
    if (!office || current === frame) continue;
    const url = officeUrl(current.url);
    if (!url) continue;
    if (url.hostname.endsWith(".sharepoint.com") && (!container || new URL(container).hostname === "m365.cloud.microsoft")) container = url.href;
    else if (!container && url.hostname === "m365.cloud.microsoft") container = url.href;
  }
  if (container) return container;
  throw new Error("This editor has no reusable SharePoint or Microsoft 365 document container. Paste its SharePoint file link instead.");
}

function ancestorOrigins(frame) {
  const result = [];
  try {
    for (let parent = frame.parent, depth = 0; parent && depth < 24; parent = parent.parent, depth++) {
      const url = new URL(parent.url);
      if (url.protocol === "https:") result.push(url.origin);
    }
  } catch { /* Detached frame; retain only already inspected origins. */ }
  return result;
}

// Independent Office pages share a Teams session, but never its preload. The
// local shell owns the tabs; remote pages receive no application IPC API.
class DocumentTabs {
  #electron;
  #getMainWindow;
  #inspector;
  #bridge;
  #manifests;
  #documentSession;
  #locations;
  #window = null;
  #tabs = new Map();
  #selected = null;
  #closed = false;
  #documents = [];
  #listing = null;
  #showListener;
  #quitListener;

  constructor({ getMainWindow, inspector, liveBridge = null, manifests = null, documentSession = null, getWebContents = null, electron = require("electron") }) {
    this.#electron = electron;
    this.#getMainWindow = getMainWindow;
    this.#inspector = inspector;
    this.#bridge = liveBridge;
    this.#manifests = manifests;
    this.#documentSession = documentSession;
    const { DocumentLocations } = require("./documentLocations");
    this.#locations = new DocumentLocations(getWebContents || (() => electron.webContents.getAllWebContents()));
    this.#showListener = () => { try { this.show(); } catch { /* App is closing. */ } };
    this.#quitListener = () => this.close();
    electron.app.on("office-documents-show", this.#showListener);
    electron.app.on("before-quit", this.#quitListener);
    // Only the top frame of our local Documents shell can operate its tabs.
    electron.ipcMain.handle("office-tabs", async (event, request) => {
      if (!this.#window || event.sender !== this.#window.webContents ||
        event.senderFrame !== this.#window.webContents.mainFrame || event.senderFrame.url !== pathToFileURL(SHELL).href) throw new Error("Documents window only.");
      try {
        switch (request?.action) {
          case "list": return await this.list();
          case "open": return this.open({ url: request.url, editorId: request.editorId, foreground: true });
          case "select": return this.show(request.tabId);
          case "hide": return this.hide();
          case "setup": return await this.setupBridge(request.tabId, request.retry === true);
          case "close": return this.closeTab(request.tabId);
          default: throw new Error("Unknown document tab action.");
        }
      } catch { return { error: "Could not update this tab. Check the document link and current tab, then try again." }; }
    });
  }

  #sourceReport() {
    const ids = new Set([...this.#tabs.values()].map((tab) => tab.view.webContents.id));
    const available = [], unavailable = [];
    for (const entry of this.#inspector.collect().entries) {
      if (ids.has(entry.metadata.webContentsId)) continue;
      if (entry.metadata.surface !== "document-wrapper" && !["excel", "word"].includes(entry.metadata.application)) continue;
      try { available.push({ ...entry, sourceUrl: sourceDocumentUrl(entry) }); }
      catch { unavailable.push({ editorId: entry.metadata.id, application: entry.metadata.application, ancestorOrigins: ancestorOrigins(entry.frame) }); }
    }
    return { available, unavailable };
  }

  #ensureWindow() {
    if (this.#closed) throw new Error("Document tabs are closed.");
    if (this.#window) return;
    const { BrowserWindow } = this.#electron;
    this.#window = new BrowserWindow({
      title: "Documents — Teams Office", width: 1200, height: 850, minWidth: 720, minHeight: 450,
      show: false, autoHideMenuBar: true,
      webPreferences: { preload: path.join(__dirname, "tabs", "preload.js"), sandbox: true, contextIsolation: true, nodeIntegration: false, webviewTag: false },
    });
    this.#window.removeMenu();
    this.#window.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    this.#window.webContents.on("will-navigate", (event) => event.preventDefault());
    this.#window.webContents.on("did-finish-load", () => { this.#layout(); this.#publish(); });
    this.#window.on("resize", () => this.#layout());
    // Closing the workspace hides it. Only closing individual tabs unloads
    // Office, so a user dismissing this window doesn't interrupt an agent.
    this.#window.on("close", (event) => { if (!this.#closed) { event.preventDefault(); this.#window.hide(); } });
    void this.#window.loadFile(SHELL).catch(() => {});
  }

  #layout() {
    if (!this.#window) return;
    const [width, height] = this.#window.getContentSize();
    for (const tab of this.#tabs.values()) {
      tab.view.setBounds({ x: 0, y: HEADER_HEIGHT, width, height: Math.max(1, height - HEADER_HEIGHT) });
      tab.view.setVisible(tab.id === this.#selected);
    }
  }

  #state() {
    const sources = this.#sourceReport();
    return {
      visible: this.#window?.isVisible() === true, selectedTabId: this.#selected, bridgeAvailable: !!this.#bridge,
      tabs: [...this.#tabs.values()].map((tab) => {
        const documents = this.#documents.filter((doc) => doc.webContentsId === tab.view.webContents.id);
        return { id: tab.id, title: tab.title, webContentsId: tab.view.webContents.id,
          status: tab.status, blockedNavigation: tab.blockedNavigation, connected: documents.length > 0,
          bridgeSetup: tab.bridgeSetup ? { ...tab.bridgeSetup } : null,
          documents: documents.map(({ id, application, name }) => ({ id, application, name })) };
      }),
      sources: sources.available.map(({ metadata }) => ({ editorId: metadata.id, application: metadata.application, webContentsId: metadata.webContentsId })),
      unavailableSources: sources.unavailable,
    };
  }

  #publish() {
    if (this.#window && !this.#window.isDestroyed()) this.#window.webContents.send("office-tabs-state", this.#state());
  }

  async list() {
    if (this.#bridge) {
      // Shell polling and MCP callers share discovery instead of racing it.
      this.#listing ??= this.#bridge.listDocuments().finally(() => { this.#listing = null; });
      this.#documents = (await this.#listing).documents;
    }
    return this.#state();
  }

  open({ url, editorId, foreground = false } = {}) {
    if ((typeof url === "string") === (typeof editorId === "string")) throw new Error("Choose either a document URL or a listed editor ID.");
    const source = editorId ? this.#sourceReport().available.find((entry) => entry.metadata.id === editorId) : null;
    if (editorId && !source) throw new Error("Select a current Excel/Word editor or document wrapper from office_list_editors, or paste its SharePoint file link.");
    const link = source ? source.sourceUrl : documentUrl(url);
    const main = this.#getMainWindow();
    // An editor ID binds to its own account; pasted links use the main account.
    const contents = source ? this.#electron.webContents.fromId(source.metadata.webContentsId) : main?.webContents;
    return this.#open(link, contents, foreground);
  }

  #open(link, contents, foreground = false) {
    if (!contents || contents.isDestroyed()) throw new Error("The signed-in Teams session is unavailable.");
    const existing = [...this.#tabs.values()].find((tab) => tab.url === link && tab.view.webContents.session === contents.session);
    if (existing) {
      if (foreground) this.show(existing.id);
      return { tabId: existing.id, reused: true, ...this.#state() };
    }
    if (this.#tabs.size >= LIMIT) throw new Error("Sixteen document tabs are already open. Close one before opening another.");
    this.#ensureWindow();
    const view = new this.#electron.WebContentsView({ webPreferences: {
      session: contents.session, sandbox: true, contextIsolation: true, nodeIntegration: false,
      webviewTag: false, backgroundThrottling: false, spellcheck: true,
    } });
    let title = "Office document";
    try {
      const filename = decodeURIComponent(new URL(link).pathname.split("/").pop());
      if (/\.(xlsx|xlsm|xlsb|xls|docx|docm|doc)$/i.test(filename)) title = filename.slice(0, 160);
    } catch { /* Sharing links may not contain a readable file name. */ }
    const tab = { id: randomUUID(), url: link, view, title, status: "loading", blockedNavigation: false };
    this.#tabs.set(tab.id, tab);
    this.#window.contentView.addChildView(view);
    if (!this.#selected || foreground) this.#selected = tab.id;
    const target = view.webContents;
    const origin = new URL(link).origin;
    this.#documentSession?.watch(contents.session, link);
    target.setAudioMuted(true);
    target.setWindowOpenHandler(() => { tab.blockedNavigation = true; this.#publish(); return { action: "deny" }; });
    const guard = (event, nextUrl) => {
      // Redirects also fire for Office's authentication/add-in subframes;
      // restrict top-level navigation without breaking their own redirects.
      if (event.isMainFrame === false || allowedNavigation(event.url ?? nextUrl, origin)) return;
      event.preventDefault(); tab.blockedNavigation = true; this.#publish();
    };
    target.on("will-navigate", guard);
    target.on("will-redirect", guard);
    target.on("page-title-updated", (_event, title) => { if (title.trim()) tab.title = title.slice(0, 160); this.#publish(); });
    target.on("did-start-loading", () => { tab.status = "loading"; this.#forgetDocuments(target.id); this.#publish(); });
    target.on("did-finish-load", () => { tab.status = "loaded"; this.#publish(); });
    target.on("render-process-gone", () => { tab.status = "crashed"; this.#forgetDocuments(target.id); this.#publish(); });
    target.on("did-start-navigation", (event, _url, _inPlace, isMainFrame) => {
      if ((event.isMainFrame ?? isMainFrame) && !["waiting_for_editor", "waiting_for_sign_in", "selecting_account"].includes(tab.bridgeSetup?.stage)) tab.setupController?.abort();
    });
    target.once("destroyed", () => {
      tab.setupController?.abort();
      this.#tabs.delete(tab.id);
      this.#forgetDocuments(target.id);
      if (this.#selected === tab.id) this.#selected = this.#tabs.keys().next().value ?? null;
      if (!this.#closed) { this.#window.contentView.removeChildView(view); this.#layout(); this.#publish(); }
    });
    this.#layout();
    void target.loadURL(link, { userAgent: contents.getUserAgent() }).catch(() => {
      if (!this.#tabs.has(tab.id)) return;
      tab.status = "failed"; this.#forgetDocuments(target.id); this.#publish();
    });
    if (foreground) this.show(tab.id);
    this.#publish();
    return { tabId: tab.id, reused: false, ...this.#state() };
  }

  listLocations() { return this.#locations.list(); }

  resolveLocation(args) { return this.#locations.resolve(args); }

  listLocationFolders(locationId) { return this.#locations.listFolders(locationId); }

  listLocationFiles(locationId, name) { return this.#locations.listFiles(locationId, name); }

  async createDocument({ locationId, name }) {
    if (this.#closed || !this.#bridge || !this.#manifests) throw new Error("Enable the Office Bridge and restart Teams before creating a document.");
    if (this.#tabs.size >= LIMIT) throw new Error("Close a document tab before creating another document. No file was created.");
    const created = await this.#locations.create(locationId, name);
    // Creation is already successful. Never turn a later open/setup failure
    // into an ambiguous creation error that invites another upload.
    try {
      const opened = this.#open(documentUrl(created.item.url), created.contents);
      const setup = await this.setupBridge(opened.tabId);
      return { created: true, locationId: created.locationId, item: created.item, tabId: opened.tabId, bridgeSetup: setup.bridgeSetup };
    } catch {
      return { created: true, locationId: created.locationId, item: created.item, code: "DOCUMENT_CREATED_OPEN_INCOMPLETE",
        next: "The file exists. Do not create it again. Open it in the same Teams account and use office_open_document / office_setup_bridge." };
    }
  }

  #forgetDocuments(id) { this.#documents = this.#documents.filter((doc) => doc.webContentsId !== id); }

  async inspectBridgeSetup(tabId) {
    const tab = this.#tabs.get(tabId);
    if (!tab || this.#closed || tab.view.webContents.isDestroyed()) throw new Error("Document tab is no longer available. List tabs again.");
    const { inspectBridgeSetup } = require("./bridgeSetup");
    const { inspectDocumentAuthentication } = require("./documentAuthentication");
    const { inspectDocumentAccount } = require("./documentAccount");
    const authentication = await inspectDocumentAuthentication(tab.view.webContents.session, tab.url);
    return { tabId, window: { visible: this.#window.isVisible(), minimized: this.#window.isMinimized(),
      focused: this.#window.isFocused?.() === true, selected: this.#selected === tabId },
    authentication,
    accountPicker: await inspectDocumentAccount(tab.view.webContents, new URL(tab.url).origin),
    ...(await inspectBridgeSetup(tab.view.webContents)) };
  }

  async bridgeSetupStep(tabId, action, interaction = "pointer", retry = false) {
    const tab = this.#tabs.get(tabId);
    if (!tab || this.#closed) throw new Error("Document tab is no longer available. List tabs again.");
    if (tab.setupPromise) throw new Error("Bridge setup is running. Wait for its result before taking a single step.");
    if (!this.#bridge || !this.#manifests) throw new Error("Bridge setup is unavailable in this run.");
    if (action === "submit" && tab.bridgeSetup?.submitted && !retry) throw new Error("An upload was already attempted. Inspect its outcome before explicitly retrying.");
    const submitted = tab.bridgeSetup?.submitted === true || action === "submit";
    const stage = `step_${action}`;
    tab.bridgeSetup = { status: "running", stage, submitted };
    const { executeBridgeSetupStep } = require("./bridgeSetup");
    tab.setupPromise = executeBridgeSetupStep(tab.view.webContents, this.#manifests, action, interaction)
      .catch(() => ({ performed: false, code: "SETUP_UI_UNAVAILABLE" }));
    this.#publish();
    try {
      const result = await tab.setupPromise;
      if (this.#tabs.get(tabId) === tab && !this.#closed) {
        tab.bridgeSetup = { status: "needs_attention", stage, submitted, code: result.code || "STEP_COMPLETE" };
        this.#publish();
      }
      return { tabId, action, interaction, ...result };
    } finally { tab.setupPromise = null; }
  }

  async setupBridge(tabId, retry = false, accountHint) {
    const tab = this.#tabs.get(tabId);
    if (!tab || this.#closed) throw new Error("Document tab is no longer available. List tabs again.");
    if (!this.#bridge || !this.#manifests) throw new Error("Bridge setup is unavailable. Enable the Office Bridge and restart Teams.");
    if (tab.setupPromise || (tab.bridgeSetup && !retry)) return { tabId, bridgeSetup: { ...tab.bridgeSetup } };
    const { runBridgeSetup } = require("./bridgeSetup");
    const controller = new AbortController();
    tab.setupController = controller;
    tab.bridgeSetup = { status: "running", stage: "waiting_for_editor", submitted: false };
    const update = (state) => {
      if (this.#tabs.get(tabId) === tab && !this.#closed) { tab.bridgeSetup = state; this.#publish(); }
    };
    // Return immediately; callers poll office_list_tabs while Office loads its
    // menus and manifest. Repeated calls reuse this job instead of uploading twice.
    tab.setupPromise = Promise.resolve().then(() => runBridgeSetup({
      contents: tab.view.webContents, manifests: this.#manifests, signal: controller.signal, accountScope: new URL(tab.url).origin,
      getAccount: () => {
        const { getRememberedDocumentAccount } = require("./documentAccount");
        return getRememberedDocumentAccount(tab.view.webContents, new URL(tab.url).origin, accountHint);
      },
      listDocuments: async () => {
        const id = tab.view.webContents.id;
        const result = await this.#bridge.listDocuments({ webContentsId: id });
        this.#documents = [...this.#documents.filter((document) => document.webContentsId !== id), ...result.documents];
        return result;
      }, update,
    })).then(update).catch(() => update({ status: "needs_attention", stage: "stopped", code: "SETUP_UI_UNAVAILABLE" })).finally(() => { tab.setupPromise = null; });
    this.#publish();
    return { tabId, bridgeSetup: { ...tab.bridgeSetup } };
  }

  show(tabId) {
    if (tabId !== undefined && !this.#tabs.has(tabId)) throw new Error("Document tab is no longer available. List tabs again.");
    this.#ensureWindow();
    if (tabId) this.#selected = tabId;
    this.#layout();
    if (this.#window.isMinimized()) this.#window.restore();
    this.#window.show();
    this.#window.focus();
    this.#publish();
    return this.#state();
  }

  closeTab(tabId) {
    const tab = this.#tabs.get(tabId);
    if (!tab) throw new Error("Document tab is no longer available. List tabs again.");
    tab.setupController?.abort();
    // Office may cancel the close for an unsaved document. Returning from
    // this method is not a save confirmation or proof that the tab closed.
    tab.view.webContents.close({ waitForBeforeUnload: true });
    return { requested: true, tabId };
  }

  hide() {
    this.#window?.hide();
    this.#publish();
    return this.#state();
  }

  close() {
    if (this.#closed) return;
    this.#closed = true;
    this.#documentSession?.close();
    this.#locations.close();
    this.#electron.app.removeListener("office-documents-show", this.#showListener);
    this.#electron.app.removeListener("before-quit", this.#quitListener);
    this.#electron.ipcMain.removeHandler("office-tabs");
    for (const tab of this.#tabs.values()) {
      tab.setupController?.abort();
      if (!tab.view.webContents.isDestroyed()) tab.view.webContents.close();
    }
    this.#tabs.clear();
    this.#window?.destroy();
    this.#window = null;
  }
}

module.exports = { DocumentTabs, documentUrl, allowedNavigation, sourceDocumentUrl };
