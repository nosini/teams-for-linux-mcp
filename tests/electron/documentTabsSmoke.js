// Real Electron tabs, hidden renderers, TLS and MCP; Office.js is simulated.
const { app, BrowserWindow, session, webContents } = require("electron");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const https = require("node:https");
const path = require("node:path");
const { Client } = require("@modelcontextprotocol/client");
const { StdioClientTransport } = require("@modelcontextprotocol/client/stdio");
const { OfficeMcp } = require("../../app/mcp");
const { excelFixture } = require("../fixtures/excelApi");
const { wordFixture } = require("../fixtures/wordApi");
const { officeImportFixture } = require("../fixtures/officeImport");

const profile = path.resolve(__dirname, "../../test-results/tabs-profile");
app.setPath("userData", profile);
app.disableHardwareAcceleration();
const testingSetup = process.env.DOCUMENT_TABS_SETUP === "1";
const testingNative = process.env.DOCUMENT_TABS_NATIVE === "1";
const deadline = setTimeout(() => app.exit(1), testingSetup ? 65000 : 45000);

app.whenReady().then(async () => {
  let client, mcp, main;
  try {
    main = new BrowserWindow({ show: false, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false } });
    mcp = new OfficeMcp({ config: { mcp: { enabled: true, addin: { enabled: true, port: 0 } } }, userDataPath: profile, getWebContents: () => webContents.getAllWebContents(), getMainWindow: () => main });
    await mcp.start();
    const manifest = await fs.readFile(path.join(profile, "mcp/excel-addin.xml"), "utf8");
    const origin = /DefaultValue="(https:\/\/127\.0\.0\.1:\d+)\/taskpane\.html"/.exec(manifest)[1];
    const cert = await fs.readFile(path.join(profile, "mcp/addin/localhost.pem"));
    const fixtureSession = session.defaultSession;
    fixtureSession.protocol.handle("https", async (request) => {
      const url = new URL(request.url);
      if (url.hostname === "teams.cloud.microsoft") return new Response(`<h1>Main Teams fixture</h1><script>globalThis.mainPageUnchanged = 123;</script><iframe src='${process.env.DOCUMENT_TABS_DIRECT_M365 === "1" ? "https://m365.cloud.microsoft/document/" : "https://tenant.sharepoint.com/_layouts/15/teamslogon.aspx?fixture=source"}'></iframe>`, { headers: { "Content-Type": "text/html" } });
      if (url.hostname === "tenant.sharepoint.com") {
        if (url.pathname.endsWith("teamslogon.aspx")) return new Response("<title>source.xlsx</title><iframe src='https://m365.cloud.microsoft/document/'></iframe>", { headers: { "Content-Type": "text/html" } });
        const application = url.pathname.endsWith(".docx") ? "word" : "excel";
        const file = url.pathname.split("/").pop();
        return new Response(`<iframe style="width:100%;height:100%;border:0" src="https://euc-${application}.officeapps.live.com/x/edit?file=${file}"></iframe>`, { headers: { "Content-Type": "text/html" } });
      }
      if (url.hostname === "m365.cloud.microsoft") return new Response("<iframe src='https://euc-excel.officeapps.live.com/x/edit?file=source.xlsx'></iframe>", { headers: { "Content-Type": "text/html" } });
      if (url.hostname === "store.office.com") return new Response(officeImportFixture({ origin, application: url.searchParams.get("application"), file: url.searchParams.get("file"), store: true, native: testingNative }), { headers: { "Content-Type": "text/html" } });
      if (url.hostname === "res.public.onecdn.static.microsoft") return new Response(officeImportFixture({ origin, application: url.searchParams.get("application"), file: url.searchParams.get("file"), panel: true, native: testingNative }), { headers: { "Content-Type": "text/html" } });
      if (url.hostname.endsWith(".officeapps.live.com")) {
        const application = url.hostname.includes("word") ? "word" : "excel", file = url.searchParams.get("file");
        const html = testingSetup ? officeImportFixture({ origin, application, file, native: testingNative }) : `<iframe style="width:100%;height:100%;border:0" src="${origin}/taskpane.html?application=${application}&file=${file}"></iframe>`;
        return new Response(html, { headers: { "Content-Type": "text/html" } });
      }
      if (url.hostname === "appsforoffice.microsoft.com") return new Response(`
        const application = new URL(location.href).searchParams.get('application');
        const file = new URL(location.href).searchParams.get('file');
        const clone = value => JSON.parse(JSON.stringify(value));
        const rangeShape = address => globalThis.TeamsExcelOperations.rangeShape(address);
        const fixture = application === 'word' ? (${wordFixture.toString()})() : (${excelFixture.toString()})();
        globalThis.Excel = fixture.excel; globalThis.Word = fixture.word;
        globalThis.Office = {HostType: {Excel: 'Excel', Word: 'Word'},
          context: {document: {url: 'https://tenant.sharepoint.com/' + file}, requirements: {isSetSupported: () => true}},
          onReady: async () => ({host: application === 'word' ? 'Word' : 'Excel'})};
      `, { headers: { "Content-Type": "text/javascript", "Cache-Control": "no-store" } });
      if (url.origin === origin) return new Promise((resolve, reject) => {
        https.get(request.url, { ca: cert }, (response) => {
          const chunks = []; response.on("data", (chunk) => chunks.push(chunk));
          response.once("end", () => resolve(new Response(Buffer.concat(chunks), { status: response.statusCode, headers: response.headers })));
        }).once("error", reject);
      });
      return new Response("No external network in fixture", { status: 403 });
    });
    await main.loadURL("https://teams.cloud.microsoft/");
    await fixtureSession.cookies.set({ url: "https://tenant.sharepoint.com", name: "session-fixture", value: "shared", secure: true });
    client = new Client({ name: "tabs-smoke", version: "1" });
    await client.connect(new StdioClientTransport({ command: "node", args: [path.resolve(__dirname, "../../app/mcp/stdio.js"), path.join(profile, "mcp/bridge.sock")] }));
    const call = async (name, args = {}) => {
      const response = await client.callTool({ name, arguments: args });
      assert.ok(!response.isError, response.content?.[0]?.text);
      return response.structuredContent;
    };
    const ids = [];
    const { editors } = await call("office_list_editors");
    const embedded = editors.find((editor) => editor.application === "excel" && editor.webContentsId === main.webContents.id);
    assert.ok(embedded, "The live-style embedding should expose an Office editor");
    assert.equal(editors.some((editor) => editor.surface === "document-wrapper" && editor.webContentsId === main.webContents.id), false);
    ids.push((await call("office_open_document", { editorId: embedded.id })).tabId);
    for (const file of ["second.xlsx", "third.docx"]) ids.push((await call("office_open_document", { url: `https://tenant.sharepoint.com/${file}` })).tabId);
    const workspace = BrowserWindow.getAllWindows().find((win) => win !== main);
    assert.ok(workspace);
    assert.equal(workspace.isVisible(), false);
    if (testingSetup) {
      const before = await call("office_list_tabs");
      assert.ok(before.tabs.every((tab) => !tab.connected), "No prior add-in panes or installation in this fixture");
      const selected = before.selectedTabId;
      let inspection;
      const inspectUntil = Date.now() + 5000;
      do {
        inspection = await client.callTool({ name: "office_inspect_bridge_setup", arguments: { tabId: ids[0] } });
        if (!inspection.isError && inspection.structuredContent.frames.some((frame) => frame.controls?.some((control) => control.label === "home"))) break;
        await new Promise((resolve) => setTimeout(resolve, 50));
      } while (Date.now() < inspectUntil);
      assert.ok(!inspection.isError);
      assert.equal(inspection.structuredContent.window.visible, false);
      assert.equal(inspection.structuredContent.window.selected, true);
      assert.ok(inspection.structuredContent.frames.some((frame) => frame.controls?.some((control) => control.label === "home")), "Inspection must reveal the real hidden ribbon controls");
      assert.doesNotMatch(JSON.stringify(inspection.structuredContent), /private cell|private document|private input|taskpane\.html|file=/);
      assert.equal((await call("office_list_tabs")).visible, false);
      assert.equal((await call("office_list_tabs")).selectedTabId, selected);
      assert.equal((await call("office_bridge_setup_step", { tabId: ids[0], action: "home", interaction: testingNative ? "auto" : "pointer" })).performed, true);
      assert.equal((await call("office_bridge_setup_step", { tabId: ids[0], action: "addins", interaction: "click" })).performed, true);
      const afterClick = await call("office_inspect_bridge_setup", { tabId: ids[0] });
      assert.ok(afterClick.frames.some((frame) => frame.actions?.includes("addins")), "Click-only must leave the fixture's pointer-sensitive menu closed");
      assert.equal((await call("office_bridge_setup_step", { tabId: ids[0], action: "addins", interaction: testingNative ? "nativeMouse" : "pointer" })).performed, true);
      const started = await Promise.all(ids.map((tabId) => call("office_setup_bridge", { tabId, retry: true })));
      assert.ok(started.every((result) => result.bridgeSetup.status === "running"));
      const repeated = await call("office_setup_bridge", { tabId: ids[0] });
      assert.equal(repeated.bridgeSetup.status, "running");
      assert.equal((await call("office_list_tabs")).selectedTabId, selected);
      const rejected = await client.callTool({ name: "office_setup_bridge", arguments: { tabId: ids[0], manifestPath: "/etc/passwd" } });
      assert.equal(rejected.isError, true, "Caller paths must never reach import automation");
    }
    let listing;
    const until = Date.now() + (testingSetup ? 35000 : 8000);
    do {
      listing = await call("office_list_tabs");
      if (listing.tabs.length === 3 && listing.tabs.every((tab) => tab.connected)) break;
      await new Promise((resolve) => setTimeout(resolve, 30));
    } while (Date.now() < until);
    if (testingSetup && listing.tabs.some((tab) => !tab.connected)) {
      for (const tab of listing.tabs) {
        for (const frame of webContents.fromId(tab.webContentsId).mainFrame.framesInSubtree) {
          console.log("Import fixture diagnostic", new URL(frame.url).origin, await frame.executeJavaScript("({ready:document.readyState, controls:[...document.querySelectorAll('button')].map(e=>({text:e.textContent,rects:e.getClientRects().length})), audit:globalThis.importAudit, error:globalThis.fixtureError})"));
        }
      }
    }
    assert.equal(listing.tabs.filter((tab) => tab.connected).length, 3, JSON.stringify(listing));
    assert.equal(listing.visible, false);
    if (testingSetup) {
      for (const tab of listing.tabs) {
        const contents = webContents.fromId(tab.webContentsId);
        const inspection = await call("office_inspect_bridge_setup", { tabId: tab.id });
        assert.ok(inspection.frames.some((frame) => frame.origin === "https://res.public.onecdn.static.microsoft" && frame.eligible), "The modern add-ins CDN panel must be inspected under the selected Office editor");
        const word = tab.id === ids[2];
        const storeOrigin = word ? "https://store.office.com/" : "https://res.public.onecdn.static.microsoft/";
        const store = contents.mainFrame.framesInSubtree.find((frame) => frame.url.startsWith(storeOrigin));
        assert.ok(store);
        const audit = await store.executeJavaScript("importAudit");
        assert.equal(audit.uploads, 1, "Exactly one upload per job, including after duplicate tool calls");
        assert.equal(audit.browse, 0, "No native picker opened");
        assert.equal(audit.file, word ? "word-addin.xml" : "excel-addin.xml");
        assert.equal(audit.type, word ? "Document" : "Workbook");
        const editor = contents.mainFrame.framesInSubtree.find((frame) => frame.url.includes(".officeapps.live.com/x/edit"));
        const sourceAudit = await editor.executeJavaScript("importAudit");
        assert.equal(sourceAudit.advanced, 1, "Advanced must be activated before opening the upload window");
        assert.equal(sourceAudit.gesture, true, "Authorized setup commands must have browser user activation");
        assert.equal(sourceAudit.blocked, tab.id === ids[0] ? 1 : 0);
        if (testingNative) assert.equal(sourceAudit.trusted, true, "Native activation must produce trusted click events");
        assert.equal(contents.debugger.isAttached(), false);
      }
      const again = await call("office_setup_bridge", { tabId: ids[0] });
      // Discovery may see the pane just before the job's next poll finishes.
      assert.ok(["running", "connected"].includes(again.bridgeSetup.status));
    }
    assert.equal(main.webContents.getURL(), "https://teams.cloud.microsoft/");
    assert.equal(await main.webContents.executeJavaScript("mainPageUnchanged"), 123);
    const docs = ids.map((id) => listing.tabs.find((tab) => tab.id === id).documents[0]);
    const targets = await Promise.all(docs.slice(0, 2).map(async (doc) => ({ documentId: doc.id, worksheetId: (await call("excel_list_worksheets", { documentId: doc.id })).worksheets[0].id, address: "A1" })));
    const reads = await Promise.all(targets.map((target) => call("excel_read_range", target)));
    const written = await Promise.all(targets.map((target, i) => call("excel_write_range", { ...target, readToken: reads[i].readToken, data: [[i + 41]] })));
    assert.deepEqual(written.map((result) => result.values), [[[41]], [[42]]]);
    const paragraphs = await call("word_read_paragraphs", { documentId: docs[2].id });
    const paragraph = paragraphs.paragraphs[0];
    const edited = await call("word_replace_text", { documentId: docs[2].id, paragraphId: paragraph.paragraphId, readToken: paragraph.readToken, findText: "Hello", replacementText: "Background" });
    assert.equal(edited.text, "Background from Word.");
    const firstContents = webContents.fromId(listing.tabs[0].webContentsId);
    assert.equal(firstContents.session, fixtureSession);
    assert.equal(firstContents.getBackgroundThrottling(), false);
    assert.equal(await firstContents.executeJavaScript("typeof globalThis.documentTabs"), "undefined");
    assert.equal((await firstContents.session.cookies.get({ name: "session-fixture" }))[0].value, "shared");
    await call("office_show_tab", { tabId: ids[2] });
    assert.equal(workspace.isVisible(), true);
    await workspace.webContents.executeJavaScript("documentTabs.request({action:'list'})");
    assert.match(await workspace.webContents.executeJavaScript("document.getElementById('status').textContent"), /Bridge connected/);
    await new Promise((resolve) => setTimeout(resolve, 500));
    const screenshot = await workspace.capturePage();
    await fs.writeFile(path.join(profile, "tabs-preview.png"), screenshot.toPNG());
    // BrowserWindow.capturePage captures its own webContents, excluding native
    // child views. Optional X11 capture checks the composed document surface.
    if (process.env.DOCUMENT_TABS_CAPTURE === "1") {
      const { promisify } = require("node:util");
      const { execFile } = require("node:child_process");
      await promisify(execFile)("import", ["-window", "root", path.join(profile, "tabs-native-preview.png")]);
    }
    // User closing the window hides it without closing any editor.
    workspace.close();
    assert.equal(workspace.isVisible(), false);
    assert.equal((await call("office_list_tabs")).tabs.filter((tab) => tab.connected).length, 3);
    assert.deepEqual((await call("excel_read_range", targets[0])).values, [[41]]);
    await call("office_show_tab", { tabId: ids[0] });
    await call("office_hide_tabs");
    assert.equal(workspace.isVisible(), false);
    const secondContents = webContents.fromId(listing.tabs[1].webContentsId);
    await secondContents.executeJavaScript("onbeforeunload = () => false; undefined", true);
    await call("office_close_tab", { tabId: ids[1] });
    await new Promise((resolve) => setTimeout(resolve, 100));
    assert.equal((await call("office_list_tabs")).tabs.length, 3, "Office's before-unload must be able to cancel a close");
    await secondContents.executeJavaScript("onbeforeunload = null");
    await call("office_close_tab", { tabId: ids[1] });
    await new Promise((resolve) => setTimeout(resolve, 100));
    assert.equal((await call("office_list_tabs")).tabs.length, 2);
    assert.deepEqual((await call("excel_read_range", targets[0])).values, [[41]]);
    console.log(`Document tabs smoke passed: real MCP startup/tools, three hidden independent Office pages, ${testingSetup ? "first-time import through English/Dutch UI and cross-origin dialogs, correct manifests, no native picker or duplicate uploads, " : ""}shared session, parallel Excel edits, Word edit, isolated remote pages, unchanged main Teams page, shell IPC, show/hide, close-one-preserve-others. Office.js and import UI were simulated.`);
  } finally {
    await client?.close();
    await mcp?.close();
    main?.destroy();
    clearTimeout(deadline);
  }
}).then(() => app.exit(0), (error) => { console.error(error); app.exit(1); });
