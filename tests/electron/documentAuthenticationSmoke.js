// Isolated fake authentication cookies; never connects to Microsoft. Run once
// with DOCUMENT_AUTH_PHASE=write, then read, against the same isolated profile.
const { app, BrowserWindow, session } = require("electron");
const assert = require("node:assert/strict");
const { testPath } = require("../fixtures/testPaths");
const { inspectDocumentAuthentication } = require("../../app/office/documentAuthentication");
const { DocumentTabs } = require("../../app/office/documentTabs");
const { RememberDocumentSession } = require("../../app/office/rememberDocumentSession");

app.setPath("userData", testPath(`document-auth-profile${process.env.DOCUMENT_AUTH_PERSIST === "1" ? "-persistent" : ""}`));
app.disableHardwareAcceleration();
const deadline = setTimeout(() => app.exit(1), 15000);
const persist = process.env.DOCUMENT_AUTH_PERSIST === "1";
app.whenReady().then(async () => {
  let window, tabs;
  try {
    const browserSession = session.fromPartition("persist:document-auth-fixture");
    const link = "https://tenant.sharepoint.com/test.docx";
    const keeper = persist ? new RememberDocumentSession({ days: 7 }) : null;
    keeper?.watch(browserSession, link);
    await keeper?.settle();
    if (process.env.DOCUMENT_AUTH_PHASE === "write") {
      for (const [name, url, persistent] of [
        ["ESTSAUTH", "https://login.microsoftonline.com/", false],
        ["ESTSAUTHPERSISTENT", "https://login.microsoftonline.com/", true],
        ["FedAuth", "https://tenant.sharepoint.com/", false],
        ["rtFa", "https://tenant.sharepoint.com/", false],
        ["FedAuth2", "https://tenant.sharepoint.com/", true],
        ["FedAuth1", "https://tenant.sharepoint.com/", false],
        ["FedAuth", "https://other.sharepoint.com/", false],
      ]) await browserSession.cookies.set({ url, name, value: "fixture-only", secure: true, httpOnly: true,
        ...(name === "rtFa" ? { domain: ".sharepoint.com", path: "/" } : {}),
        ...(persistent ? { expirationDate: Date.now() / 1000 + 3600 } : {}) });
      await keeper?.settle();
      const state = await inspectDocumentAuthentication(browserSession, link);
      assert.equal(state.persistentSession, true);
      assert.deepEqual(state.cookies, {
        microsoftSession: "session", microsoftPersistent: "persistent", sharepointSite: persist ? "persistent" : "session", sharepointRoot: persist ? "persistent" : "session",
      });
      assert.deepEqual(state.cookieAttributes.sharepointSite, { entries: [{ persistence: persist ? "persistent" : "session",
        scope: "opened_host", secure: true, httpOnly: true, rootPath: true, partitioned: false }], truncated: false });
      assert.deepEqual(state.cookieAttributes.sharepointRoot, { entries: [{ persistence: persist ? "persistent" : "session",
        scope: "parent_domain", secure: true, httpOnly: true, rootPath: true, partitioned: false }], truncated: false });
      await browserSession.cookies.flushStore();
    } else {
      assert.equal(process.env.DOCUMENT_AUTH_PHASE, "read");
      const state = await inspectDocumentAuthentication(browserSession, link);
      assert.equal(state.persistentSession, true);
      assert.deepEqual(state.cookies, {
        microsoftSession: "missing", microsoftPersistent: "persistent", sharepointSite: persist ? "persistent" : "missing", sharepointRoot: persist ? "persistent" : "missing",
      });
      assert.deepEqual(state.cookieAttributes.sharepointRoot, { entries: persist ? [{ persistence: "persistent", scope: "parent_domain",
        secure: true, httpOnly: true, rootPath: true, partitioned: false }] : [], truncated: false });
      assert.doesNotMatch(JSON.stringify(state), /fixture-only|tenant|token|value/);
      assert.equal((await browserSession.cookies.get({ url: "https://tenant.sharepoint.com/", name: "FedAuth1" })).length, persist ? 1 : 0);
      assert.equal((await browserSession.cookies.get({ url: "https://other.sharepoint.com/", name: "FedAuth" })).length, 0, "An unopened site's session cookie must disappear");
      assert.equal((await browserSession.cookies.get({ url: link, name: "FedAuth2" })).length, 1, "Existing server-persistent cookie stays intact");
    }
    let redirected = false, loaded = false;
    browserSession.protocol.handle("https", (request) => {
      const url = new URL(request.url);
      if (url.hostname === "login.microsoftonline.com") return new Response(`<title>Signing in</title>
        <script>setTimeout(() => location.replace('https://tenant.sharepoint.com/test.docx?web=1'), 500);</script>`, { headers: { "Content-Type": "text/html" } });
      if (!redirected) { redirected = true; return Response.redirect("https://login.microsoftonline.com/common/authorize"); }
      loaded = true;
      return new Response("<title>Word fixture</title><h1>Signed in</h1>", { headers: { "Content-Type": "text/html" } });
    });
    window = new BrowserWindow({ show: false, webPreferences: { session: browserSession,
      sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false } });
    let target;
    tabs = new DocumentTabs({ getMainWindow: () => window, inspector: { collect: () => ({ entries: [] }) },
      liveBridge: { listDocuments: async () => ({ documents: loaded && !target.isLoadingMainFrame() ? [{ id: "fixture-word", webContentsId: target.id, application: "word" }] : [] }) },
      manifests: { word: "fixture manifest" },
      documentSession: keeper,
    });
    const { tabId } = tabs.open({ url: link });
    target = require("electron").webContents.fromId((await tabs.list()).tabs[0].webContentsId);
    assert.equal(target.session, browserSession);
    await tabs.setupBridge(tabId);
    const updates = [], until = Date.now() + 5000;
    let listing;
    do {
      listing = await tabs.list();
      updates.push(listing.tabs[0].bridgeSetup);
      assert.equal(listing.visible, false);
      if (listing.tabs[0].bridgeSetup.status !== "running") break;
      await new Promise((resolve) => setTimeout(resolve, 20));
    } while (Date.now() < until);
    const result = listing.tabs[0].bridgeSetup;
    assert.equal(result.status, "connected");
    assert.equal(result.submitted, false);
    assert.ok(updates.some((state) => state.stage === "waiting_for_sign_in"));
    const inspection = await tabs.inspectBridgeSetup(tabId);
    assert.equal(inspection.authentication.persistentSession, true);
    assert.equal(inspection.application, null, "Read-only auth diagnostics also work without Office.js");
    assert.doesNotMatch(JSON.stringify(inspection), /fixture-only|test.docx|authorize/);
    assert.ok(BrowserWindow.getAllWindows().every((item) => !item.isVisible()));
    console.log(`Document authentication ${process.env.DOCUMENT_AUTH_PHASE}: cookie persistence and hidden silent redirect passed`);
    tabs.close();
    window.destroy();
    clearTimeout(deadline);
    app.quit();
  } catch (error) {
    console.error(error);
    tabs?.close();
    window?.destroy();
    app.exit(1);
  }
});
