// Real hidden Electron account selection against simulated Microsoft pages.
const { app, BrowserWindow, session, webContents } = require("electron");
const assert = require("node:assert/strict");
const { testPath } = require("../fixtures/testPaths");
const { DocumentTabs } = require("../../app/office/documentTabs");
const { inspectDocumentAccount, selectDocumentAccount, getRememberedDocumentAccount } = require("../../app/office/documentAccount");
const { documentAccountUi } = require("../../app/office/documentAccountUi");

app.setPath("userData", testPath("document-account-profile"));
app.disableHardwareAcceleration();
const deadline = setTimeout(() => app.exit(1), 45000);
app.whenReady().then(async () => {
  let main, tabs;
  try {
    let activeCase = "match", graphRequests = 0;
    const selections = [], loaded = new Set();
    const fixtureSession = session.fromPartition("persist:account-fixture");
    const html = (body) => new Response(body, { headers: { "Content-Type": "text/html" } });
    fixtureSession.protocol.handle("https", (request) => {
      const url = new URL(request.url);
      if (url.hostname === "teams.cloud.microsoft") return html(`<title>Teams fixture</title><script>
        globalThis.authRequests=0; globalThis.teamsForLinuxReactHandler = {acquireToken: async () => {authRequests++;throw new Error('Token requests forbidden in this test');}};</script>`);
      if (url.hostname === "graph.microsoft.com") { graphRequests++; throw new Error("Graph requests forbidden"); }
      if (url.hostname.endsWith(".sharepoint.com")) {
        const file = url.pathname.split("/").pop();
        if (!url.searchParams.has("signedin")) return Response.redirect(`https://login.microsoftonline.com/common/authorize?file=${file}`);
        selections.push(url.searchParams.get("picked"));
        loaded.add(file);
        return html("<title>Word fixture</title><h1>Signed in</h1>");
      }
      if (url.hostname === "login.microsoftonline.com") {
        const file = url.searchParams.get("file");
        const users = activeCase === "wrong" ? ["not-user@example.test"] : activeCase === "duplicate" ? ["USER@example.test", "user@example.test"] : ["not-user@example.test", "USER@example.test"];
        return html(`<title>Pick an account</title><div id="tilesHolder">${users.map((user, index) => `<div class="tile" role="button" tabindex="0" data-index="${index}"><span>Friendly name</span><span>${user}</span></div>`).join("")}</div>
          <input type="password" style="display:none">
          <script>globalThis.accountAudit={keys:0,trusted:false};
          for(const tile of document.querySelectorAll('.tile')) tile.addEventListener('keydown', event => {
            if(event.key!=='Enter')return; accountAudit.keys++; accountAudit.trusted=event.isTrusted;
            if(event.isTrusted)location.replace('https://tenant.sharepoint.com/${file}?signedin=1&picked='+tile.dataset.index);
          });</script>`);
      }
      throw new Error("Unexpected fixture destination");
    });
    main = new BrowserWindow({ show: true, webPreferences: { session: fixtureSession, sandbox: true, contextIsolation: true, nodeIntegration: false } });
    await main.loadURL("https://teams.cloud.microsoft/");
    main.focus();
    const focusedWindow = BrowserWindow.getFocusedWindow()?.id;
    const focusedContents = webContents.getFocusedWebContents()?.id;
    tabs = new DocumentTabs({ getMainWindow: () => main, inspector: { collect: () => ({ entries: [] }) }, manifests: { word: "fixture manifest" },
      liveBridge: { listDocuments: async () => ({ documents: webContents.getAllWebContents().filter((contents) => {
        const url = new URL(contents.getURL() || "about:blank");
        return url.hostname === "tenant.sharepoint.com" && loaded.has(url.pathname.split("/").pop()) && !contents.isLoadingMainFrame();
      }).map((contents) => ({ id: `fixture-${contents.id}`, webContentsId: contents.id, application: "word" })) }) },
    });
    for (const [name, expected] of [["match", "connected"], ["wrong", "ACCOUNT_NOT_FOUND"], ["duplicate", "ACCOUNT_AMBIGUOUS"], ["unavailable", "DOCUMENT_ACCOUNT_NOT_REMEMBERED"]]) {
      activeCase = name;
      const site = name === "unavailable" ? "https://empty.sharepoint.com" : "https://tenant.sharepoint.com";
      const { tabId } = tabs.open({ url: `${site}/${name}.docx` });
      await tabs.setupBridge(tabId, false, name === "match" && process.env.DOCUMENT_ACCOUNT_PHASE !== "read" ? "user@example.test" : undefined);
      let state, listing;
      const until = Date.now() + 8000;
      do {
        listing = await tabs.list();
        state = listing.tabs.find((tab) => tab.id === tabId);
        assert.equal(listing.visible, false);
        if (state.bridgeSetup.status !== "running") break;
        await new Promise((resolve) => setTimeout(resolve, 30));
      } while (Date.now() < until);
      assert.equal(state.bridgeSetup.code || state.bridgeSetup.status, expected, JSON.stringify(state.bridgeSetup));
      assert.equal(BrowserWindow.getFocusedWindow()?.id, focusedWindow);
      assert.equal(webContents.getFocusedWebContents()?.id, focusedContents);
      const inspection = await tabs.inspectBridgeSetup(tabId);
      assert.doesNotMatch(JSON.stringify(inspection), /user@example|fixture-token|authorize|signedin/);
      if (name === "match") {
        assert.equal(state.bridgeSetup.accountSelection, "dispatched");
        assert.equal(state.bridgeSetup.submitted, false);
        assert.deepEqual(selections, ["1"], "Choose the exact Teams identity, not the first tile");
      } else {
        const target = webContents.fromId(state.webContentsId);
        assert.deepEqual(await target.executeJavaScript("accountAudit"), { keys: 0, trusted: false });
        const oldProbe = await target.executeJavaScript(`(${documentAccountUi.toString()})(${JSON.stringify({ type: "probe", url: target.getURL(), deadline: Date.now() + 2000 })})`);
        assert.equal(oldProbe.picker, false, "Old token-lookup builds cannot enter account selection through the hot-reloaded probe");
      }
    }
    assert.equal(graphRequests, 0, "Never request Graph identity");
    assert.equal(await main.webContents.executeJavaScript("authRequests"), 0, "Never touch Teams' token provider");
    const current = await tabs.list();
    const pending = webContents.fromId(current.tabs.find((tab) => tab.bridgeSetup.code === "DOCUMENT_ACCOUNT_NOT_REMEMBERED").webContentsId);
    await pending.executeJavaScript("document.querySelector('.tile').click()");
    assert.equal(await getRememberedDocumentAccount(pending, "https://empty.sharepoint.com"), null, "Synthetic clicks cannot remember a choice");

    // An actual credential page is observed passively, never focused/submitted.
    const credential = new BrowserWindow({ show: false, webPreferences: { session: fixtureSession, sandbox: true } });
    await fixtureSession.protocol.unhandle("https");
    fixtureSession.protocol.handle("https", () => html('<input type="password"><div id="tilesHolder"><button>user@example.test</button></div>'));
    await credential.loadURL("https://login.microsoftonline.com/common/login");
    assert.deepEqual(await inspectDocumentAccount(credential.webContents), { picker: false, credentials: true, tileCount: 1 });
    assert.deepEqual(await selectDocumentAccount(credential.webContents, async () => ({ identifiers: ["user@example.test"], isCurrent: () => true }),
      new AbortController().signal, Date.now() + 2000), { performed: false, code: "ACCOUNT_PICKER_UNAVAILABLE" });
    credential.destroy();
    console.log("Hidden account selection passed: remembered site account, trusted Enter, no Teams token/Graph requests, preserved focus, duplicate/mismatch/missing-preference refusal and credential-page refusal");
    tabs.close(); main.destroy(); clearTimeout(deadline); app.quit();
  } catch (error) {
    console.error(error); tabs?.close(); main?.destroy(); app.exit(1);
  }
});
