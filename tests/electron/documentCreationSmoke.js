// Uses a fake SharePoint service and an isolated profile. Never contacts
// Microsoft or creates real cloud files. No external dependencies/downloads.
const { app, BrowserWindow, session, webContents } = require("electron");
const assert = require("node:assert/strict");
const { testPath } = require("../fixtures/testPaths");
const { DocumentTabs } = require("../../app/office/documentTabs");

app.setPath("userData", testPath("document-create-profile"));
app.disableHardwareAcceleration();
const deadline = setTimeout(() => app.exit(1), 20000);
app.whenReady().then(async () => {
  let main, source, tabs;
  try {
    const browserSession = session.fromPartition("persist:document-create-fixture");
    const origin = "https://tenant.sharepoint.com", web = origin + "/sites/team";
    const folder = "/sites/team/Shared Documents/Work";
    const sourceAgent = "SharePointSourceFixture/1";
    const files = new Map(); let posts = 0, requests = 0;
    const metadata = (name, parent = folder) => ({ Name: name, ServerRelativeUrl: parent + "/" + name, UniqueId: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" });
    browserSession.protocol.handle("https", async (request) => {
      const url = new URL(request.url), pathname = decodeURIComponent(url.pathname);
      assert.equal(url.origin, origin, "No requests to Teams token providers, Graph or other sites");
      if (pathname.includes("/_api/")) {
        requests++;
        const parent = pathname.match(/GetFolderByServerRelativePath\(decodedurl='((?:''|[^'])*)'\)/)?.[1].replace(/''/g, "'") || folder;
        assert.equal(request.headers.get("User-Agent"), sourceAgent, "Use the source frame's browser identity");
        if (pathname.endsWith("/contextinfo")) return Response.json({ FormDigestValue: "fixture-digest" });
        if (pathname.includes("/Files/AddUsingPath")) {
          posts++;
          assert.equal(request.method, "POST");
          assert.equal(request.headers.get("X-RequestDigest"), "fixture-digest");
          assert.ok(pathname.includes(",overwrite=false)"));
          const name = pathname.match(/AddUsingPath\(decodedurl='(.*)',overwrite=false\)/)[1].replace(/''/g, "'");
          const key = parent + "/" + name;
          if (files.has(key)) return new Response("private-server-error", { status: 409 });
          const bytes = Buffer.from(await request.arrayBuffer());
          assert.equal(bytes.subarray(0, 4).toString("hex"), "504b0304");
          files.set(key, metadata(name, parent));
          return Response.json(metadata(name, parent));
        }
        if (pathname.endsWith("/Files")) return Response.json({ value: [...files.values()].filter((item) => item.ServerRelativeUrl.slice(0, item.ServerRelativeUrl.lastIndexOf("/")) === parent) });
        if (pathname.endsWith("/Folders")) return Response.json({ value: [metadata("Child", parent)] });
        if (parent.endsWith("/Missing")) return new Response("private-missing-body", { status: 404 });
        if (parent.endsWith("/Denied")) return new Response("private-denied-body", { status: 403 });
        return Response.json({ Name: parent.split("/").at(-1), ServerRelativeUrl: parent });
      }
      return new Response("<title>SharePoint fixture</title><script>globalThis._spPageContextInfo={webAbsoluteUrl:" +
        JSON.stringify(origin) + "};const nativeFetch=globalThis.fetch;globalThis.fixtureRequestCount=0;" +
        "globalThis.fetch=(url,options)=>{if(options.credentials!=='include'||options.mode!=='same-origin')throw Error('Browser context missing');" +
        "globalThis.fixtureRequestCount++;return nativeFetch(url,options)};</script><p>Fixture</p>", { headers: { "Content-Type": "text/html" } });
    });
    await browserSession.cookies.set({ url: origin, name: "FedAuth", value: "fixture-only", secure: true, httpOnly: true });
    main = new BrowserWindow({ show: false });
    source = new BrowserWindow({ show: false, webPreferences: { session: browserSession, sandbox: true, contextIsolation: true, nodeIntegration: false } });
    source.webContents.setUserAgent(sourceAgent);
    await source.loadURL(origin + "/_layouts/15/filebrowser.aspx?id=" + encodeURIComponent(folder));
    let target;
    tabs = new DocumentTabs({ getMainWindow: () => main, getWebContents: () => [source.webContents], inspector: { collect: () => ({ entries: [] }) },
      liveBridge: { listDocuments: async () => ({ documents: target && !target.isLoadingMainFrame() ? [{ id: "fixture-created-word", application: "word", webContentsId: target.id }] : [] }) },
      manifests: { word: "fixture manifest" },
    });
    const report = await tabs.listLocations();
    assert.equal(report.locations.length, 1);
    assert.equal(report.locations[0].folderPath, folder);
    const created = await tabs.createDocument({ locationId: report.locations[0].locationId, name: "New report #1% O'Brien" });
    assert.equal(created.created, true);
    assert.equal(created.item.name, "New report #1% O'Brien.docx");
    target = webContents.fromId((await tabs.list()).tabs[0].webContentsId);
    assert.equal(target.session, browserSession, "Creation and opening retain the source account");
    assert.notEqual(target.session, main.webContents.session);
    const until = Date.now() + 5000;
    let state;
    do {
      state = await tabs.list();
      assert.equal(state.visible, false);
      if (state.tabs[0].bridgeSetup.status !== "running") break;
      await new Promise((resolve) => setTimeout(resolve, 20));
    } while (Date.now() < until);
    assert.equal(state.tabs[0].bridgeSetup.status, "connected");
    assert.equal(state.tabs[0].bridgeSetup.submitted, false, "The simulated existing Bridge needs no upload");
    await assert.rejects(tabs.createDocument({ locationId: report.locations[0].locationId, name: created.item.name }), /DOCUMENT_NAME_EXISTS/);
    assert.equal(files.size, 1);
    assert.equal(posts, 2, "One creation and one explicitly requested conflicting attempt, no retries");
    const listed = await tabs.listLocationFiles(report.locations[0].locationId, created.item.name);
    assert.equal(listed.files[0].name, created.item.name);
    assert.doesNotMatch(JSON.stringify([report, created, listed]), /fixture-digest|fixture-only|private-server-error/);
    assert.ok(BrowserWindow.getAllWindows().every((window) => !window.isVisible()));
    assert.ok(requests >= 5);
    const otherWeb = origin + "/sites/other/subsite", otherFolder = "/sites/other/subsite/Documents/Reports #1% O'Brien";
    const resolved = (await tabs.resolveLocation({ sourceLocationId: report.locations[0].locationId, folderPath: otherFolder, webUrl: otherWeb })).location;
    const children = await tabs.listLocationFolders(resolved.locationId);
    assert.equal(children.folders[0].folderPath, otherFolder + "/Child");
    const childCreated = await tabs.createDocument({ locationId: children.folders[0].locationId, name: "Other site report" });
    assert.equal(childCreated.created, true);
    assert.equal(decodeURIComponent(new URL(childCreated.item.url).pathname), otherFolder + "/Child/Other site report.docx");
    assert.equal(source.webContents.getURL(), origin + "/_layouts/15/filebrowser.aspx?id=" + encodeURIComponent(folder), "Targeting another site never navigates the source view");
    const secondTab = (await tabs.list()).tabs.find((tab) => tab.id === childCreated.tabId);
    assert.equal(webContents.fromId(secondTab.webContentsId).session, browserSession);
    assert.equal((await tabs.listLocationFiles(children.folders[0].locationId)).files[0].name, childCreated.item.name);
    assert.equal(files.size, 2);
    assert.equal(posts, 3, "One additional creation in a browsed destination, without retries");
    await assert.rejects(tabs.resolveLocation({ sourceLocationId: report.locations[0].locationId, folderPath: folder + "/Missing" }), /SHAREPOINT_NOT_FOUND/);
    await assert.rejects(tabs.resolveLocation({ sourceLocationId: report.locations[0].locationId, folderPath: folder + "/Denied" }), /SHAREPOINT_ACCESS_DENIED/);
    await assert.rejects(tabs.resolveLocation({ sourceLocationId: report.locations[0].locationId, folderPath: otherFolder, webUrl: "https://other.sharepoint.com/sites/other/subsite" }), /webUrl/);
    assert.equal(posts, 3, "Resolution failures never upload files");
    assert.ok(BrowserWindow.getAllWindows().every((window) => !window.isVisible()));
    assert.equal(await source.webContents.executeJavaScript("globalThis.fixtureRequestCount"), requests,
      "Every REST request executes inside the original browser frame with explicit same-origin credentials");
    await source.loadURL(web + "/_layouts/15/Unrecognized.aspx?private=secret");
    const unsupported = await tabs.listLocations();
    assert.equal(unsupported.locations.length, 0);
    assert.equal(unsupported.unavailable[0].code, "FOLDER_NOT_IDENTIFIED");
    assert.equal(unsupported.unavailable[0].pageName, "Unrecognized.aspx");
    assert.doesNotMatch(JSON.stringify(unsupported), /private|secret/);
    console.log("Document creation passed: live Electron browser-frame fetch with source User-Agent and same-origin credentials, discovered folder, valid Word package, same-account hidden open, automatic setup, collision refusal, arbitrary-folder resolution, child browsing, other-site creation and private request digest");
    tabs.close(); source.destroy(); main.destroy(); clearTimeout(deadline); app.quit();
  } catch (error) {
    console.error(error); tabs?.close(); source?.destroy(); main?.destroy(); app.exit(1);
  }
});
