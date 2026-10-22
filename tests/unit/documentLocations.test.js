const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const { DocumentLocations, locationSource, wordDocumentName } = require("../../app/office/documentLocations");

const origin = "https://tenant.sharepoint.com", web = origin + "/sites/team";
const folder = "/sites/team/Shared Documents/Work";
const fileId = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const folderUrl = web + "/Shared%20Documents/Forms/AllItems.aspx?id=" + encodeURIComponent(folder);
const documentUrl = web + "/_layouts/15/Doc.aspx?sourcedoc=%7B" + fileId + "%7D&private=secret";
const file = (name) => ({ Name: name, ServerRelativeUrl: folder + "/" + name, UniqueId: fileId, private: "private-token" });

function fixture(urls = [folderUrl]) {
  const calls = [], frames = urls.map((url) => ({ url, executeJavaScript: async () => ({ webUrl: web }) }));
  const contents = { id: 1, isDestroyed: () => false, mainFrame: { framesInSubtree: frames } };
  const state = { files: [], onRequest: null, pageWeb: web };
  for (const frame of frames) frame.executeJavaScript = async (script) => vm.runInNewContext(script, {
    location: new URL(frame.url), URL, Date, _spPageContextInfo: { webAbsoluteUrl: state.pageWeb },
    fetch: (url, options) => contents.session.fetch(url, options), AbortController, Uint8Array, TextDecoder, atob, setTimeout, clearTimeout,
  });
  contents.session = { fetch: async (url, options) => {
    calls.push({ url, options });
    if (state.onRequest) { const response = await state.onRequest(url, options); if (response) return response; }
    const path = decodeURIComponent(new URL(url).pathname);
    if (path.endsWith("/contextinfo")) return Response.json({ GetContextWebInformation: { FormDigestValue: "private-digest" } });
    if (path.includes("/Files/AddUsingPath")) {
      const name = path.match(/AddUsingPath\(decodedurl='(.*)',overwrite=false\)/)[1].replace(/''/g, "'");
      const created = file(name); state.files.push(created); return Response.json(created);
    }
    if (path.endsWith("/Files")) return Response.json({ value: state.files });
    if (path.includes("GetFileBy")) return Response.json(file("Existing.docx"));
    return Response.json({ Name: "Work", ServerRelativeUrl: folder, UniqueId: fileId });
  } };
  const locations = new DocumentLocations(() => [contents]);
  return { locations, contents, frames, calls, state };
}

describe("document folder discovery and creation", () => {
  it("recognizes folder URLs, document IDs and direct document parents without accepting arbitrary sites", () => {
    assert.equal(locationSource(folderUrl).folderPath, folder);
    assert.equal(locationSource(documentUrl).fileId, fileId);
    assert.equal(locationSource(origin + "/:w:/r/sites/team/_layouts/15/Doc.aspx?sourcedoc=" + fileId).webUrl, web);
    assert.equal(locationSource(web + "/Shared%20Documents/Existing.docx").filePath, "/sites/team/Shared Documents/Existing.docx");
    assert.equal(locationSource(web + "/Shared%20Documents/Forms/AllItems.aspx").folderPath, "/sites/team/Shared Documents");
    for (const url of ["https://evilsharepoint.com/a.docx", "http://tenant.sharepoint.com/a.docx", "https://user:password@tenant.sharepoint.com/a.docx",
      "https://tenant.sharepoint.com:8443/a.docx", "https://tenant.sharepoint.com.evil.test/a.docx", "file:///a.docx",
      web + "/_layouts/15/Doc.aspx?sourcedoc=not-a-guid", web + "/Shared%20Documents/Forms/AllItems.aspx?id=https://evil.test/"])
      assert.equal(locationSource(url), null);
  });

  it("discovers and verifies a folder using only its own session, without leaking request secrets", async (t) => {
    const f = fixture(); t.after(() => f.locations.close());
    const report = await f.locations.list();
    assert.equal(report.locations.length, 1);
    assert.equal(report.locations[0].folderPath, folder);
    assert.equal(report.locations[0].source, "folder");
    assert.ok(f.calls.every(({ url, options }) => url.startsWith(web + "/_api/") && options.credentials === "include" && options.redirect === "error" && options.mode === "same-origin"));
    assert.doesNotMatch(JSON.stringify(report), /secret|private|cookie|digest|sourcedoc/);
    assert.equal((await f.locations.list()).locations[0].locationId, report.locations[0].locationId);
  });

  it("resolves an open document's parent instead of asking the agent for a link", async (t) => {
    const f = fixture([documentUrl]); t.after(() => f.locations.close());
    const report = await f.locations.list();
    assert.equal(report.locations[0].folderPath, folder);
    assert.equal(report.locations[0].source, "document_parent");
    assert.match(f.calls[0].url, /GetFileById/);
    assert.doesNotMatch(JSON.stringify(report), /private|secret|sourcedoc|Existing/);
  });

  it("uploads a valid app template once, refuses overwrite and keeps the request digest private", async (t) => {
    const f = fixture(); t.after(() => f.locations.close());
    const created = await f.locations.create(undefined, "Report #1% O'Brien");
    assert.equal(created.contents, f.contents);
    assert.equal(created.item.name, "Report #1% O'Brien.docx");
    assert.equal(decodeURIComponent(new URL(created.item.url).pathname), folder + "/Report #1% O'Brien.docx");
    const writes = f.calls.filter(({ options }) => options.body);
    assert.equal(writes.length, 1);
    assert.match(writes[0].url, /overwrite=false/);
    assert.match(writes[0].url, /O%27%27Brien/);
    assert.equal(writes[0].options.headers["X-RequestDigest"], "private-digest");
    assert.equal(Buffer.from(writes[0].options.body).subarray(0, 4).toString("hex"), "504b0304");
    assert.doesNotMatch(JSON.stringify(created.item), /digest|private|token/);
    assert.equal((await f.locations.listFiles(created.locationId, created.item.name)).files[0].name, created.item.name);
    assert.match(f.calls.at(-1).url, /\$filter=Name%20eq%20/);
  });

  it("requires a folder choice for ambiguous contexts and binds it to its original view", async (t) => {
    const f = fixture([folderUrl, documentUrl]); t.after(() => f.locations.close());
    const report = await f.locations.list();
    assert.equal(report.locations.length, 1, "Same folder in one account is deduplicated");
    const ambiguous = fixture([folderUrl]); t.after(() => ambiguous.locations.close());
    ambiguous.contents.mainFrame.framesInSubtree.push({ url: documentUrl, executeJavaScript: async () => ({ webUrl: web }) });
    const second = { ...ambiguous.contents, id: 2, session: { ...ambiguous.contents.session } };
    const both = new DocumentLocations(() => [ambiguous.contents, second]); t.after(() => both.close());
    await assert.rejects(both.create(undefined, "Report"), /ambiguous/);
    assert.equal(ambiguous.calls.filter(({ options }) => options.body).length, 0);
    const chosen = (await both.list()).locations[0];
    assert.equal((await both.create(chosen.locationId, "Report")).contents, ambiguous.contents);
  });

  it("rejects stale navigation, expiry and close before issuing a creation request", async (t) => {
    t.mock.timers.enable({ apis: ["Date"] });
    for (const invalidate of [(f) => { f.frames[0].url = "https://teams.cloud.microsoft/"; },
      (f) => { f.contents.mainFrame.framesInSubtree = []; },
      () => t.mock.timers.tick(5 * 60 * 1000), (f) => f.locations.close()]) {
      const f = fixture(); t.after(() => f.locations.close());
      const id = (await f.locations.list()).locations[0].locationId;
      invalidate(f);
      await assert.rejects(f.locations.create(id, "Report"), /LOCATION_CHANGED/);
      assert.equal(f.calls.filter(({ options }) => options.method === "POST").length, 0);
    }
  });

  it("does not write if navigation changes while obtaining the request digest", async (t) => {
    const f = fixture(); t.after(() => f.locations.close());
    f.state.onRequest = async (url) => { if (url.includes("contextinfo")) f.frames[0].url = "https://teams.cloud.microsoft/"; };
    await assert.rejects(f.locations.create(undefined, "Report"), /LOCATION_CHANGED/);
    assert.equal(f.calls.filter(({ options }) => options.body).length, 0);
  });

  it("sanitizes write failures and never retries an uncertain creation", async (t) => {
    for (const failure of [() => { throw new Error("private-token https://secret.test/?credential=value"); },
      () => new Response("private-server-body", { status: 500 }), () => new Response("private-server-body"),
      () => new Response(Buffer.alloc(300000, "a"))]) {
      const f = fixture(); t.after(() => f.locations.close());
      f.state.onRequest = async (_url, options) => options.body ? failure() : null;
      await assert.rejects(f.locations.create(undefined, "Report"), (error) => {
        assert.match(error.message, /CREATION_OUTCOME_UNKNOWN.*office_list_location_files/);
        assert.doesNotMatch(error.message, /private|secret|credential|value/); return true;
      });
      assert.equal(f.calls.filter(({ options }) => options.body).length, 1);
    }
  });

  it("bounds a stalled upload and reports its outcome as unknown without retrying", async (t) => {
    const f = fixture(); t.after(() => f.locations.close());
    const id = (await f.locations.list()).locations[0].locationId;
    t.mock.timers.enable({ apis: ["setTimeout"] });
    let started;
    const ready = new Promise((resolve) => { started = resolve; });
    f.state.onRequest = async (_url, options) => {
      if (!options.body) return null;
      started();
      return new Promise((_resolve, reject) => options.signal.addEventListener("abort", () => reject(new Error("private-token")), { once: true }));
    };
    const pending = f.locations.create(id, "Report");
    await ready;
    t.mock.timers.tick(15000);
    await assert.rejects(pending, /CREATION_OUTCOME_UNKNOWN/);
    assert.equal(f.calls.filter(({ options }) => options.body).length, 1);
  });

  it("refuses redirects, missing permissions and cross-origin page context without exposing server output", async (t) => {
    for (const response of [new Response(null, { status: 302, headers: { location: "https://secret.test/login?token=secret" } }),
      new Response("private-server-body", { status: 403 })]) {
      const f = fixture(); t.after(() => f.locations.close());
      f.state.onRequest = async () => response;
      const report = await f.locations.list();
      assert.equal(report.locations.length, 0);
      assert.doesNotMatch(JSON.stringify(report), /secret|private|token/);
    }
    const f = fixture(); t.after(() => f.locations.close()); f.state.pageWeb = "https://evil.test/";
    assert.equal((await f.locations.list()).locations.length, 0);
    assert.equal(f.calls.length, 0);
  });

  it("marks partial listings and rejects invalid names before touching the network", async (t) => {
    const f = fixture(); t.after(() => f.locations.close());
    for (const name of ["../Report", "name/other", "x\\y", "a".repeat(252), "~$Report", ".docx", "name.", "a\u007fb"])
      await assert.rejects(f.locations.create(undefined, name), /Word document name/);
    assert.equal(f.calls.length, 0);
    assert.equal(wordDocumentName("Report"), "Report.docx");
    f.state.onRequest = async (url) => new URL(url).pathname.endsWith("/Files") ? Response.json({ value: [], "@odata.nextLink": "https://private.test/?secret=token" }) : null;
    const result = await f.locations.listFiles((await f.locations.list()).locations[0].locationId);
    assert.equal(result.truncated, true);
    assert.doesNotMatch(JSON.stringify(result), /private|secret|token/);
  });
});
