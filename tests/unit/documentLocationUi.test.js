const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const { documentLocationUi } = require("../../app/office/documentLocationUi");

const origin = "https://tenant.sharepoint.com", web = origin + "/sites/team";
const folder = "/sites/team/Shared Documents/Work";
function probe(url, context = { webAbsoluteUrl: web }, request = {}) {
  const forbidden = new Proxy({}, { get: () => assert.fail("Do not inspect document content, cookies or storage") });
  return JSON.parse(JSON.stringify(vm.runInNewContext("(" + documentLocationUi.toString() + ")(request)", {
    location: new URL(url), URL, Date, request: { url, deadline: Date.now() + 1000, ...request },
    _spPageContextInfo: context, document: forbidden, localStorage: forbidden, sessionStorage: forbidden,
  })));
}

describe("fixed SharePoint folder location probe", () => {
  it("resolves Teams embedded libraries and gives explicit subfolders precedence", () => {
    const url = web + "/_layouts/15/listallitems.aspx?app=teamsPage&listUrl=" + encodeURIComponent("/sites/team/Shared Documents");
    assert.equal(probe(url).folderPath, "/sites/team/Shared Documents");
    assert.equal(probe(url + "&RootFolder=" + encodeURIComponent(folder)).folderPath, folder);
    assert.equal(probe(web + "/_layouts/15/DocLibEmbed.aspx?rootFolder=" + encodeURIComponent(folder)).folderPath, folder);
    assert.equal(probe(web + "/_layouts/15/filebrowser.aspx?id=" + encodeURIComponent(folder)).folderPath, folder);
  });

  it("unwraps only same-origin Teams login destinations", () => {
    const dest = web + "/_layouts/15/listallitems.aspx?listUrl=" + encodeURIComponent(folder);
    const url = origin + "/_layouts/15/teamslogon.aspx?spfx=true&dest=" + encodeURIComponent(dest);
    const result = probe(url, null);
    assert.equal(result.webUrl, web);
    assert.equal(result.folderPath, folder);
    assert.equal(result.pageKind, "teamslogon");
    assert.deepEqual(result.queryKeys, ["dest"]);
    assert.equal(probe(origin + "/_layouts/15/teamslogon.aspx?dest=" + encodeURIComponent("https://evil.test/?id=" + folder), null).folderPath, null);
  });

  it("uses the addressed folder's site rather than the Teams file browser's host web", () => {
    const url = origin + "/_layouts/15/filebrowser.aspx?id=" + encodeURIComponent(folder);
    assert.equal(probe(url, { webAbsoluteUrl: origin }).webUrl, web);
    assert.equal(probe(url, null).webUrl, web);
    const nested = web + "/subweb";
    const nestedFolder = "/sites/team/subweb/Documents/Work";
    assert.equal(probe(origin + "/_layouts/15/filebrowser.aspx?id=" + encodeURIComponent(nestedFolder), { webAbsoluteUrl: nested }).webUrl, nested);
  });

  it("returns bounded page classifications without URL queries, values or content", () => {
    const url = web + "/_layouts/15/unrecognized.aspx?id=private-token&listId=private-guid&email=private-account";
    const result = probe(url);
    assert.equal(result.folderPath, null);
    assert.equal(result.pageKind, "other");
    assert.deepEqual(result.queryKeys, ["id", "listId"]);
    assert.doesNotMatch(JSON.stringify(result), /private|token|guid|account|email/);
  });

  it("does not guess a root for OneDrive home or inspect an expired/wrong-origin page", () => {
    assert.equal(probe(web + "/_layouts/15/onedrive.aspx").folderPath, null);
    assert.equal(probe(web + "/_layouts/15/onedrive.aspx?id=" + encodeURIComponent(folder)).folderPath, folder);
    assert.equal(probe(web + "/_layouts/15/onedrive.aspx", undefined, { deadline: 0 }), null);
    assert.equal(probe("https://evilsharepoint.com/_layouts/15/listallitems.aspx"), null);
  });
});
