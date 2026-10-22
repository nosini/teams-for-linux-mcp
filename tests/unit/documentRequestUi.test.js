const { it } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const { documentRequestUi } = require("../../app/office/documentRequestUi");

const page = "https://tenant.sharepoint.com/_layouts/15/filebrowser.aspx";
const web = "https://tenant.sharepoint.com/sites/team";
function run(changes, fetch) {
  const request = { pageUrl: page, webUrl: web, url: web + "/_api/contextinfo", method: "POST", deadline: Date.now() + 1000, ...changes };
  return vm.runInNewContext(`(${documentRequestUi.toString()})(${JSON.stringify(request)})`, {
    location: new URL(page), URL, Date, fetch, AbortController, Uint8Array, TextDecoder, atob, setTimeout, clearTimeout,
  });
}

it("refuses stale pages, other origins, paths outside the selected web and unsupported methods before fetch", async () => {
  for (const changes of [{ pageUrl: page + "?changed" }, { deadline: 0 }, { url: "https://other.sharepoint.com/_api/contextinfo" },
    { webUrl: "https://other.sharepoint.com" }, { url: web + "/Shared%20Documents/private.docx" },
    { url: "https://tenant.sharepoint.com/sites/other/_api/contextinfo" }, { method: "DELETE" }]) {
    const result = await run(changes, () => assert.fail("Invalid requests must not reach the network"));
    assert.ok(["LOCATION_CHANGED", "INVALID_LOCATION"].includes(result.code));
  }
});

it("redacts server error bodies, redirect headers and fetch exceptions", async () => {
  for (const fetch of [async () => Response.json({ error: { code: "-1, System.UnauthorizedAccessException", message: "private-token" } }, { status: 403 }),
    async () => new Response("private-token", { status: 302, headers: { location: "https://private.test/?secret=token" } }),
    async () => { throw new Error("private-token"); }]) {
    const result = await run({}, fetch);
    assert.ok(result.status || result.code);
    assert.doesNotMatch(JSON.stringify(result), /private|secret|token|location|headers/);
    assert.equal(result.text, undefined);
  }
});
