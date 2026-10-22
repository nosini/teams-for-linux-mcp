const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { inspectDocumentAuthentication } = require("../../app/office/documentAuthentication");

describe("document sign-in diagnostics", () => {
  it("reports session versus persistent cookies without exposing credentials or profile data", async () => {
    const queries = [];
    const session = { isPersistent: () => true, cookies: { get: async (filter) => {
      queries.push(filter);
      return ({ ESTSAUTH: [{ session: true, value: "private-token", domain: "private-account" }],
        ESTSAUTHPERSISTENT: [], FedAuth: [{ session: false, value: "secret-token" }],
        rtFa: [{ session: true }, { session: false }] })[filter.name];
    } } };
    const result = await inspectDocumentAuthentication(session, "https://tenant.sharepoint.com/sites/private/Doc.aspx?token=secret");
    assert.equal(result.persistentSession, true);
    assert.deepEqual(result.cookies, {
      microsoftSession: "session", microsoftPersistent: "missing", sharepointSite: "persistent", sharepointRoot: "mixed",
    });
    assert.deepEqual(queries, [
      { url: "https://login.microsoftonline.com/", name: "ESTSAUTH" },
      { url: "https://login.microsoftonline.com/", name: "ESTSAUTHPERSISTENT" },
      { url: "https://tenant.sharepoint.com", name: "FedAuth" },
      { url: "https://tenant.sharepoint.com", name: "rtFa" },
    ]);
    assert.doesNotMatch(JSON.stringify(result), /secret|token|account|tenant|private|"value"|"domain"|"path"/i);
    assert.deepEqual(result.cookieAttributes.sharepointSite, { entries: [{ persistence: "persistent", scope: "unknown",
      secure: null, httpOnly: null, rootPath: null, partitioned: false }], truncated: false });
  });

  it("reports only bounded SharePoint attributes that explain retention exclusions", async () => {
    const samples = [
      { domain: ".tenant.sharepoint.com", path: "/", secure: true, httpOnly: true, session: true },
      { domain: ".sharepoint.com", path: "/private-path", secure: true, httpOnly: false, session: true, partitionKey: "private-partition" },
      { domain: "other-private.example", path: "/", secure: false, httpOnly: true, session: false },
    ];
    const result = await inspectDocumentAuthentication({ isPersistent: () => true, cookies: { get: async ({ name }) =>
      name === "FedAuth" ? [] : name === "rtFa" ? [...samples, ...Array(10).fill(samples[0])] : [] } },
    "https://tenant.sharepoint.com/private-document");
    const attrs = result.cookieAttributes;
    assert.deepEqual(attrs.sharepointSite, { entries: [], truncated: false });
    assert.equal(attrs.sharepointRoot.entries.length, 8);
    assert.equal(attrs.sharepointRoot.truncated, true);
    assert.deepEqual(attrs.sharepointRoot.entries.slice(0, 3), [
      { persistence: "session", scope: "opened_host", secure: true, httpOnly: true, rootPath: true, partitioned: false },
      { persistence: "session", scope: "parent_domain", secure: true, httpOnly: false, rootPath: false, partitioned: true },
      { persistence: "persistent", scope: "other", secure: false, httpOnly: true, rootPath: true, partitioned: false },
    ]);
    assert.doesNotMatch(JSON.stringify(result), /tenant|sharepoint\.com|private|example/);
  });

  it("does not guess another tenant from a cloud link or untrusted destination", async () => {
    for (const url of ["https://word.cloud.microsoft/open/private", "https://evilsharepoint.com/file.docx", "https://tenant.sharepoint.com.evil.test/file.docx", "https://user:password@tenant.sharepoint.com/file.docx", "http://tenant.sharepoint.com/file.docx"]) {
      const queries = [];
      const result = await inspectDocumentAuthentication({ isPersistent: () => false,
        cookies: { get: async (filter) => { queries.push(filter); return []; } } }, url);
      assert.equal(result.persistentSession, false);
      assert.equal(result.cookies.sharepointSite, "unavailable");
      assert.equal(result.cookies.sharepointRoot, "unavailable");
      assert.deepEqual(result.cookieAttributes, { sharepointSite: null, sharepointRoot: null });
      assert.equal(queries.length, 2);
      assert.ok(queries.every((query) => query.url === "https://login.microsoftonline.com/"));
    }
  });

  it("sanitizes cookie-store failures", async () => {
    const result = await inspectDocumentAuthentication({ cookies: { get: async () => { throw new Error("secret-token"); } } }, "https://tenant.sharepoint.com/file.docx");
    assert.equal(result.persistentSession, null);
    assert.ok(Object.values(result.cookies).every((value) => value === "unavailable"));
    assert.deepEqual(result.cookieAttributes, { sharepointSite: null, sharepointRoot: null });
    assert.doesNotMatch(JSON.stringify(result), /secret|token/);
  });

  it("bounds an unresponsive cookie store", async (t) => {
    t.mock.timers.enable({ apis: ["setTimeout"] });
    const pending = inspectDocumentAuthentication({ isPersistent: () => true,
      cookies: { get: () => new Promise(() => {}) } }, "https://tenant.sharepoint.com/file.docx");
    t.mock.timers.tick(1500);
    const result = await pending;
    assert.equal(result.persistentSession, true);
    assert.ok(Object.values(result.cookies).every((value) => value === "unavailable"));
  });
});
