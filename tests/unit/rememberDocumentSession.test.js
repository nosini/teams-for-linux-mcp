const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const { RememberDocumentSession } = require("../../app/office/rememberDocumentSession");

function fixture(initial = []) {
  const key = (cookie) => `${cookie.domain}|${cookie.name}`;
  const jar = new Map(initial.map((cookie) => [key(cookie), cookie]));
  const cookies = new EventEmitter(), writes = [];
  let flushed = 0;
  cookies.get = async ({ url, name }) => [...jar.values()].filter((cookie) => {
    const host = new URL(url).hostname, domain = cookie.domain.replace(/^\./, "");
    return (host === domain || host.endsWith(`.${domain}`)) && (!name || name === cookie.name);
  });
  cookies.set = async (details) => {
    writes.push(details);
    const cookie = { ...details, domain: details.domain || new URL(details.url).hostname,
      hostOnly: !details.domain, session: details.expirationDate === undefined };
    const old = jar.get(key(cookie));
    if (old) cookies.emit("changed", {}, old, "overwrite", true);
    jar.set(key(cookie), cookie);
    cookies.emit("changed", {}, cookie, "explicit", false);
  };
  cookies.remove = async (url, name) => {
    for (const cookie of await cookies.get({ url, name })) {
      jar.delete(key(cookie)); cookies.emit("changed", {}, cookie, "explicit", true);
    }
  };
  cookies.flushStore = async () => { flushed++; };
  const session = { isPersistent: () => true, cookies };
  return { jar, cookies, session, writes, flushed: () => flushed };
}

const cookie = (name = "FedAuth", extra = {}) => ({ name, value: "private-token", domain: ".tenant.sharepoint.com",
  hostOnly: false, path: "/", secure: true, httpOnly: true, sameSite: "no_restriction", session: true, ...extra });
const site = "https://tenant.sharepoint.com/file.docx";

describe("opt-in document session persistence", () => {
  it("preserves only exact opened-site auth cookies and their original attributes", async (t) => {
    const f = fixture([cookie(), cookie("rtFa", { domain: "tenant.sharepoint.com", hostOnly: true }), cookie("FedAuth1"),
      cookie("FedAuth2", { session: false, expirationDate: 123 }), cookie("other"), cookie("ESTSAUTH"),
      cookie("FedAuth", { domain: ".other.sharepoint.com" }), cookie("FedAuth", { domain: ".sharepoint.com" }),
      cookie("FedAuth3", { secure: false }), cookie("FedAuth4", { httpOnly: false }),
      cookie("FedAuth5", { partitionKey: "https://other.test" }), cookie("FedAuth6", { path: "/different" })]);
    const keeper = new RememberDocumentSession({ days: 3 }); t.after(() => keeper.close());
    const before = Date.now() / 1000;
    keeper.watch(f.session, site); await keeper.settle();
    assert.deepEqual(f.writes.map((item) => item.name).sort(), ["FedAuth", "FedAuth1", "rtFa"]);
    for (const details of f.writes) {
      assert.equal(details.value, "private-token");
      assert.equal(details.secure, true); assert.equal(details.httpOnly, true);
      assert.equal(details.path, "/"); assert.equal(details.sameSite, "no_restriction");
      assert.ok(details.expirationDate >= Math.floor(before) + 3 * 86400);
      assert.ok(details.expirationDate <= Math.floor(Date.now() / 1000) + 3 * 86400);
      assert.equal(details.domain, details.name === "rtFa" ? undefined : ".tenant.sharepoint.com");
    }
    assert.equal(f.flushed(), 3);
    assert.equal(f.jar.get(".tenant.sharepoint.com|FedAuth2").expirationDate, 123, "Server-issued expiry is never extended");
  });

  it("observes future login cookies once and removes listeners on close", async () => {
    const f = fixture(), keeper = new RememberDocumentSession();
    keeper.watch(f.session, site); keeper.watch(f.session, site); await keeper.settle();
    assert.equal(f.cookies.listenerCount("changed"), 1);
    const fresh = cookie(); f.jar.set(`${fresh.domain}|${fresh.name}`, fresh);
    f.cookies.emit("changed", {}, fresh, "explicit", false);
    await keeper.settle();
    assert.equal(f.writes.length, 1);
    keeper.close(); assert.equal(f.cookies.listenerCount("changed"), 0);
    f.cookies.emit("changed", {}, fresh, "explicit", false);
    assert.equal(f.writes.length, 1);
  });

  it("retains the shared rtFa root cookie without broadening other cookie scopes", async (t) => {
    const root = cookie("rtFa", { domain: ".sharepoint.com" });
    const f = fixture([root, cookie("FedAuth", { domain: ".sharepoint.com" }),
      cookie("rtFa", { domain: ".com" }), cookie("rtFa", { domain: ".other.sharepoint.com" }),
      cookie("rtFa1", { domain: ".sharepoint.com" }), cookie("ESTSAUTH", { domain: ".sharepoint.com" })]);
    const keeper = new RememberDocumentSession(); t.after(() => keeper.close());
    keeper.watch(f.session, site); await keeper.settle();
    assert.equal(f.writes.length, 1);
    assert.equal(f.writes[0].name, "rtFa");
    assert.equal(f.writes[0].url, "https://sharepoint.com/");
    assert.equal(f.writes[0].domain, root.domain);
    assert.equal(f.writes[0].secure, root.secure);
    assert.equal(f.writes[0].httpOnly, root.httpOnly);
    assert.equal(f.writes[0].sameSite, root.sameSite);
    assert.equal(f.writes[0].path, root.path);
    assert.equal(f.writes[0].value, root.value);
    await f.cookies.remove("https://sharepoint.com/", "rtFa");
    await keeper.settle();
    assert.equal(f.jar.has(".sharepoint.com|rtFa"), false, "Explicit root sign-out stays removed");
    assert.equal(f.writes.length, 1);
  });

  it("applies the same security and partition restrictions to the shared root", async (t) => {
    for (const extra of [{ secure: false }, { httpOnly: false }, { path: "/other" },
      { partitionKey: "https://other.test" }, { hostOnly: true }]) {
      const f = fixture([cookie("rtFa", { domain: ".sharepoint.com", ...extra })]);
      const keeper = new RememberDocumentSession(); t.after(() => keeper.close());
      keeper.watch(f.session, site); await keeper.settle();
      assert.equal(f.writes.length, 0);
    }
  });

  it("does not resurrect sign-out during the initial scan", async (t) => {
    const f = fixture([cookie()]), keeper = new RememberDocumentSession(); t.after(() => keeper.close());
    const get = f.cookies.get;
    let release;
    f.cookies.get = () => { f.cookies.get = get; return new Promise((resolve) => { release = resolve; }); };
    keeper.watch(f.session, site);
    await f.cookies.remove("https://tenant.sharepoint.com/", "FedAuth");
    release([cookie()]); await keeper.settle();
    assert.equal(f.writes.length, 0); assert.equal(f.jar.size, 0);
  });

  it("lets an explicit sign-out win while a promotion is pending", async (t) => {
    const f = fixture([cookie()]), keeper = new RememberDocumentSession(); t.after(() => keeper.close());
    const set = f.cookies.set;
    let release, started;
    const pending = new Promise((resolve) => { started = resolve; });
    f.cookies.set = async (details) => { started(); await new Promise((resolve) => { release = resolve; }); await set(details); };
    keeper.watch(f.session, site); await pending;
    await f.cookies.remove("https://tenant.sharepoint.com/", "FedAuth");
    release(); await keeper.settle();
    assert.equal(f.jar.size, 0);
  });

  it("does not overwrite a token that changed during initial scanning", async (t) => {
    const f = fixture([cookie()]), keeper = new RememberDocumentSession(); t.after(() => keeper.close());
    const get = f.cookies.get;
    f.cookies.get = async () => {
      f.cookies.get = get;
      f.jar.set(".tenant.sharepoint.com|FedAuth", cookie("FedAuth", { value: "new-token" }));
      return [cookie()];
    };
    keeper.watch(f.session, site); await keeper.settle();
    assert.equal(f.writes.length, 0);
    assert.equal(f.jar.get(".tenant.sharepoint.com|FedAuth").value, "new-token");
  });

  it("ignores private sessions and non-document sites without accessing cookies", () => {
    const keeper = new RememberDocumentSession();
    const session = { isPersistent: () => true, cookies: new Proxy({}, { get: () => assert.fail("No cookie access") }) };
    for (const url of ["https://login.microsoftonline.com/", "https://evilsharepoint.com/a.docx", "https://tenant.sharepoint.com.evil.test/a.docx", "http://tenant.sharepoint.com/a.docx", "https://user:pass@tenant.sharepoint.com/a.docx"]) keeper.watch(session, url);
    keeper.watch({ ...session, isPersistent: () => false }, site);
    keeper.close();
  });

  it("does not retry failed cookie writes or expose their errors", async (t) => {
    const f = fixture([cookie()]), keeper = new RememberDocumentSession(); t.after(() => keeper.close());
    let writes = 0;
    f.cookies.set = async () => { writes++; throw new Error("private-token"); };
    keeper.watch(f.session, site); await keeper.settle();
    assert.equal(writes, 1);
    assert.equal(f.jar.get(".tenant.sharepoint.com|FedAuth").session, true);
  });
});
