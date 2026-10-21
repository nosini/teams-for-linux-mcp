const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const https = require("node:https");
const { startAddinHost } = require("../../app/office/addinHost");
const { parseStringPromise } = require("xml2js");

async function fixture(t) {
  const root = path.resolve(__dirname, "../../test-results/mcp");
  await fs.mkdir(root, { recursive: true });
  const directory = await fs.mkdtemp(path.join(root, "host-"));
  const host = await startAddinHost({ directory, port: 0 });
  const cert = await fs.readFile(path.join(directory, "addin/localhost.pem"));
  t.after(async () => { await host.close(); await fs.rm(directory, { recursive: true, force: true }); });
  const request = (name, options = {}) => new Promise((resolve, reject) => {
    const req = https.request(`${host.origin}${name}`, { ca: cert, ...options }, (response) => {
      let data = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => { data += chunk; });
      response.once("end", () => resolve({ status: response.statusCode, headers: response.headers, data }));
    });
    req.once("error", reject);
    req.end();
  });
  return { directory, host, cert, request };
}

describe("Office add-in HTTPS host", () => {
  it("serves only static assets on loopback and creates a usable Excel manifest", async (t) => {
    const { directory, host, request } = await fixture(t);
    assert.match(host.origin, /^https:\/\/127\.0\.0\.1:\d+$/);
    const page = await request("/taskpane.html");
    assert.equal(page.status, 200);
    assert.match(page.data, /appsforoffice.microsoft.com\/lib\/1\/hosted\/office.js/);
    assert.equal(page.headers["cache-control"], "no-store");
    assert.equal((await request("/localhost-key.pem")).status, 404);
    assert.equal((await request("/../addin/localhost-key.pem")).status, 404);
    assert.equal((await request("/taskpane.html", { method: "POST" })).status, 403);
    assert.equal((await request("/taskpane.html", { servername: "localhost", headers: { Host: "evil.test" } })).status, 403);
    assert.equal((await fs.stat(path.join(directory, "addin/localhost-key.pem"))).mode & 0o777, 0o600);
    const manifest = (await parseStringPromise(await fs.readFile(host.manifestPath, "utf8"))).OfficeApp;
    assert.equal(manifest.Hosts[0].Host[0].$.Name, "Workbook");
    assert.equal(manifest.DefaultSettings[0].SourceLocation[0].$.DefaultValue, `${host.origin}/taskpane.html`);
    assert.equal(manifest.Permissions[0], "ReadWriteDocument");
  });

  it("pins the exact certificate, address, port, error and validity without broad trust", async (t) => {
    const { host, cert } = await fixture(t);
    const other = await fixture(t);
    const peer = { data: cert.toString() };
    const authority = "net::ERR_CERT_AUTHORITY_INVALID";
    assert.equal(host.acceptsCertificate(`${host.origin}/taskpane.html`, authority, peer), true);
    assert.equal(host.acceptsCertificate(`${other.host.origin}/taskpane.html`, authority, peer), false);
    assert.equal(host.acceptsCertificate("https://evil.test/", authority, peer), false);
    assert.equal(host.acceptsCertificate(host.origin.replace("127.0.0.1", "127.0.0.1.evil.test"), authority, peer), false);
    assert.equal(host.acceptsCertificate(host.origin, "net::ERR_CERT_DATE_INVALID", peer), false);
    assert.equal(host.acceptsCertificate(host.origin, "net::ERR_CERT_REVOKED", peer), false);
    assert.equal(host.acceptsCertificate(host.origin, authority, { data: other.cert.toString() }), false);
    assert.equal(host.acceptsCertificate(host.origin, authority, { data: "invalid" }), false);
  });

  it("reuses its certificate between runs", async (t) => {
    const { host, directory, cert } = await fixture(t);
    await host.close();
    const second = await startAddinHost({ directory, port: 0 });
    t.after(() => second.close());
    assert.equal(second.acceptsCertificate(second.origin, "net::ERR_CERT_AUTHORITY_INVALID", { data: cert.toString() }), true);
  });
});
