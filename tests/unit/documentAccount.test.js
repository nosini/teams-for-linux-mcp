const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { getRememberedDocumentAccount, inspectDocumentAccount, selectDocumentAccount } = require("../../app/office/documentAccount");
const { runBridgeSetup } = require("../../app/office/bridgeSetup");

function contents(url, session = {}) {
  return { id: 1, session, isDestroyed: () => false, mainFrame: { url, executeJavaScript: async () => ["USER@example.test", "user@example.test"] } };
}

describe("Office document account selection", () => {
  it("reads only the document site's remembered choice and invalidates navigation", async () => {
    const login = contents("https://login.microsoftonline.com/common/authorize");
    login.mainFrame.executeJavaScript = async (script) => {
      assert.ok(script.includes('"type":"remembered"'));
      assert.ok(script.includes('"scope":"https://tenant.sharepoint.com"'));
      return "USER@example.test";
    };
    const account = await getRememberedDocumentAccount(login, "https://tenant.sharepoint.com");
    assert.deepEqual(account.identifiers, ["user@example.test"]);
    assert.equal(account.isCurrent(), true);
    login.mainFrame.url += "different";
    assert.equal(account.isCurrent(), false);
  });

  it("accepts an explicit hint without reading Teams or requesting tokens", async () => {
    const login = contents("https://login.microsoftonline.com/common/authorize");
    login.mainFrame.executeJavaScript = () => assert.fail("An explicit hint requires no source lookup");
    assert.deepEqual((await getRememberedDocumentAccount(login, "https://tenant.sharepoint.com", "USER@example.test")).identifiers, ["user@example.test"]);
    const teams = contents("https://teams.cloud.microsoft/");
    teams.mainFrame.executeJavaScript = () => assert.fail("Never inspect the main Teams session");
    assert.equal(await getRememberedDocumentAccount(teams, "https://tenant.sharepoint.com"), null);
  });

  it("rejects invalid hints, untrusted login pages and stale preference reads", async () => {
    for (const url of ["https://login.microsoftonline.com.evil.test/", "http://login.microsoftonline.com/", "https://user:password@login.microsoftonline.com/", "https://login.microsoftonline.com:444/"]) {
      const source = contents(url);
      source.mainFrame.executeJavaScript = () => assert.fail("Untrusted source must not be read");
      assert.equal(await getRememberedDocumentAccount(source, "https://tenant.sharepoint.com"), null);
    }
    const login = contents("https://login.microsoftonline.com/common/authorize");
    for (const hint of ["not-an-identity", "bad\n@example.test"]) assert.equal(await getRememberedDocumentAccount(login, "https://tenant.sharepoint.com", hint), null);
    login.mainFrame.executeJavaScript = async () => { login.mainFrame.url += "moved"; return "user@example.test"; };
    assert.equal(await getRememberedDocumentAccount(login, "https://tenant.sharepoint.com"), null);
    login.mainFrame.executeJavaScript = async () => { throw new Error("private-token"); };
    assert.equal(await getRememberedDocumentAccount(login, "https://tenant.sharepoint.com"), null);
  });

  it("bounds picker diagnostics and drops page-supplied text and errors", async () => {
    const login = contents("https://login.microsoftonline.com/common/authorize?private=secret");
    login.mainFrame.executeJavaScript = async () => ({ picker: true, tileCount: 200, credentials: false, email: "private@example.test", token: "secret" });
    assert.deepEqual(await inspectDocumentAccount(login), { picker: true, tileCount: 17, credentials: false });
    login.mainFrame.executeJavaScript = async () => { throw new Error("private-token"); };
    assert.deepEqual(await inspectDocumentAccount(login), { picker: false, tileCount: 0, credentials: false });
    assert.deepEqual(await selectDocumentAccount(login, async () => null, new AbortController().signal, Date.now() + 1000), { performed: false, code: "DOCUMENT_ACCOUNT_NOT_REMEMBERED" });
  });

  it("dispatches an account selection once and still stops if credentials are required afterward", async () => {
    const login = contents("https://login.microsoftonline.com/common/authorize");
    login.mainFrame.framesInSubtree = [login.mainFrame];
    let keys = 0, profiles = 0;
    login.debugger = { isAttached: () => login.attached, attach: () => { login.attached = true; },
      detach: () => { login.attached = false; }, sendCommand: async (method, args) => {
        if (method === "Input.dispatchKeyEvent" && args.type === "keyDown") keys++;
      } };
    login.mainFrame.executeJavaScript = async (script) => script.includes('"type":"probe"') ? { picker: keys === 0, tileCount: 1, credentials: keys > 0 } : { performed: true };
    const updates = [];
    const result = await runBridgeSetup({ contents: login, manifests: {}, signal: new AbortController().signal,
      listDocuments: async () => ({ documents: [] }), update: (value) => updates.push(value),
      getAccount: async () => { profiles++; return { identifiers: ["private@example.test"], isCurrent: () => true }; },
    }, { timeoutMs: 100, signInGraceMs: 15, pollMs: 1 });
    assert.equal(result.code, "SIGN_IN_REQUIRED");
    assert.equal(result.accountSelection, "dispatched");
    assert.equal(keys, 1);
    assert.equal(profiles, 1);
    assert.equal(login.attached, false);
    assert.doesNotMatch(JSON.stringify([result, updates]), /private|example|token/);
  });
});
