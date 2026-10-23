const { it } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const { officeUrl } = require("../../app/office/urlPolicy");
const { allowedNavigation } = require("../../app/office/documentTabs");
const { createBridgeSetupUiLoader } = require("../../app/office/bridgeSetupUiLoader");

it("keeps navigation and the serialized Microsoft account probe on the same strict URL checks", () => {
  const script = createBridgeSetupUiLoader({ filename: require.resolve("../../app/office/documentAccountUi"), exportName: "documentAccountUi" });
  const loginHosts = ["login.microsoftonline.com", "login.microsoft.com", "login.live.com"];
  const urls = loginHosts.map((host) => `https://${host}/authorize`).concat([
    "http://login.microsoftonline.com/", "https://login.microsoftonline.com.evil.test/",
    "https://user:password@login.microsoftonline.com/", "https://login.microsoftonline.com:8443/", "file:///etc/passwd",
  ]);
  for (const url of urls) {
    const allowed = loginHosts.some((host) => url === `https://${host}/authorize`);
    assert.equal(!!officeUrl(url, "login"), allowed);
    assert.equal(allowedNavigation(url, "https://tenant.sharepoint.com"), allowed);
    const result = vm.runInNewContext(script({ url, type: "probe", deadline: Date.now() + 1000 }), {
      URL, location: new URL(url), Date, document: { querySelectorAll: () => [], getElementById: () => null },
    });
    assert.equal(result.code, allowed ? undefined : "ACCOUNT_PAGE_CHANGED");
  }
  assert.equal(officeUrl("https://tenant.sharepoint.com/", "sharepoint").origin, "https://tenant.sharepoint.com");
  for (const url of ["https://sharepoint.com/", "https://evilsharepoint.com/", "https://tenant.sharepoint.com:8443/", "https://tenant.sharepoint.com.evil.test/"])
    assert.equal(officeUrl(url, "sharepoint"), null);
});
