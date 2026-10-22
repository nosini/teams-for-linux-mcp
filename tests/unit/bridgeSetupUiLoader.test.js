const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createBridgeSetupUiLoader } = require("../../app/office/bridgeSetupUiLoader");

function fixture(t) {
  const parent = path.resolve(__dirname, "../../test-results");
  fs.mkdirSync(parent, { recursive: true });
  const directory = fs.mkdtempSync(path.join(parent, "ui-loader-"));
  const filename = path.join(directory, "ui.cjs");
  let tick = 1;
  const write = (source) => { fs.writeFileSync(filename, source); fs.utimesSync(filename, tick, tick); tick++; };
  t.after(() => { delete require.cache[filename]; fs.rmSync(directory, { recursive: true }); });
  write("exports.bridgeSetupUi = function(request) { return request.action + ':one'; };");
  return { filename, write };
}

describe("fixed Office UI probe development reload", () => {
  it("reloads a fixed named export for the account picker", (t) => {
    const { filename, write } = fixture(t);
    write("exports.documentAccountUi = function() { return 'one'; };");
    const script = createBridgeSetupUiLoader({ filename, exportName: "documentAccountUi", development: true });
    assert.equal(vm.runInNewContext(script({})), "one");
    write("exports.documentAccountUi = function() { return 'two'; };");
    assert.equal(vm.runInNewContext(script({})), "two");
  });
  it("reloads a changed fixed module in development without accepting request source", (t) => {
    const { filename, write } = fixture(t);
    const script = createBridgeSetupUiLoader({ filename, development: true });
    assert.equal(vm.runInNewContext(script({ action: "addins" })), "addins:one");
    write("exports.bridgeSetupUi = function(request) { return request.action + ':two'; };");
    assert.equal(vm.runInNewContext(script({ action: "addins", filename: "/ignored", source: "throw 1" })), "addins:two");
  });

  it("keeps packaged behavior fixed and stops on broken development changes until repaired", (t) => {
    const { filename, write } = fixture(t);
    const packaged = createBridgeSetupUiLoader({ filename, development: false });
    const development = createBridgeSetupUiLoader({ filename, development: true });
    write("exports.bridgeSetupUi = null;");
    assert.equal(vm.runInNewContext(packaged({ action: "addins" })), "addins:one");
    assert.throws(() => development({ action: "addins" }), /could not reload/);
    write("exports.bridgeSetupUi = function(request) { return request.action + ':repaired'; };");
    assert.equal(vm.runInNewContext(development({ action: "addins" })), "addins:repaired");
  });
});
