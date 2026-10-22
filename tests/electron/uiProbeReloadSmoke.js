// Verify source-mode UI probe updates in one real Electron process/window.
const { app, BrowserWindow } = require("electron");
const fs = require("node:fs");
const path = require("node:path");
const { testPath } = require("../fixtures/testPaths");
const assert = require("node:assert/strict");
const { createBridgeSetupUiLoader } = require("../../app/office/bridgeSetupUiLoader");

const parent = testPath();
app.setPath("userData", path.join(parent, "reload-profile"));
app.disableHardwareAcceleration();
const timer = setTimeout(() => app.exit(1), 15000);
app.whenReady().then(async () => {
  let window, directory;
  try {
    assert.equal(process.defaultApp, true, "The checkout launch must enable development-only reload");
    fs.mkdirSync(parent, { recursive: true });
    directory = fs.mkdtempSync(path.join(parent, "reload-probe-"));
    const filename = path.join(directory, "probe.cjs");
    fs.writeFileSync(filename, "exports.bridgeSetupUi = function() { return document.querySelector('button').textContent + ':first'; };");
    const script = createBridgeSetupUiLoader({ filename });
    window = new BrowserWindow({ show: false, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false } });
    await window.loadURL("data:text/html,<button>Add-ins</button>");
    assert.equal(await window.webContents.executeJavaScript(script({ type: "inspect" })), "Add-ins:first");
    fs.writeFileSync(filename, "exports.bridgeSetupUi = function() { return document.querySelector('button').textContent + ':changed'; };");
    assert.equal(await window.webContents.executeJavaScript(script({ type: "inspect" })), "Add-ins:changed");
    assert.equal(window.isVisible(), false);
    console.log("UI probe reload passed in one hidden Electron window without restarting; fixed local source only, no real Office UI.");
  } finally {
    window?.destroy();
    if (directory) fs.rmSync(directory, { recursive: true });
    clearTimeout(timer);
  }
}).then(() => app.exit(0), (error) => { console.error(error); app.exit(1); });
