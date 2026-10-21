// Actual Electron/TLS/task-pane integration with a simulated Office.js host.
// No Microsoft account or external HTTPS response is used in this fixture.
const { app, BrowserWindow, session, webContents } = require("electron");
const fs = require("node:fs/promises");
const https = require("node:https");
const path = require("node:path");
const assert = require("node:assert/strict");
const { startAddinHost } = require("../../app/office/addinHost");
const { LiveOfficeBridge } = require("../../app/office/liveBridge");
const { excelFixture } = require("../fixtures/excelApi");

app.setPath("userData", path.resolve(__dirname, "../../test-results/mcp/addin-electron-profile"));
app.disableHardwareAcceleration();
const deadline = setTimeout(() => app.exit(1), 30000);

app.whenReady().then(async () => {
  const host = await startAddinHost({ directory: path.join(app.getPath("userData"), "mcp"), port: 0 });
  const windows = [];
  try {
    let certificateChecks = 0;
    app.on("certificate-error", (event, _contents, url, error, certificate, callback) => {
      const accepted = host.acceptsCertificate(url, error, certificate);
      if (accepted) { certificateChecks++; event.preventDefault(); }
      callback(accepted);
    });
    // This session has no protocol interception: it proves real Chromium TLS
    // acceptance of the exact local certificate, independently of the fixtures.
    const tlsWindow = new BrowserWindow({ show: false, webPreferences: { partition: "addin-tls-test", sandbox: true } });
    windows.push(tlsWindow);
    await tlsWindow.loadURL(`${host.origin}/taskpane.css`);
    assert.ok(certificateChecks > 0);
    assert.match(await tlsWindow.webContents.executeJavaScript("document.body.textContent"), /color-scheme/);

    const cert = await fs.readFile(path.join(app.getPath("userData"), "mcp/addin/localhost.pem"));
    const fixtureSession = session.fromPartition("addin-office-fixture");
    const officeScript = `
      const clone = value => JSON.parse(JSON.stringify(value));
      const rangeShape = address => globalThis.TeamsExcelOperations.rangeShape(address);
      const fixture = (${excelFixture.toString()})();
      globalThis.Excel = fixture.excel;
      globalThis.__excelFixture = fixture.model;
      globalThis.Office = {
        HostType: {Excel: 'Excel'},
        context: {document: {url: 'https://tenant.sharepoint.com/book.xlsx'},
          requirements: {isSetSupported: name => name === 'ExcelApi'}},
        onReady: async () => ({host: 'Excel'})
      };
    `;
    fixtureSession.protocol.handle("https", async (request) => {
      const url = new URL(request.url);
      if (url.hostname === "teams.cloud.microsoft") return new Response('<iframe src="https://euc-excel.officeapps.live.com/x/edit"></iframe>', { headers: { "Content-Type": "text/html" } });
      if (url.hostname === "euc-excel.officeapps.live.com") return new Response(`<iframe src="${host.origin}/taskpane.html?_host_Info=fixture"></iframe>`, { headers: { "Content-Type": "text/html" } });
      if (url.hostname === "appsforoffice.microsoft.com") return new Response(officeScript, { headers: { "Content-Type": "text/javascript" } });
      if (url.origin === host.origin) {
        // Forward only our own static assets using ordinary CA/name validation.
        return new Promise((resolve, reject) => {
          https.get(request.url, { ca: cert }, (response) => {
            const chunks = [];
            response.on("data", (chunk) => chunks.push(chunk));
            response.once("end", () => resolve(new Response(Buffer.concat(chunks), { status: response.statusCode, headers: response.headers })));
          }).once("error", reject);
        });
      }
      return new Response("No network access in fixture", { status: 403 });
    });
    const window = new BrowserWindow({ show: false, webPreferences: { session: fixtureSession, sandbox: true, contextIsolation: true, nodeIntegration: false } });
    windows.push(window);
    await window.loadURL("https://teams.cloud.microsoft/");
    const bridge = new LiveOfficeBridge(() => webContents.getAllWebContents(), host.origin);
    let documents = [];
    const until = Date.now() + 5000;
    while (documents.length === 0 && Date.now() < until) {
      documents = (await bridge.listDocuments()).documents;
      if (!documents.length) await new Promise((resolve) => setTimeout(resolve, 25));
    }
    assert.equal(documents.length, 1);
    const id = documents[0].id;
    const worksheets = await bridge.execute(id, "list_worksheets", {});
    const target = { worksheetId: worksheets.worksheets[0].id, address: "A1" };
    const read = await bridge.execute(id, "read_range", target);
    assert.deepEqual(read.values, [[1]]);
    const written = await bridge.execute(id, "write_range", { ...target, readToken: read.readToken, data: [[7]] });
    assert.equal(written.applied, true);
    assert.deepEqual(written.values, [[7]]);
    await assert.rejects(bridge.execute(id, "write_range", { ...target, readToken: read.readToken, data: [[8]] }), /token expired/);
    const formula = await bridge.execute(id, "set_formulas", { ...target, readToken: written.readToken, data: [["=1+2"]] });
    assert.deepEqual(formula.values, [[3]]);
    const frame = window.webContents.mainFrame.framesInSubtree.find((candidate) => candidate.url.startsWith(host.origin));
    assert.ok(frame);
    assert.match(await frame.executeJavaScript("document.getElementById('status').textContent"), /Ready for local agent/);
    await frame.executeJavaScript("document.getElementById('disconnect').click()");
    await assert.rejects(bridge.execute(id, "read_range", target), /document changed/);
    assert.equal((await bridge.listDocuments()).documents.length, 0);
    console.log("Office add-in Electron smoke passed: exact certificate pin, HTTPS task pane, nested frame discovery, read/write/formula workflow, stale token rejection, and disconnect. Office.js was simulated.");
  } finally {
    for (const window of windows) window.destroy();
    await host.close();
    clearTimeout(deadline);
  }
}).then(() => app.exit(0), (error) => { console.error(error); app.exit(1); });
