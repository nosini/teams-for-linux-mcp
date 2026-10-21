// Run with Electron, on an isolated display if needed. All HTTPS responses are
// local fixtures; this test never signs in or contacts Microsoft services.
const { app, BrowserWindow, session, webContents } = require("electron");
const assert = require("node:assert/strict");
const path = require("node:path");
const { OfficeInspector } = require("../../app/office/inspector");
const { startLocalEndpoint } = require("../../app/mcp/localEndpoint");
const { createOfficeMcpServer } = require("../../app/mcp/server");

app.setPath("userData", path.resolve(__dirname, "../../test-results/mcp/electron-profile"));
app.disableHardwareAcceleration();
const deadline = setTimeout(() => app.exit(1), 20000);

app.whenReady().then(async () => {
  let endpoint;
  let window;
  try {
    const fixtures = new Map([
      ["teams.cloud.microsoft", '<iframe src="https://tenant.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=private"></iframe>'],
      ["tenant.sharepoint.com", '<iframe src="https://weu-excel.officeapps.live.com/x/edit?access_token=secret"></iframe>'],
      ["weu-excel.officeapps.live.com", `<script>
        globalThis.Office = {context: {document: {}, requirements: {
          isSetSupported: (name) => name === 'ExcelApi'
        }}};
        globalThis.Excel = {run: () => { throw new Error('must not access workbook'); }};
      </script>`],
    ]);
    session.defaultSession.protocol.handle("https", (request) => new Response(
      fixtures.get(new URL(request.url).hostname) || "", { headers: { "Content-Type": "text/html" } },
    ));
    window = new BrowserWindow({ show: false, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false } });
    await window.loadURL("https://teams.cloud.microsoft/");
    const inspector = new OfficeInspector(() => webContents.getAllWebContents());
    const { editors } = inspector.listEditors();
    assert.equal(editors.length, 2);
    assert.ok(editors.some((editor) => editor.surface === "document-wrapper"));
    const excel = editors.find((editor) => editor.application === "excel");
    assert.ok(excel);
    assert.doesNotMatch(JSON.stringify(editors), /secret|private|access_token|sourcedoc/);
    const result = await inspector.inspectEditor(excel.id);
    assert.equal(result.probeStatus, "completed");
    assert.equal(result.apiPresence.excelApi11, true);
    assert.equal(result.apiPresence.excelRun, true);
    assert.equal(result.apiPresence.wordApi11, false);
    assert.equal(result.bridgeStatus, "not-connected");
    endpoint = await startLocalEndpoint({
      directory: path.join(app.getPath("userData"), "mcp"),
      createServer: () => createOfficeMcpServer(inspector),
    });
    console.log("Office MCP Electron fixture smoke passed: nested cross-origin frames, fixed capability probe, SDK endpoint startup.");
  } finally {
    await endpoint?.close();
    window?.destroy();
    clearTimeout(deadline);
  }
}).then(() => app.exit(0), (error) => {
  console.error(error);
  app.exit(1);
});
