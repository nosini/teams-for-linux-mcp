const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const fs = require("node:fs/promises");
const path = require("node:path");
const { randomUUID } = require("node:crypto");
const { execFile } = require("node:child_process");
const { promisify } = require("node:util");
const { Client } = require("@modelcontextprotocol/client");
const { StdioClientTransport } = require("@modelcontextprotocol/client/stdio");
const { LiveOfficeBridge } = require("../../app/office/liveBridge");
const { ExcelSession } = require("../../app/office/addin/operations");
const { createOfficeMcpServer } = require("../../app/mcp/server");
const { startLocalEndpoint } = require("../../app/mcp/localEndpoint");
const { excelFixture } = require("../fixtures/excelApi");
const execFileAsync = promisify(execFile);

function fixture() {
  const { excel, model } = excelFixture();
  const id = randomUUID();
  const origin = "https://127.0.0.1:39847";
  const session = new ExcelSession(excel, { documentId: id, getDocumentUrl: () => "book.xlsx" });
  const scope = vm.createContext({ TeamsOfficeBridge: {
    describe: () => ({ documentId: id, application: "excel", ready: true }),
    execute: (request) => session.execute(request),
  } });
  const frame = {
    url: `${origin}/taskpane.html?_host_Info=private`, parent: { url: "https://euc-excel.officeapps.live.com/x/edit" },
    executeJavaScript: (source, userGesture) => {
      assert.equal(userGesture, false);
      return Promise.resolve(vm.runInContext(source, scope));
    },
  };
  const contents = [{ id: 1, isDestroyed: () => false, mainFrame: { framesInSubtree: [frame] } }];
  const bridge = new LiveOfficeBridge(() => contents, origin);
  return { bridge, frame, model, id, scope };
}

describe("Live Office bridge", () => {
  it("requires discovery and targets only the local task pane under an Excel ancestor", async () => {
    const { bridge, frame, id } = fixture();
    await assert.rejects(bridge.execute(id, "list_worksheets", {}), /document changed/);
    assert.equal((await bridge.listDocuments()).documents[0].id, id);
    assert.equal((await bridge.execute(id, "list_worksheets", {})).worksheets[0].id, "sheet-id");
    frame.parent.url = "https://euc-excel.officeapps.live.com.evil.test/";
    assert.equal((await bridge.listDocuments()).documents.length, 0);
    frame.parent.url = "https://excel.cloud.microsoft/";
    frame.url = "https://evil.test/taskpane.html";
    assert.equal((await bridge.listDocuments()).documents.length, 0);
  });

  it("encodes cell strings as data and rejects disconnected or navigated documents", async () => {
    const { bridge, model, frame, id, scope } = fixture();
    await bridge.listDocuments();
    const target = { worksheetId: "sheet-id", address: "A1" };
    const read = await bridge.execute(id, "read_range", target);
    const value = 'x"); globalThis.compromised = true; //';
    await bridge.execute(id, "write_range", { ...target, readToken: read.readToken, data: [[value]] });
    assert.equal(model.values[0][0], value);
    assert.equal(scope.compromised, undefined);
    frame.url = "https://evil.test/";
    await assert.rejects(bridge.execute(id, "read_range", target), /document changed/);
  });

  it("does not forward renderer errors or allow overlapping commands", async () => {
    const { bridge, frame, id } = fixture();
    await bridge.listDocuments();
    let reject;
    frame.executeJavaScript = () => new Promise((_, rejectPromise) => { reject = rejectPromise; });
    const first = bridge.execute(id, "read_range", {});
    await assert.rejects(bridge.execute(id, "write_range", {}), /Another operation/);
    reject(new Error("private document contents"));
    await assert.rejects(first, (error) => !error.message.includes("private") && /Excel could not complete/.test(error.message));
  });

  it("serves the full read/write workflow through the official MCP client", async (t) => {
    const root = path.resolve(__dirname, "../../test-results/mcp");
    await fs.mkdir(root, { recursive: true });
    const directory = await fs.mkdtemp(path.join(root, "live-"));
    const { bridge, model, id } = fixture();
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer({ listEditors: () => ({ editors: [] }) }, bridge) });
    const client = new Client({ name: "live-test", version: "1" });
    t.after(async () => { await client.close(); await endpoint.close(); await fs.rm(directory, { recursive: true, force: true }); });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [path.resolve(__dirname, "../../app/mcp/stdio.js"), endpoint.socketPath] }));
    const call = (name, args = {}) => client.callTool({ name, arguments: args });
    const docs = await call("office_list_documents");
    assert.equal(docs.structuredContent.documents[0].id, id);
    const command = await execFileAsync(process.execPath, [
      path.resolve(__dirname, "../../scripts/callOfficeMcp.js"), "office_list_documents", "{}", endpoint.socketPath,
    ], { timeout: 5000 });
    assert.equal(JSON.parse(command.stdout).documents[0].id, id);
    const diagnostic = await execFileAsync(process.execPath, [
      path.resolve(__dirname, "../../scripts/diagnoseOfficeMcp.js"), endpoint.socketPath,
    ], { timeout: 5000 });
    assert.equal(JSON.parse(diagnostic.stdout).documents[0].id, id);
    const target = { documentId: id, worksheetId: "sheet-id", address: "A1" };
    const first = await call("excel_read_range", target);
    const updated = await call("excel_write_range", { ...target, readToken: first.structuredContent.readToken, data: [[7]] });
    assert.equal(updated.structuredContent.applied, true);
    assert.deepEqual(model.values, [[7]]);
    const stale = await call("excel_write_range", { ...target, readToken: first.structuredContent.readToken, data: [[9]] });
    assert.equal(stale.isError, true);
    const formula = await call("excel_set_formulas", { ...target, readToken: updated.structuredContent.readToken, data: [["=1+2"]] });
    assert.equal(formula.structuredContent.values[0][0], 3);
    const invalid = await call("excel_write_range", { ...target, readToken: formula.structuredContent.readToken, data: [["=1+3"]] });
    assert.equal(invalid.isError, true);
    assert.equal(model.writes, 2);
  });
});
