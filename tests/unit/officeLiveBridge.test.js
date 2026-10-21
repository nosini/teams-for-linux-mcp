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
const { WordSession } = require("../../app/office/addin/wordOperations");
const { createOfficeMcpServer } = require("../../app/mcp/server");
const { startLocalEndpoint } = require("../../app/mcp/localEndpoint");
const { excelFixture } = require("../fixtures/excelApi");
const { wordFixture } = require("../fixtures/wordApi");
const execFileAsync = promisify(execFile);

function fixture(application = "excel") {
  const api = application === "word" ? wordFixture() : excelFixture();
  const { model } = api;
  const id = randomUUID();
  const origin = "https://127.0.0.1:39847";
  const options = { documentId: id, getDocumentUrl: () => application === "word" ? "document.docx" : "book.xlsx" };
  const session = application === "word" ? new WordSession(api.word, options) : new ExcelSession(api.excel, options);
  const scope = vm.createContext({ TeamsOfficeBridge: {
    describe: () => ({ documentId: id, application, ready: true }),
    execute: (request) => session.execute(request),
  } });
  const frame = {
    url: `${origin}/taskpane.html?_host_Info=private`, parent: { url: `https://euc-${application}.officeapps.live.com/x/edit` },
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
  it("discovers Word task panes, binds their application, and rejects tools for the wrong host", async () => {
    const { bridge, frame, id } = fixture("word");
    assert.equal((await bridge.listDocuments()).documents[0].application, "word");
    assert.equal((await bridge.execute(id, "read_paragraphs", {})).paragraphs[0].text, "Hello from Word.");
    await assert.rejects(bridge.execute(id, "read_range", {}), /does not match/);
    frame.parent.url = "https://euc-excel.officeapps.live.com/";
    await assert.rejects(bridge.execute(id, "read_paragraphs", {}), /document changed/);
    assert.equal((await bridge.listDocuments()).documents.length, 0);
    frame.parent.url = "https://euc-word.officeapps.live.com.evil.test/";
    assert.equal((await bridge.listDocuments()).documents.length, 0);
  });

  it("keeps replacement text as data and sanitizes Word renderer errors", async () => {
    const { bridge, model, scope, frame, id } = fixture("word");
    await bridge.listDocuments();
    const before = (await bridge.execute(id, "read_paragraphs", {})).paragraphs[0];
    const replacementText = 'x"); globalThis.compromised = true; //';
    await bridge.execute(id, "replace_text", { ...before, findText: "Hello", replacementText });
    assert.equal(model.paragraphs[0].text, `${replacementText} from Word.`);
    assert.equal(scope.compromised, undefined);
    frame.executeJavaScript = () => Promise.reject(new Error("private document contents"));
    await assert.rejects(bridge.execute(id, "read_paragraph", before), (error) => /Word could not complete/.test(error.message) && !error.message.includes("private"));
  });

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

  it("serves Word reads and replacements through the official MCP client and validates inputs", async (t) => {
    const root = path.resolve(__dirname, "../../test-results/mcp");
    await fs.mkdir(root, { recursive: true });
    const directory = await fs.mkdtemp(path.join(root, "word-"));
    const { bridge, model, id } = fixture("word");
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer({ listEditors: () => ({ editors: [] }) }, bridge) });
    const client = new Client({ name: "word-test", version: "1" });
    t.after(async () => { await client.close(); await endpoint.close(); await fs.rm(directory, { recursive: true, force: true }); });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [path.resolve(__dirname, "../../app/mcp/stdio.js"), endpoint.socketPath] }));
    const call = (name, args = {}) => client.callTool({ name, arguments: args });
    assert.equal((await call("office_list_documents")).structuredContent.documents[0].application, "word");
    const tools = (await client.listTools()).tools;
    assert.equal(tools.find((tool) => tool.name === "word_read_paragraphs").annotations.readOnlyHint, true);
    assert.equal(tools.find((tool) => tool.name === "word_replace_text").annotations.destructiveHint, true);
    const before = (await call("word_read_paragraphs", { documentId: id, count: 1 })).structuredContent.paragraphs[0];
    const target = { documentId: id, paragraphId: before.paragraphId };
    const fresh = (await call("word_read_paragraph", target)).structuredContent;
    const result = await call("word_replace_text", { ...target, readToken: fresh.readToken, findText: "Hello", replacementText: "Updated" });
    assert.equal(result.structuredContent.applied, true);
    assert.equal(result.structuredContent.text, "Updated from Word.");
    const stale = await call("word_replace_text", { ...target, readToken: fresh.readToken, findText: "Updated", replacementText: "Again" });
    assert.equal(stale.isError, true);
    const invalid = await call("word_replace_text", { ...target, readToken: result.structuredContent.readToken, findText: "^p", replacementText: "" });
    assert.equal(invalid.isError, true);
    assert.equal((await call("excel_list_worksheets", { documentId: id })).isError, true);
    assert.equal(model.writes, 1);
  });
});
