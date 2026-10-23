const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const { testPath } = require("../fixtures/testPaths");
const { Client } = require("@modelcontextprotocol/client");
const { StdioClientTransport } = require("@modelcontextprotocol/client/stdio");
const { OfficeMcp } = require("../../app/mcp");

async function userData(t) {
  const root = testPath("mcp");
  await fs.mkdir(root, { recursive: true });
  const directory = await fs.mkdtemp(path.join(root, "startup-"));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  return directory;
}

function create(userDataPath, mcp, graphApiClient = null) {
  return new OfficeMcp({ config: { mcp: { enabled: true, ...mcp } }, userDataPath, getWebContents: () => [], getMainWindow: () => null, graphApiClient });
}

const authority = "net::ERR_CERT_AUTHORITY_INVALID";

describe("Office MCP startup", () => {

  it("slow Office MCP startup leaves shortcuts, updater and download handling initialized", async () => {
    const vm = require("node:vm");
    const source = await fs.readFile(path.resolve(__dirname, "../../app/index.js"), "utf8");
    const start = source.indexOf("    initializeGraphApiClient();", source.indexOf("async function handleAppReady()"));
    const end = source.indexOf("    console.info('[IPC Security]", start);
    assert.ok(start >= 0 && end > start);
    let complete;
    const pending = new Promise((resolve) => { complete = resolve; });
    const calls = [];
    class OfficeMcp {
      static isEnabled() { return true; }
      start() { calls.push("mcp"); return pending; }
    }
    const startup = vm.runInNewContext(`(async () => { let officeMcp; ${source.slice(start, end)} })()`, {
      OfficeMcp, config: {}, app: { getPath: () => "fixture" }, mainAppWindow: {}, webContents: {}, graphApiClient: {}, ipcMain: {}, teamsControlService: null,
      initializeGraphApiClient() {}, registerGraphApiHandlers() {},
      initializeQuickChat: () => calls.push("chat"), registerGlobalShortcuts: () => calls.push("shortcuts"),
      initializeAutoUpdater: () => calls.push("updater"), session: { fromPartition() {} },
      downloadManager: { initialize: () => calls.push("downloads") },
    });
    try { assert.deepEqual(calls, ["chat", "shortcuts", "updater", "downloads", "mcp"]); }
    finally { complete(); await startup; }
  });

  it("serves tools with the app's Graph client and keeps the endpoint when the add-in host fails", async (t) => {
    const directory = await userData(t);
    let tokenRequests = 0;
    const graph = { acquireToken: async () => { tokenRequests++; return { success: false }; } };
    const mcp = create(directory, { addin: { enabled: true, port: 1 }, files: { enabled: true } }, graph);
    t.after(() => mcp.close());
    await mcp.start();
    const client = new Client({ name: "startup-test", version: "1" });
    t.after(() => client.close());
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [path.resolve(__dirname, "../../app/mcp/stdio.js"), path.join(directory, "mcp/bridge.sock")] }));
    const { tools } = await client.listTools();
    assert.ok(tools.some(({ name }) => name === "teams_list_teams"));
    assert.ok(!tools.some(({ name }) => name === "office_list_documents"));
    const result = await client.callTool({ name: "teams_list_teams", arguments: {} });
    assert.equal(result.isError, true);
    assert.equal(tokenRequests, 1);
  });

  it("accepts the add-in certificate only while running and closes a startup in progress", async (t) => {
    const directory = await userData(t);
    const mcp = create(directory, { addin: { enabled: true, port: 0 } });
    await mcp.start();
    const cert = await fs.readFile(path.join(directory, "mcp/addin/localhost.pem"), "utf8");
    const manifest = await fs.readFile(path.join(directory, "mcp/excel-addin.xml"), "utf8");
    const origin = /DefaultValue="(https:\/\/127\.0\.0\.1:\d+)\/taskpane\.html"/.exec(manifest)[1];
    assert.equal(mcp.acceptsCertificate(`${origin}/taskpane.html`, authority, { data: cert }), true);
    await mcp.close();
    assert.equal(mcp.acceptsCertificate(`${origin}/taskpane.html`, authority, { data: cert }), false);
    await assert.rejects(fs.stat(path.join(directory, "mcp/bridge.sock")));

    const early = create(directory, { addin: { enabled: true, port: 0 } });
    const starting = early.start();
    await early.close();
    await starting;
    assert.equal(early.acceptsCertificate(`${origin}/taskpane.html`, authority, { data: cert }), false);
    await assert.rejects(fs.stat(path.join(directory, "mcp/bridge.sock")));
  });
});
