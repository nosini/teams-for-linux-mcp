const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const net = require("node:net");
const path = require("node:path");
const { testPath } = require("../fixtures/testPaths");
const { once } = require("node:events");
const { spawn, execFile } = require("node:child_process");
const { promisify } = require("node:util");
const { createInterface } = require("node:readline");
const { startLocalEndpoint } = require("../../app/mcp/localEndpoint");
const { createOfficeMcpServer } = require("../../app/mcp/server");
const { OfficeInspector } = require("../../app/office/inspector");
const { findSocket } = require("../../scripts/diagnoseOfficeMcp");
const { exportConnection } = require("../../scripts/setupOfficeMcp");
const execFileAsync = promisify(execFile);
const diagnosticScript = path.resolve(__dirname, "../../scripts/diagnoseOfficeMcp.js");

async function testDirectory(t) {
  const root = testPath("mcp");
  await fs.mkdir(root, { recursive: true });
  const directory = await fs.mkdtemp(path.join(root, "run-"));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  return directory;
}

function rpcReader(input, output) {
  let id = 0;
  const pending = new Map();
  const reader = createInterface({ input });
  reader.on("line", (line) => {
    const message = JSON.parse(line);
    if (pending.has(message.id)) {
      const resolve = pending.get(message.id);
      pending.delete(message.id);
      resolve(message);
    }
  });
  return {
    async request(method, params = {}) {
      const requestId = ++id;
      const response = new Promise((resolve) => pending.set(requestId, resolve));
      output.write(`${JSON.stringify({ jsonrpc: "2.0", id: requestId, method, params })}\n`);
      return response;
    },
    notify(method) { output.write(`${JSON.stringify({ jsonrpc: "2.0", method })}\n`); },
    close() { reader.close(); },
  };
}

async function initialize(client) {
  const response = await client.request("initialize", {
    protocolVersion: "2025-11-25", capabilities: {}, clientInfo: { name: "test-agent", version: "1" },
  });
  assert.equal(response.result.serverInfo.name, "teams-for-linux-office");
  client.notify("notifications/initialized");
}

describe("Office MCP endpoint", { timeout: 10000 }, () => {
  it("exports a standalone working relay without overwriting existing agent configuration", async (t) => {
    const directory = await testDirectory(t);
    const endpoint = await startLocalEndpoint({ directory: path.join(directory, "endpoint"), createServer: () => createOfficeMcpServer(new OfficeInspector(() => [])) });
    t.after(() => endpoint.close());
    const output = path.join(directory, "export");
    const result = await exportConnection(endpoint.socketPath, output);
    assert.equal(result.live, false);
    const config = JSON.parse(await fs.readFile(result.configPath, "utf8"));
    assert.deepEqual(config.mcpServers["teams-office"], { command: process.execPath, args: [result.relayPath, endpoint.socketPath] });
    assert.deepEqual((await fs.readdir(output)).sort(), ["mcp.json", "teams-office-stdio.js"]);
    const child = spawn(process.execPath, [result.relayPath, endpoint.socketPath], { stdio: ["pipe", "pipe", "pipe"] });
    t.after(() => child.kill());
    const client = rpcReader(child.stdout, child.stdin);
    t.after(() => client.close());
    await initialize(client);
    assert.equal((await client.request("tools/list")).result.tools.some((tool) => tool.name === "office_list_editors"), true);
    await assert.rejects(exportConnection(endpoint.socketPath, output), { code: "EEXIST" });
    assert.deepEqual(JSON.parse(await fs.readFile(result.configPath, "utf8")), config);
    // A conflicting config must also roll back the newly copied relay.
    await fs.unlink(result.relayPath);
    await assert.rejects(exportConnection(endpoint.socketPath, output), { code: "EEXIST" });
    await assert.rejects(fs.stat(result.relayPath), { code: "ENOENT" });
    const publicOutput = path.join(directory, "public");
    await fs.mkdir(publicOutput, { mode: 0o755 });
    await assert.rejects(exportConnection(endpoint.socketPath, publicOutput), /private/);
    await assert.rejects(exportConnection("relative.sock", output), /absolute/);
  });

  it("collects editor diagnostics with the SDK client and distinguishes an empty editor list", async (t) => {
    const directory = await testDirectory(t);
    let contents = [];
    const inspector = new OfficeInspector(() => contents);
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer(inspector) });
    t.after(() => endpoint.close());
    const options = { timeout: 5000 };
    const empty = await execFileAsync(process.execPath, [diagnosticScript, endpoint.socketPath], options);
    assert.deepEqual(JSON.parse(empty.stdout), { connected: true, editors: [], truncated: false });
    assert.match(empty.stderr, /Connected to Teams, but no recognized Office editor/);
    contents = [{
      id: 1, isDestroyed: () => false,
      mainFrame: { framesInSubtree: [{
        processId: 2, routingId: 3, url: "https://excel.cloud.microsoft/?secret=token",
        executeJavaScript: async () => ({ officeGlobal: true }),
      }] },
    }];
    const populated = await execFileAsync(process.execPath, [diagnosticScript, endpoint.socketPath], options);
    const result = JSON.parse(populated.stdout);
    assert.equal(result.connected, true);
    assert.equal(result.editors.length, 1);
    assert.equal(result.editors[0].apiPresence.officeGlobal, true);
    assert.equal(result.editors[0].bridgeStatus, "not-connected");
    assert.doesNotMatch(populated.stdout, /secret|token/);
    assert.match(populated.stderr, /MCP socket:/);
    assert.doesNotMatch(populated.stderr, /Could not collect/);
  });

  it("discovers the source-run socket, supports the packaged location, and rejects ambiguity", async (t) => {
    const directory = await testDirectory(t);
    const createServer = () => createOfficeMcpServer(new OfficeInspector(() => []));
    await assert.rejects(findSocket(directory), { code: "ENOENT" });
    const source = await startLocalEndpoint({ directory: path.join(directory, "Electron", "mcp"), createServer });
    t.after(() => source.close());
    assert.equal(await findSocket(directory), source.socketPath);
    const result = await execFileAsync(process.execPath, [diagnosticScript], {
      timeout: 5000, env: { ...process.env, XDG_CONFIG_HOME: directory },
    });
    assert.equal(JSON.parse(result.stdout).connected, true);
    assert.ok(result.stderr.includes(source.socketPath));
    const packaged = await startLocalEndpoint({ directory: path.join(directory, "teams-for-linux", "mcp"), createServer });
    t.after(() => packaged.close());
    await assert.rejects(findSocket(directory), { code: "AMBIGUOUS_SOCKET" });
    await source.close();
    assert.equal(await findSocket(directory), packaged.socketPath);
  });

  it("serves a real stdio agent through the socket, validates tools and cleans up", async (t) => {
    const directory = await testDirectory(t);
    const inspector = new OfficeInspector(() => [{
      id: 1, isDestroyed: () => false,
      mainFrame: { framesInSubtree: [{
        processId: 2, routingId: 3, url: "https://excel.cloud.microsoft/?secret=token",
        executeJavaScript: async () => ({ officeGlobal: true }),
      }] },
    }]);
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer(inspector) });
    const child = spawn(process.execPath, [path.resolve(__dirname, "../../app/mcp/stdio.js"), endpoint.socketPath]);
    const exited = once(child, "exit");
    const errors = [];
    child.stderr.on("data", (chunk) => errors.push(chunk.toString()));
    const client = rpcReader(child.stdout, child.stdin);
    t.after(async () => { client.close(); child.kill(); await endpoint.close(); });

    assert.equal((await fs.stat(directory)).mode & 0o777, 0o700);
    assert.equal((await fs.stat(endpoint.socketPath)).mode & 0o777, 0o600);
    await initialize(client);
    const listing = await client.request("tools/list");
    assert.deepEqual(listing.result.tools.map((tool) => tool.name).sort(), ["office_inspect_editor", "office_list_editors"]);
    assert.ok(listing.result.tools.every((tool) => tool.annotations.readOnlyHint));
    const editors = await client.request("tools/call", { name: "office_list_editors", arguments: {} });
    assert.equal(editors.result.structuredContent.editors[0].id, "1:2:3");
    assert.doesNotMatch(JSON.stringify(editors), /secret|token/);
    const inspected = await client.request("tools/call", { name: "office_inspect_editor", arguments: { editorId: "1:2:3" } });
    assert.equal(inspected.result.structuredContent.apiPresence.officeGlobal, true);
    assert.equal(inspected.result.structuredContent.bridgeStatus, "not-connected");
    const missing = await client.request("tools/call", { name: "office_inspect_editor", arguments: { editorId: "9:9:9" } });
    assert.equal(missing.result.isError, true);
    const invalid = await client.request("tools/call", { name: "office_inspect_editor", arguments: { editorId: "javascript:alert(1)" } });
    assert.ok(invalid.error || invalid.result?.isError);
    child.stdin.end();
    assert.deepEqual(await exited, [0, null]);
    assert.deepEqual(errors, []);
  });

  it("refuses permissive directories, symlinks and non-socket files", async (t) => {
    const directory = await testDirectory(t);
    const createServer = () => createOfficeMcpServer(new OfficeInspector(() => []));
    await fs.chmod(directory, 0o755);
    await assert.rejects(startLocalEndpoint({ directory, createServer }), /private directory/);
    await fs.chmod(directory, 0o700);
    const link = path.join(directory, "link");
    await fs.symlink(directory, link);
    await assert.rejects(startLocalEndpoint({ directory: link, createServer }), /private directory/);
    const socketPath = path.join(directory, "bridge.sock");
    await fs.writeFile(socketPath, "keep me");
    await assert.rejects(startLocalEndpoint({ directory, createServer }), /not a socket/);
    assert.equal(await fs.readFile(socketPath, "utf8"), "keep me");
  });

  it("does not unlink a running endpoint and closes idle clients on shutdown", async (t) => {
    const directory = await testDirectory(t);
    const createServer = () => createOfficeMcpServer(new OfficeInspector(() => []));
    const endpoint = await startLocalEndpoint({ directory, createServer });
    t.after(() => endpoint.close());
    const original = await fs.stat(endpoint.socketPath);
    await assert.rejects(startLocalEndpoint({ directory, createServer }), /already running/);
    assert.equal((await fs.stat(endpoint.socketPath)).ino, original.ino);
    const socket = net.createConnection(endpoint.socketPath);
    await once(socket, "connect");
    const client = rpcReader(socket, socket);
    t.after(() => { client.close(); socket.destroy(); });
    await initialize(client);
    assert.deepEqual((await client.request("tools/call", { name: "office_list_editors" })).result.structuredContent.editors, []);
    const closed = once(socket, "close");
    await endpoint.close();
    await closed;
    await assert.rejects(fs.stat(endpoint.socketPath), { code: "ENOENT" });
  });

  it("recovers the socket left by an abruptly terminated process", async (t) => {
    const directory = await testDirectory(t);
    const socketPath = path.join(directory, "bridge.sock");
    const child = spawn(process.execPath, ["-e", "require('node:net').createServer().listen(process.argv[1], () => process.stdout.write('ready\\n'))", socketPath]);
    t.after(() => child.kill());
    await once(child.stdout, "data");
    const exited = once(child, "exit");
    child.kill("SIGKILL");
    await exited;
    assert.ok((await fs.stat(socketPath)).isSocket());
    await assert.rejects(
      execFileAsync(process.execPath, [diagnosticScript, socketPath], { timeout: 5000 }),
      (error) => error.code === 1 && /socket exists, but no Teams MCP server is listening/.test(error.stderr),
    );
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer(new OfficeInspector(() => [])) });
    await endpoint.close();
    await assert.rejects(fs.stat(socketPath), { code: "ENOENT" });
  });

  it("serves channel uploads larger than 64 KiB through the actual MCP socket", async (t) => {
    const { ChannelFiles } = require("../../app/mcp/channelFiles");
    const directory = await testDirectory(t);
    const uploaded = [];
    const files = new ChannelFiles({ acquireToken: async () => ({ success: true, token: "private" }) }, async (_url, options) => {
      uploaded.push(options.body);
      return Response.json({ id: "new-file", name: "test.txt", file: {}, eTag: '"one"', "@microsoft.graph.downloadUrl": "private-link" });
    });
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer(new OfficeInspector(() => []), null, files) });
    const socket = net.createConnection(endpoint.socketPath);
    await once(socket, "connect");
    const client = rpcReader(socket, socket);
    t.after(async () => { client.close(); socket.destroy(); await endpoint.close(); });
    await initialize(client);
    const listing = await client.request("tools/list");
    assert.equal(listing.result.tools.filter((tool) => tool.name.startsWith("teams_")).length, 14);
    const bytes = Buffer.alloc(128 * 1024, 65);
    const response = await client.request("tools/call", { name: "teams_create_file", arguments: { driveId: "d", parentId: "p", name: "test.txt", contentBase64: bytes.toString("base64") } });
    assert.equal(response.result.isError, undefined);
    assert.equal(response.result.structuredContent.item.id, "new-file");
    assert.deepEqual(uploaded, [bytes]);
    assert.doesNotMatch(JSON.stringify(response), /private/);
    const invalid = await client.request("tools/call", { name: "teams_create_file", arguments: { driveId: "d", parentId: "p", name: "../bad", contentBase64: "" } });
    assert.ok(invalid.error || invalid.result?.isError);
    assert.equal(uploaded.length, 1);
  });

  it("round-trips a file larger than 1 MiB in chunks through the MCP socket", async (t) => {
    const { ChannelFiles } = require("../../app/mcp/channelFiles");
    const { CHUNK_SIZE } = require("../../app/mcp/fileTransfers");
    const directory = await testDirectory(t);
    const source = Buffer.alloc(CHUNK_SIZE * 2 + 17, 93), received = [];
    let stored, published = false;
    const item = { id: "file", name: "chunked.bin", size: source.length, eTag: '"one"', file: {} };
    const files = new ChannelFiles({ acquireToken: async () => ({ success: true, token: "private-token" }) }, async (url, options) => {
      if (url.endsWith("/createUploadSession")) return Response.json({ uploadUrl: "https://tenant.sharepoint.com/upload?private=session", expirationDateTime: new Date(Date.now() + 3600000).toISOString() });
      if (url.includes("sharepoint")) {
        assert.equal(options.headers.Authorization, undefined);
        if (options.method === "PUT") {
          received.push(options.body);
          const offset = Buffer.concat(received).length;
          return Response.json({ nextExpectedRanges: offset === source.length ? [] : [`${offset}-${source.length - 1}`] }, { status: 202 });
        }
        if (options.method === "POST") {
          stored = Buffer.concat(received); published = true;
          return Response.json(item, { status: 201 });
        }
        const [, start, end] = /^bytes=(\d+)-(\d+)$/.exec(options.headers.Range);
        return new Response(stored.subarray(Number(start), Number(end) + 1), { status: 206, headers: { "Content-Range": `bytes ${start}-${end}/${stored.length}` } });
      }
      assert.equal(published, true);
      return Response.json({ ...item, "@microsoft.graph.downloadUrl": "https://tenant.sharepoint.com/download?private=link" });
    });
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer(new OfficeInspector(() => []), null, files) });
    const socket = net.createConnection(endpoint.socketPath);
    await once(socket, "connect");
    const client = rpcReader(socket, socket);
    t.after(async () => { client.close(); socket.destroy(); await endpoint.close(); });
    await initialize(client);
    const call = async (name, args) => {
      const response = await client.request("tools/call", { name, arguments: args });
      assert.equal(response.error, undefined);
      assert.equal(response.result.isError, undefined, JSON.stringify(response));
      assert.doesNotMatch(JSON.stringify(response), /private/);
      return response.result.structuredContent;
    };
    const upload = await call("teams_begin_upload", { driveId: "d", parentId: "p", name: "chunked.bin", size: source.length });
    for (let offset = 0; offset < source.length; offset += CHUNK_SIZE) await call("teams_upload_chunk", { transferId: upload.transferId, offset, contentBase64: source.subarray(offset, offset + CHUNK_SIZE).toString("base64") });
    assert.equal(published, false);
    assert.equal((await call("teams_finish_upload", { transferId: upload.transferId })).status, "completed");
    assert.deepEqual(stored, source);
    const download = await call("teams_begin_download", { driveId: "d", itemId: "file" });
    const chunks = [];
    for (let offset = 0; offset < source.length;) {
      const chunk = await call("teams_read_file_chunk", { transferId: download.transferId, offset });
      chunks.push(Buffer.from(chunk.contentBase64, "base64"));
      offset = chunk.nextOffset;
    }
    assert.deepEqual(Buffer.concat(chunks), source);
  });

  it("fails the fifth relay with a diagnostic and accepts another after a slot is freed", async (t) => {
    const directory = await testDirectory(t);
    const endpoint = await startLocalEndpoint({ directory, createServer: () => createOfficeMcpServer(new OfficeInspector(() => [])) });
    const sockets = [];
    t.after(async () => { for (const socket of sockets) socket.destroy(); await endpoint.close(); });
    for (let index = 0; index < 4; index++) {
      const socket = net.createConnection(endpoint.socketPath); sockets.push(socket);
      await once(socket, "connect");
      const client = rpcReader(socket, socket);
      await initialize(client); client.close();
    }
    const child = spawn(process.execPath, [path.resolve(__dirname, "../../app/mcp/stdio.js"), endpoint.socketPath]);
    t.after(() => child.kill());
    const errors = []; child.stderr.on("data", (data) => errors.push(data.toString()));
    assert.deepEqual(await once(child, "exit"), [1, null]);
    assert.match(errors.join(""), /connection limit/);
    sockets.pop().destroy();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const socket = net.createConnection(endpoint.socketPath); sockets.push(socket);
    await once(socket, "connect");
    const client = rpcReader(socket, socket); t.after(() => client.close());
    await initialize(client);
  });

  it("fails the stdio adapter clearly when Teams is not running", async (t) => {
    const directory = await testDirectory(t);
    const child = spawn(process.execPath, [path.resolve(__dirname, "../../app/mcp/stdio.js"), path.join(directory, "missing.sock")]);
    const errors = [];
    child.stderr.on("data", (chunk) => errors.push(chunk.toString()));
    assert.deepEqual(await once(child, "exit"), [1, null]);
    assert.match(errors.join(""), /Cannot connect to Teams MCP/);
    await assert.rejects(
      execFileAsync(process.execPath, [diagnosticScript, path.join(directory, "missing.sock")], { timeout: 5000 }),
      (error) => error.code === 1 && /No MCP socket was found/.test(error.stderr),
    );
  });
});
