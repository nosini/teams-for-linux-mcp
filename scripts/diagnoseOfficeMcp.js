#!/usr/bin/env node
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");
const { Client } = require("@modelcontextprotocol/client");
const { StdioClientTransport } = require("@modelcontextprotocol/client/stdio");

async function findSocket(configHome) {
  // `electron ./app` has no app/package.json, so it keeps Electron's default
  // userData location. Packaged builds load the root package metadata instead.
  const candidates = ["Electron", "teams-for-linux"].map((name) =>
    path.join(configHome, name, "mcp", "bridge.sock"));
  const found = [];
  for (const candidate of candidates) {
    try {
      if ((await fs.stat(candidate)).isSocket()) found.push(candidate);
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  }
  if (found.length === 1) return found[0];
  const error = new Error(found.length > 1 ? "Multiple MCP sockets found" : "No MCP socket found");
  error.code = found.length > 1 ? "AMBIGUOUS_SOCKET" : "ENOENT";
  error.socketPaths = found.length > 1 ? found : candidates;
  throw error;
}

async function diagnoseOfficeMcp(socketPath) {
  if (!path.isAbsolute(socketPath)) throw new Error("Socket path must be absolute");
  if (!(await fs.stat(socketPath)).isSocket()) {
    const error = new Error("MCP endpoint is not a socket");
    error.code = "ENOTSOCK";
    throw error;
  }
  const client = new Client({ name: "teams-office-diagnostic", version: "0.1.0" });
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [path.resolve(__dirname, "../app/mcp/stdio.js"), socketPath],
    stderr: "pipe",
  });
  let connectionError;
  transport.stderr.on("data", (chunk) => {
    // The proxy emits only local connection diagnostics. Preserve an errno,
    // not arbitrary stderr text or MCP response contents.
    const match = /\b(ENOENT|ECONNREFUSED|EACCES|EPERM|ETIMEDOUT)\b/.exec(chunk.toString());
    if (match) connectionError = match[1];
  });
  const call = async (name, args = {}) => {
    const result = await client.callTool({ name, arguments: args }, { timeout: 5000 });
    if (result.isError) throw new Error("Diagnostic tool failed");
    return result.structuredContent;
  };
  try {
    await client.connect(transport, { timeout: 5000 });
    const listing = await call("office_list_editors");
    const editors = [];
    for (const editor of listing.editors) {
      try {
        editors.push(await call("office_inspect_editor", { editorId: editor.id }));
      } catch {
        editors.push({ ...editor, probeStatus: "unavailable", bridgeStatus: "not-connected" });
      }
    }
    return { connected: true, editors, truncated: listing.truncated };
  } catch (error) {
    if (connectionError) error.code = connectionError;
    throw error;
  } finally {
    await client.close();
  }
}

async function main() {
  const arg = process.argv[2];
  if (arg === "--help") {
    console.log("Usage: npm run mcp:diagnose -- [/absolute/path/to/mcp/bridge.sock]");
    console.log("Keep Teams running with --mcp.enabled and a document open in Edit in Teams.");
    return;
  }
  const configHome = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), ".config");
  let socketPath = arg;
  try {
    socketPath ||= await findSocket(configHome);
    console.error(`MCP socket: ${JSON.stringify(socketPath)}`);
    const result = await diagnoseOfficeMcp(socketPath);
    console.log(JSON.stringify(result, null, 2));
    if (result.editors.length === 0) {
      console.error("Connected to Teams, but no recognized Office editor frames were found. Keep a document open in Edit in Teams.");
    }
  } catch (error) {
    const explanations = {
      ENOENT: "No MCP socket was found. The running client may have MCP disabled or use a custom user data directory.",
      ECONNREFUSED: "The socket exists, but no Teams MCP server is listening there.",
      EACCES: "Access to the MCP socket was denied.",
      EPERM: "Access to the MCP socket was denied.",
      ENOTSOCK: "The selected path is not a Unix socket.",
      ETIMEDOUT: "Connecting to the MCP socket timed out.",
      AMBIGUOUS_SOCKET: "More than one MCP socket exists. Pass the socket path printed by the running source client.",
    };
    console.error(`Could not collect Teams MCP diagnostics. ${explanations[error.code] || "MCP communication failed after selecting the socket."}`);
    if (error.socketPaths) {
      for (const candidate of error.socketPaths) console.error(`Socket candidate: ${JSON.stringify(candidate)}`);
    }
    console.error("Check the first terminal for '[MCP] Office diagnostics available at ...'. Pass that exact socket path after npm run mcp:diagnose -- if needed.");
    process.exitCode = 1;
  }
}

if (require.main === module) void main();

module.exports = { diagnoseOfficeMcp, findSocket };
