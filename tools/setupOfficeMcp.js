#!/usr/bin/env node
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { findSocket, withOfficeClient } = require("./diagnoseOfficeMcp");

async function exportConnection(socketPath, outputDirectory) {
  if (!path.isAbsolute(socketPath) || !path.isAbsolute(outputDirectory)) {
    throw new Error("Socket and output paths must be absolute");
  }
  if (!(await fs.stat(socketPath)).isSocket()) throw new Error("The selected endpoint is not a Unix socket");
  // Check the running server before exporting a configuration for it.
  const { tools } = await withOfficeClient(socketPath, (client) => client.listTools({}, { timeout: 5000 }));
  const live = tools.some((tool) => tool.name === "office_list_documents");
  await fs.mkdir(outputDirectory, { recursive: true, mode: 0o700 });
  const stat = await fs.lstat(outputDirectory);
  if (!stat.isDirectory() || stat.uid !== process.getuid() || (stat.mode & 0o077)) {
    throw new Error("Output directory must be private and owned by this user");
  }
  const relayPath = path.join(outputDirectory, "teams-office-stdio.js");
  const configPath = path.join(outputDirectory, "mcp.json");
  const config = { mcpServers: { "teams-office": { command: process.execPath, args: [relayPath, socketPath] } } };
  // Refuse to overwrite files, including symlinks. This directory is an export,
  // never an agent's existing settings directory.
  const created = [];
  try {
    await fs.copyFile(path.resolve(__dirname, "../app/mcp/stdio.js"), relayPath, fs.constants.COPYFILE_EXCL);
    created.push(relayPath);
    await fs.writeFile(configPath, `${JSON.stringify(config, null, 2)}\n`, { flag: "wx", mode: 0o600 });
    created.push(configPath);
  } catch (error) {
    for (const file of created) await fs.unlink(file);
    throw error;
  }
  return { socketPath, relayPath, configPath, live, manifests: {
    excel: path.join(path.dirname(socketPath), "excel-addin.xml"),
    word: path.join(path.dirname(socketPath), "word-addin.xml"),
  } };
}

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 1 || args.length > 2 || args.includes("--help")) {
    console.log("Usage: npm run mcp:setup -- /absolute/export-directory [/absolute/path/to/mcp/bridge.sock]");
    console.log("Teams must be running with --mcp.enabled --mcp.addin.enabled. Exports a standalone relay and MCP JSON without changing agent settings.");
    if (!args.includes("--help")) process.exitCode = 1;
    return;
  }
  try {
    const socketPath = args[1] || await findSocket(process.env.XDG_CONFIG_HOME || path.join(os.homedir(), ".config"));
    const result = await exportConnection(socketPath, args[0]);
    console.log(JSON.stringify(result, null, 2));
    console.error(`Merge the teams-office entry from ${JSON.stringify(result.configPath)} into your agent's MCP settings, then reconnect the agent.`);
    if (!result.live) console.error("Teams is running in diagnostics-only mode. Restart with --mcp.addin.enabled for document tools.");
    console.error("Container agents need the relay and socket explicitly mounted by their host launcher. See docs/office-agents.md.");
  } catch (error) {
    console.error(`Could not export Teams Office setup: ${error.code || error.message}`);
    process.exitCode = 1;
  }
}

if (require.main === module) void main();
module.exports = { exportConnection };
