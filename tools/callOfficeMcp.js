#!/usr/bin/env node
const os = require("node:os");
const path = require("node:path");
const { findSocket, withOfficeClient } = require("./diagnoseOfficeMcp");

async function main() {
  const [name, input, explicitSocket] = process.argv.slice(2);
  if (name === "--help" || !name || !input) {
    console.log("Usage: npm run mcp:call -- <tool-name> '<JSON arguments>' [/absolute/path/to/mcp/bridge.sock]");
    if (name !== "--help") process.exitCode = 1;
    return;
  }
  try {
    const args = JSON.parse(input);
    if (!args || typeof args !== "object" || Array.isArray(args)) throw new Error("Arguments must be a JSON object");
    const configHome = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), ".config");
    const socketPath = explicitSocket || await findSocket(configHome);
    const result = await withOfficeClient(socketPath, (client) => client.callTool({ name, arguments: args }, { timeout: 20000 }));
    console.log(JSON.stringify(result.structuredContent || result, null, 2));
    if (result.isError) process.exitCode = 1;
  } catch {
    console.error("MCP call failed. Check the tool name, JSON arguments and the socket path. Use npm run mcp:diagnose to check the connection. Read cells before retrying any write; its outcome may be unknown.");
    process.exitCode = 1;
  }
}

void main();
