#!/usr/bin/env node
// Agent-launched stdio adapter. The running Electron app owns the MCP server;
// this process only forwards bytes and never launches a second Teams instance.
const net = require("node:net");
const path = require("node:path");

function run(socketPath) {
  if (!socketPath || !path.isAbsolute(socketPath)) {
    console.error("Usage: node app/mcp/stdio.js /absolute/path/to/mcp/bridge.sock");
    process.exitCode = 1;
    return;
  }
  const socket = net.createConnection(socketPath);
  let stopping = false, failed = false;
  const unexpectedClose = () => {
    console.error("Teams MCP connection closed unexpectedly. The connection limit may have been reached or the app stopped.");
    process.exitCode = 1;
  };
  const stop = () => {
    stopping = true;
    process.stdin.unpipe(socket);
    process.stdin.pause();
    socket.destroy();
  };
  const timeout = setTimeout(() => {
    const error = new Error("Connection timeout");
    error.code = "ETIMEDOUT";
    socket.destroy(error);
  }, 5000);
  socket.once("connect", () => {
    clearTimeout(timeout);
    process.stdin.pipe(socket);
    socket.pipe(process.stdout);
    if (process.stdin.readableEnded) socket.end();
  });
  socket.once("error", (error) => {
    failed = true;
    const code = ["ENOENT", "ECONNREFUSED", "EACCES", "EPERM", "ETIMEDOUT"].includes(error.code) ? error.code : "CONNECTION_FAILED";
    if (["ECONNRESET", "EPIPE"].includes(error.code)) unexpectedClose();
    else console.error(`Cannot connect to Teams MCP (${code}). Socket: ${JSON.stringify(socketPath)}`);
    process.exitCode = 1;
    stop();
  });
  socket.once("close", () => {
    clearTimeout(timeout);
    if (!stopping && !failed && !process.stdin.readableEnded) {
      unexpectedClose();
    }
    stop();
  });
  process.stdin.once("error", stop);
  process.stdout.once("error", stop);
  process.once("SIGINT", stop);
  process.once("SIGTERM", stop);
}

if (require.main === module) run(process.argv[2]);
