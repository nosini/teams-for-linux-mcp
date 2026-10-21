const fs = require("node:fs/promises");
const net = require("node:net");
const path = require("node:path");
const { StdioServerTransport, serveStdio } = require("@modelcontextprotocol/server/stdio");

async function removeStaleSocket(socketPath) {
  let stat;
  try {
    stat = await fs.lstat(socketPath);
  } catch (error) {
    if (error.code === "ENOENT") return;
    throw error;
  }
  if (!stat.isSocket()) throw new Error("MCP endpoint path is not a socket");
  await new Promise((resolve, reject) => {
    const probe = net.createConnection(socketPath);
    probe.setTimeout(1000, () => probe.destroy(new Error("MCP endpoint is busy")));
    probe.once("connect", () => {
      probe.destroy();
      reject(new Error("MCP endpoint is already running"));
    });
    probe.once("error", (error) => error.code === "ECONNREFUSED" ? resolve() : reject(error));
  });
  const current = await fs.lstat(socketPath);
  if (!current.isSocket() || current.ino !== stat.ino || current.dev !== stat.dev) {
    throw new Error("MCP endpoint changed during startup");
  }
  await fs.unlink(socketPath);
}

async function startLocalEndpoint({ directory, createServer }) {
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const stat = await fs.lstat(directory);
  if (!stat.isDirectory() || stat.uid !== process.getuid() || (stat.mode & 0o077) !== 0) {
    throw new Error("MCP directory must be a private directory owned by this user");
  }
  const socketPath = path.join(directory, "bridge.sock");
  if (Buffer.byteLength(socketPath) > 103) throw new Error("MCP socket path is too long");
  await removeStaleSocket(socketPath);
  const connections = new Map();
  const listener = net.createServer((socket) => {
    if (connections.size >= 4) {
      socket.destroy();
      return;
    }
    // The SDK supports separate streams, including a duplex Unix socket. Its
    // stdio binding negotiates both legacy and current MCP clients for us.
    const transport = new StdioServerTransport(socket, socket, { maxBufferSize: 64 * 1024 });
    const handle = serveStdio(createServer, {
      transport,
      onerror: () => socket.destroy(),
    });
    const closeTransport = transport.onclose;
    transport.onclose = () => {
      closeTransport?.();
      socket.destroy();
    };
    connections.set(socket, handle);
    socket.on("error", () => {});
    socket.once("close", () => {
      connections.delete(socket);
      void handle.close().catch(() => {});
    });
  });
  await new Promise((resolve, reject) => {
    listener.once("error", reject);
    listener.listen(socketPath, resolve);
  });
  try {
    await fs.chmod(socketPath, 0o600);
  } catch (error) {
    listener.close();
    throw error;
  }
  // Keep runtime transport errors out of Teams' fatal uncaught-error handler.
  listener.on("error", () => {});
  let closing;
  return {
    socketPath,
    close() {
      if (closing) return closing;
      for (const [socket, handle] of connections) {
        socket.destroy();
        void handle.close().catch(() => {});
      }
      closing = new Promise((resolve, reject) => listener.close((error) => error ? reject(error) : resolve()));
      return closing;
    },
  };
}

module.exports = { startLocalEndpoint };
