const path = require("node:path");
const { tildify } = require("../utils/homePath");

// Owns the local MCP endpoint and its optional add-in host, channel-file and
// CodeMirror bridges for one app run. Startup failures are logged, never thrown.
class OfficeMcp {
  #config;
  #directory;
  #getWebContents;
  #getMainWindow;
  #graphApiClient;
  #addinHost = null;
  #endpoint = null;
  #closed = false;

  constructor({ config, userDataPath, getWebContents, getMainWindow, graphApiClient = null }) {
    this.#config = config;
    this.#directory = path.join(userDataPath, "mcp");
    this.#getWebContents = getWebContents;
    this.#getMainWindow = getMainWindow;
    this.#graphApiClient = graphApiClient;
  }

  static isEnabled(config) {
    return config.mcp?.enabled === true && process.platform === "linux";
  }

  async start() {
    const mcp = this.#config.mcp;
    try {
      const { OfficeInspector } = require("../office/inspector");
      const { SurfaceInspector } = require("../editor/inspector");
      const { createOfficeMcpServer } = require("./server");
      const { startLocalEndpoint } = require("./localEndpoint");
      const inspector = new OfficeInspector(this.#getWebContents);
      const surfaceInspector = new SurfaceInspector(this.#getWebContents);
      const liveBridge = mcp.addin?.enabled === true ? await this.#startAddin(mcp.addin) : null;
      if (this.#closed) return;
      let channelFiles = null;
      if (mcp.files?.enabled === true) {
        const { ChannelFiles } = require("./channelFiles");
        channelFiles = new ChannelFiles(this.#graphClient());
      }
      let codeMirrorBridge = null;
      if (mcp.codemirror?.enabled === true) {
        const { CodeMirrorBridge } = require("../editor/codeMirrorBridge");
        codeMirrorBridge = new CodeMirrorBridge(surfaceInspector);
      }
      const endpoint = await startLocalEndpoint({
        directory: this.#directory,
        createServer: () => createOfficeMcpServer(inspector, liveBridge, channelFiles, surfaceInspector, codeMirrorBridge),
      });
      if (this.#closed) {
        await endpoint.close();
        return;
      }
      this.#endpoint = endpoint;
      console.info(`[MCP] Office diagnostics available at ${tildify(endpoint.socketPath)}`);
    } catch (error) {
      await this.close();
      // These are local startup/transport errors, never document content or
      // renderer exceptions. Keep their reason so startup is diagnosable.
      console.warn("[MCP] Office diagnostics could not start:", { message: tildify(error.message), code: error.code });
    }
  }

  // Without the add-in host only the live Office tools are missing;
  // diagnostics, channel files and CodeMirror tools still start.
  async #startAddin({ port }) {
    let host = null;
    try {
      const { startAddinHost } = require("../office/addinHost");
      const { LiveOfficeBridge } = require("../office/liveBridge");
      host = await startAddinHost({ directory: this.#directory, port });
      if (this.#closed) {
        await host.close();
        return null;
      }
      const liveBridge = new LiveOfficeBridge(this.#getWebContents, host.origin);
      this.#addinHost = host;
      console.info(`[MCP] Upload this Excel add-in manifest: ${tildify(host.manifestPath)}`);
      console.info(`[MCP] Upload this Word add-in manifest: ${tildify(host.wordManifestPath)}`);
      return liveBridge;
    } catch (error) {
      await host?.close();
      this.#addinHost = null;
      console.warn("[MCP] Office add-in host could not start:", { message: tildify(error.message), code: error.code });
      return null;
    }
  }

  // Channel files share the app's Graph client and its token cache. With the
  // Graph API option off there is none, so the bridge gets its own.
  #graphClient() {
    if (this.#graphApiClient) return this.#graphApiClient;
    const GraphApiClient = require("../graphApi");
    const client = new GraphApiClient({ graphApi: { enabled: true } });
    client.initialize(this.#getMainWindow());
    return client;
  }

  // Only this app instance's own loopback service and exact leaf certificate.
  acceptsCertificate(url, error, certificate) {
    return this.#addinHost?.acceptsCertificate(url, error, certificate) === true;
  }

  async close() {
    this.#closed = true;
    const endpoint = this.#endpoint, addinHost = this.#addinHost;
    this.#endpoint = null;
    this.#addinHost = null;
    await Promise.allSettled([endpoint?.close(), addinHost?.close()]);
  }
}

module.exports = { OfficeMcp };
