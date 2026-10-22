# Local MCP Endpoint

Exposes live Office, CodeMirror and Teams channel-file tools to local agents
through a private Unix socket. Linux only, disabled by default. See the
[user guide](../../docs-site/docs/office-agents.md) and the
[developer notes](../../docs-site/docs/development/office-mcp.md).

## Files

- `index.js`: `OfficeMcp`, the entry point `app/index.js` uses. It starts the
  socket, the optional add-in host and bridges, accepts the add-in host's own
  loopback certificate, and closes everything on quit. A failing add-in host
  only removes the live Office tools.
- `localEndpoint.js`: the private `mcp/bridge.sock` listener.
- `server.js`: the MCP server and its diagnostic and live Office tools.
- `editingTools.js`, `codeMirrorTools.js`: Office editing and CodeMirror tools.
- `documentTabTools.js`: background document tab opening, discovery, showing and
  closing. `app/office/documentTabs.js` owns the separate Documents window and
  its sandboxed Office views. Closing the window hides it; individual tab closes
  unload a document. All views are disposed when Teams quits.
- `app/office/bridgeSetup.js`, `bridgeSetupUi.js`: bounded first-time add-in
  import through Office's named UI controls, exposed as `office_setup_bridge`.
  Only this app's generated manifest is uploaded; no caller scripts or file paths
  are accepted. Setup progress is included in tab discovery.
- `app/office/bridgeSetupUiLoader.js`: reloads the fixed UI probe after source
  changes in a checkout launch; packaged builds keep normal module caching.
- `app/office/nativeSetupInput.js`: Chromium activation of a validated named
  setup control, with native mouse input for Add-ins, Enter for later steps,
  temporary focus emulation and debugger cleanup. `setupMousePoint.js` validates
  frame geometry and `setupMouseSession.js` routes to the selected child frame.
  No remote debugging port or caller-provided keys/coordinates/debugger commands.
- `channelFiles.js`, `channelFileTypes.js`, `fileTransfers.js`: Teams channel
  files over Microsoft Graph, including chunked transfers. They use the app's
  `GraphApiClient` when `graphApi.enabled` is set, and their own otherwise.
- `tools.js`: `registerJsonTool`, the shared tool result, error and annotation
  wrapper. Every tool states whether it reaches Microsoft services.
- `evaluateInFrame.js`: runs the fixed bridge and probe scripts in a renderer
  frame with a timeout.
- `stdio.js`: the relay an agent runs to connect its stdio to the socket.

The add-in host and live bridge live in `app/office/`, the CodeMirror bridge in
`app/editor/`.

## Configuration

```json
{
  "mcp": {
    "enabled": false,
    "addin": { "enabled": false, "port": 39847 },
    "files": { "enabled": false },
    "codemirror": { "enabled": false }
  }
}
```
