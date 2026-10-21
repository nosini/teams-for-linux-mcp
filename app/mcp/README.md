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
- `channelFiles.js`, `channelFileTypes.js`, `fileTransfers.js`: Teams channel
  files over Microsoft Graph, including chunked transfers. They use the app's
  `GraphApiClient` when `graphApi.enabled` is set, and their own otherwise.
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
