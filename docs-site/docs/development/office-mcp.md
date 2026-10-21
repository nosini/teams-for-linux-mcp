# Office MCP bridge

The Office MCP bridge lets a local agent discover Word and Excel editor frames
inside the running Teams client. It currently provides read-only diagnostics;
it cannot read cells, read document text, or make edits. The purpose of the
diagnostic is to establish how the embedded editor can connect to a supported
Office add-in before implementing document operations.

## Run from source

Use Node.js 22 or later. Install the dependencies with `npm ci` in the repository,
then add the following setting to the client's `config.json`, merging it with
any existing settings:

```json
{
  "mcp": { "enabled": true }
}
```

Alternatively, enable the bridge for a single source run with
`npm start -- --mcp.enabled`. Open an Excel or Word file using **Edit in
Teams**. The app prints the socket location as `[MCP] Office diagnostics available
at …/mcp/bridge.sock`. The source command runs `electron ./app`, so its default
location is `~/.config/Electron/mcp/bridge.sock`. Packaged builds normally use
`~/.config/teams-for-linux/mcp/bridge.sock`. A custom user data directory changes
this location. The bridge currently supports Linux source runs. Packaged and
Flatpak agent connections have not been verified.

With the document open, collect the diagnostic results in another terminal:

```sh
npm run mcp:diagnose
```

This connects to the running app, lists the candidate editors, inspects each
one, prints the results as JSON, and exits. It does not require configuring an
agent. It checks the `Electron` and `teams-for-linux` directories under
`XDG_CONFIG_HOME` or `~/.config`. If both contain sockets, pass an explicit path
to select the instance. A connection failure reports the socket location and
whether it is missing, not listening, or inaccessible.
For a custom user data directory, pass the socket location printed by Teams:

```sh
npm run mcp:diagnose -- /absolute/path/to/user-data/mcp/bridge.sock
```

Add a stdio server to your agent's MCP configuration using absolute paths:

```json
{
  "mcpServers": {
    "teams-office": {
      "command": "node",
      "args": [
        "/absolute/path/to/teams-for-linux/app/mcp/stdio.js",
        "/absolute/path/to/user-data/mcp/bridge.sock"
      ]
    }
  }
}
```

The adapter attaches to the running client; it does not launch Teams. The server
provides these tools:

| Tool | Arguments | Result |
| --- | --- | --- |
| `office_list_editors` | None | Candidate frames, temporary IDs, application hints and origins |
| `office_inspect_editor` | `editorId` from discovery | Office.js API presence and probe status |

Ask the agent to list the editors and inspect each candidate. IDs can change
when the document reloads. A SharePoint document wrapper and the editor inside
it may both appear; wrappers are identified separately. An empty list means no
recognized frame is open, not that the client cannot edit documents. Discovery
currently recognizes public-cloud Office Apps, Word/Excel cloud hosts and
SharePoint document wrappers; sovereign-cloud and Defender proxy URLs are not
recognized yet.

`probeStatus: "completed"` means the diagnostic ran. It does not mean that the
document is accessible through Office.js. `bridgeStatus: "not-connected"` is
expected until the live document adapter is implemented. Office.js normally
provides document APIs inside an Office add-in; globals in the surrounding
editor are not a supported substitute for that add-in context. The probe neither
loads Office.js nor calls `Office.onReady`, `Excel.run` or `Word.run`.

## Local connection and implementation

MCP is disabled by default and requires a restart when enabled or disabled. The
app serves MCP through a Unix socket with mode `0600` in a user-owned directory
with mode `0700`. It opens no TCP port and exposes no new renderer IPC methods.
Only local processes with access to that directory can attach. These permissions
do not distinguish different applications running as the same user.

Discovery returns origins rather than full document URLs. The diagnostic never
returns document content, sharing tokens, or renderer error messages. It runs
one fixed script in recognized frames, with a timeout; callers cannot supply
JavaScript. Results are observations from page code, not a trusted authorization
decision for future editing tools.

`app/mcp/server.js` defines the tools using the official MCP SDK.
`app/mcp/localEndpoint.js` binds SDK transports to local connections, and
`app/mcp/stdio.js` forwards an agent's stdio connection to the app.
`app/office/inspector.js` owns frame discovery and capability inspection. Document
operations should be implemented behind a separate live adapter, using stable
document targeting and checking the current content before writes. Local-file
servers are not used to overwrite documents being coauthored in Teams.

Run the focused checks with:

```sh
node --test tests/unit/officeInspector.test.js tests/unit/officeMcp.test.js
npm run lint
```

To check the Electron frame APIs against nested cross-origin fixture pages, run
`./node_modules/.bin/electron tests/electron/officeMcpSmoke.js`. A headless Linux
test machine can use `xvfb-run -a` before that command. This test serves local
fixture responses for every HTTPS request and does not contact Microsoft or
verify a real signed-in Office editor.
