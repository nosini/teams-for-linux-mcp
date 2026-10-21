# Office MCP bridge

The Office MCP bridge connects local agents to documents inside the running
Teams client. Diagnostics discover Word and Excel frames. The optional Excel
task pane provides live worksheet reads, value writes and formula writes through
Office.js. The Word task pane reads paragraphs and replaces exact text inside a
chosen paragraph using the same local bridge.

For installation, agent configuration and container connections, see
[Connect an agent to Word and Excel](../office-agents.md).

## Connect an Excel workbook

From the repository, start the source client with both MCP and the add-in enabled:

```sh
npm start -- --mcp.enabled --mcp.addin.enabled
```

Open the workbook using **Edit in Teams**. In Excel, choose **Home > Add-ins >
More Settings > Upload My Add-in** and upload the file printed at startup as
`[MCP] Upload this Excel add-in manifest: …/mcp/excel-addin.xml`. The default
source-run location is `~/.config/Electron/mcp/excel-addin.xml`.

The **Teams Office Bridge** task pane should say **Ready for local agent access**.
Keep the pane open. Its **Disconnect this workbook** button revokes that pane's
access; reopening the add-in creates a new document connection. A document that
changes its URL must also reconnect. Sideload support has to be checked in the
target tenant; a local fixture test does not prove a real Office session works.

In another terminal, run `npm run mcp:diagnose`. The `documents` list contains
the live connection IDs and workbook names. The separate `editors` diagnostics
may still report no Office.js globals in the Microsoft editor itself: the API
context belongs to the task pane.

The client serves only the add-in's static files over HTTPS on
`127.0.0.1:39847`. OpenSSL must be on PATH the first time it generates its local
certificate and key. The client pins that exact certificate for that exact
loopback origin; it makes no changes to system or browser trust stores. The
add-in is intended to run inside this client. A separate browser does not
inherit this trust. No document commands or content are served over HTTPS.
The add-in loads the Office.js SDK from Microsoft's CDN. Its script policy
permits Microsoft's Office/Ajax CDNs and the evaluation used by the legacy
Microsoft Ajax runtime for type resolution.

For persistent configuration, merge this into the client's config.json:

```json
{
  "mcp": {
    "enabled": true,
    "addin": { "enabled": true, "port": 39847 }
  }
}
```

Changing the port requires uploading the regenerated manifest. Certificate and
key files live under the private `mcp/addin` directory in the app's user data.
Do not copy the private key or add that directory to source control.

## Live Excel tools

| Tool | Purpose |
| --- | --- |
| `office_list_documents` | Discover initialized task panes and their temporary document IDs |
| `excel_list_worksheets` | Get stable worksheet IDs and names for a chosen document |
| `excel_read_range` | Read values, formulas and displayed text; obtain a read token |
| `excel_write_range` | Write a matching matrix of literal values using a read token |
| `excel_set_formulas` | Write a matching matrix of formulas starting with `=` using a read token |

Always list documents first and target an explicit document ID and worksheet ID.
Ranges use A1 rectangles such as `A1` or `B2:D8`, limited to 1,000 cells and a
32 KB result. Writes use `data` as a two-dimensional matrix, limited to 32 KB;
empty strings clear cells. Literal text starting with `=`, `+` or `-` is rejected
to avoid interpreting it as a formula. Use numeric values for negative numbers.

Every read returns a single-use `readToken` valid for two minutes. Writes must
include that token, the same worksheet ID and the same range address. The add-in
reads the cells again and rejects a write if either their values or formulas
changed. Excel does not provide an atomic compare-and-set, so a coauthor can
still change cells between the comparison and write. Serialize operations and
avoid working on the same cells at the same time.

After an uncertain or timed-out write, read the range again rather than retrying
the old request. A successful write returns its resulting cells and a new read
token. The bridge does not download and replace the workbook file; Excel handles
calculation, saving and coauthoring.

For source testing without an agent, call a tool from the terminal:

```sh
npm run mcp:call -- office_list_documents '{}'
npm run mcp:call -- excel_list_worksheets '{"documentId":"DOCUMENT_ID"}'
npm run mcp:call -- excel_read_range '{"documentId":"DOCUMENT_ID","worksheetId":"WORKSHEET_ID","address":"A1"}'
```

Replace the IDs with those returned by discovery. To test a write in a test
workbook, use the token returned by the read:

```sh
npm run mcp:call -- excel_write_range '{"documentId":"DOCUMENT_ID","worksheetId":"WORKSHEET_ID","address":"A1","readToken":"READ_TOKEN","data":[[7]]}'
```

`mcp:call` also accepts a socket path after the JSON arguments. These commands
use the same MCP tools an agent uses; the write command changes the document.

## Connect a Word document

Start Teams with the same `--mcp.enabled --mcp.addin.enabled` flags and open a
Word document using **Edit in Teams**. Upload the separate Word manifest printed
at startup: `…/mcp/word-addin.xml` (normally
`~/.config/Electron/mcp/word-addin.xml` when running from source). The existing
Excel manifest still targets Excel. Each application has its own manifest and
uses the same HTTPS task pane and local MCP socket.

Word requires WordApi 1.6 for its session-local paragraph IDs. The pane checks
this requirement before connecting and explains when the host is unsupported.
Keep the pane open and call `office_list_documents`; the entry must report
`application: "word"`. Reconnect your agent after restarting Teams to discover
the additional tools.

| Tool | Purpose |
| --- | --- |
| `word_read_paragraphs` | Read a page of main-body paragraphs, their IDs and read tokens |
| `word_read_paragraph` | Refresh one paragraph by its ID |
| `word_replace_text` | Replace one exact text match within an unchanged paragraph |

```sh
npm run mcp:call -- word_read_paragraphs '{"documentId":"DOCUMENT_ID","start":0,"count":10}'
npm run mcp:call -- word_read_paragraph '{"documentId":"DOCUMENT_ID","paragraphId":"PARAGRAPH_ID"}'
npm run mcp:call -- word_replace_text '{"documentId":"DOCUMENT_ID","paragraphId":"PARAGRAPH_ID","readToken":"READ_TOKEN","findText":"Hello","replacementText":"Greetings"}'
```

Discovery pages contain up to 50 paragraphs (20 by default), bounded to 32 KB,
and return `nextStart` or `null` at the end. Paragraph insertion can shift page
offsets during coauthoring. Writes always target a paragraph ID, which remains
valid only within that add-in session. Headers, footers, text boxes and other
document stories are outside the discovery scope. Large paragraphs may exceed
the result limit even when read individually.

Each paragraph read returns a two-minute, single-use token. Replacement checks
the paragraph's text again after searching and rejects changes or multiple
matches. Searches are literal and case-sensitive, limited to 255 characters;
caret codes and control characters are rejected. Replacement is plain text,
limited to 8,192 characters with no paragraph breaks; an empty string deletes
the matched text. Only the matched range is replaced, leaving surrounding
content intact. Formatting within the replacement follows Word's behavior;
this tool does not provide rich-text editing or paragraph creation. The token
compares text, not formatting. Word's own save and coauthoring workflow remains
responsible for the document.

As with Excel, the comparison and write are not atomic against other editors.
Read back after an uncertain write; do not retry the previous request. Test
against a disposable document before using these tools on real work.

The API choices follow Microsoft's reference for
[paragraph IDs and search](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph?view=word-js-preview)
and [paged collection loads](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption?view=word-js-preview).

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
expected for editor frames; live access appears separately in `documents`. Office.js normally
provides document APIs inside an Office add-in; globals in the surrounding
editor are not a supported substitute for that add-in context. The probe neither
loads Office.js nor calls `Office.onReady`, `Excel.run` or `Word.run`.

## Local connection and implementation

MCP is disabled by default and requires a restart when enabled or disabled. The
app serves MCP through a Unix socket with mode `0600` in a user-owned directory
with mode `0700`. Diagnostics alone open no TCP port. The optional add-in enables
the loopback HTTPS asset host. There are no new renderer IPC methods.
Only local processes with access to that directory can attach. These permissions
do not distinguish different applications running as the same user.

Editor discovery returns origins rather than full document URLs. The diagnostic never
returns document content, sharing tokens, or renderer error messages. It runs
one fixed script in recognized frames, with a timeout; callers cannot supply
JavaScript. Results are observations from page code, not a trusted authorization
decision for future editing tools.

`app/mcp/server.js` defines the tools using the official MCP SDK.
`app/mcp/localEndpoint.js` binds SDK transports to local connections, and
`app/mcp/stdio.js` forwards an agent's stdio connection to the app.
`app/office/inspector.js` owns frame discovery and capability inspection.
`app/office/liveBridge.js` finds our own task-pane frames under Excel/Word ancestors
and sends structured commands through Electron's frame API. The Office.js
operations run in `app/office/addin/operations.js` and `wordOperations.js`. The app does not inject
Office.js into Microsoft's editor or expose arbitrary JavaScript execution as
an MCP tool. Local-file servers are not used to overwrite coauthored documents.

Run the focused checks with:

```sh
node --test tests/unit/excelSession.test.js tests/unit/wordSession.test.js tests/unit/officeAddinHost.test.js tests/unit/officeLiveBridge.test.js tests/unit/officeInspector.test.js tests/unit/officeMcp.test.js
npm run lint
```

To check the Electron frame APIs against nested cross-origin fixture pages, run
`./node_modules/.bin/electron tests/electron/officeMcpSmoke.js`. A headless Linux
test machine can use `xvfb-run -a` before that command. This test serves local
fixture responses for every HTTPS request and does not contact Microsoft or
verify a real signed-in Office editor. Also run
`./node_modules/.bin/electron tests/electron/officeAddinSmoke.js` to check the real
loopback certificate pin and Excel/Word task-pane reads and writes with simulated
Office.js hosts. It also checks that older Word hosts remain disconnected. This
remains separate from testing against Microsoft Office and another coauthor.
