# Office MCP bridge

The Office MCP bridge connects local agents to documents inside the running
Teams client. Diagnostics discover Word and Excel frames. The optional Excel
task pane provides live worksheet reads, discovery, formatting, value/formula
writes and structural edits through Office.js. The Word task pane provides
paragraph and table inspection, text and style edits, comments and tracked-change
review using the same local bridge.

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
Do not copy the private key or add that directory to source control. The client
replaces a pair that is missing, mismatched, readable by others or within a day
of expiry. If the add-in host cannot start, the live Office tools are absent
and the other MCP tools still start.

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
this literal replacement tool does not create paragraphs; separate tools below
provide whole-paragraph edits, formatting and paragraph insertion. The token
compares text, not formatting. Word's own save and coauthoring workflow remains
responsible for the document.

As with Excel, the comparison and write are not atomic against other editors.
Read back after an uncertain write; do not retry the previous request. Test
against a disposable document before using these tools on real work.

The API choices follow Microsoft's reference for
[paragraph IDs and search](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph?view=word-js-preview)
and [paged collection loads](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption?view=word-js-preview).

## Extended live-document tools

Registered with the live add-in bridge alongside the original tools. All take
`documentId` from `office_list_documents`. Pages use zero-based `start` and
`count` (default 20, maximum 50), with `nextStart` to continue. Results and
arguments are capped at 32 KiB. A single unusually large paragraph, cell or
table row can exceed this limit.

### Excel

| Tool | Additional arguments | Result or effect |
| --- | --- | --- |
| `excel_inspect_worksheet` | `worksheetId`, optional page | Used bounds, paged structured tables and layout token |
| `excel_search_range` | `worksheetId`, `address`, `query`, optional `searchIn` (`text`/`formulas`), `matchCase`, `offset` | Up to 50 literal matches; continue with `nextOffset` in the same window |
| `excel_read_range_format` | `worksheetId`, `address` | Values, formulas, text, number formats, basic formatting and format token |
| `excel_format_range` | Same range, `readToken`, `format` | Uniform formatting; checks content and inspected formatting |
| `excel_create_worksheet` | `name`, `readToken` from `excel_list_worksheets` | New worksheet ID/name/position; checks the sheet list |
| `excel_insert_rows_columns` | `worksheetId`, `axis` (`rows`/`columns`), `index`, `count`, layout `readToken` | Insert whole rows/columns before a zero-based index |

Search and formatting use A1 rectangles of at most 1,000 cells. Search scans a
bounded window, treating wildcard characters literally. Use used-range metadata
to plan windows across large sheets. Search results are not write tokens.

Excel `format` accepts `numberFormat` (one invariant format string applied to all
cells), `font`, `fillColor`, `horizontalAlignment`, `verticalAlignment` and
`wrapText`. Font accepts `bold`, `italic`, `name`, `size` (1–200 points) and
`color`; colors use `#RRGGBB`. Unspecified properties remain unchanged. Mixed
properties may return null/empty values. Number formats are a per-cell matrix.

Worksheet creation needs a non-truncated worksheet-list token (maximum 200
sheets) and a unique valid Excel name of at most 31 characters. Insertion allows
up to 100 rows or 20 columns per call. Layout tokens compare worksheet ID and
used-range bounds, not all cell content or table structure. Excel shifts data
and references. Coordinate structural edits with coauthors. These operations
invalidate all earlier pane tokens. Used-range inspection/insertion require
ExcelApi 1.4; other new Excel tools require ExcelApi 1.1.

### Word

| Tool | Additional arguments | Result or effect |
| --- | --- | --- |
| `word_read_structure` | Optional page | Paragraph IDs/text, styles/alignment and ordinary text tokens |
| `word_read_paragraph_details` | `paragraphId` | Text, style, alignment, font and format token |
| `word_list_tables` | Optional page | Main-body table indexes and row counts |
| `word_read_table` | `tableIndex`, optional `rowStart`, `rowCount` | Paged row/cell text, cell counts and `nextRowStart` |
| `word_replace_paragraph` | `paragraphId`, ordinary `readToken`, `text` | Replace complete paragraph content, excluding its mark |
| `word_insert_paragraph` | Anchor `paragraphId`, ordinary `readToken`, `text`, `position` (`Before`/`After`), optional `styleBuiltIn` | New paragraph ID/text/token |
| `word_delete_paragraph` | `paragraphId`, ordinary `readToken` | Delete paragraph and content; tracking may preserve a revision |
| `word_format_paragraph` | `paragraphId`, format `readToken`, `format` | Style, alignment and/or font edits |
| `word_read_comments` | `paragraphId`, optional page | Comment threads/replies and token per thread |
| `word_add_comment` | `paragraphId`, ordinary `readToken`, `text` | Add Office comment anchored to paragraph content |
| `word_reply_comment` | `paragraphId`, `commentId`, thread `readToken`, `text` | Reply to thread |
| `word_resolve_comment` | `paragraphId`, `commentId`, thread `readToken`, `resolved` | Resolve or reopen thread |
| `word_get_change_tracking` | None | Document-wide mode and mode token |
| `word_set_change_tracking` | Mode `readToken`, `mode` | Explicitly set `Off`, `TrackAll` or `TrackMineOnly` |
| `word_read_tracked_changes` | `paragraphId` | Ordered revisions, Original/Current reviewed text and revision token |
| `word_resolve_tracked_change` | `paragraphId`, revision `readToken`, `changeIndex`, `action` (`accept`/`reject`) | Resolve one inspected revision |

All new Word tools require WordApi 1.6, already required for stable paragraph
IDs. Discovery remains in the main body; headers, footers and text boxes are
outside scope. Table indexes are positional, read-only references; relist after
structural changes. Merged cells may produce unequal cell counts. Table
inspection does not expose table editing.

Whole-paragraph text is limited to 8,192 characters without paragraph/control
characters. It bypasses the 255-character search limit but can replace inline
formatting, links, fields and other rich content. Prefer `word_replace_text` for
small corrections. Insertion/deletion/whole replacement invalidate earlier pane
tokens; insertion/replacement return a fresh paragraph token. Word handles
saving and tracks edits according to the document's current tracking mode.

Word `format` accepts `style` (localized/custom name) OR `styleBuiltIn`,
`alignment` (`Left`, `Centered`, `Right`, `Justified`) and `font` with the same
fields as Excel. Supported built-in styles: Normal, Title, Subtitle,
Heading1–Heading9, Quote, Caption and ListParagraph. Use built-in styles across
Office languages. Formatting targets the complete paragraph; mixed inspection
properties may be null/empty.

Comments/replies accept single-line plain text up to 4,096 characters. Tokens
compare paragraph text and the selected thread's metadata/replies. Each thread
returns up to 20 replies, with `repliesTruncated`; truncated threads and
paragraphs with more than 50 comments cannot be modified through reply/resolve.
Use `comment.id` as `commentId`. No @mention markup is generated. Office controls
normal collaboration notifications associated with comments.

Tracking mode is a shared document setting, not an agent-local preference. No
mode change or accept/reject operation happens implicitly. Revision reads allow
up to 50 changes per paragraph; larger sets fail explicitly. Tokens compare the
whole ordered list (author, date, text, type), paragraph text and Original/Current
reviewed text. Resolution invalidates earlier pane tokens. These checks reduce
accidental retargeting but cannot lock out coauthors atomically.

All edit tokens expire after two minutes and are consumed on attempted writes.
Comparison and mutation are separate Office batches; no operation claims atomic
protection against another editor. After an uncertain write, inspect current
state before deciding whether another write is needed.

Reference APIs:
[Excel ranges](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.9),
[Word paragraphs](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph?view=word-js-preview),
[Word tables](https://learn.microsoft.com/en-us/javascript/api/word/word.table?view=word-js-preview),
[comments](https://learn.microsoft.com/en-us/javascript/api/word/word.comment?view=word-js-preview),
[tracked changes](https://learn.microsoft.com/en-us/javascript/api/word/word.trackedchange?view=word-js-preview).

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

## Non-Office editor diagnostics

`editor_list_frames` lists HTTPS frame IDs and origins, including file editors
that the Word/Excel diagnostics do not recognize. It does not return document
paths, names or full URLs. Only Microsoft-hosted origins are marked
`inspectable`; third-party frames are listed by origin but are not probed.

`editor_inspect_frame` takes a listed `frameId` and runs a fixed read-only probe
for editor signatures: Monaco, CodeMirror, ProseMirror, Lexical, Slate, Quill,
editable-element counts and limited adapter-handle presence. It returns only
validated counts, booleans and allowlisted tag names. It does not read text,
call editor/document/save methods, inspect storage or expose general JavaScript
execution. Results are discarded if the frame navigates or detaches during the
probe. Presence does not establish a working live-editing or coauthor bridge.

These diagnostics help investigate the native SharePoint/OneDrive Markdown
editor. The separate experimental adapter below enables bounded model editing. Restart Teams and reconnect the
agent after updating the code, then open the test Markdown editor before
listing/probing its frames.

## Experimental CodeMirror adapter

Enable `mcp.codemirror.enabled` with `mcp.enabled` and restart. This adds three
tools independently of the Word/Excel add-ins:

| Tool | Arguments | Behavior |
| --- | --- | --- |
| `codemirror_list_editors` | None | Discover supported running CodeMirror 6 models in SharePoint HTTPS frames |
| `codemirror_read_text` | `editorId`, optional `start` (0), `count` (4096, max 8192) | Bounded live model window and two-minute single-use read token |
| `codemirror_replace_text` | `editorId`, `readToken`, `start`, `end`, `text` | Native transaction replacing half-open UTF-16 range; equal offsets insert, empty text deletes |

Use a listed editor ID, not a positional DOM index. Discovery returns editor
length and read-only status, never text or document URLs. IDs are session-local,
not file IDs or names. Identify the intended file/editor with the user. Discovery
bounds frame fanout at 16 and returned editors at 32, reporting truncation and
unavailable frames/instances. Unsupported editors have no file-upload fallback.

The renderer adapter in `app/editor/codeMirrorCommand.js` supports known `cmView`
and `cmTile` DOM-link generations, following bounded own-data parent links to
find an existing view. It verifies the root/content DOM identities and a usable
CodeMirror document/configuration. It does not import its own editor library,
reconfigure the host, rewrite DOM text or accept executable caller code.
`app/editor/codeMirrorBridge.js` limits access to HTTPS SharePoint frames and
retains full frame URLs only internally for navigation checks. Registry and
snapshot state stays in the target renderer under a per-bridge random key.

Reads use the model's `sliceString` rather than virtualized visible DOM text.
Windows and edits use UTF-16 code units, LF line endings and a maximum 8,192
units per operation, capped at 32 KiB in JSON. A read window's end is shortened
when necessary to avoid splitting a surrogate pair; a starting offset inside a
pair is rejected. Edits cannot split pairs, insert unpaired surrogates, CR/NUL or
Unicode paragraph separators. The edited range must lie entirely inside the
read window. Larger documents are read using `nextStart`.

Snapshot tokens bind editor/root/view identity, URL, the immutable document
reference, editor configuration and read-window text. Selection-only changes
are allowed; any document revision or configuration change invalidates the
snapshot. Snapshots expire after two minutes; at most 64 are retained. A valid
write attempt consumes its token before dispatch. Read-only state and IME
composition block writes. Comparison and dispatch run synchronously in one
renderer invocation, with no awaited work between them. This does not imply a
server-side lock against remote coauthors or a persisted cloud revision.

Edits call the existing view's `dispatch` with a bounded changes specification
and `userEvent: "input"`. Host transaction filters, dispatch overrides and update
listeners remain responsible for processing edits. Model readback must confirm
the inserted text and resulting length; filtered/deferred/throwing operations
report an uncertain outcome, with no retry or force-write fallback. Successful
results confirm local application only and always report `saveStatus:
"unverified"`. Cloud autosave and coauthor behavior require live validation.

Focused checks: `node --test tests/unit/codeMirrorBridge.test.js` and
`npm run lint`. The optional `tests/integration/codeMirrorModel.js` runs against
actual `@codemirror/view`, `@codemirror/state` and `jsdom` installed in a separate
test environment; these are not production dependencies. It checks native
updates/listeners, stale revisions, selection-only changes, read-only state and
transaction filtering. It does not test SharePoint's persistence/transport.

A live test on 2026-10-21 verified model reads and native heading replacement
in the SharePoint Markdown editor, followed by an independent Graph download
confirming autosaved text and a changed eTag. Reusing the consumed token was
rejected. A subsequent edit in a separate browser reached the already-open
Teams model without reloading, including text newer than the saved snapshot.
This verifies incoming browser-to-Teams live updates. Unused-token rejection
after a coauthor edit, simultaneous outbound propagation and other native editor
variants still need dedicated tests.
The per-operation `saveStatus` remains unverified until the caller independently
checks persistence.

API reference:
[CodeMirror view source](https://github.com/codemirror/view/blob/main/src/editorview.ts),
[older DOM links](https://github.com/codemirror/view/blob/6.38.6/src/contentview.ts),
[immutable state](https://github.com/codemirror/state/blob/main/src/state.ts).

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
operations run in `app/office/addin/operations.js`, `wordOperations.js` and
`editingOperations.js`; `app/mcp/editingTools.js` registers the extended tools. The app does not inject
Office.js into Microsoft's editor or expose arbitrary JavaScript execution as
an MCP tool. Local-file servers are not used to overwrite coauthored documents.

Run the focused checks with:

```sh
node --test tests/unit/excelSession.test.js tests/unit/wordSession.test.js tests/unit/officeEditing.test.js tests/unit/officeAddinHost.test.js tests/unit/officeLiveBridge.test.js tests/unit/officeInspector.test.js tests/unit/officeMcp.test.js
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
