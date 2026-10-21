# Connect an agent to Word and Excel

The Office bridge lets an MCP agent read and edit Word and Excel documents open
inside Teams for Linux. Excel supports discovery, cell reads, values, formulas,
formatting and worksheet structure edits. Word supports paragraph and table
inspection, text and style edits, comments and tracked-change review. Documents stay in
Office's normal saving and coauthoring workflow.

## Enable the bridge

The [Teams Office Flatpak](https://github.com/nosini/teams-office-flatpak)
packages this version with the bridge always enabled; its README lists the
add-in and socket paths it uses. To run it from source instead, install
dependencies in this repository with `npm ci`, then start Teams:

```sh
npm start -- --mcp.enabled --mcp.addin.enabled
```

To enable it on subsequent launches, merge this setting into Teams' existing
`config.json`, then restart:

```json
{ "mcp": { "enabled": true, "addin": { "enabled": true } } }
```

OpenSSL must be installed for the first launch. Startup prints the exact socket
and add-in manifest paths. Source launches normally use `~/.config/Electron/mcp`;
packaged launches normally use `~/.config/teams-for-linux/mcp`. Use the printed
paths when you have a custom profile.

## Open the add-in

Open a document using **Edit in Teams**. In Office's **Add-ins** menu, choose
**Upload My Add-in** and select the manifest printed at startup:

- Excel: `excel-addin.xml`
- Word: `word-addin.xml`

Your organization's Office policy must allow uploading add-ins. Keep the
**Teams Office Bridge** pane open and wait for **Ready for local agent access**.
Word requires WordApi 1.6. Each document needs its own open pane.

The pane runs inside Teams for Linux. A separate browser does not inherit the
client's trust for the local HTTPS certificate and may show an add-in startup
error. Other editors can close that pane and coauthor normally in their browser;
they do not need the bridge installed or working there.

## Connect a local agent

Run this on the desktop where Teams is running, choosing a new private export
directory:

```sh
npm run mcp:setup -- "$HOME/.local/share/teams-office-agent"
```

The command checks the running MCP server and exports a standalone Node.js relay
and `mcp.json`. It does not change agent settings or copy certificates. Merge the
`teams-office` entry into your agent's MCP server configuration. The generated
JSON uses the `mcpServers` format; clients with other settings formats need the
same command and argument values entered in their own format. Reconnect the
agent, then ask it to list connected Office documents.

If multiple Teams profiles have sockets, specify the intended socket explicitly:

```sh
npm run mcp:setup -- "$HOME/.local/share/teams-office-agent" /absolute/profile/mcp/bridge.sock
```

An existing export is never overwritten. Choose a new directory to regenerate
it. The relay only requires Node.js, so it can run independently of the checkout
and its npm dependencies. Teams itself must remain running.

## Connect a container agent

The host launcher must explicitly expose the socket and relay. Run setup on the
host first. Give the launcher these two files:

- The `socketPath` printed by setup, mounted as a Unix socket at a fixed container path.
- The exported `teams-office-stdio.js`, mounted read-only at a fixed container path.

Inside the container, configure the MCP command as `node` with the mounted relay
path and mounted socket path as its two arguments. Node.js must be installed in
the image. Use the launcher's existing user mapping so the container process can
access the private socket; keep the host socket's permissions unchanged.

Expose only the socket file, rather than its parent `mcp` directory: that directory
also contains the add-in TLS private key. A file bind retains the old socket
inode when Teams restarts. After restarting Teams, recreate the agent container
and reconnect the agent. Restarting only the MCP relay will not refresh that bind.

For Zed launchers already configured with separate **Codex + Teams** or
**Claude + Teams** entries, select the Teams-enabled entry. Those launchers inject
the relay configuration themselves; a second agent MCP entry is unnecessary.
Mount changes belong in the host launcher's configuration and installation flow.

## Check the connection

From the Teams repository on the host:

```sh
npm run mcp:diagnose
```

A connected entry in `documents` confirms the add-in is ready. Entries under
`editors` alone only show that an Office frame was found. If `documents` is empty,
open the correct add-in pane in Teams and wait for its ready message. If the
document tools are missing, enable both MCP options and restart Teams and the
agent. For a missing or refused socket, check the path printed by the current
Teams process; container users should also recreate their container after a
Teams restart.

Reads return short-lived, single-use tokens. If another editor changes the same
cells or paragraph, read again before writing. Comparison and writing are not
atomic against coauthors, so coordinate edits to the same content. After a timed
out write, read back before attempting another write.

See the [development guide](development/office-mcp.md) for tool arguments, limits,
and implementation details.

## Inspect and edit live documents

Ask the agent to list connected documents first. The new tools use the same
add-ins and configuration. After updating Teams, reopen the add-in panes and
reconnect the agent to refresh its catalog. Container users should recreate
their container after the Teams restart.

For Excel, the agent can discover used-range bounds and structured tables,
search text or formulas in windows of up to 1,000 cells, inspect number formats
and basic formatting, and apply a uniform format to a selected range. It can
also create worksheets and insert whole rows or columns. Insertion checks the
sheet ID and used-range bounds, not every cell's content or every table;
coordinate structural edits with other editors.

For Word, the agent can inspect paragraph styles and headings, read table rows,
inspect fonts and alignment, replace whole paragraphs, insert or delete a
paragraph, and apply basic styles and formatting. Full paragraph replacement
can remove inline formatting, fields and links. Prefer exact text replacement
for small corrections where surrounding rich content should remain intact.

Word review tools read comments and replies, add comments, reply to threads,
and resolve or reopen them. Separate tools inspect and explicitly set the
shared document's tracking mode and inspect, accept or reject individual
revisions within a selected paragraph. The bridge never changes tracking mode
or accepts/rejects revisions implicitly.

Font options are bold, italic, font name, size and RGB color. Word additionally
supports paragraph styles and alignment; Excel supports number formats, fill
color, alignment and wrapping. Mixed formatting may return null or empty
properties. Only explicitly supplied formatting properties change.

Write tokens expire after two minutes and are consumed on attempted writes.
Formatting edits require a formatting-inspection token; comment and revision
edits require their own review tokens. Structural operations and revision
resolution invalidate earlier document tokens. Coauthor edits can still occur
between comparison and writing. Inspect results after an uncertain write
before trying again.

New Word tools use WordApi 1.6, matching the existing Word requirement. Excel
used-range inspection and row/column insertion require ExcelApi 1.4; other new
Excel tools use ExcelApi 1.1. Unsupported hosts return a clear error. See the
[tool reference](development/office-mcp.md#extended-live-document-tools) for
names, arguments and limits. The new tools have automated and simulated
Electron coverage; real Office compatibility still needs live testing.

## Experimental live Markdown editing

The native SharePoint Markdown editor uses CodeMirror. A separate opt-in enables
MCP reads and edits of the running CodeMirror 6 model, without an Office add-in:

```json
{ "mcp": { "enabled": true, "codemirror": { "enabled": true } } }
```

Or launch from this checkout with:

```sh
npm start -- --mcp.enabled --mcp.addin.enabled --mcp.files.enabled --mcp.codemirror.enabled
```

Restart Teams, reopen the Markdown file, and reconnect the agent. Container
users need a recreated container to refresh the socket mount. Ask the agent to
call `codemirror_list_editors`, identify the intended editor, and read it with
`codemirror_read_text`. Unsupported editor instances are reported as unavailable.
IDs identify running editors, not persistent file names.

`codemirror_replace_text` edits a selected range through the existing editor's
transaction method. It requires a recent read token and rejects changes to the
entire document revision since that read. It also rejects read-only editors and
active IME composition. It never replaces the file through Graph or rewrites
visible DOM text. Reads include unsaved content from the editor model.

Offsets use UTF-16 code units, with half-open `[start,end)` edit ranges. Text
uses LF newlines; insertion/replacement is limited to 8,192 characters and must
fit the window that was read. Equal offsets insert; empty text deletes. Reads
return `nextStart` for larger files and avoid splitting Unicode surrogate pairs.
Tokens expire after two minutes and are single-use. Any document text change,
including one outside the selected window, requires a fresh read.

This adapter uses version-specific editor handles. Automated tests against real
CodeMirror verify native transaction dispatch and update listeners, not the
SharePoint collaboration service. Successful edits return `saveStatus:
"unverified"`: local model readback is confirmed, but cloud saving and other
users' views must be checked in a live disposable file. Never automatically
retry a timed-out, filtered or otherwise uncertain edit; read current state
first. No new synchronization service or file-upload fallback is added.

## Read and write channel files

Channel file access is a separate opt-in. With this version running, merge into
`config.json` and restart Teams:

```json
{ "mcp": { "enabled": true, "files": { "enabled": true } } }
```

From the source checkout you can instead launch with:

```sh
npm start -- --mcp.enabled --mcp.files.enabled
```

Keep `addin.enabled` enabled too if you use live Word and Excel editing. File
tools do not require an add-in or `graphApi.enabled`. Reconnect your agent after
restarting Teams; container agents need their socket bind recreated as described
above. An installed Flatpak must first be rebuilt with this version of the code.

Ask the agent to list teams, list a team's channels, resolve a channel's Files
folder, and list its files. The tools are `teams_list_teams`,
`teams_list_channels`, `teams_get_channel_folder`, `teams_list_files`,
`teams_read_file`, `teams_create_file` and `teams_replace_file`. Listings return
`nextCursor`; pass it back as `cursor` with the same team/folder to continue.
If team/channel discovery is denied, the agent can still try resolving the file
folder with team and channel IDs from a Teams channel link you provide.

Single-call transfers use base64 bytes and are limited to **1 MiB per file**.
Larger files use the chunked tools below, without a bridge-level total file size
limit; Microsoft storage limits and the session's permissions still apply. These tools can
read binary formats, but interpreting or modifying a DOCX/XLSX package requires
format-aware tooling. Use the existing Word/Excel tools for edits in open Office
documents. Replacement uploads replace the complete file, rather than merging
Office changes. A replacement requires `expectedETag` from a recent file read;
the server checks it again and sends `If-Match`. This safeguard and creation's
conflict-failure behavior still require live validation against your tenant;
coordinate edits while testing. Do not automatically repeat a timed-out upload:
list/read the destination first to determine whether it succeeded.

### Chunked downloads and uploads

| Tool | Purpose |
| --- | --- |
| `teams_begin_download` | Bind a download to a selected `driveId`, `itemId`, size and eTag. |
| `teams_read_file_chunk` | Read the next byte range as base64, with a chunk SHA-256. |
| `teams_begin_upload` | Create an upload session for a fixed positive total byte size. |
| `teams_upload_chunk` | Send the next sequential base64 fragment. |
| `teams_get_transfer` | Inspect progress and reconcile uncertain chunk requests. |
| `teams_finish_upload` | Explicitly publish the staged file. |
| `teams_cancel_transfer` | Stop a download or cancel unfinished upload staging. |

For a download, begin with `driveId` and `itemId`, then pass its `transferId` and
`nextOffset` to `teams_read_file_chunk` until `done` is true. Decode each chunk
and concatenate the bytes in offset order. Each chunk contains up to **640 KiB**
of raw bytes, keeping MCP messages within the existing transport limit. The
server requests exact HTTP byte ranges and checks the original eTag and size
before and after each range. If the file changes, discard all assembled bytes
and begin again. The most recent chunk can be reread after a lost response.

For an upload, begin with `driveId`, total `size` in bytes, and either
`parentId` + `name` to create a new file, or `itemId` + `expectedETag` to replace
one. Creation requests conflict failure. Optionally provide `expectedSha256`
(the SHA-256 of the complete source bytes). Send chunks at `nextOffset`, at most
640 KiB each; every nonfinal chunk must be a multiple of 320 KiB. The server
checks the source hash before sending the final fragment, then reports it as
`sourceSha256`. Empty files use the single-call tools.

The server requests Microsoft's deferred commit mode: the final chunk normally
returns `ready`, then `teams_finish_upload` publishes the file. If session
creation is rejected with a definite HTTP 400, the server tries once without the
optional deferred-commit property. This fallback returns `commitMode: automatic`:
**the final chunk publishes the file**. The normal mode is `commitMode: deferred`.
Conflict failure and eTag preconditions remain in both session requests. The
request serializer puts OData annotations before ordinary properties, as required
by affected SharePoint parsers ([Microsoft SDK issue](https://github.com/microsoftgraph/msgraph-sdk-java/issues/1351)). There
is no fallback after a timeout, conflict, or fragment/commit failure. If Microsoft
instead reports that a fragment already published the file, the transfer returns
`completed` and finish will not publish it again. Replacement rechecks the
original eTag before each fragment and immediately before finish, and sends
`If-Match` when creating and finishing the session. This is not an atomic lock
against coauthors. Deferred commit was verified in one SharePoint tenant with
a disposable binary file; validate your own tenant and file types before relying
on the same behavior.
SharePoint can rewrite Office package metadata, so `sourceSha256` verifies the
uploaded source rather than promising an identical hash for the saved package.
Use the live editing tools for collaboration in open documents.

After a chunk request fails, call `teams_get_transfer` before continuing. It
queries Microsoft's progress and reconciles whether that chunk was accepted.
An identical repeat of the last acknowledged chunk returns its receipt without
sending it again. An uncertain finish is never automatically retried. If a
session disappears, inspect the destination: disappearance can mean expiration
or successful publication. Cancelling cannot delete or revert a saved file.

Transfer IDs survive agent reconnects while Teams remains running, expire after
30 idle minutes (or Microsoft's earlier expiry), and are lost on Teams restart.
Up to eight transfers can be active at once. Completed receipts remain available
for ten minutes. Only one operation can run on a given transfer at a time.
The server retains at most one pending chunk per upload, with no local temporary
file or whole-file buffer. Provider staging expires separately if a local
transfer is abandoned. Temporary upload/download URLs remain inside Teams and
requests to those URLs never carry the Graph bearer token. See Microsoft's
[upload session protocol](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0)
and [ranged download instructions](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0).

The server reuses the signed-in Teams session. It does not request new Microsoft
permissions and does not expose its access token or temporary download links to
the agent. HTTP 403 means the session lacks permission or organization policy
blocks the operation; enabling MCP does not grant additional SharePoint rights.
Discovery, reading and writing can have different permissions. Microsoft's
[folder lookup](https://learn.microsoft.com/en-us/graph/api/channel-get-filesfolder?view=graph-rest-1.0)
and [upload documentation](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0)
list the required permissions.

A live SharePoint test on 2026-10-21 verified a 1,310,737-byte binary upload in
three chunks, explicit deferred commit, byte-for-byte download and SHA-256,
replacement using a fresh eTag followed by another verified download, stale
eTag rejection, and cancellation after staging the first chunk with no resulting
file. A separate 7,995,415-byte download passed in thirteen chunks. These tests
do not establish an atomic coauthor lock, live network-failure recovery,
create-name conflict handling, or Office-package byte preservation.

For a first live test, choose a test channel, list its files, then ask the agent
to create a uniquely named `mcp-file-test.txt`, read it back, replace it using the
returned eTag, and read it back again. Use a disposable file for testing conflict
handling with another editor. File deletion is not exposed by these tools.
