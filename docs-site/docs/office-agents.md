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

OpenSSL must be installed for the first launch. Startup prints the socket and
add-in manifest paths, with your home directory shown as `~` so the log does
not record your account name. Source launches normally use `~/.config/Electron/mcp`;
packaged launches normally use `~/.config/teams-for-linux/mcp`. Use the printed
paths when you have a custom profile.

## Open the add-in

In a Documents tab, choose **Connect Bridge** to import the matching add-in
through Office's menus. Agents can do the same with `office_setup_bridge` while
the window stays hidden. This uses the app's own manifest; it does not need
admin deployment when your organization already allows manual add-in uploads.
The automation recognizes English and Dutch add-in controls. First-time imports
and hidden document operations have been verified in real Excel and Word.
If setup stops, finish sign-in or use the manual steps below.

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

## Create a document in a SharePoint folder

An agent can create a blank Word document without asking for a file link. Open
the destination folder in Teams, or open a Word/Excel document already in that
folder, then ask the agent to create the new document there.

The agent calls `office_list_locations` to discover the open SharePoint folder
or the parent of the open document. It then calls `office_create_document` with
a name and the selected `locationId`. The app adds `.docx` if needed, creates
the file, opens it hidden in the originating account and starts Bridge setup.
The agent checks `office_list_tabs`, lists the connected document and reads
before editing. No pasted link or agent-generated Office package is needed.

The destination can also be another existing folder without navigating Teams.
The agent uses `office_list_location_folders` to browse child folders, whose
returned `locationId` values work directly for creation. For a known path, it
calls `office_resolve_location` with a `sourceLocationId` to select the signed-in
account and a decoded `folderPath`, such as `/sites/team/Documents/Reports`.
The source location's `webUrl` is used by default. For another site or subsite
on the same SharePoint host, the agent supplies that site's owning `webUrl`.
The app verifies the folder before returning a destination `locationId`.
Folders on another SharePoint host need a source location on that host.
These tools read metadata and do not create missing folders.
Resolving the parent library, browsing its children, creating in a sibling
folder, importing the Bridge and verifying an edit have also passed in a live
session with the document window hidden throughout.

If several destinations are available, the agent must select the one matching
your request. It can omit `locationId` only when exactly one folder is available.
Locations expire after five minutes or when the source page navigates; list
them again if needed. Unsupported SharePoint views include bounded page-type
diagnostics without URL query values. The app never overwrites an existing filename. After an
uncertain creation, the agent uses `office_list_location_files` with the exact
filename before considering another attempt. A result with `created: true`
means the file exists even when opening or Bridge setup needs attention.

Creation uses SharePoint's normal permissions and the source view's session
inside Electron. Cookies and request digests stay inside the app. It does not
request a Teams token or require the separate Graph channel-file option.
This currently creates Word `.docx` files; Excel workbook creation is not
included. Sign-in and Office add-in policy can still require user interaction.
A fresh Word creation, automatic first-time Bridge import, edit and readback
have passed after restarting the app with a retained session, with the document
window hidden throughout and without a sign-in prompt or setup retry.

## Keep multiple documents open

With MCP enabled, choose **Documents** from the Teams for Linux menu. This
opens a separate window with persistent document tabs. Paste an Excel or Word
SharePoint link and choose **Open tab**, or open a file in Teams first and
choose **Open current file** in Documents. Each tab is an independent Office
page, so opening another file in the main Teams window no longer unloads it.
Up to 16 tabs can stay open. Tabs are kept for this app run; they are not
restored after quitting Teams.

Choose **Connect Bridge**, or open the **Teams Office Bridge** add-in manually
in each tab, and wait for the window to say **Bridge connected**. A page being
loaded does not establish agent access. Opening a URL alone does not install
the add-in. The setup operation imports it through Office's UI without showing
the tab. Tenant restrictions still apply, and sign-in may need your interaction.

Document tabs share the originating Teams account's browser session. Setup
waits up to 15 seconds on Microsoft's sign-in page for an automatic redirect
before reporting `SIGN_IN_REQUIRED`; it keeps the window hidden during this
wait. A temporary visit to a sign-in URL does not itself mean you logged out.
If Microsoft displays its account picker, setup can select the document account
previously chosen for this site and browser profile while keeping the window
hidden. Choose your account once manually, or have the agent supply an explicit
`accountHint` email to `office_setup_bridge`. Only the selected username is
remembered locally; passwords, tokens and cookie expiry are not changed. No
identity request is made through the main Teams session. Setup stops if no
account is remembered, no tile matches, or several tiles match. Account
selection runs at most once per setup attempt. Passwords, MFA and consent
prompts still require you. Microsoft controls whether the resulting login is
accepted and how long it remains valid.
If Microsoft offers **Stay signed in**, choosing it allows Microsoft to issue
persistent cookies that survive an app restart, subject to your organization's
session policies. Using the same launch/profile each time also matters;
source and packaged launches can have different browser storage.

For a repeated prompt, `office_inspect_bridge_setup` reports whether that tab's
Electron session is persistent and whether the known Microsoft/SharePoint
sign-in cookies are missing, session-only or persistent. It works before Word
or Excel loads, and does not return cookie values or account identifiers.
It also reports whether a recognized account picker or credential fields are
present, without returning field values or account names.
These diagnostics do not establish that Microsoft will still accept a cookie.

If SharePoint's sign-in cookies disappear on every restart and Microsoft offers
no **Stay signed in** option, the app has a separate opt-in setting:

```json
{
  "mcp": {
    "enabled": true,
    "addin": { "enabled": true },
    "documentSession": { "enabled": true, "days": 7 }
  }
}
```

After restarting with this setting, open the document and finish sign-in once.
The app gives session-only `FedAuth` and `rtFa` cookies for that exact SharePoint
site, plus the shared `sharepoint.com` `rtFa` root cookie, a local expiry so
Electron can retain them across restarts. This also covers numbered `FedAuth`
fragments. The root cookie keeps its existing shared scope; other site cookies
remain limited to exact opened hosts. It retains the sign-in cookies on disk for
up to seven days by default; choose 1–30 days. Passwords and cookie values are
never copied to app settings, diagnostics or repository files. Microsoft login
cookies, unrelated site cookies and cookies with an existing expiry are left unchanged.
Explicit sign-out still removes the cookies. Microsoft's server-side expiry,
revocation and MFA requirements still apply, so this cannot guarantee that
every subsequent open will be silent. Turning the setting off stops further
promotions; sign out of the site to remove cookies already retained.

Use **Hide window**, or close the Documents window, to keep working in Teams
while the agent edits. Both actions keep all tabs running. To unload a document,
close its individual tab. Closing a tab does not itself confirm Office has
saved changes; finish saving first. Quitting Teams ends every tab and bridge.
Background tabs still use memory and CPU, and need a running Teams app and
network connection. Your ordinary browser's tabs are separate from these.

Agents can use these additional tools:

| Tool | Purpose |
| --- | --- |
| `office_open_document` | Open a document in a background tab by default; leaves your main Teams page and selected tab alone. |
| `office_list_tabs` | Check loading state and discover connected Office document IDs for each tab. |
| `office_setup_bridge` | Start Bridge import in a hidden tab; poll `office_list_tabs` for setup progress and the connected document ID. |
| `office_inspect_bridge_setup` | Read bounded Office control names and frame origins to diagnose an import that stopped. Leaves the tab hidden and unchanged. |
| `office_bridge_setup_step` | Perform one named add-in setup action after inspection, without showing the tab. |
| `office_show_tab` | Show a tab for inspection or sign-in/add-in setup when requested. |
| `office_hide_tabs` | Hide the Documents window without closing any documents. |
| `office_close_tab` | Request closing a specific tab, allowing Office to prevent unloading. |

`office_open_document` accepts either a SharePoint file URL or a current Excel,
Word or document-wrapper `editorId` from `office_list_editors`. Embedded editors
resolve through their own SharePoint or Microsoft 365 container, including
layouts without any SharePoint wrapper frame. If an editor cannot be reopened,
`unavailableSources` reports its ancestor origins without document URLs or query
strings. A pasted URL uses the main Teams
account's session. An editor ID uses that editor's account. It returns a tab ID;
use `office_list_tabs` to obtain its connected document ID before calling Excel
or Word tools. An agent can run operations on different documents concurrently;
each individual document accepts only one operation at a time.

`office_setup_bridge` takes a `tabId`. It returns immediately while the import
runs for up to 45 seconds. The tab's `bridgeSetup` reports `running`, `connected`,
or `needs_attention`, with a stage, measured `elapsedMs`, and an error code when
setup stops.
Elapsed time starts when setup is requested; document loading before that is
separate. Setup checks the next control immediately after each action and polls
every 150 ms only while waiting for Office.
An existing live Bridge needs no import. Repeating the call reuses the same job
or result, so it cannot accidentally upload twice. After inspecting a stopped
attempt, `retry: true` starts a new attempt. An upload with an uncertain outcome
is never retried automatically. Only a live connected document ID permits edits.
If import reports `IMPORT_UI_NOT_FOUND`, `office_inspect_bridge_setup` can show
the current Office control names without clicking anything or returning document
bodies, input values or full URLs. Control names are untrusted UI data.
Diagnostics also report the Documents window's visibility, minimization and focus
state, plus button dimensions, clipping and the hit element's tag/role. Blocker
text and identifiers are excluded. This helps investigate `CONTROL_OBSCURED`
without sending another click.
`office_bridge_setup_step` accepts only named setup actions such as `home`,
`ribbonOverflow`, `addins`, `advanced`, `moreSettings`, `uploadEntry`,
`selectManifest`, `submit`
and `trust`. The Advanced flow is Add-ins → Advanced → Upload My Add-in →
select this app's manifest → Upload. Word's compact Home ribbon can place
Add-ins under More options; `ribbonOverflow` opens that menu only within the
named ribbon toolbar. File selection supplies the manifest to
the dialog's file input without opening the native Browse picker.
The default `auto` interaction uses Chromium Enter for recognized controls.
`nativeMouse` selects mouse input explicitly; `native` also uses Enter.
`pointer` dispatches DOM pointer/mouse events
followed by a click; `click` dispatches the DOM click alone. Native input does
not show the window or focus it. Automatic setup uses `auto` too.
It refuses to run alongside automatic setup, and a repeated upload
requires an explicit retry after inspecting the result. `performed: true`
confirms dispatch only; inspect the resulting UI or connected document afterward.

Document pages share their account's signed-in session but do not receive the
Teams preload or the local Documents window's controls. Tabs only accept
recognized SharePoint/Office document links. Unrelated top-level navigation
and popups are blocked; finish any additional authentication in Teams if needed.

Opening a fresh SharePoint Excel workbook, importing its Bridge for the first
time, and a temporary formula write/read/clear have now been verified with the
Documents window hidden throughout. The new tab remained unselected while a
previous workbook stayed connected. The agent selected the app's generated
manifest without a native file picker or user interaction.

That live test used the earlier mouse-first setup and the agent recovered an
Add-ins error using native Enter. Automatic setup now uses Enter directly,
prepares a collapsed Home ribbon, and advances as controls become ready. The
faster automatic path has passed Electron tests; its fresh-file live timing
still needs verification. The original test reused an already signed-in browser
session. A new sign-in may still require user interaction.

In a separate real Excel test, literal writes and formula calculation worked
with the Documents window hidden. Another open view received the formula and
its result, then the cleared test cell. Reopening that configured workbook in
a fresh hidden tab restored its existing Bridge. Authentication and first-time
add-in setup are separate requirements for fully unattended opening.
Copying a Teams Microsoft 365 container alone did not reopen the workbook in
that test, so use the actual file link when needed.

The tab manager has also been tested in Electron with two Excel documents and
one Word document using simulated Office.js hosts, including parallel edits,
showing/hiding the window, and closing one tab without unloading the others.
A real SharePoint Word document has also passed first-time Bridge import,
paragraph reads, insertion, literal replacement and deletion of a temporary test
paragraph. Import and editing ran hidden after the window was shown at the
Microsoft sign-in page. The original paragraphs were unchanged afterward.
Word's overflow-menu path needed a UI correction during that test; the reported
37.8-second successful import run included debugging and is not a clean startup
benchmark. A later fresh Word creation passed fully hidden after an app restart
using a retained sign-in session, including first-time import and a verified
edit. Initial authentication and simultaneous edits to multiple real files
remain separate from that verified path.

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
sending it again. An uncertain finish is never automatically retried. A finish
is uncertain after a network failure, timeout or server error. When Microsoft
answers with a definite client error instead, nothing was saved: a failed eTag
precondition reports `stale`, other rejections such as a name conflict report
`rejected`, and a locked file or throttling keeps the upload `ready` to finish
again later. If a session disappears, inspect the destination: disappearance
can mean expiration or successful publication. Cancelling cannot delete or
revert a saved file.

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
