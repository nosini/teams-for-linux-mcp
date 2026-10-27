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

## Persistent background documents

`OfficeMcp` also owns `DocumentTabs` when a main Teams window is available.
Its separate local shell hosts up to 16 `WebContentsView` document pages. The
shell has its own sandboxed preload; document pages have no preload, use
`sandbox: true` and `contextIsolation: true`, and share the source Teams session.
`backgroundThrottling: false` keeps Office commands active while the window is
hidden or another tab is selected. It does not establish an Office connection:
readiness comes only from `LiveOfficeBridge` discovery in that view.

The `office-tabs` IPC handler validates the exact shell webContents, top frame
and local page URL. Its main-process manager performs the same URL validation
for UI and MCP opens. Renderer exceptions, full document URLs and signed query
strings are not returned in tab metadata. Top-level navigation is restricted
to the original document origin and supported Microsoft Office/login hosts;
popups are denied. Closing the shell hides it; explicit individual closes honor
before-unload protection. App shutdown disposes every renderer and IPC handler.

`tests/electron/documentTabsSmoke.js` exercises the actual startup module, MCP
tools, shell IPC, hidden/inactive Office views and cross-document edits with
simulated Office.js. It does not establish Microsoft add-in reopening behavior
or successful cloud saves. The user guide describes these live-test limits.

### First-time Bridge import

`office_setup_bridge(tabId, retry=false)` runs `bridgeSetup.js` in the background
for up to 45 seconds. `office_list_tabs` exposes its fixed progress stages and
error codes in each tab's `bridgeSetup`. Repeated calls reuse a job/result;
`retry: true` explicitly starts another stopped attempt. Closing a tab or quitting
cancels its job. Main-frame navigation cancels setup once it has started using
the editor. An existing live Bridge completes without changing add-in setup.

The add-in host provides the exact generated Excel and Word manifest strings
to the tab manager. The agent cannot supply a path, manifest, JavaScript or DOM
selector. `bridgeSetupUi.js` operates on named English/Dutch add-in controls in
one Office editor and allowlisted descendant Office/store frames. It opens
the add-in menu, recognizes Advanced/Geavanceerd and opens Upload My Add-in,
then assigns a `File` through `DataTransfer` to
the dialog's file input, then clicks Upload. It never opens a native file picker.
Before submission, the selected file's contents must match this app's manifest.
Only an add-in consent dialog identifying Teams Office Bridge may be accepted.
Sign-in is not automated; unknown or ambiguous controls stop setup.
Named-control inspection and actions traverse up to 32 open shadow roots,
excluding document, grid and editable surfaces through shadow ancestors.
Native focus validation follows shadow-root active elements. Closed components
remain unsupported; reaching the scan limit stops actions instead of acting
on an incomplete set of controls.

The job dispatches Upload at most once. After a timeout or uncertain result,
inspect the live tab before explicitly retrying. Uploading successfully is not
readiness: only `LiveOfficeBridge` discovery in that tab completes setup. No
document content or private Office APIs are used by the import operation;
Office may persist the add-in configuration in the document.

Run the Electron fixture with `DOCUMENT_TABS_SETUP=1` to exercise first-time
import, application-specific manifests, hidden cross-origin dialogs, English
and Dutch controls, stale file replacement, duplicate calls and subsequent
Excel/Word access. Both its import UI and Office.js are simulated. A live test
must use a document with no previous Bridge setup; a new tab of an already
configured file is not evidence of first-time setup.

A fresh real SharePoint Excel workbook loaded in a hidden, nonselected tab.
The agent expanded Home, then automatic setup stopped at Add-ins with
`SETUP_UI_UNAVAILABLE`; a second mouse step also failed. Native Enter opened
the actual Add-ins menu while hidden. After inspecting that menu, the agent
resumed automatic setup, which selected the manifest, submitted once and
reported `connected_after_upload`. Worksheet discovery and a temporary formula
write/read/clear succeeded hidden; concurrent discovery in the previously open
workbook also succeeded. The window stayed hidden and its selected tab never
changed. No user interaction was requested.

This verifies the agent-driven flow with an already signed-in browser session;
cold sign-in remains untested. Automatic setup now uses native Enter directly,
matching the successful real input path. The fixture accepts native Enter;
its separate pointer-only case tests explicit mouse capability rather than
assuming Microsoft requires pointer input.

Setup no longer sleeps for 800 ms after every action. It probes immediately,
pausing for 150 ms only when waiting for a menu, an enabled Upload control or
a live Bridge. Selecting a manifest keeps the job inside the upload dialog
until submission is ready. Delayed controls do not cause duplicate dispatch,
ribbon switching or a second Upload. Each progress/result includes monotonic
`elapsedMs` from the start of that setup job, separate from document loading.

A before/after run of the same three-document Electron import fixture measured
7,352 ms before these changes and 1,839 ms afterward. These are simulated Office
hosts with a 1,200 ms menu delay; they measure the orchestration overhead and do
not predict Microsoft's network or workbook load time. After restarting the
final code, a fresh real Word file's automatic import connected in 1,472 ms,
followed by a successful hidden paragraph read. This measures setup after the
document loaded, excluding its earlier sign-in and initial page loading.

Document creation is also handled inside Electron. `office_list_locations`
recognizes open SharePoint library views and document wrappers/direct file
links, including Teams' `listallitems.aspx`, `filebrowser.aspx` and `DocLibEmbed.aspx` views and
same-origin `teamslogon.aspx` destinations. Explicit folder parameters take
precedence over the embedded library's `listUrl`. A fixed read-only probe reads
the public page-context web URL and known folder-address parameters;
SharePoint REST resolves document IDs and verifies the parent folder. Temporary
location IDs bind the source frame URL, browser session and verified folder for
five minutes. Matching folders deduplicate only within the same session.
Ambiguous, incomplete or stale discovery requires explicit selection/relisting.
Unsupported SharePoint views return a bounded ASPX page name, page class and
known query-key names without query values or page content. The fixed location
probe hot-reloads during source development; MCP callers cannot supply its
path or code. This lets supported page formats be corrected without repeated
main-process restarts once the location probe is installed.

`office_resolve_location` verifies an explicit decoded, server-relative
`folderPath`, using a discovered `sourceLocationId` to bind the account and
browser frame. Its optional `webUrl` selects an owning site or subsite on the
same origin; it defaults to the source web. The requested path must lie within
that web, and the folder metadata must match before a destination ID is issued.
No Teams navigation or folder creation occurs. Cross-origin destinations need
a source on their own SharePoint host; browser requests retain same-origin mode.
Explicit destinations inherit their source's remaining five-minute lifetime
and navigation binding. Registering them does not change current-folder defaults.

`office_list_location_folders` reads up to 50 immediate children and returns
usable destination IDs with names, paths and owning web URLs. It validates the
page's folder metadata before registering destinations and marks truncated
listings. Neither resolution nor browsing uploads a document. Missing and denied
folders return bounded diagnostics without server response bodies. Tests cover
sibling and child destinations, another site/subsite, account binding, expiry,
source navigation and rejected destination URLs. Live testing verified resolving
an open document's parent library, browsing its child folders, then creating in
a sibling of the source folder without navigating Teams or showing Documents.
First-time Bridge import completed in 7,404 ms after creation; a paragraph write
and independent readback passed. Sign-in and setup retries were unnecessary. Creation and verified editing in
the reverse sibling direction also passed with the window hidden.
Another site's creation remains covered by the isolated Electron fixture rather
than this live sibling-folder test.

`office_create_document` uploads the bundled blank Word package through
SharePoint's ResourcePath `AddUsingPath` API with `overwrite=false` and a fresh
request digest. A fixed app-owned function runs browser `fetch` in the originating
SharePoint frame with `credentials: "include"` and `mode: "same-origin"`. This
preserves its browser request context, including User-Agent, Origin and referrer,
without reading/exporting cookie values or calling Teams' token provider.
Requests time out and responses are bounded; redirects, permission
failures and uncertain creation have fixed diagnostics. File names and paths
are escaped as OData literals and URL components, including apostrophes, `%`
and `#`. No destination path, URL, file bytes or caller code enters this tool.

After successful creation, DocumentTabs opens the new URL using the same source
account and starts existing Bridge setup. Opening/setup failure returns
`created: true` with the file metadata rather than inviting another creation.
`office_list_location_files` supports an exact-name lookup for recovery and
reports truncated listings. The document template contains one empty paragraph,
A4 page settings and no author, macros or external relationships. Creation and
hidden opening passed an isolated Electron service fixture, including the source
frame's User-Agent and explicit same-origin credentials. Live testing also
verified folder discovery from Teams' `filebrowser.aspx`, exact-name listing,
Word creation without a supplied link, and discovery of the created file.
After a full app restart with the retained sign-in session, a fresh Word file
was created, opened hidden, imported its Bridge for the first time, and passed
a paragraph write and fresh readback in one uninterrupted attempt. Bridge setup
reported 7,670 ms, including loading after creation and excluding subsequent
reads and edits. No window was shown, sign-in interaction requested, or setup
retry made during that run. An earlier attempt needed the account picker shown,
so this verifies the retained-session path rather than initial authentication.
Cloud save propagation was not checked.

See Microsoft's [ResourcePath API](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/supporting-and-in-file-and-folder-with-the-resourcepath-api)
and [SharePoint REST operations](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints).

Setup allows up to 15 seconds for Microsoft login redirects to complete before
returning `SIGN_IN_REQUIRED`, within the existing overall setup deadline. It
reports `waiting_for_sign_in` during this wait. Top-level navigation does not
cancel the job while waiting for sign-in, selecting the account or waiting for
the initial editor; navigation after editor binding retains cancellation.

A recognized Microsoft account picker uses an account choice remembered for
that document site and Electron profile, or an explicit `accountHint` email on
`office_setup_bridge`. A passive listener remembers only the username on a
trusted account-tile selection in an app-specific localStorage entry on the
Microsoft login page, scoped by the document's origin. Synthetic events cannot
set the preference. No password, token or cookie is stored by this feature.
Account setup makes no calls to Teams' token provider or Graph and does not
inspect Teams' authentication storage. The main Teams session supplies only
the document view's existing Electron session, as it does for ordinary tabs.

The fixed picker script matches the complete remembered identifier and refuses
missing or multiple matches. No remembered choice returns
`DOCUMENT_ACCOUNT_NOT_REMEMBERED`, allowing a first manual account choice or an
explicit hint instead of guessing. One trusted Enter selects the matched tile
through the bounded native input helper while preserving window visibility and
native focus. Progress reports `selecting_account` and
`accountSelection: dispatched`; dispatch alone is not proof of successful login.
Password/OTP fields are checked for presence only, never read, filled or
submitted. Unknown, credential or consent UI is left for the user. An uncertain
selection is not automatically repeated. Account-picker selectors hot-reload
from their fixed local module during source development, like the add-in probe.
Older main-process builds are prevented from entering the removed token-lookup
path by the updated probe until a restart loads the replacement handler.

The inspection tool also works before Office loads, reporting only frame
origins, persistent-session status and fixed cookie-state labels for ESTSAUTH,
ESTSAUTHPERSISTENT, FedAuth and rtFa, plus bounded picker/credential-presence
flags and a tile count. For the two SharePoint cookies it also reports up to
eight attribute summaries: session or persistent, scope relative to the opened
host (without domain names), Secure, HttpOnly, root-path and partition presence.
Missing cookies have an empty list; unavailable inspection has no attributes.
These explain why a cookie may be excluded from optional retention without
changing that policy. Cookie values remain inside Electron.
This does not promote session cookies or extend their expiry. A persistent
partition alone does not make session cookies survive app termination;
Microsoft's **Stay signed in** flow controls their normal persistence.

`mcp.documentSession.enabled` separately opts into local persistence for secure,
HttpOnly, root-path `FedAuth` (including numbered fragments) and `rtFa` cookies
on exact SharePoint hosts opened in Documents, in that view's persistent
Electron session. It also retains a domain-scoped `rtFa` on exactly
`sharepoint.com` once a SharePoint document is opened in that session. This
exception does not cover `FedAuth`, other cookie names or arbitrary parent
domains; the same Secure, HttpOnly, root-path and nonpartitioned rules apply.
It preserves the cookie value, scope, SameSite and security
attributes and adds an expiry only to session cookies. Existing persistent
cookie expiry is never extended. The default is off; retention defaults to
seven days, with 1–30 accepted. Cookie values stay in Electron's normal store.
The helper watches future cookie changes and scans existing cookies, rechecking
the current value before promotion so stale scans cannot restore a removed or
rotated token. Sign-out during a pending promotion removes only that exact
promotion. Writes are flushed to disk; listeners are removed on close. It does
not alter Microsoft login cookies, access Teams' token provider or change the
server-side validity of SharePoint's signed session. Failures leave ordinary
sign-in available without retrying or forwarding raw cookie-store errors.

A real cold-restart test verified that both the site's `FedAuth` and the shared
`rtFa` survived as persistent cookies with their scope and security flags
preserved. The already-configured Word document reopened hidden, restored its
Bridge in 4,483 ms and passed a paragraph read without account selection or
interactive sign-in. This measures reopening an existing Bridge, not first-time
import. Earlier retention of only `FedAuth` had still prompted for credentials:
the exact-host filter excluded the parent-domain `rtFa`. Retaining that root
cookie resolved the prompt in this test; server expiry and revocation still
apply.

A first real Word test reached a live Bridge and verified session-local
paragraph IDs, paragraph reads, insertion, exact literal replacement and deleting
the inserted paragraph. Fresh reads confirmed the original paragraphs and that
the test paragraph no longer existed. Tracking was Off and the target had no
comments or revisions; no review settings were changed. The window was shown
at Microsoft's sign-in page, then hidden for import and all document operations.
The successful import job reported 37,766 ms while UI fixes were being applied,
so this is not a clean benchmark of the final implementation.

Word's compact Home ribbon put Add-ins under More options. The fixed
`ribbonOverflow` action recognizes English/Dutch More options only in the named
Ribbon/Lint toolbar and only when Add-ins is absent. It is distinct from
`moreSettings` inside the add-in catalog, so both can occur in one job. Add-ins
menu items can inherit `pointer-events:none` from their portal while accepting
native Enter. Home preparation and its keyboard obstruction guard apply to
ribbon commands, excluding menu/dialog descendants. Explicit mouse mode retains
all hit tests. The Electron Word fixture covers this overflow path before the
existing Advanced/catalog/manifest flow; unrelated More options buttons remain
excluded. The focused smoke covers hidden menu keyboard activation and the
existing obstruction guard for ribbon commands.

Live diagnosis found two UI differences now covered by the Electron fixture:
Office renders the visible Add-ins menu beneath `visibility:hidden` portal
containers, with `visibility:visible` on descendants. Visibility checks must use
the control's computed visibility and still reject any `display:none` ancestor.
The upload dialog uses `<input type="button" value="Upload">`; setup recognizes
button/submit input values as labels, without reading file/text input values.
Inspection includes recognized submit and consent actions after resolving their
dialogs. A collapsed ribbon also left Add-ins underneath the formula bar;
automatic setup detects that obstruction in its probe, expands Home once,
and waits for the command to become reachable. A failed input preflight can
also trigger this preparation if the layout changed after the probe. It never
bypasses the hit test.

`office_inspect_bridge_setup` reports frame origins, eligibility and up to 240
controls across the selected editor's allowed descendants. It omits document,
grid and editable content, text/file input values, HTML and full URLs. No caller
script or selector is accepted. Inspection does not click controls or change
visibility; names remain untrusted UI data.

`office_bridge_setup_step` executes one named action with the same frame,
manifest and consent checks. It accepts `auto` (default), `pointer`, `click`,
`native` (Enter) or `nativeMouse` interaction modes,
but no caller scripts, selectors, keys, coordinates or file paths. A tab serializes
steps with automatic jobs, records upload attempts before dispatch, and rejects
a repeated submission unless `retry: true` explicitly requests it. Dispatch is
not UI-success confirmation. Inspect after each step when debugging live UI.

Automatic setup uses `nativeSetupInput.js`: the validated named control is
focused, then receives a fixed Enter key press through Electron's
in-process Debugger API. This produces trusted browser activation without
showing the tab or focusing its native window. No remote debugging port is
opened. The operation refuses an existing debugger, checks the document again
before input, temporarily emulates page focus, releases pressed input, restores focus
emulation and detaches after each action. Losing debugger ownership stops input
without detaching another session. No caller CDP command or key is accepted.
Read-only diagnostics remain passive; file selection still supplies only the
app's generated manifest to the upload input.
Explicit `nativeMouse` mode remains available for diagnostics.
`setupMousePoint.js` validates the target's visible rectangle and hit tests each
parent frame, rejecting overlays, ambiguous owners and rotated/skewed frames.
`setupMouseSession.js` routes through one matching descendant Chromium session
for a frame in another renderer process. Without this routing, native mouse
input did not reach nonselected/cross-origin frames in the fixture. Session
discovery is bounded and all owned sessions end with the root debugger detach.
No caller coordinate, target ID, selector or protocol command is accepted.

`tests/electron/nativeSetupInputSmoke.js` verifies that DOM click/pointer events
are untrusted and that native activation produces a trusted click in a hidden
cross-origin frame, preserves native window focus and releases its debugger.
Run `documentTabsSmoke.js` with `DOCUMENT_TABS_SETUP=1 DOCUMENT_TABS_NATIVE=1`
to require trusted clicks throughout its representative import flow. Both
fixtures pass. The full import fixture covers visible menus overriding hidden
portal ancestors, legacy Upload input buttons, exact manifest selection, one
submission, and two Excel documents plus one Word document. Its UI and Office.js
are simulated; the separate live Excel check establishes actual import/read
behavior only for the tested workbook.

Setup permits `res.public.onecdn.static.microsoft` only within the selected
editor's descendants. Sibling frames, lookalike hosts and other hosts under the
CDN domain remain excluded. The Office host creates the Bridge task pane.
Read-only inspection exposes up to 40 named-control hit-test records across all
frames: target/descendant/ancestor/other/no-hit classification, hit element
tag/role, target pointer-events, dimensions and clipping. It omits blocker text,
identifiers, HTML, input values and coordinates. Document-tab inspection also
reports native window visibility, minimization, focus and tab selection. These
diagnostics do not relax the mouse hit test.
The point check tries the centre and four interior points, selecting only a
point whose hit is the named control or its descendant. Diagnostics retain the
centre relation separately. A full overlay still prevents input.

When running a checkout with Electron (`process.defaultApp === true`),
`bridgeSetupUiLoader.js` reloads the fixed local `bridgeSetupUi.js` module when
its file changes. This lets UI-probe corrections take effect without restarting
the app after the loader itself is installed. It cannot reload arbitrary caller
source or paths. Packaged builds keep the cached module; main-process changes
still require a restart. Invalid development changes stop setup and can be
repaired for the next inspection. `uiProbeReloadSmoke.js` verifies this in one
hidden Electron process without restarting.
Capture/visibility changes and mouse-only activation for every step
were explored but are not part of the implementation.

## Live Excel tools

| Tool | Purpose |
| --- | --- |
| `office_list_documents` | Discover initialized task panes and their temporary document IDs |
| `excel_list_worksheets` | Get stable worksheet IDs and names for a chosen document |
| `excel_read_range` | Read values and sparse formula/display differences; obtain a read token |
| `excel_write_range` | Write a matching matrix of literal values using a read token |
| `excel_set_formulas` | Write a matching matrix of formulas starting with `=` using a read token |

Always list documents first and target an explicit document ID and worksheet ID.
Ranges use A1 rectangles such as `A1` or `B2:D8`, limited to 1,000 cells and a
32 KB result. Writes use `data` as a two-dimensional matrix, limited to 32 KB;
empty strings clear cells. Literal text starting with `=`, `+` or `-` is rejected
to avoid interpreting it as a formula. Use numeric values for negative numbers.

A cell-range read returns a single-use `readToken` valid for fifteen minutes. Writes must
include that token, the same worksheet ID and the same range address. The add-in
reads the cells again and rejects a write if either their values or formulas
changed. Excel does not provide an atomic compare-and-set, so a coauthor can
still change cells between the comparison and write. Serialize operations and
avoid working on the same cells at the same time.

After an uncertain or timed-out write, read the range again rather than retrying
the old request. A successful write returns a new read token and compact
confirmation; `responseMode: "full"` also returns the resulting cells.
The bridge does not download and replace the workbook file; Excel handles
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

Each paragraph read returns a fifteen-minute, single-use token. Replacement checks
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
| `excel_read_range_format` | `worksheetId`, `address`, optional `responseMode` | Formatting and a format token; full mode also returns cell content |
| `excel_format_range` | Same range, `readToken`, `format` | Uniform formatting; checks content and inspected formatting |
| `excel_create_worksheet` | `name`, `readToken` from `excel_list_worksheets` | New worksheet ID/name/position; checks the sheet list |
| `excel_insert_rows_columns` | `worksheetId`, `axis` (`rows`/`columns`), `index`, `count`, layout `readToken` | Insert whole rows/columns before a zero-based index |

Search and formatting use A1 rectangles of at most 1,000 cells. Search scans a
bounded window, treating wildcard characters literally. Use used-range metadata
to plan windows across large sheets. Search results are not write tokens.

Excel `format` accepts `numberFormat` (one invariant format string applied to all
cells), `font`, `fillColor`, `horizontalAlignment`, `verticalAlignment` and
`wrapText`. Sizing accepts `columnWidth` (nonnegative points), `rowHeight`
(0–409 points), `autoFitColumns: true` and `autoFitRows: true`. Zero size hides
the intersecting columns/rows. An explicit size and AutoFit cannot be combined
for the same axis. Width and height changes affect whole columns/rows, even
when the inspected rectangle covers only part of them. AutoFit uses the cells
in that rectangle, after applying any other supplied formatting.

On ExcelApi 1.2 hosts, formatting reads return `format.columnWidths` and
`format.rowHeights` arrays in points, ordered left-to-right and top-to-bottom.
Tokens compare each individual dimension to detect coauthor resizing even
when the range has mixed sizes. Both arrays are omitted on 1.1 hosts, which
continue to support basic formatting but reject sizing requests.

For example, read the format of `B1:B20`, then pass its token with
`format: { "columnWidth": 150 }` to widen column B, or with
`format: { "autoFitColumns": true }` to fit it to those twenty cells.

Font accepts `bold`, `italic`, `name`, `size` (1–200 points) and
`color`; colors use `#RRGGBB`. Unspecified properties remain unchanged. Mixed
properties may return null/empty values. Number formats are a per-cell matrix.

Worksheet creation needs a non-truncated worksheet-list token (maximum 200
sheets) and a unique valid Excel name of at most 31 characters. Insertion allows
up to 100 rows or 20 columns per call. Layout tokens compare worksheet ID and
used-range bounds, not all cell content or table structure. Excel shifts data
and references. Coordinate structural edits with coauthors. These operations
invalidate all earlier pane tokens. Used-range inspection/insertion require
ExcelApi 1.4; sizing requires ExcelApi 1.2, and other new Excel tools require
ExcelApi 1.1. See Microsoft’s [RangeFormat reference](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeformat?view=office-js)
for the sizing APIs.

### Word

| Tool | Additional arguments | Result or effect |
| --- | --- | --- |
| `word_read_structure` | Optional page | Paragraph IDs/text, styles/alignment and ordinary text tokens |
| `word_read_paragraph_details` | `paragraphId` | Text, style, alignment, font and format token |
| `word_list_tables` | Optional page | Main-body table indexes and row counts |
| `word_read_table` | `tableIndex`, optional `rowStart`, `rowCount` | Paged row/cell text, cell counts and `nextRowStart` |
| `word_replace_paragraph` | `paragraphId`, ordinary `readToken`, `text` | Replace complete paragraph content, excluding its mark |
| `word_insert_paragraph` | Anchor `paragraphId`, ordinary `readToken`, `text`, `position` (`Before`/`After`), optional `styleBuiltIn` | New paragraph ID/token; text in full mode |
| `word_insert_paragraphs` | Anchor `paragraphId`, ordinary `readToken`, `position`, `paragraphs` array of `{text, styleBuiltIn?}` | Insert 1–20 paragraphs in array order with one token; returns inserted IDs/tokens |
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
small corrections. Paragraph edits invalidate tokens for the affected paragraph
and preserve tokens for other stable paragraph IDs. Their current text is still
checked before each write. Insertion/replacement return fresh paragraph tokens. Word handles
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

Office edit tokens expire after fifteen minutes and are consumed on attempted writes.
Comparison and mutation are separate Office batches; no operation claims atomic
protection against another editor. After an uncertain write, inspect current
state before deciding whether another write is needed.

### Compact responses and batched edits

Office cell reads and paragraph/range writes default to `responseMode: "compact"`.
Use `"full"` for the previous dense response, including resulting content after
a write. Complete snapshots remain in the task pane for comparison with live
state; compact output does not weaken write checks.

`excel_read_range` returns the values matrix once. Optional `formulaCells` and
`textCells` arrays encode differences as `[rowOffset, columnOffset, value]`, with
zero-based offsets relative to the requested rectangle. Without an entry, the
formula is the cell value and the displayed text is `String(value ?? "")`.
`excel_read_range_format` returns formatting without cell matrices. A uniform
`format.numberFormat` is a single string; mixed formats remain a matrix. Column
widths and row heights remain arrays. Full mode preserves all original matrices.
Successful projected writes return `applied`, target IDs/address and a fresh
`readToken`; they do not echo cell/paragraph content or formatting.

Responses are capped at 32 KiB and retained comparison snapshots at 128 KiB.
There are at most 512 tokens per document, expiring after fifteen minutes.
Large cell content can still require smaller ranges. Read relevant ranges or
pages once and reuse the tokens; do not reread merely to confirm `applied: true`.

`word_batch_edit(documentId, edits)` supports `replace_text`, `replace_paragraph`,
`insert_paragraph`, `delete_paragraph` and `format_paragraph`.
`excel_batch_edit(documentId, edits)` supports `write_range`, `set_formulas` and
`format_range`. Each edit contains `operation` plus the single tool's arguments,
omitting `documentId` and `responseMode`. Batches require distinct read tokens,
contain 1–20 edits and fit within 32 KiB. All argument schemas and cross-field
checks run before the first edit. Each edit then runs sequentially through the
ordinary bridge, application, read-token and coauthor checks.

Results contain `completed` and indexed compact `results`. On failure they also
contain `failedIndex`, `error`, `uncertain: true` and `notAttempted`. Earlier
writes remain applied; inspect the failed target before considering another
write and never replay the entire batch. After thirty seconds the server stops
before starting another edit, returning `stopped: "time_budget"`, `nextIndex`
and `notAttempted`. An in-flight edit keeps its ordinary bridge deadline. For
multiple new paragraphs at one anchor, use `word_insert_paragraphs` to insert
the section in one Office operation (1–20 paragraphs, 32 KiB total).

Reference APIs:
[Excel ranges](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.9),
[Word paragraphs](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph?view=word-js-preview),
[Word tables](https://learn.microsoft.com/en-us/javascript/api/word/word.table?view=word-js-preview),
[comments](https://learn.microsoft.com/en-us/javascript/api/word/word.comment?view=word-js-preview),
[tracked changes](https://learn.microsoft.com/en-us/javascript/api/word/word.trackedchange?view=word-js-preview).

## Run from source

Use Node.js 24. Install the dependencies with `npm run install:app` in the repository,
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
this location. The bridge supports Linux source runs and the Teams Office Flatpak.

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
        "/absolute/path/to/teams-for-linux-mcp/build/teams-office/app/mcp/stdio.js",
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

Run the focused checks from the assembled app with:

```sh
cd build/teams-office
node --test tests/unit/excelSession.test.js tests/unit/wordSession.test.js tests/unit/officeEditing.test.js tests/unit/officeAddinHost.test.js tests/unit/officeLiveBridge.test.js tests/unit/officeInspector.test.js tests/unit/officeMcp.test.js
cd ../..
npm run lint
```

From `build/teams-office`, check the Electron frame APIs against nested cross-origin fixture pages with
`./node_modules/.bin/electron tests/electron/officeMcpSmoke.js`. A headless Linux
test machine can use `xvfb-run -a` before that command. This test serves local
fixture responses for every HTTPS request and does not contact Microsoft or
verify a real signed-in Office editor. Also run
`./node_modules/.bin/electron tests/electron/officeAddinSmoke.js` to check the real
loopback certificate pin and Excel/Word task-pane reads and writes with simulated
Office.js hosts. It also checks that older Word hosts remain disconnected. This
remains separate from testing against Microsoft Office and another coauthor.
