# Connect an agent to Word and Excel

The Office bridge lets an MCP agent read and edit Word and Excel documents open
inside Teams for Linux. Excel supports cell reads, value writes and formulas;
Word supports paragraph reads and exact text replacements. Documents stay in
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

Transfers use base64 bytes and are limited to **1 MiB per file**. These tools can
read binary formats, but interpreting or modifying a DOCX/XLSX package requires
format-aware tooling. Use the existing Word/Excel tools for edits in open Office
documents. Replacement uploads replace the complete file, rather than merging
Office changes. A replacement requires `expectedETag` from a recent file read;
the server checks it again and sends `If-Match`. This safeguard and creation's
conflict-failure behavior still require live validation against your tenant;
coordinate edits while testing. Do not automatically repeat a timed-out upload:
list/read the destination first to determine whether it succeeded.

The server reuses the signed-in Teams session. It does not request new Microsoft
permissions and does not expose its access token or temporary download links to
the agent. HTTP 403 means the session lacks permission or organization policy
blocks the operation; enabling MCP does not grant additional SharePoint rights.
Discovery, reading and writing can have different permissions. Microsoft's
[folder lookup](https://learn.microsoft.com/en-us/graph/api/channel-get-filesfolder?view=graph-rest-1.0)
and [upload documentation](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0)
list the required permissions.

For a first live test, choose a test channel, list its files, then ask the agent
to create a uniquely named `mcp-file-test.txt`, read it back, replace it using the
returned eTag, and read it back again. Use a disposable file for testing conflict
handling with another editor. File deletion is not exposed by these tools.
